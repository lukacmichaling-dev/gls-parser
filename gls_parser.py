#!/usr/bin/env python3
"""
GLS Parser – Konverzia CAMT.053 bankového výpisu do formátu MoneyData
s doplnením VS a zákazníka z GLS XLSX súborov.
"""

import re
import glob
import os
import uuid
import warnings
from datetime import datetime
from xml.etree import ElementTree as ET
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom

import sys
import openpyxl
warnings.filterwarnings('ignore')

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QSpinBox,
    QGroupBox, QTextEdit, QMessageBox, QProgressBar,
    QTableWidget, QTableWidgetItem, QHeaderView, QSplitter, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor


# ─── Konfigurácia spoločnosti ─────────────────────────────────────────────────
GLS_IBAN = 'SK9602000000002802401454'

FIXED = {
    'ICAgendy':   '36470449',
    'PrKont':     'O',
    'Cleneni':    '21U 00',
    'Cinnost':    'MOREPNEU',
    'ZpVypDPH':   '1',
    'SSazba':     '5',
    'ZSazba':     '23',
    'DRada':      'SLSrr',
    'Vyst':       'Mgr. Lukáčová',
}

BANK_ACCOUNT = {
    'Zkrat': 'SLSP',
    'Ucet':  '5212572513',
    'BKod':  '0900',
    'BNazev':'Slovenská sporiteľňa, a.s.',
}

DALSI_SAZBA = {
    'Popis':      'druhá znížená',
    'HladinaDPH': '1',
    'Sazba':      '19',
    'Zaklad':     '0',
    'DPH':        '0',
}
# ─────────────────────────────────────────────────────────────────────────────

NS = 'urn:iso:std:iso:20022:tech:xsd:camt.053.001.02'


# ─── Pomocné funkcie ──────────────────────────────────────────────────────────

def load_xlsx_files(xlsx_dir: str) -> dict:
    """Načíta GLS XLSX súbory, vráti slovník: round(total,2) -> {vs, addr_str}"""
    xlsx_map = {}
    for fpath in sorted(glob.glob(os.path.join(xlsx_dir, '*.xlsx'))):
        try:
            wb = openpyxl.load_workbook(fpath)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 10:
                continue
            total = rows[-1][4]
            if not isinstance(total, (int, float)):
                continue
            data_rows = [r for r in rows[8:-1] if r[0] is not None]
            if not data_rows:
                continue
            first = data_rows[0]
            vs   = str(first[2]).strip() if first[2] else ''
            addr = str(first[6]).strip() if first[6] else ''
            key  = round(float(total), 2)
            xlsx_map[key] = {'vs': vs, 'addr_str': addr, 'file': os.path.basename(fpath)}
        except Exception as e:
            print(f'XLSX chyba {fpath}: {e}')
    return xlsx_map


def parse_address(addr_str: str) -> dict:
    """Parsuje adresu GLS: '{Meno} [telefon] SK-{PSC} {Mesto} {Ulica}'"""
    addr_str = addr_str.strip()
    m = re.match(r'^(.*?)\s+(?:\d{9,15}\s+)?SK-(\d+)\s+(.+)$', addr_str)
    if not m:
        return {'ObchNazev': addr_str}
    name     = m.group(1).strip()
    psc      = str(int(m.group(2)))
    location = m.group(3).strip()
    words    = location.split()
    street_start = len(words)
    for i, w in enumerate(words):
        if re.search(r'\d', w):
            street_start = max(1, i - 1)
            break
    city   = ' '.join(words[:street_start])
    street = ' '.join(words[street_start:])
    return {
        'ObchNazev': name,
        'Ulice':     street,
        'Misto':     city,
        'PSC':       psc,
        'Stat':      'Slovenská republika',
    }


def extract_from_eid(eid: str):
    """Extrahuje VS, SS, KS z EndToEndId formátu /VS.../SS.../KS..."""
    if not eid:
        return '', '', ''
    def g(pattern):
        m = re.search(pattern, eid)
        return m.group(1) if m else ''
    return g(r'/VS(\d+)'), g(r'/SS(\d+)'), g(r'/KS(\d+)')


def build_popis(vs: str, vydej: int, ustrd: str, month: int, year: int) -> str:
    if ustrd:
        if re.match(r'^/VS', ustrd) and 'EUR' in ustrd:
            return 'Transakčná daň'
        if ustrd.startswith('NOTPROVIDED'):
            return 'Transakčná daň'
        if ustrd == 'Transakcie_500':
            return 'Transakcie_500'
        if 'Poplatok za vedenie uctu' in ustrd:
            return f'Poplatok za vedenie uctu {month}/{year}'
        if 'Riadny debetny urok' in ustrd:
            return f'Riadny debetny urok {month}/{year}'
        if ustrd.startswith('Poplatok'):
            return ustrd
    if vs and vs != '0000000000':
        smer = 'prijatej' if vydej else 'vystavenej'
        return f'{vs} - Úhrada faktúry {smer}'
    return ustrd or ''


def make_doklad(year: int, dcislo: int) -> str:
    return f'SLS{str(year)[-2:]}{dcislo:05d}'


def make_id_polozky(account: str, year: int, month: int, seq: int) -> str:
    return f'{account}-{year}-{month:02d}-3000{month:02d}{seq:05d}'


def add_el(parent: Element, tag: str, text) -> Element:
    el = SubElement(parent, tag)
    if text is not None:
        el.text = str(text)
    return el


# ─── Načítanie pôvodného XML (preview) ───────────────────────────────────────

def read_original_xml(xml_path: str) -> list:
    """
    Vráti zoznam riadkov pre zobrazenie pôvodného XML.
    Každý riadok: (datum, suma, typ, vs, popis)
    """
    ns_map = {'ns': NS}
    tree   = ET.parse(xml_path)
    root   = tree.getroot()
    rows   = []
    for entry in root.findall('.//ns:Ntry', ns_map):
        dt    = entry.find('.//ns:BookgDt/ns:Dt', ns_map)
        amt   = entry.find('ns:Amt', ns_map)
        cdi   = entry.find('ns:CdtDbtInd', ns_map)
        eid   = entry.find('.//ns:Refs/ns:EndToEndId', ns_map)
        ustrd = entry.find('.//ns:RmtInf/ns:Ustrd', ns_map)
        vs, _, _ = extract_from_eid(eid.text if eid is not None else '')
        rows.append({
            'datum': dt.text    if dt    is not None else '',
            'suma':  amt.text   if amt   is not None else '0',
            'typ':   cdi.text   if cdi   is not None else '',
            'vs':    vs or '',
            'popis': (ustrd.text or '')[:60] if ustrd is not None else '',
        })
    return rows


# ─── Hlavná konverzia ─────────────────────────────────────────────────────────

def convert(xml_path: str, xlsx_dir: str, start_dcislo: int,
            hosp_rok_od: str, hosp_rok_do: str,
            output_path: str, log_fn=None) -> tuple:
    """
    Skonvertuje CAMT.053 XML + GLS XLSX -> MoneyData XML.
    Vráti (počet záznamov, zoznam riadkov pre výstupnú tabuľku).
    """
    def log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    xlsx_map = load_xlsx_files(xlsx_dir)
    log(f'Načítaných XLSX súborov: {len(xlsx_map)}')
    for total, info in xlsx_map.items():
        log(f'  {info["file"]}: {total} EUR  VS={info["vs"]}')

    tree    = ET.parse(xml_path)
    root    = tree.getroot()
    ns_map  = {'ns': NS}
    entries = root.findall('.//ns:Ntry', ns_map)
    log(f'\nPočet záznamov v XML: {len(entries)}')

    now        = datetime.now()
    money_data = Element('MoneyData', {
        'ICAgendy':     FIXED['ICAgendy'],
        'KodAgendy':    '',
        'HospRokOd':    hosp_rok_od,
        'HospRokDo':    hosp_rok_do,
        'description':  'bankové doklady',
        'ExpZkratka':   '_BD',
        'ExpDate':      now.strftime('%Y-%m-%d'),
        'ExpTime':      now.strftime('%H:%M:%S'),
        'JazykVerze':   'SK',
        'VyberZaznamu': '0',
        'GUID':         '{' + str(uuid.uuid4()).upper() + '}',
    })
    seznam = SubElement(money_data, 'SeznamBankDokl')

    dcislo      = start_dcislo
    seq         = 1
    gls_matches = 0
    out_rows    = []   # pre tabuľku výstupu

    for i, entry in enumerate(entries):
        cdi_el   = entry.find('ns:CdtDbtInd', ns_map)
        vydej    = 1 if (cdi_el is not None and cdi_el.text == 'DBIT') else 0

        amt_el   = entry.find('ns:Amt', ns_map)
        amt_str  = amt_el.text if amt_el is not None else '0'
        try:
            amount = round(float(amt_str), 2)
        except ValueError:
            amount = 0.0

        dt_el    = entry.find('.//ns:BookgDt/ns:Dt', ns_map)
        date_str = dt_el.text if dt_el is not None else ''
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            date_obj = now
        year  = date_obj.year
        month = date_obj.month

        eid_el = entry.find('.//ns:Refs/ns:EndToEndId', ns_map)
        eid    = eid_el.text if eid_el is not None else ''
        vs, ss, ks = extract_from_eid(eid)

        ustrd_el = entry.find('.//ns:RmtInf/ns:Ustrd', ns_map)
        ustrd    = ustrd_el.text if ustrd_el is not None else ''

        if vydej == 0:
            iban_el = entry.find('.//ns:DbtrAcct/ns:Id/ns:IBAN', ns_map)
            bic_el  = entry.find('.//ns:RltdAgts/ns:DbtrAgt/ns:FinInstnId/ns:BIC', ns_map)
        else:
            iban_el = entry.find('.//ns:CdtrAcct/ns:Id/ns:IBAN', ns_map)
            bic_el  = entry.find('.//ns:RltdAgts/ns:CdtrAgt/ns:FinInstnId/ns:BIC', ns_map)
        ad_ucet = iban_el.text if iban_el is not None else ''
        ad_kod  = bic_el.text  if bic_el  is not None else ''

        dbtr_iban_el = entry.find('.//ns:DbtrAcct/ns:Id/ns:IBAN', ns_map)
        dbtr_iban    = dbtr_iban_el.text if dbtr_iban_el is not None else ''

        addr_data: dict = {}
        is_gls_match    = False

        if vydej == 0 and dbtr_iban == GLS_IBAN and amount in xlsx_map:
            xlsx      = xlsx_map[amount]
            vs        = xlsx['vs']
            if xlsx['addr_str']:
                addr_data = parse_address(xlsx['addr_str'])
            is_gls_match = True
            gls_matches += 1
            log(f'  GLS zhoda [{i}]: {amount} EUR → VS={vs} ({xlsx["file"]})')

        if not addr_data:
            cdtr_nm_el = entry.find('.//ns:Cdtr/ns:Nm', ns_map)
            if cdtr_nm_el is not None and cdtr_nm_el.text:
                addr_data = {'ObchNazev': cdtr_nm_el.text.strip()}

        vs_final = vs if vs else '0000000000'
        popis    = build_popis(vs_final, vydej, ustrd, month, year)
        doklad   = make_doklad(year, dcislo)
        id_pol   = make_id_polozky(BANK_ACCOUNT['Ucet'], year, month, seq)

        doc = SubElement(seznam, 'BankDokl')
        add_el(doc, 'Vydej',     vydej)
        add_el(doc, 'Doklad',    doklad)
        add_el(doc, 'Popis',     popis)
        add_el(doc, 'DatUcPr',   date_str)
        add_el(doc, 'DatVyst',   date_str)
        add_el(doc, 'DatPlat',   date_str)
        add_el(doc, 'DatPln',    date_str)
        add_el(doc, 'Vypis',     month)
        add_el(doc, 'IDPolozky', id_pol)
        add_el(doc, 'AdUcet',    ad_ucet)
        add_el(doc, 'AdKod',     ad_kod)
        add_el(doc, 'VarSym',    vs_final)
        add_el(doc, 'ParSym',    vs_final)
        if ks:
            add_el(doc, 'KonSym',  ks)
        if ss:
            add_el(doc, 'SpecSym', ss)

        if addr_data:
            adresa   = SubElement(doc, 'Adresa')
            if 'ObchNazev' in addr_data:
                add_el(adresa, 'ObchNazev', addr_data['ObchNazev'])
            obch_adr = SubElement(adresa, 'ObchAdresa')
            if 'Ulice' in addr_data:
                add_el(obch_adr, 'Ulice', addr_data['Ulice'])
            if 'Misto' in addr_data:
                add_el(obch_adr, 'Misto', addr_data['Misto'])
            if 'PSC' in addr_data:
                add_el(obch_adr, 'PSC',   addr_data['PSC'])
            add_el(obch_adr, 'Stat', addr_data.get('Stat', 'Slovensko'))

        add_el(doc, 'PrKont',   FIXED['PrKont'])
        add_el(doc, 'Cleneni',  FIXED['Cleneni'])
        add_el(doc, 'Cinnost',  FIXED['Cinnost'])
        add_el(doc, 'ZpVypDPH', FIXED['ZpVypDPH'])
        add_el(doc, 'SSazba',   FIXED['SSazba'])
        add_el(doc, 'ZSazba',   FIXED['ZSazba'])

        sdph = SubElement(doc, 'SouhrnDPH')
        add_el(sdph, 'Zaklad0',  amt_str)
        add_el(sdph, 'Zaklad5',  '0')
        add_el(sdph, 'Zaklad22', '0')
        add_el(sdph, 'DPH5',     '0')
        add_el(sdph, 'DPH22',    '0')
        if vydej == 0:
            sds = SubElement(sdph, 'SeznamDalsiSazby')
            ds  = SubElement(sds,  'DalsiSazba')
            for k, v in DALSI_SAZBA.items():
                add_el(ds, k, v)

        add_el(doc, 'Celkem', amt_str)
        add_el(doc, 'DRada',  FIXED['DRada'])
        add_el(doc, 'DCislo', dcislo)
        add_el(doc, 'Vyst',   FIXED['Vyst'])

        ucet = SubElement(doc, 'Ucet')
        for k, v in BANK_ACCOUNT.items():
            add_el(ucet, k, v)

        out_rows.append({
            'datum':    date_str,
            'doklad':   doklad,
            'vydej':    vydej,
            'vs':       vs_final,
            'zakaznik': addr_data.get('ObchNazev', ''),
            'suma':     amt_str,
            'gls':      is_gls_match,
        })

        dcislo += 1
        seq    += 1

    xml_str = tostring(money_data, encoding='unicode')
    dom     = minidom.parseString(xml_str)
    pretty  = dom.toprettyxml(indent='  ', encoding='UTF-8')
    lines   = [l for l in pretty.decode('utf-8').splitlines() if l.strip()]
    output  = '\n'.join(lines)

    with open(output_path, 'w', encoding='utf-8-sig') as f:
        f.write(output)

    log(f'\n✓ Výstup uložený: {output_path}')
    log(f'  Celkom záznamov : {len(entries)}')
    log(f'  GLS XLSX zhody  : {gls_matches}')
    return len(entries), out_rows


# ─── Worker thread ────────────────────────────────────────────────────────────

class WorkerThread(QThread):
    log_signal   = pyqtSignal(str)
    done_signal  = pyqtSignal(int, list)
    error_signal = pyqtSignal(str)

    def __init__(self, xml_path, xlsx_dir, start_dcislo,
                 hosp_rok_od, hosp_rok_do, output_path):
        super().__init__()
        self.xml_path     = xml_path
        self.xlsx_dir     = xlsx_dir
        self.start_dcislo = start_dcislo
        self.hosp_rok_od  = hosp_rok_od
        self.hosp_rok_do  = hosp_rok_do
        self.output_path  = output_path

    def run(self):
        try:
            count, out_rows = convert(
                self.xml_path, self.xlsx_dir, self.start_dcislo,
                self.hosp_rok_od, self.hosp_rok_do,
                self.output_path, log_fn=self.log_signal.emit
            )
            self.done_signal.emit(count, out_rows)
        except Exception as e:
            import traceback
            self.error_signal.emit(traceback.format_exc())


# ─── GUI ──────────────────────────────────────────────────────────────────────

# Tmavá paleta
BG_DARK      = QColor('#1e1e2e')   # pozadie tabuľky
BG_ROW_EVEN  = QColor('#1e1e2e')   # párny riadok
BG_ROW_ODD   = QColor('#252535')   # nepárny riadok
BG_GLS       = QColor('#2e1a28')   # tmavo-ružová – GLS XLSX zhoda
BG_DBIT      = QColor('#3a2a1a')   # tmavo-oranžová – výdaj
FG_DEFAULT   = QColor('#e0e0f0')   # svetlé písmo
FG_DIM       = QColor('#888899')   # tlmené (prázdne VS)
FG_GLS       = QColor('#ff80c0')   # ružová – GLS zhoda
FG_DBIT      = QColor('#ffb347')   # oranžová – výdaj

TBL_STYLE = """
QTableWidget {
    background-color: #1e1e2e;
    color: #e0e0f0;
    gridline-color: #333350;
    border: 1px solid #333350;
    outline: none;
}
QTableWidget::item:selected {
    background-color: #3a3a5e;
    color: #ffffff;
}
QHeaderView::section {
    background-color: #12122a;
    color: #a0a0cc;
    padding: 4px;
    border: none;
    border-bottom: 1px solid #444466;
    font-weight: bold;
}
QScrollBar:vertical {
    background: #1e1e2e;
    width: 10px;
}
QScrollBar::handle:vertical {
    background: #444466;
    border-radius: 4px;
}
"""

def _tbl_item(text: str, align=Qt.AlignLeft,
              fg: QColor = None, bg: QColor = None) -> QTableWidgetItem:
    item = QTableWidgetItem(str(text))
    item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
    item.setTextAlignment(align | Qt.AlignVCenter)
    item.setForeground(fg or FG_DEFAULT)
    item.setBackground(bg or BG_ROW_EVEN)
    return item


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('GLS Parser – MoneyData konverzia')
        self.setMinimumSize(1100, 750)
        self._build_ui()

    # ── Zostav UI ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setSpacing(8)
        root_layout.setContentsMargins(12, 12, 12, 12)

        # ── Horný formulár ────────────────────────────────────────────────────
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(6)

        # Riadok 1: XML + XLSX
        row1 = QHBoxLayout()

        grp_xml = QGroupBox('XML bankový výpis')
        l_xml = QHBoxLayout(grp_xml)
        self.xml_edit = QLineEdit()
        self.xml_edit.setPlaceholderText('Vyber CAMT.053 XML súbor...')
        btn_xml = QPushButton('Prehľadávať')
        btn_xml.setFixedWidth(110)
        btn_xml.clicked.connect(self._pick_xml)
        l_xml.addWidget(self.xml_edit)
        l_xml.addWidget(btn_xml)

        grp_xlsx = QGroupBox('Priečinok GLS XLSX')
        l_xlsx = QHBoxLayout(grp_xlsx)
        self.xlsx_edit = QLineEdit()
        self.xlsx_edit.setPlaceholderText('Priečinok s XLSX súbormi...')
        btn_xlsx = QPushButton('Prehľadávať')
        btn_xlsx.setFixedWidth(110)
        btn_xlsx.clicked.connect(self._pick_xlsx)
        l_xlsx.addWidget(self.xlsx_edit)
        l_xlsx.addWidget(btn_xlsx)

        row1.addWidget(grp_xml, 3)
        row1.addWidget(grp_xlsx, 2)
        form_layout.addLayout(row1)

        # Riadok 2: Nastavenia + Výstup + Tlačidlo
        row2 = QHBoxLayout()

        grp_set = QGroupBox('Nastavenia')
        l_set = QHBoxLayout(grp_set)
        l_set.addWidget(QLabel('DCislo štart:'))
        self.dcislo_spin = QSpinBox()
        self.dcislo_spin.setRange(1, 999999)
        self.dcislo_spin.setValue(1)
        self.dcislo_spin.setFixedWidth(80)
        l_set.addWidget(self.dcislo_spin)
        l_set.addSpacing(12)
        l_set.addWidget(QLabel('Hosp. rok od:'))
        self.hosp_od_edit = QLineEdit(f'{datetime.now().year}-01-01')
        self.hosp_od_edit.setFixedWidth(95)
        l_set.addWidget(self.hosp_od_edit)
        l_set.addWidget(QLabel('do:'))
        self.hosp_do_edit = QLineEdit(f'{datetime.now().year}-12-31')
        self.hosp_do_edit.setFixedWidth(95)
        l_set.addWidget(self.hosp_do_edit)
        l_set.addStretch()

        grp_out = QGroupBox('Výstupný súbor')
        l_out = QHBoxLayout(grp_out)
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText('Cesta k výstupnému XML...')
        btn_out = QPushButton('Prehľadávať')
        btn_out.setFixedWidth(110)
        btn_out.clicked.connect(self._pick_output)
        l_out.addWidget(self.out_edit)
        l_out.addWidget(btn_out)

        self.btn_run = QPushButton('▶  Spustiť konverziu')
        self.btn_run.setFixedHeight(48)
        self.btn_run.setFixedWidth(160)
        f = QFont(); f.setBold(True); self.btn_run.setFont(f)
        self.btn_run.clicked.connect(self._run)

        row2.addWidget(grp_set, 2)
        row2.addWidget(grp_out, 3)
        row2.addWidget(self.btn_run)
        form_layout.addLayout(row2)

        root_layout.addWidget(form_widget)

        # ── Progress bar ──────────────────────────────────────────────────────
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedHeight(6)
        self.progress.setVisible(False)
        root_layout.addWidget(self.progress)

        # ── Dve tabuľky (splitter) ────────────────────────────────────────────
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(6)

        # Ľavý panel – pôvodný XML
        left_panel  = self._make_panel(
            'Pôvodný XML (CAMT.053)',
            ['Dátum', 'Suma (EUR)', 'Typ', 'VS', 'Popis'],
            [90, 100, 50, 115, 250],
        )
        self.tbl_orig, self.lbl_total_orig = left_panel[0], left_panel[1]

        # Pravý panel – upravený MoneyData
        right_panel = self._make_panel(
            'Upravený XML (MoneyData)',
            ['Dátum', 'Doklad', 'Vydej', 'VS', 'Zákazník', 'Suma (EUR)'],
            [90, 95, 50, 115, 200, 100],
        )
        self.tbl_out, self.lbl_total_out = right_panel[0], right_panel[1]

        splitter.addWidget(left_panel[2])    # wrapper widget
        splitter.addWidget(right_panel[2])
        splitter.setSizes([550, 550])

        root_layout.addWidget(splitter, 1)

        # ── Log ───────────────────────────────────────────────────────────────
        grp_log = QGroupBox('Priebeh')
        grp_log.setMaximumHeight(110)
        l_log = QVBoxLayout(grp_log)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont('Courier New', 9))
        l_log.addWidget(self.log_text)
        root_layout.addWidget(grp_log)

    def _make_panel(self, title: str, headers: list, widths: list):
        """Vytvorí panel s nadpisom, tabuľkou a celkovou sumou."""
        wrapper = QWidget()
        layout  = QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(3)

        # Nadpis
        lbl_title = QLabel(title)
        f = QFont(); f.setBold(True); f.setPointSize(10)
        lbl_title.setFont(f)
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet(
            'background: #12122a; color: #a0a0cc; padding: 6px;'
            'border: 1px solid #333350; border-bottom: none;')
        layout.addWidget(lbl_title)

        # Tabuľka
        tbl = QTableWidget(0, len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.verticalHeader().setDefaultSectionSize(22)
        tbl.verticalHeader().setVisible(False)
        tbl.setSelectionBehavior(tbl.SelectRows)
        tbl.setEditTriggers(tbl.NoEditTriggers)
        tbl.setAlternatingRowColors(False)
        tbl.setStyleSheet(TBL_STYLE)
        tbl.setFont(QFont('Menlo', 9))
        for i, w in enumerate(widths):
            tbl.setColumnWidth(i, w)
        layout.addWidget(tbl, 1)

        # Celková suma
        lbl_total = QLabel('Celková suma: –')
        lbl_total.setAlignment(Qt.AlignRight)
        f2 = QFont(); f2.setBold(True); f2.setPointSize(9)
        lbl_total.setFont(f2)
        lbl_total.setStyleSheet(
            'padding: 5px 8px; background: #12122a; color: #a0c8ff;'
            'border: 1px solid #333350; border-top: none;')
        layout.addWidget(lbl_total)

        return tbl, lbl_total, wrapper

    # ── Výber súborov ─────────────────────────────────────────────────────────
    def _pick_xml(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Vyber XML bankový výpis', '', 'XML súbory (*.xml)')
        if path:
            self.xml_edit.setText(path)
            base = os.path.splitext(path)[0]
            self.out_edit.setText(base + '_moneydata.xml')
            if not self.xlsx_edit.text():
                self.xlsx_edit.setText(os.path.dirname(path))
            self._load_original_preview(path)

    def _pick_xlsx(self):
        path = QFileDialog.getExistingDirectory(
            self, 'Vyber priečinok s GLS XLSX súbormi')
        if path:
            self.xlsx_edit.setText(path)

    def _pick_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, 'Ulož výstupný XML súbor', '', 'XML súbory (*.xml)')
        if path:
            self.out_edit.setText(path)

    # ── Načítaj náhľad pôvodného XML ──────────────────────────────────────────
    def _load_original_preview(self, xml_path: str):
        try:
            rows = read_original_xml(xml_path)
        except Exception as e:
            self.log_text.append(f'Chyba pri načítaní XML: {e}')
            return

        tbl = self.tbl_orig
        tbl.setRowCount(0)
        crdt_sum = 0.0
        dbit_sum = 0.0

        for r in rows:
            row_idx = tbl.rowCount()
            tbl.insertRow(row_idx)
            is_dbit = r['typ'] == 'DBIT'
            bg = BG_DBIT if is_dbit else (BG_ROW_ODD if row_idx % 2 else BG_ROW_EVEN)
            fg_sum = FG_DBIT if is_dbit else FG_DEFAULT
            tbl.setItem(row_idx, 0, _tbl_item(r['datum'], Qt.AlignCenter, bg=bg))
            tbl.setItem(row_idx, 1, _tbl_item(r['suma'],  Qt.AlignRight,  fg=fg_sum, bg=bg))
            tbl.setItem(row_idx, 2, _tbl_item(r['typ'],   Qt.AlignCenter, bg=bg))
            vs_fg = FG_DIM if not r['vs'] else FG_DEFAULT
            tbl.setItem(row_idx, 3, _tbl_item(r['vs'] or '–', fg=vs_fg, bg=bg))
            tbl.setItem(row_idx, 4, _tbl_item(r['popis'], bg=bg))

            try:
                val = float(r['suma'])
            except ValueError:
                val = 0.0
            if r['typ'] == 'CRDT':
                crdt_sum += val
            else:
                dbit_sum += val

        net = crdt_sum - dbit_sum
        self.lbl_total_orig.setText(
            f'Príjmy: {crdt_sum:,.2f} €   |   Výdaje: {dbit_sum:,.2f} €   |   '
            f'Rozdiel: {net:,.2f} €   |   Záznamy: {len(rows)}'
        )

    # ── Spusti konverziu ──────────────────────────────────────────────────────
    def _run(self):
        xml_path  = self.xml_edit.text().strip()
        xlsx_dir  = self.xlsx_edit.text().strip()
        out_path  = self.out_edit.text().strip()
        dcislo    = self.dcislo_spin.value()
        hosp_od   = self.hosp_od_edit.text().strip()
        hosp_do   = self.hosp_do_edit.text().strip()

        errors = []
        if not xml_path or not os.path.isfile(xml_path):
            errors.append('XML súbor neexistuje.')
        if not xlsx_dir or not os.path.isdir(xlsx_dir):
            errors.append('Priečinok XLSX neexistuje.')
        if not out_path:
            errors.append('Zadaj cestu k výstupnému súboru.')
        if not re.match(r'\d{4}-\d{2}-\d{2}', hosp_od):
            errors.append('Hospodársky rok od – nesprávny formát.')
        if not re.match(r'\d{4}-\d{2}-\d{2}', hosp_do):
            errors.append('Hospodársky rok do – nesprávny formát.')
        if errors:
            QMessageBox.warning(self, 'Chyba vstupu', '\n'.join(errors))
            return

        self.log_text.clear()
        self.tbl_out.setRowCount(0)
        self.lbl_total_out.setText('Celková suma: –')
        self.btn_run.setEnabled(False)
        self.progress.setVisible(True)

        self._worker = WorkerThread(
            xml_path, xlsx_dir, dcislo, hosp_od, hosp_do, out_path)
        self._worker.log_signal.connect(self.log_text.append)
        self._worker.done_signal.connect(self._on_done)
        self._worker.error_signal.connect(self._on_error)
        self._worker.start()

    def _on_done(self, count: int, out_rows: list):
        self.progress.setVisible(False)
        self.btn_run.setEnabled(True)
        self._fill_output_table(out_rows)
        QMessageBox.information(
            self, 'Hotovo',
            f'Konverzia dokončená.\nSpracovaných záznamov: {count}')

    def _fill_output_table(self, rows: list):
        tbl = self.tbl_out
        tbl.setRowCount(0)
        vydaj_sum = 0.0
        prijem_sum = 0.0

        for r in rows:
            row_idx = tbl.rowCount()
            tbl.insertRow(row_idx)
            is_gls   = r['gls']
            is_vydaj = r['vydej']
            if is_gls:
                bg     = BG_GLS
                fg_sum = FG_GLS
            elif is_vydaj:
                bg     = BG_DBIT
                fg_sum = FG_DBIT
            else:
                bg     = BG_ROW_ODD if row_idx % 2 else BG_ROW_EVEN
                fg_sum = FG_DEFAULT

            vydej_str = 'Výdaj' if is_vydaj else 'Príjem'
            fg_vydej  = FG_DBIT if is_vydaj else FG_DEFAULT
            vs_fg     = FG_GLS if is_gls else (FG_DIM if r['vs'] == '0000000000' else FG_DEFAULT)
            nm_fg     = FG_GLS if is_gls else FG_DEFAULT

            tbl.setItem(row_idx, 0, _tbl_item(r['datum'],    Qt.AlignCenter, bg=bg))
            tbl.setItem(row_idx, 1, _tbl_item(r['doklad'],                   bg=bg))
            tbl.setItem(row_idx, 2, _tbl_item(vydej_str,     Qt.AlignCenter, fg=fg_vydej, bg=bg))
            tbl.setItem(row_idx, 3, _tbl_item(r['vs'],                       fg=vs_fg,    bg=bg))
            tbl.setItem(row_idx, 4, _tbl_item(r['zakaznik'],                 fg=nm_fg,    bg=bg))
            tbl.setItem(row_idx, 5, _tbl_item(r['suma'],     Qt.AlignRight,  fg=fg_sum,   bg=bg))

            try:
                val = float(r['suma'])
            except ValueError:
                val = 0.0
            if is_vydaj:
                vydaj_sum  += val
            else:
                prijem_sum += val

        net = prijem_sum - vydaj_sum
        self.lbl_total_out.setText(
            f'Príjmy: {prijem_sum:,.2f} €   |   Výdaje: {vydaj_sum:,.2f} €   |   '
            f'Rozdiel: {net:,.2f} €   |   Záznamy: {len(rows)}'
        )

    def _on_error(self, msg: str):
        self.progress.setVisible(False)
        self.btn_run.setEnabled(True)
        self.log_text.append(f'\n❌ CHYBA:\n{msg}')
        QMessageBox.critical(self, 'Chyba', msg[:400])


# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
