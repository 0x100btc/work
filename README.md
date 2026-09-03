"""Mouser BOM Search Tool  Rev0.2"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import urllib.request
import urllib.error
import urllib.parse
import ssl
import json
import sqlite3
import re
import threading
import time
import os
import sys
import subprocess
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlrd  # legacy .xls (BIFF/OLE2) — openpyxl only reads .xlsx
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


def _load_excel_sheets(path):
    """Read every sheet of an .xlsx or legacy .xls file into {sheet_name: [row_tuple, ...]}."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xls':
        if not HAS_XLRD:
            raise RuntimeError('讀取舊版 .xls 檔需要 xlrd：pip install xlrd')
        book = xlrd.open_workbook(path)
        sheets = {}
        for sheet in book.sheets():
            rows = []
            for r in range(sheet.nrows):
                row = []
                for c in range(sheet.ncols):
                    v = sheet.cell_value(r, c)
                    if isinstance(v, float) and v.is_integer():
                        v = int(v)
                    row.append(v)
                rows.append(tuple(row))
            sheets[sheet.name] = rows
        return sheets
    if not HAS_OPENPYXL:
        raise RuntimeError('請先安裝 openpyxl：pip install openpyxl')
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = {name: list(wb[name].iter_rows(values_only=True)) for name in wb.sheetnames}
    wb.close()
    return sheets


API_KEY  = 'a29293ca-86c1-417a-b079-1d2848f5eccb'
BASE_URL = 'https://api.mouser.com/api/v2'
TIMEOUT  = 15

DK_CLIENT_ID     = 'apasNExKXJrGtGhYpL5XTlRmlQQ7yFAAiNl39lVeeVDATgHw'
DK_CLIENT_SECRET = 'KAloJpMLThNz96EDD8UJUQbbkAlZrpqcRDfyU0xTKAMd6l8m5tZsUpGGLgDz3uww'
DK_BASE_URL      = 'https://api.digikey.com'
DK_LOCALE_HDRS   = {
    'X-DIGIKEY-Locale-Site':     'US',
    'X-DIGIKEY-Locale-Language': 'en',
    'X-DIGIKEY-Locale-Currency': 'USD',
}

_BASE    = os.path.dirname(sys.executable if getattr(sys, 'frozen', False)
                           else os.path.abspath(__file__))
NOTE_FOLDER   = os.path.join(_BASE, 'note')
DB_PATH       = os.path.join(_BASE, 'bom_parts.db')
POWER_DB_PATH = os.path.join(_BASE, 'power.db')
TODO_DB_PATH  = os.path.join(_BASE, 'todo.db')

_PRIO_CLR  = {'high': '#FF3B30', 'medium': '#FF9500', 'low': '#007AFF'}
_TODO_BG   = '#F2F2F7'
_LIST_PAL  = ['#FF3B30', '#FF9500', '#FFCC00', '#34C759',
              '#007AFF', '#5856D6', '#FF2D55', '#AF52DE']
APP_CFG      = os.path.join(_BASE, 'app_config.json')
RISK_DB_PATH = os.path.join(_BASE, 'risk_scan.db')

_RISK_CLR = {
    'CRITICAL': '#FFB3B3',
    'HIGH':     '#FFD9B3',
    'MEDIUM':   '#FFFACC',
    'LOW':      '#D4FFCC',
}

# File tag colors: {key: (bg_color, menu_label)}
_PROJ_TAGS = {
    'red':    ('#FFCDD2', '🔴  Red'),
    'orange': ('#FFE0B2', '🟠  Orange'),
    'yellow': ('#FFFDE7', '🟡  Yellow'),
    'green':  ('#DCEDC8', '🟢  Green'),
    'blue':   ('#BBDEFB', '🔵  Blue'),
    'purple': ('#E1BEE7', '🟣  Purple'),
    'gray':   ('#EEEEEE', '⚫  Gray'),
}

FONT_UI = ('Microsoft JhengHei UI', 9)
FONT_TX = ('Courier New', 9)


# ── Data class ────────────────────────────────────────────────────────────────

class OrcadItem:
    __slots__ = ('code', 'value', 'vendor_pn', 'description',
                 'footprint', 'qty', 'locations', 'vendor')

    def __init__(self, code, value, description, footprint, qty, locations,
                 vendor='', vendor_pn=''):
        self.code        = code
        self.value       = value
        self.vendor_pn   = vendor_pn
        self.description = description
        self.footprint   = footprint
        self.qty         = qty
        self.locations   = locations
        self.vendor      = vendor


# ── BOM Parser ────────────────────────────────────────────────────────────────

class BomParser:
    CODE_COL = 1; VAL_COL = 2; VPN_COL = 3; VEN_COL = 4
    DSC_COL  = 5; FP_COL  = 6; QTY_COL = 7; LOC_COL = 8; NC_COL = 9

    IG_VAL  = ('NC_',)
    IG_LOC  = ('TP', 'H', 'G')

    @classmethod
    def parse(cls, path):
        text  = cls._read(path)
        lines = text.splitlines()
        start = next((i + 1 for i, ln in enumerate(lines)
                      if ln.strip().startswith('____')), -1)
        if start < 0:
            return []
        items, cur = [], None
        for ln in lines[start:]:
            if not ln.strip():
                continue
            parts = ln.split('\t')
            is_cont = ln.startswith('\t') or (
                parts[0].strip() == '' and len(parts) > cls.LOC_COL)
            if is_cont:
                if cur and len(parts) > cls.LOC_COL:
                    extra = parts[cls.LOC_COL].strip()
                    cur.locations += [l for l in extra.split(',')
                                      if l.strip() and not cls._ig_loc(l.strip())]
                continue
            if len(parts) <= cls.CODE_COL:
                continue

            def g(i):
                return parts[i].strip() if len(parts) > i else ''

            nc = g(cls.NC_COL).upper()
            if nc:
                cur = None  # reset so this NC item's continuation lines don't leak to prev item
                continue
            val = g(cls.VAL_COL)
            if any(val.upper().startswith(p) for p in cls.IG_VAL):
                cur = None
                continue
            try:
                qty = int(g(cls.QTY_COL))
            except ValueError:
                qty = 0
            loc_s = g(cls.LOC_COL)
            # Handle malformed lines where location is embedded in Item column
            # e.g. "U5189" instead of item=189 + location=U5
            if not loc_s and len(parts) <= cls.QTY_COL:
                m = re.match(r'^([A-Za-z]+\d+)\d{3}$', parts[0].strip())
                if m:
                    loc_s = m.group(1)
                    if not qty:
                        qty = 1
            locs  = [l for l in loc_s.split(',')
                     if l.strip() and not cls._ig_loc(l.strip())]
            cur = OrcadItem(g(cls.CODE_COL), val, g(cls.DSC_COL), g(cls.FP_COL),
                            qty, locs, g(cls.VEN_COL), g(cls.VPN_COL))
            items.append(cur)
        return items

    @classmethod
    def _ig_loc(cls, loc):
        return any(loc.upper().startswith(p) for p in cls.IG_LOC)

    @staticmethod
    def _read(path):
        for enc in ('utf-8', 'big5', 'cp950', 'latin-1'):
            try:
                with open(path, encoding=enc) as f:
                    return f.read()
            except (UnicodeDecodeError, LookupError):
                pass
        with open(path, encoding='utf-8', errors='replace') as f:
            return f.read()

    @staticmethod
    def valid_code(code):
        """Return True if 10-Code is complete (not empty/TBD/COM, ends with M)."""
        if not code or not code.strip():
            return False
        c = code.strip().upper()
        if c in ('TBD', 'NEW10CODE'):
            return False
        if c.startswith('COM'):
            return False
        if not c.endswith('M'):
            return False
        return True


# ── EXP Parser ───────────────────────────────────────────────────────────────

class ExpParser:
    COL_PART_REF = 2   # Part Reference (RefDes)
    COL_VALUE    = 3   # Value
    COL_VALUE    = 3   # Value
    COL_DESC     = 17  # DESCRIPTION  (uppercase — same data as col 21)
    COL_DESC2    = 21  # Description  (mixed-case — Orcad display field)
    COL_NC       = 36  # NC
    COL_PART_NO  = 39  # PART_NUMBER (10-code)
    COL_FP       = 40  # PCB Footprint
    COL_VALUE1   = 68  # VALUE1
    COL_VENDOR   = 69  # VENDOR  (uppercase — same data as col 72)
    COL_VENDOR2  = 72  # Vendor  (mixed-case — Orcad display field)
    COL_VPN      = 70  # VENDOR_PN

    def parse(self, path):
        """Return (design_line, headers, rows) where rows is list[list[str]]."""
        text  = BomParser._read(path)
        lines = [ln.rstrip('\r\n') for ln in text.splitlines() if ln.strip()]
        design_line = lines[0]
        def _split(ln):
            return [f.strip('"') for f in ln.split('\t')]
        headers = _split(lines[1])
        rows    = [_split(ln) for ln in lines[2:]]
        return design_line, headers, rows

    def write(self, path, design_line, headers, rows):
        def _q(v): return f'"{v}"'
        ncols = len(headers)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(design_line + '\n')
            f.write('\t'.join(_q(h) for h in headers) + '\n')
            for row in rows:
                # Normalize: pad short rows with empty strings, trim long rows
                normalized = (list(row) + [''] * ncols)[:ncols]
                f.write('\t'.join(_q(v) for v in normalized) + '\n')


# ── Mouser API ────────────────────────────────────────────────────────────────

class MouserAPIError(Exception):
    pass


class MouserClient:
    def __init__(self, api_key=API_KEY):
        self.api_key = api_key

    def _post(self, endpoint, body, _retries=3):
        url  = f'{BASE_URL}/{endpoint}?apiKey={self.api_key}'
        data = json.dumps(body).encode('utf-8')
        ctx  = ssl.create_default_context()
        for attempt in range(_retries):
            req = urllib.request.Request(
                url, data=data, method='POST',
                headers={'Content-Type': 'application/json',
                         'Accept': 'application/json'})
            try:
                with urllib.request.urlopen(req, context=ctx, timeout=TIMEOUT) as r:
                    return json.loads(r.read())
            except urllib.error.HTTPError as e:
                if e.code == 403 and attempt < _retries - 1:
                    time.sleep(3 * (attempt + 1))   # 3s, 6s, 9s
                    continue
                raise MouserAPIError(f'HTTP {e.code}: {e.reason}')
            except urllib.error.URLError as e:
                raise MouserAPIError(f'Network error: {e.reason}')

    def _check(self, data):
        errs = data.get('Errors') or []
        if errs:
            raise MouserAPIError(errs[0].get('Message', str(errs)))

    def search_by_part(self, mpn, _retries=3):
        data = self._post('search/partnumber', {
            'SearchByPartRequest': {
                'mouserPartNumber': mpn.strip(),
                'partSearchOptions': ''}},
            _retries=_retries)
        self._check(data)
        return (data.get('SearchResults') or {}).get('Parts') or []

    def search_by_keyword(self, kw):
        data = self._post('search/keyword', {
            'SearchByKeywordRequest': {
                'keyword': kw.strip(), 'records': 20,
                'startingRecord': 0, 'searchOptions': ''}})
        self._check(data)
        return (data.get('SearchResults') or {}).get('Parts') or []

    @staticmethod
    def price_at(part, target_qty=1000):
        """Return (price_str, currency) for the best bracket <= target_qty."""
        pb = part.get('PriceBreaks') or []
        price, currency = '', ''
        for brk in reversed(pb):
            try:
                q = int(str(brk.get('Quantity', 0)).replace(',', ''))
                if q <= target_qty:
                    price    = brk.get('Price', '')
                    currency = brk.get('Currency', '')
                    break
            except Exception:
                pass
        if not price and pb:
            price    = pb[-1].get('Price', '')
            currency = pb[-1].get('Currency', '')
        return price, currency

    @staticmethod
    def price_best(part):
        """Return (price_str, currency) at the lowest unit price (highest qty break)."""
        pb = part.get('PriceBreaks') or []
        if not pb:
            return '', ''
        brk = pb[-1]
        return brk.get('Price', ''), brk.get('Currency', '')

    @staticmethod
    def _parse_price(price_str):
        try:
            return float(price_str.replace('$', '').replace(',', ''))
        except Exception:
            return 0.0

    @staticmethod
    def fmt_detail(part):
        pb = part.get('PriceBreaks') or []
        if pb:
            pb_lines = ['  Price Breaks :']
            for b in pb:
                pb_lines.append(
                    f"    {b.get('Quantity',''):>8} qty  {b.get('Price',''):>8} {b.get('Currency','')}")
        else:
            pb_lines = ['  Price Breaks : N/A']
        lines = (
            ['┌' + '─' * 76 + '┐',
             f"  Manufacturer : {part.get('Manufacturer','')}",
             f"  MPN          : {part.get('ManufacturerPartNumber','')}",
             f"  Mouser PN    : {part.get('MouserPartNumber','')}",
             f"  Description  : {part.get('Description','')}",
             f"  Availability : {part.get('Availability','')}",
             f"  Min / Mult   : {part.get('Min','')} / {part.get('Mult','')}"]
            + pb_lines
            + [f"  RoHS         : {part.get('ROHSStatus','')}",
               '└' + '─' * 76 + '┘',
               '']
        )
        return '\n'.join(lines)


# ── DigiKey API ───────────────────────────────────────────────────────────────

class DigiKeyError(Exception):
    pass


class DigiKeyClient:
    def __init__(self, client_id=DK_CLIENT_ID, client_secret=DK_CLIENT_SECRET):
        self._id     = client_id
        self._sec    = client_secret
        self._token  = ''
        self._exp_at = 0.0   # unix timestamp

    def _fetch_token(self):
        url  = f'{DK_BASE_URL}/v1/oauth2/token'
        body = urllib.parse.urlencode({
            'grant_type':    'client_credentials',
            'client_id':     self._id,
            'client_secret': self._sec,
        }).encode('utf-8')
        req = urllib.request.Request(
            url, data=body, method='POST',
            headers={'Content-Type': 'application/x-www-form-urlencoded'})
        ctx = ssl.create_default_context()
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=TIMEOUT) as r:
                d = json.loads(r.read())
            self._token  = d['access_token']
            self._exp_at = time.time() + d.get('expires_in', 1800) - 60
        except urllib.error.HTTPError as e:
            msg = e.read().decode('utf-8', errors='replace')
            raise DigiKeyError(f'Token HTTP {e.code}: {msg[:200]}')
        except urllib.error.URLError as e:
            raise DigiKeyError(f'Token network error: {e.reason}')

    def _ensure_token(self):
        if not self._token or time.time() >= self._exp_at:
            self._fetch_token()

    def _headers(self):
        return {
            'Authorization':        f'Bearer {self._token}',
            'X-DIGIKEY-Client-Id':  self._id,
            'Content-Type':         'application/json',
            'Accept':               'application/json',
            **DK_LOCALE_HDRS,
        }

    def _post_json(self, path, body):
        self._ensure_token()
        url  = DK_BASE_URL + path
        data = json.dumps(body).encode('utf-8')
        req  = urllib.request.Request(url, data=data, method='POST',
                                      headers=self._headers())
        ctx  = ssl.create_default_context()
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=TIMEOUT) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            msg = e.read().decode('utf-8', errors='replace')
            raise DigiKeyError(f'HTTP {e.code}: {msg[:200]}')
        except urllib.error.URLError as e:
            raise DigiKeyError(f'Network error: {e.reason}')

    def search_by_mpn(self, mpn, limit=10):
        data = self._post_json('/products/v4/search/keyword', {
            'Keywords': mpn.strip(), 'Limit': limit, 'Offset': 0})
        return data.get('Products') or []

    def search_by_keyword(self, kw, limit=20):
        data = self._post_json('/products/v4/search/keyword', {
            'Keywords': kw.strip(), 'Limit': limit, 'Offset': 0})
        return data.get('Products') or []

    @staticmethod
    def _best_variation(product):
        """Pick Cut Tape (MOQ=1) variation with most stock; fall back to any."""
        variations = product.get('ProductVariations') or []
        if not variations:
            return None
        low_moq = [v for v in variations if v.get('MinimumOrderQuantity', 9999) == 1]
        pool    = low_moq if low_moq else variations
        return max(pool, key=lambda v: v.get('QuantityAvailableforPackageType', 0))

    @staticmethod
    def price_at(product, target_qty=1000):
        """Return (unit_price: float, 'USD') for best bracket <= target_qty."""
        var = DigiKeyClient._best_variation(product)
        if not var:
            return 0.0, 'USD'
        pricing = var.get('StandardPricing') or []
        price   = 0.0
        for brk in reversed(pricing):
            if brk.get('BreakQuantity', 0) <= target_qty:
                price = float(brk.get('UnitPrice', 0.0))
                break
        if not price and pricing:
            price = float(pricing[0].get('UnitPrice', 0.0))
        return price, 'USD'

    @staticmethod
    def price_best(product):
        """Return (lowest unit price: float, 'USD') at highest qty break."""
        var = DigiKeyClient._best_variation(product)
        if not var:
            return 0.0, 'USD'
        pricing = var.get('StandardPricing') or []
        if not pricing:
            return 0.0, 'USD'
        return float(pricing[-1].get('UnitPrice', 0.0)), 'USD'

    @staticmethod
    def fmt_detail(product):
        mpn    = product.get('ManufacturerProductNumber', '')
        mfr    = (product.get('Manufacturer') or {}).get('Name', '')
        desc   = (product.get('Description') or {}).get('ProductDescription', '')
        qty    = product.get('QuantityAvailable', 0)
        rohs   = (product.get('Classifications') or {}).get('RohsStatus', '')
        status = (product.get('ProductStatus') or {}).get('Status', '')

        var    = DigiKeyClient._best_variation(product)
        dkpn   = var.get('DigiKeyProductNumber', '') if var else ''
        moq    = var.get('MinimumOrderQuantity', 1) if var else 1
        pkg    = (var.get('PackageType') or {}).get('Name', '') if var else ''
        stock  = var.get('QuantityAvailableforPackageType', qty) if var else qty
        avail_s = f'{stock:,}' if stock else '0'

        pricing = (var.get('StandardPricing') or []) if var else []
        if pricing:
            pb_lines = ['  Price Breaks :']
            for b in pricing:
                pb_lines.append(
                    f"    {b.get('BreakQuantity', 0):>8} qty  "
                    f"${float(b.get('UnitPrice', 0)):>8.4f} USD")
        else:
            pb_lines = ['  Price Breaks : N/A']

        lines = (
            ['┌' + '─' * 76 + '┐',
             f'  MPN          : {mpn}',
             f'  DigiKey PN   : {dkpn}  [{pkg}]',
             f'  Manufacturer : {mfr}',
             f'  Description  : {desc}',
             f'  Availability : {avail_s} in stock  (MOQ: {moq})',
             f'  Status       : {status}  |  {rohs}',
             ]
            + pb_lines
            + ['└' + '─' * 76 + '┘', '']
        )
        return '\n'.join(lines)


# ── Parts DB ─────────────────────────────────────────────────────────────────

class PartsDB:
    _DDL = [
        '''CREATE TABLE IF NOT EXISTS projects (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL UNIQUE,
            source_file TEXT,
            imported_at TEXT DEFAULT (datetime('now','localtime')))''',
        '''CREATE TABLE IF NOT EXISTS parts (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            internal_pn  TEXT NOT NULL DEFAULT '',
            vendor_pn    TEXT NOT NULL DEFAULT '',
            description  TEXT DEFAULT '',
            manufacturer TEXT DEFAULT '')''',
        '''CREATE TABLE IF NOT EXISTS bom_entries (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
            part_id    INTEGER NOT NULL REFERENCES parts(id),
            qty        INTEGER DEFAULT 0,
            reference  TEXT DEFAULT '',
            level      INTEGER DEFAULT 0,
            lf_flag    TEXT DEFAULT '',
            UNIQUE(project_id, part_id))''',
        'CREATE INDEX IF NOT EXISTS idx_vpn ON parts(vendor_pn COLLATE NOCASE)',
        'CREATE INDEX IF NOT EXISTS idx_ipn ON parts(internal_pn COLLATE NOCASE)',
        '''CREATE TABLE IF NOT EXISTS part_attachments (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id   INTEGER NOT NULL REFERENCES parts(id) ON DELETE CASCADE,
            file_path TEXT NOT NULL,
            label     TEXT DEFAULT "",
            added_at  TEXT DEFAULT "")''',
        'CREATE INDEX IF NOT EXISTS idx_pa_part ON part_attachments(part_id)',
    ]

    def __init__(self, path=DB_PATH):
        self.path = path
        with self._conn() as c:
            for stmt in self._DDL:
                c.execute(stmt)
            existing = {row[1] for row in c.execute('PRAGMA table_info(bom_entries)')}
            for col, dflt in (('level', '0'), ('lf_flag', "''")):
                if col not in existing:
                    c.execute(f'ALTER TABLE bom_entries ADD COLUMN {col} '
                              f'{"INTEGER" if col == "level" else "TEXT"} DEFAULT {dflt}')
            existing_p = {row[1] for row in c.execute('PRAGMA table_info(parts)')}
            if 'priority' not in existing_p:
                c.execute('ALTER TABLE parts ADD COLUMN priority INTEGER DEFAULT NULL')

    def _conn(self):
        conn = sqlite3.connect(self.path)
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        return conn

    def stats(self):
        with self._conn() as c:
            p = c.execute('SELECT COUNT(*) FROM parts').fetchone()[0]
            j = c.execute('SELECT COUNT(*) FROM projects').fetchone()[0]
        return p, j

    def get_projects(self):
        with self._conn() as c:
            return c.execute('''
                SELECT pr.id, pr.name, pr.imported_at,
                       COUNT(be.id) AS part_count
                FROM projects pr
                LEFT JOIN bom_entries be ON be.project_id = pr.id
                GROUP BY pr.id ORDER BY pr.imported_at DESC
            ''').fetchall()

    def delete_project(self, project_id):
        with self._conn() as c:
            c.execute('DELETE FROM bom_entries WHERE project_id=?', (project_id,))
            c.execute('DELETE FROM projects WHERE id=?', (project_id,))

    def import_rows(self, rows, project_name, source_file=''):
        with self._conn() as c:
            c.execute('INSERT OR IGNORE INTO projects(name,source_file) VALUES(?,?)',
                      (project_name, source_file))
            c.execute('UPDATE projects SET source_file=?, '
                      'imported_at=datetime("now","localtime") WHERE name=?',
                      (source_file, project_name))
            proj_id = c.execute('SELECT id FROM projects WHERE name=?',
                                (project_name,)).fetchone()[0]
            c.execute('DELETE FROM bom_entries WHERE project_id=?', (proj_id,))

            inserted = 0
            for row in rows:
                ipn  = str(row.get('internal_pn') or '').strip()
                vpn  = str(row.get('vendor_pn')   or '').strip()
                desc = str(row.get('description')  or '').strip()
                mfr  = str(row.get('manufacturer') or '').strip()
                qty  = row.get('qty') or 0
                ref  = str(row.get('reference')    or '').strip()
                try:
                    prio = int(row.get('priority')) if row.get('priority') not in (None, '') else None
                except (ValueError, TypeError):
                    prio = None
                if not ipn and not vpn:
                    continue
                try:
                    qty = int(str(qty).replace(',', '')) if qty else 0
                except (ValueError, TypeError):
                    qty = 0

                existing = c.execute(
                    'SELECT id FROM parts WHERE internal_pn=? AND vendor_pn=?',
                    (ipn, vpn)).fetchone()
                if existing:
                    part_id = existing[0]
                    c.execute(
                        'UPDATE parts SET '
                        'description=COALESCE(NULLIF(description,""),?), '
                        'manufacturer=COALESCE(NULLIF(manufacturer,""),?), '
                        'priority=COALESCE(?,priority) WHERE id=?',
                        (desc, mfr, prio, part_id))
                else:
                    cur = c.execute(
                        'INSERT INTO parts(internal_pn,vendor_pn,description,manufacturer,priority)'
                        ' VALUES(?,?,?,?,?)', (ipn, vpn, desc, mfr, prio))
                    part_id = cur.lastrowid

                level   = int(row.get('level')   or 0)
                lf_flag = str(row.get('lf_flag') or '').strip()
                c.execute(
                    'INSERT OR REPLACE INTO bom_entries'
                    '(project_id,part_id,qty,reference,level,lf_flag)'
                    ' VALUES(?,?,?,?,?,?)',
                    (proj_id, part_id, qty, ref, level, lf_flag))
                inserted += 1
        return inserted

    def lookup_vendor_pn_exact(self, vpn):
        """Exact case-insensitive match; falls back to LIKE if no exact hit."""
        _sql = '''
            SELECT p.internal_pn, p.vendor_pn, p.description, p.manufacturer,
                   p.priority,
                   COUNT(DISTINCT be.project_id) AS proj_count,
                   SUM(be.qty)                  AS total_qty,
                   GROUP_CONCAT(pr.name, ', ')  AS projects
            FROM parts p
            JOIN bom_entries be ON be.part_id = p.id
            JOIN projects    pr ON pr.id       = be.project_id
            WHERE {}
            GROUP BY p.id
            ORDER BY proj_count DESC, total_qty DESC
        '''
        with self._conn() as c:
            rows = c.execute(_sql.format('p.vendor_pn = ? COLLATE NOCASE'),
                             (vpn,)).fetchall()
            if not rows:
                rows = c.execute(
                    _sql.format('p.vendor_pn LIKE ? COLLATE NOCASE'),
                    (f'%{vpn}%',)).fetchall()
        return rows

    def search_vendor_pn(self, vpn):
        with self._conn() as c:
            return c.execute('''
                SELECT p.internal_pn, p.vendor_pn, p.description, p.manufacturer,
                       p.priority,
                       COUNT(DISTINCT be.project_id) AS proj_count,
                       SUM(be.qty)                  AS total_qty,
                       GROUP_CONCAT(pr.name, ', ')  AS projects
                FROM parts p
                JOIN bom_entries be ON be.part_id  = p.id
                JOIN projects    pr ON pr.id        = be.project_id
                WHERE p.vendor_pn LIKE ? COLLATE NOCASE
                   OR p.internal_pn LIKE ? COLLATE NOCASE
                GROUP BY p.id
                ORDER BY proj_count DESC, total_qty DESC
            ''', (f'%{vpn}%', f'%{vpn}%')).fetchall()

    def popular_parts(self, keywords=None, limit=50, project_id=None):
        where, params = [], []
        for kw in (keywords or []):
            if kw:
                k = f'%{kw}%'
                where.append('(p.description LIKE ? OR p.vendor_pn LIKE ? OR p.internal_pn LIKE ?)')
                params += [k, k, k]
        if project_id:
            where.append('be.project_id = ?')
            params.append(project_id)
        w = ('WHERE ' + ' AND '.join(where)) if where else ''
        params.append(limit)
        with self._conn() as c:
            return c.execute(f'''
                SELECT p.internal_pn, p.vendor_pn, p.description, p.manufacturer,
                       p.priority,
                       COUNT(DISTINCT be.project_id) AS proj_count,
                       SUM(be.qty)                  AS total_qty,
                       GROUP_CONCAT(pr.name, ', ')  AS projects
                FROM parts p
                JOIN bom_entries be ON be.part_id = p.id
                JOIN projects    pr ON pr.id       = be.project_id
                {w}
                GROUP BY p.id
                ORDER BY proj_count DESC, total_qty DESC
                LIMIT ?
            ''', params).fetchall()

    def add_attachment(self, part_id, file_path, label=''):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with self._conn() as c:
            c.execute('INSERT INTO part_attachments(part_id,file_path,label,added_at)'
                      ' VALUES(?,?,?,?)', (part_id, file_path, label, ts))

    def get_attachments(self, part_id):
        with self._conn() as c:
            return c.execute(
                'SELECT id, file_path, label FROM part_attachments WHERE part_id=?'
                ' ORDER BY id', (part_id,)).fetchall()

    def delete_attachment(self, att_id):
        with self._conn() as c:
            c.execute('DELETE FROM part_attachments WHERE id=?', (att_id,))

    def get_second_sources(self, internal_pn: str, exclude_vpn: str = '') -> list:
        """Return parts rows with matching internal_pn, excluding the given vendor_pn."""
        with self._conn() as c:
            return c.execute(
                'SELECT internal_pn, vendor_pn, description, manufacturer '
                'FROM parts WHERE internal_pn=? COLLATE NOCASE '
                'AND (? = "" OR vendor_pn != ? COLLATE NOCASE)',
                (internal_pn, exclude_vpn, exclude_vpn)
            ).fetchall()


# ── Power DB ─────────────────────────────────────────────────────────────────

class PowerDB:
    _DDL = [
        '''CREATE TABLE IF NOT EXISTS pwr_parts (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_pn      TEXT UNIQUE NOT NULL,
            company_pn     TEXT DEFAULT '',
            manufacturer   TEXT DEFAULT '',
            description    TEXT DEFAULT '',
            datasheet_path TEXT DEFAULT '',
            remark         TEXT DEFAULT '',
            created_at     TEXT DEFAULT (datetime('now','localtime')),
            updated_at     TEXT DEFAULT (datetime('now','localtime')))''',
        '''CREATE TABLE IF NOT EXISTS power_rails (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id           INTEGER NOT NULL REFERENCES pwr_parts(id) ON DELETE CASCADE,
            rail_name         TEXT NOT NULL DEFAULT '',
            power_group       TEXT DEFAULT '',
            input_rail        TEXT DEFAULT '',
            rail_description  TEXT DEFAULT '',
            voltage           REAL DEFAULT 0,
            tolerance_percent REAL DEFAULT 0,
            typ_current_a     REAL DEFAULT 0,
            max_current_a     REAL DEFAULT 0,
            condition         TEXT DEFAULT '',
            pin_names         TEXT DEFAULT '')''',
        '''CREATE TABLE IF NOT EXISTS power_sources (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id        INTEGER REFERENCES pwr_parts(id) ON DELETE CASCADE,
            rail_name      TEXT DEFAULT '',
            source_rail    TEXT DEFAULT '',
            regulator_type TEXT DEFAULT '',
            efficiency     REAL DEFAULT 1.0,
            input_voltage  REAL DEFAULT 0,
            output_voltage REAL DEFAULT 0,
            remark         TEXT DEFAULT '')''',
        'CREATE INDEX IF NOT EXISTS idx_pwr_vpn ON pwr_parts(vendor_pn COLLATE NOCASE)',
    ]

    def __init__(self, path=POWER_DB_PATH):
        self.path = path
        with self._conn() as c:
            for stmt in self._DDL:
                c.execute(stmt)
            # migrate existing DBs — add input_rail if missing
            try:
                c.execute("ALTER TABLE power_rails ADD COLUMN input_rail TEXT DEFAULT ''")
            except Exception:
                pass

    def _conn(self):
        conn = sqlite3.connect(self.path)
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        return conn

    def stats(self):
        with self._conn() as c:
            p = c.execute('SELECT COUNT(*) FROM pwr_parts').fetchone()[0]
            r = c.execute('SELECT COUNT(*) FROM power_rails').fetchone()[0]
        return p, r

    def search_parts(self, keyword=''):
        with self._conn() as c:
            if keyword:
                k = f'%{keyword}%'
                return c.execute(
                    'SELECT * FROM pwr_parts WHERE vendor_pn LIKE ? OR '
                    'company_pn LIKE ? OR description LIKE ? ORDER BY vendor_pn',
                    (k, k, k)).fetchall()
            return c.execute('SELECT * FROM pwr_parts ORDER BY vendor_pn').fetchall()

    def get_part(self, part_id):
        with self._conn() as c:
            return c.execute('SELECT * FROM pwr_parts WHERE id=?',
                             (part_id,)).fetchone()

    def get_part_by_vpn(self, vpn):
        with self._conn() as c:
            row = c.execute('SELECT * FROM pwr_parts WHERE vendor_pn=? COLLATE NOCASE',
                            (vpn,)).fetchone()
            if not row:
                row = c.execute(
                    'SELECT * FROM pwr_parts WHERE vendor_pn LIKE ? COLLATE NOCASE',
                    (f'%{vpn}%',)).fetchone()
        return row

    def upsert_part(self, data: dict):
        pid = data.get('id')  # set when editing an existing part
        vpn = data.get('vendor_pn', '')
        with self._conn() as c:
            if pid:
                # Editing existing record — UPDATE by id so VPN can be renamed
                c.execute(
                    '''UPDATE pwr_parts SET vendor_pn=?,company_pn=?,manufacturer=?,
                       description=?,datasheet_path=?,remark=?,
                       updated_at=datetime('now','localtime') WHERE id=?''',
                    (vpn, data.get('company_pn', ''), data.get('manufacturer', ''),
                     data.get('description', ''), data.get('datasheet_path', ''),
                     data.get('remark', ''), pid))
                return pid
            # New part — check for duplicate VPN first
            existing = c.execute(
                'SELECT id FROM pwr_parts WHERE vendor_pn=? COLLATE NOCASE',
                (vpn,)).fetchone()
            if existing:
                part_id = existing[0]
                c.execute(
                    '''UPDATE pwr_parts SET company_pn=?,manufacturer=?,
                       description=?,datasheet_path=?,remark=?,
                       updated_at=datetime('now','localtime') WHERE id=?''',
                    (data.get('company_pn', ''), data.get('manufacturer', ''),
                     data.get('description', ''), data.get('datasheet_path', ''),
                     data.get('remark', ''), part_id))
            else:
                cur = c.execute(
                    'INSERT INTO pwr_parts(vendor_pn,company_pn,manufacturer,'
                    'description,datasheet_path,remark) VALUES(?,?,?,?,?,?)',
                    (vpn, data.get('company_pn', ''), data.get('manufacturer', ''),
                     data.get('description', ''), data.get('datasheet_path', ''),
                     data.get('remark', '')))
                part_id = cur.lastrowid
        return part_id

    def delete_part(self, part_id):
        with self._conn() as c:
            c.execute('DELETE FROM pwr_parts WHERE id=?', (part_id,))

    def get_rails(self, part_id):
        with self._conn() as c:
            return c.execute(
                'SELECT * FROM power_rails WHERE part_id=? ORDER BY id',
                (part_id,)).fetchall()

    def upsert_rail(self, data: dict):
        with self._conn() as c:
            rid = data.get('id')
            if rid:
                c.execute('''UPDATE power_rails SET rail_name=?, power_group=?,
                             input_rail=?, rail_description=?, voltage=?,
                             tolerance_percent=?, typ_current_a=?, max_current_a=?,
                             condition=?, pin_names=? WHERE id=?''',
                          (data.get('rail_name', ''), data.get('power_group', ''),
                           data.get('input_rail', ''),
                           data.get('rail_description', ''), data.get('voltage', 0),
                           data.get('tolerance_percent', 0), data.get('typ_current_a', 0),
                           data.get('max_current_a', 0), data.get('condition', ''),
                           data.get('pin_names', ''), rid))
                return rid
            cur = c.execute(
                'INSERT INTO power_rails(part_id,rail_name,power_group,input_rail,'
                'rail_description,voltage,tolerance_percent,typ_current_a,max_current_a,'
                'condition,pin_names) VALUES(?,?,?,?,?,?,?,?,?,?,?)',
                (data['part_id'], data.get('rail_name', ''), data.get('power_group', ''),
                 data.get('input_rail', ''),
                 data.get('rail_description', ''), data.get('voltage', 0),
                 data.get('tolerance_percent', 0), data.get('typ_current_a', 0),
                 data.get('max_current_a', 0), data.get('condition', ''),
                 data.get('pin_names', '')))
            return cur.lastrowid

    def delete_rail(self, rail_id):
        with self._conn() as c:
            c.execute('DELETE FROM power_rails WHERE id=?', (rail_id,))

    def get_sources(self, part_id):
        with self._conn() as c:
            return c.execute(
                'SELECT * FROM power_sources WHERE part_id=? ORDER BY id',
                (part_id,)).fetchall()

    def upsert_source(self, data: dict):
        with self._conn() as c:
            sid = data.get('id')
            if sid:
                c.execute('''UPDATE power_sources SET rail_name=?, source_rail=?,
                             regulator_type=?, efficiency=?, input_voltage=?,
                             output_voltage=?, remark=? WHERE id=?''',
                          (data.get('rail_name', ''), data.get('source_rail', ''),
                           data.get('regulator_type', ''), data.get('efficiency', 1.0),
                           data.get('input_voltage', 0), data.get('output_voltage', 0),
                           data.get('remark', ''), sid))
                return sid
            cur = c.execute(
                'INSERT INTO power_sources(part_id,rail_name,source_rail,regulator_type,'
                'efficiency,input_voltage,output_voltage,remark) VALUES(?,?,?,?,?,?,?,?)',
                (data['part_id'], data.get('rail_name', ''), data.get('source_rail', ''),
                 data.get('regulator_type', ''), data.get('efficiency', 1.0),
                 data.get('input_voltage', 0), data.get('output_voltage', 0),
                 data.get('remark', '')))
            return cur.lastrowid

    def delete_source(self, source_id):
        with self._conn() as c:
            c.execute('DELETE FROM power_sources WHERE id=?', (source_id,))

    def get_all_sources(self):
        """All power_sources rows, LEFT JOIN'd with the owning Part's vendor_pn (NULL if none)."""
        with self._conn() as c:
            return c.execute(
                'SELECT ps.*, p.vendor_pn AS owner_vpn '
                'FROM power_sources ps LEFT JOIN pwr_parts p ON ps.part_id = p.id '
                'ORDER BY ps.id').fetchall()

    def get_library_sources(self):
        """Sources not tied to any Part — reusable DC/DC templates."""
        with self._conn() as c:
            return c.execute(
                'SELECT * FROM power_sources WHERE part_id IS NULL ORDER BY id').fetchall()

    def get_all_rail_names(self):
        """Distinct rail_name values across every Part, for combobox suggestions."""
        with self._conn() as c:
            return [row[0] for row in c.execute(
                "SELECT DISTINCT rail_name FROM power_rails WHERE rail_name != '' "
                "ORDER BY rail_name COLLATE NOCASE").fetchall()]


# ── Todo DB ──────────────────────────────────────────────────────────────────

class TodoDB:
    _DDL = [
        '''CREATE TABLE IF NOT EXISTS todo_lists (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT NOT NULL UNIQUE,
            color      TEXT DEFAULT "#007AFF",
            sort_order INTEGER DEFAULT 0)''',
        '''CREATE TABLE IF NOT EXISTS todo_items (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            list_id      INTEGER NOT NULL REFERENCES todo_lists(id) ON DELETE CASCADE,
            title        TEXT NOT NULL DEFAULT "",
            notes        TEXT DEFAULT "",
            due_date     TEXT DEFAULT "",
            priority     TEXT DEFAULT "",
            completed    INTEGER DEFAULT 0,
            created_at   TEXT DEFAULT "",
            completed_at TEXT DEFAULT "")''',
        'CREATE INDEX IF NOT EXISTS idx_ti_list ON todo_items(list_id)',
        '''CREATE TABLE IF NOT EXISTS todo_attachments (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id   INTEGER NOT NULL REFERENCES todo_items(id) ON DELETE CASCADE,
            file_path TEXT NOT NULL,
            added_at  TEXT DEFAULT "")''',
        'CREATE INDEX IF NOT EXISTS idx_ta_item ON todo_attachments(item_id)',
        '''CREATE TABLE IF NOT EXISTS todo_subitems (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id    INTEGER NOT NULL REFERENCES todo_items(id) ON DELETE CASCADE,
            title      TEXT NOT NULL DEFAULT "",
            due_date   TEXT DEFAULT "",
            completed  INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT "")''',
        'CREATE INDEX IF NOT EXISTS idx_tsi_item ON todo_subitems(item_id)',
    ]
    _DEFAULTS = [('Reminders', '#FF3B30', 0),
                 ('Work',      '#007AFF', 1),
                 ('Personal',  '#34C759', 2)]

    # Columns added after the original schema — plain ALTER TABLE ADD COLUMN
    # (no REFERENCES/constraints: SQLite's handling of those via ALTER TABLE
    # is version-dependent, so cleanup for these is done explicitly in code
    # instead of relying on ON DELETE CASCADE).
    _MIGRATIONS = [
        'ALTER TABLE todo_attachments ADD COLUMN subitem_id INTEGER DEFAULT NULL',
        'ALTER TABLE todo_items ADD COLUMN linked_note_path TEXT DEFAULT ""',
    ]

    def __init__(self, path=TODO_DB_PATH):
        self.path = path
        with self._conn() as c:
            for stmt in self._DDL:
                c.execute(stmt)
            if c.execute('SELECT COUNT(*) FROM todo_lists').fetchone()[0] == 0:
                c.executemany(
                    'INSERT OR IGNORE INTO todo_lists(name,color,sort_order) VALUES(?,?,?)',
                    self._DEFAULTS)
            for stmt in self._MIGRATIONS:
                try:
                    c.execute(stmt)
                except sqlite3.OperationalError:
                    pass   # column already exists from a previous run

    def _conn(self):
        conn = sqlite3.connect(self.path)
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        return conn

    def get_lists(self):
        with self._conn() as c:
            return c.execute('''
                SELECT l.id, l.name, l.color,
                       COUNT(CASE WHEN i.completed=0 THEN 1 END) AS active_count
                FROM todo_lists l
                LEFT JOIN todo_items i ON i.list_id = l.id
                GROUP BY l.id ORDER BY l.sort_order, l.id
            ''').fetchall()

    def add_list(self, name, color='#007AFF'):
        with self._conn() as c:
            c.execute('INSERT OR IGNORE INTO todo_lists(name,color) VALUES(?,?)', (name, color))
            return c.execute('SELECT id FROM todo_lists WHERE name=?', (name,)).fetchone()[0]

    def delete_list(self, list_id):
        with self._conn() as c:
            c.execute('DELETE FROM todo_lists WHERE id=?', (list_id,))

    def rename_list(self, list_id, name, color):
        with self._conn() as c:
            c.execute('UPDATE todo_lists SET name=?, color=? WHERE id=?',
                      (name, color, list_id))

    def get_items(self, list_id, show_completed=True):
        with self._conn() as c:
            where = '' if show_completed else ' AND completed=0'
            return c.execute(f'''
                SELECT * FROM todo_items WHERE list_id=?{where}
                ORDER BY completed ASC,
                    CASE priority WHEN "high" THEN 0 WHEN "medium" THEN 1
                                  WHEN "low"  THEN 2 ELSE 3 END,
                    created_at ASC
            ''', (list_id,)).fetchall()

    def add_item(self, list_id, title, notes='', due_date='', priority=''):
        with self._conn() as c:
            cur = c.execute(
                'INSERT INTO todo_items(list_id,title,notes,due_date,priority,created_at) VALUES(?,?,?,?,?,?)',
                (list_id, title.strip(), notes, due_date, priority,
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            return cur.lastrowid

    def toggle(self, item_id):
        with self._conn() as c:
            done = c.execute('SELECT completed FROM todo_items WHERE id=?',
                             (item_id,)).fetchone()[0]
            new  = 0 if done else 1
            ts   = datetime.now().strftime('%Y-%m-%d %H:%M:%S') if new else ''
            c.execute('UPDATE todo_items SET completed=?, completed_at=? WHERE id=?',
                      (new, ts, item_id))

    def update(self, item_id, **kw):
        allowed = {'title', 'notes', 'due_date', 'priority'}
        sets = {k: v for k, v in kw.items() if k in allowed}
        if not sets:
            return
        c_str = ', '.join(f'{k}=?' for k in sets)
        with self._conn() as c:
            c.execute(f'UPDATE todo_items SET {c_str} WHERE id=?',
                      list(sets.values()) + [item_id])

    def delete_item(self, item_id):
        with self._conn() as c:
            c.execute('DELETE FROM todo_items WHERE id=?', (item_id,))

    def add_attachment(self, item_id, file_path, subitem_id=None):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with self._conn() as c:
            c.execute('INSERT INTO todo_attachments(item_id,file_path,added_at,subitem_id) '
                      'VALUES(?,?,?,?)', (item_id, file_path, ts, subitem_id))

    def get_attachments(self, item_id=None, subitem_id=None):
        with self._conn() as c:
            if subitem_id is not None:
                return c.execute(
                    'SELECT id, file_path FROM todo_attachments WHERE subitem_id=? ORDER BY id',
                    (subitem_id,)).fetchall()
            return c.execute(
                'SELECT id, file_path FROM todo_attachments '
                'WHERE item_id=? AND subitem_id IS NULL ORDER BY id',
                (item_id,)).fetchall()

    def delete_attachment(self, att_id):
        with self._conn() as c:
            c.execute('DELETE FROM todo_attachments WHERE id=?', (att_id,))

    def get_subitems(self, item_id):
        with self._conn() as c:
            return c.execute(
                'SELECT * FROM todo_subitems WHERE item_id=? ORDER BY sort_order, id',
                (item_id,)).fetchall()

    def add_subitem(self, item_id, title, due_date=''):
        with self._conn() as c:
            n = c.execute('SELECT COUNT(*) FROM todo_subitems WHERE item_id=?',
                          (item_id,)).fetchone()[0]
            cur = c.execute(
                'INSERT INTO todo_subitems(item_id,title,due_date,sort_order,created_at) '
                'VALUES(?,?,?,?,?)',
                (item_id, title.strip(), due_date, n,
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            return cur.lastrowid

    def toggle_subitem(self, subitem_id):
        with self._conn() as c:
            done = c.execute('SELECT completed FROM todo_subitems WHERE id=?',
                             (subitem_id,)).fetchone()[0]
            c.execute('UPDATE todo_subitems SET completed=? WHERE id=?',
                      (0 if done else 1, subitem_id))

    def update_subitem(self, subitem_id, **kw):
        allowed = {'title', 'due_date'}
        sets = {k: v for k, v in kw.items() if k in allowed}
        if not sets:
            return
        c_str = ', '.join(f'{k}=?' for k in sets)
        with self._conn() as c:
            c.execute(f'UPDATE todo_subitems SET {c_str} WHERE id=?',
                      list(sets.values()) + [subitem_id])

    def delete_subitem(self, subitem_id):
        with self._conn() as c:
            c.execute('DELETE FROM todo_attachments WHERE subitem_id=?', (subitem_id,))
            c.execute('DELETE FROM todo_subitems WHERE id=?', (subitem_id,))

    def search_items(self, query):
        """Cross-list search by title/notes (items) or title (sub-items)."""
        q = f'%{query}%'
        with self._conn() as c:
            items = c.execute('''
                SELECT ti.*, tl.name AS list_name, tl.color AS list_color
                FROM todo_items ti JOIN todo_lists tl ON tl.id = ti.list_id
                WHERE ti.title LIKE ? COLLATE NOCASE OR ti.notes LIKE ? COLLATE NOCASE
                ORDER BY ti.completed, ti.due_date
            ''', (q, q)).fetchall()
            subs = c.execute('''
                SELECT tsi.*, ti.title AS item_title, ti.list_id AS list_id,
                       tl.name AS list_name, tl.color AS list_color
                FROM todo_subitems tsi
                JOIN todo_items ti ON ti.id = tsi.item_id
                JOIN todo_lists tl ON tl.id = ti.list_id
                WHERE tsi.title LIKE ? COLLATE NOCASE
                ORDER BY tsi.completed
            ''', (q,)).fetchall()
            return items, subs

    def set_linked_note(self, item_id, path):
        with self._conn() as c:
            c.execute('UPDATE todo_items SET linked_note_path=? WHERE id=?', (path, item_id))

    def get_items_by_note(self, path):
        with self._conn() as c:
            return c.execute('''
                SELECT ti.*, tl.name AS list_name, tl.color AS list_color
                FROM todo_items ti JOIN todo_lists tl ON tl.id = ti.list_id
                WHERE ti.linked_note_path=?
            ''', (path,)).fetchall()

    def get_due_rollup(self, today_str):
        """Cross-list Today/Overdue rollup: active items and active sub-items
        whose due date is today or earlier, each tagged with their list."""
        with self._conn() as c:
            item_rows = c.execute('''
                SELECT ti.*, tl.name AS list_name, tl.color AS list_color
                FROM todo_items ti JOIN todo_lists tl ON tl.id = ti.list_id
                WHERE ti.completed=0 AND ti.due_date!='' AND ti.due_date<=?
                ORDER BY ti.due_date
            ''', (today_str,)).fetchall()
            sub_rows = c.execute('''
                SELECT tsi.*, ti.title AS item_title, ti.list_id AS list_id,
                       tl.name AS list_name, tl.color AS list_color
                FROM todo_subitems tsi
                JOIN todo_items ti ON ti.id = tsi.item_id
                JOIN todo_lists tl ON tl.id = ti.list_id
                WHERE tsi.completed=0 AND tsi.due_date!='' AND tsi.due_date<=?
                  AND ti.completed=0
                ORDER BY tsi.due_date
            ''', (today_str,)).fetchall()
            return item_rows, sub_rows

    def get_list_progress(self, list_id):
        """Average completion fraction across a list's items (0..1), and the
        item count. An item with sub-items contributes its checklist ratio;
        an item without sub-items contributes 0 or 1 (its own completed flag)."""
        items = self.get_items(list_id, show_completed=True)
        if not items:
            return 0.0, 0
        total = 0.0
        for it in items:
            subs = self.get_subitems(it['id'])
            if subs:
                total += sum(1 for s in subs if s['completed']) / len(subs)
            else:
                total += 1.0 if it['completed'] else 0.0
        return total / len(items), len(items)


# ── Risk Scan DB ──────────────────────────────────────────────────────────────

class RiskDB:
    _DDL = [
        '''CREATE TABLE IF NOT EXISTS risk_scans (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT NOT NULL,
            bom_source TEXT DEFAULT "",
            created_at TEXT DEFAULT "")''',
        '''CREATE TABLE IF NOT EXISTS risk_items (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_id      INTEGER NOT NULL REFERENCES risk_scans(id) ON DELETE CASCADE,
            vendor_pn    TEXT DEFAULT "",
            description  TEXT DEFAULT "",
            qty          INTEGER DEFAULT 0,
            mouser_stock INTEGER DEFAULT -1,
            dk_stock     INTEGER DEFAULT -1,
            lead_time_wk INTEGER DEFAULT -1,
            lifecycle    TEXT DEFAULT "",
            risk_inv     INTEGER DEFAULT 0,
            risk_lt      INTEGER DEFAULT 0,
            risk_src     INTEGER DEFAULT 0,
            risk_life    INTEGER DEFAULT 0,
            risk_total   INTEGER DEFAULT 0,
            risk_level   TEXT DEFAULT "")''',
        '''CREATE TABLE IF NOT EXISTS alt_parts (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            primary_vpn TEXT NOT NULL,
            alt_vpn     TEXT NOT NULL,
            alt_mfr     TEXT DEFAULT "",
            compat      TEXT DEFAULT "",
            note        TEXT DEFAULT "",
            added_at    TEXT DEFAULT "")''',
        'CREATE INDEX IF NOT EXISTS idx_ri_scan ON risk_items(scan_id)',
        'CREATE INDEX IF NOT EXISTS idx_alt_pri ON alt_parts(primary_vpn COLLATE NOCASE)',
    ]

    def __init__(self, path=RISK_DB_PATH):
        self.path = path
        with self._conn() as c:
            for stmt in self._DDL:
                c.execute(stmt)

    def _conn(self):
        conn = sqlite3.connect(self.path)
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        return conn

    def new_scan(self, name, bom_source=''):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with self._conn() as c:
            cur = c.execute(
                'INSERT INTO risk_scans(name,bom_source,created_at) VALUES(?,?,?)',
                (name, bom_source, ts))
            return cur.lastrowid

    def get_scans(self):
        with self._conn() as c:
            return c.execute(
                'SELECT * FROM risk_scans ORDER BY created_at DESC').fetchall()

    def delete_scan(self, scan_id):
        with self._conn() as c:
            c.execute('DELETE FROM risk_scans WHERE id=?', (scan_id,))

    def add_item(self, scan_id, vendor_pn, description='', qty=0):
        with self._conn() as c:
            cur = c.execute(
                'INSERT INTO risk_items(scan_id,vendor_pn,description,qty) VALUES(?,?,?,?)',
                (scan_id, vendor_pn.strip(), description.strip(), qty))
            return cur.lastrowid

    def update_item(self, item_id, **kw):
        allowed = {'mouser_stock','dk_stock','lead_time_wk','lifecycle',
                   'risk_inv','risk_lt','risk_src','risk_life','risk_total','risk_level'}
        sets = {k: v for k, v in kw.items() if k in allowed}
        if not sets:
            return
        with self._conn() as c:
            c.execute(f'UPDATE risk_items SET {", ".join(f"{k}=?" for k in sets)} WHERE id=?',
                      list(sets.values()) + [item_id])

    def get_items(self, scan_id):
        with self._conn() as c:
            return c.execute(
                'SELECT * FROM risk_items WHERE scan_id=? ORDER BY risk_total DESC, vendor_pn',
                (scan_id,)).fetchall()

    def add_alt(self, primary_vpn, alt_vpn, alt_mfr='', compat='', note=''):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with self._conn() as c:
            c.execute(
                'INSERT INTO alt_parts(primary_vpn,alt_vpn,alt_mfr,compat,note,added_at) VALUES(?,?,?,?,?,?)',
                (primary_vpn.strip(), alt_vpn.strip(), alt_mfr, compat, note, ts))

    def get_alts(self, primary_vpn=''):
        with self._conn() as c:
            if primary_vpn:
                return c.execute(
                    'SELECT * FROM alt_parts WHERE primary_vpn=? COLLATE NOCASE ORDER BY id',
                    (primary_vpn,)).fetchall()
            return c.execute('SELECT * FROM alt_parts ORDER BY primary_vpn').fetchall()

    def delete_alt(self, alt_id):
        with self._conn() as c:
            c.execute('DELETE FROM alt_parts WHERE id=?', (alt_id,))


# ── SAP BOM Parser ───────────────────────────────────────────────────────────

class SapBomParser:
    """
    Parses Delta SAP BOM text export (YRPEN019 report, fixed-width format).

    Column layout (0-indexed):
      0-6    ITEM number
      7-14   Level (STD L)
      15-32  PART NO  (company internal part number)
      33-72  MFG PART (vendor/manufacturer part number)
      73     * purchase/flag indicator
      75-76  L/F flag (AD=Active Design, AM=Active Mfg, TD=To Discontinue, …)
      89-131 DESCRIPTION
      143-155 QPA (right-justified float)
      156-161 UN  (unit)
      162+   SORT STRING (reference designators; continues on blank-prefix lines)
    """

    @classmethod
    def detect(cls, path):
        try:
            text = BomParser._read(path)
            return 'ASSEMBLY:' in text[:2000] and 'PART NO' in text[:2000]
        except Exception:
            return False

    @classmethod
    def parse(cls, path):
        """
        Returns (project_name: str, items: list[dict])
        Each item dict has keys:
            internal_pn, vendor_pn, description, manufacturer,
            qty, reference, level, lf_flag
        """
        text  = BomParser._read(path)
        lines = text.splitlines()

        # Project name from ASSEMBLY line
        project_name = os.path.splitext(os.path.basename(path))[0]
        for ln in lines[:10]:
            if 'ASSEMBLY:' in ln:
                after = ln[ln.find('ASSEMBLY:') + 9:].strip()
                project_name = after.split()[0] if after.split() else project_name
                break

        # Find column header line & separator
        hdr_line   = ''
        data_start = -1
        for i, ln in enumerate(lines):
            if 'PART NO' in ln and 'MFG PART' in ln and 'DESCRIPTION' in ln:
                hdr_line = ln
            elif hdr_line and ln.strip().startswith('---'):
                data_start = i + 1
                break

        if data_start < 0:
            raise ValueError('找不到 BOM 資料區段 — 確認格式為 YRPEN019 報表')

        # Derive column positions from the header line
        def fnd(key, default):
            idx = hdr_line.find(key)
            return idx if idx >= 0 else default

        c_partno = fnd('PART NO',      15)
        c_mfgprt = fnd('MFG PART',    34)
        c_lf     = fnd('L/F',         77)
        c_desc   = fnd('DESCRIPTION', 91)
        c_alt    = fnd('ALT',        132)
        c_qpa    = fnd('QPA',        141)
        c_un     = fnd(' UN',        151) + 1   # skip leading space
        c_sort   = fnd('SORT',       158)

        def field(ln, start, end=None):
            if len(ln) <= start:
                return ''
            chunk = ln[start:end] if end else ln[start:]
            return chunk.strip()

        items: list[dict] = []
        cur:   dict | None = None

        for ln in lines[data_start:]:
            # ── Continuation line: blank before PART NO, content at SORT col ──
            if ln[:c_mfgprt].strip() == '' and len(ln) > c_sort and ln[c_sort:].strip():
                if cur is not None:
                    cur['reference'] += ln[c_sort:].rstrip()
                continue

            # ── Main data line: integer item number in first 7 chars ──────────
            item_str = ln[:7].strip()
            if not item_str or not item_str.isdigit():
                continue

            part_no  = field(ln, c_partno, c_mfgprt)
            mfg_part = field(ln, c_mfgprt, c_lf - 2)    # stop before D-flag column
            lf_flag  = field(ln, c_lf,     c_lf + 3)

            # Description: up to ALT column
            desc = field(ln, c_desc, c_alt)

            # QPA: right-justified float in [c_qpa, c_un) region
            qpa_region = field(ln, c_qpa, c_un)
            m = re.search(r'[\d,]+\.?\d*', qpa_region)
            try:
                qty = int(round(float(m.group().replace(',', '')))) if m else 0
            except (ValueError, AttributeError):
                qty = 0

            # Level
            level_str = field(ln, 7, c_partno)
            try:
                level = int(level_str)
            except ValueError:
                level = 0

            ref_str = field(ln, c_sort)

            if not part_no and not mfg_part:
                cur = None
                continue

            cur = {
                'internal_pn':  part_no,
                'vendor_pn':    mfg_part,
                'description':  desc,
                'manufacturer': '',
                'qty':          qty,
                'reference':    ref_str,
                'level':        level,
                'lf_flag':      lf_flag,
            }
            items.append(cur)

        return project_name, items


# ── Vendor-PN column picker (for customer BOM lookup) ────────────────────────

class _VpnColPicker(tk.Toplevel):
    _HINTS = ['mpn', 'vendor', 'part no', 'partno', 'component',
              'mfr part', 'manufacturer part', '料號', '廠商料號', 'pn']

    def __init__(self, parent, columns, preview_rows):
        super().__init__(parent)
        self.title('選擇 Vendor PN 欄位')
        self.resizable(True, False)
        self.grab_set()
        self.result = None

        ttk.Label(self, text='選擇含 Vendor PN / MPN 的欄位：',
                  font=FONT_UI).pack(padx=12, pady=(10, 4), anchor='w')

        guess = next((c for c in columns
                      for h in self._HINTS if h.lower() in c.lower()), columns[0])
        self.var = tk.StringVar(value=guess)
        ttk.Combobox(self, textvariable=self.var, values=list(columns),
                     state='readonly', width=32).pack(padx=12, pady=4, anchor='w')

        ttk.Label(self, text='Preview (前 3 列)：', font=FONT_UI).pack(
            padx=12, anchor='w')
        tv = ttk.Treeview(self, columns=list(columns), show='headings', height=3)
        for col in columns:
            tv.heading(col, text=col)
            tv.column(col, width=max(60, min(160, len(str(col)) * 9)))
        for row in preview_rows[:3]:
            tv.insert('', 'end',
                      values=[str(v) if v is not None else '' for v in row])
        sb = ttk.Scrollbar(self, orient='horizontal', command=tv.xview)
        tv.configure(xscrollcommand=sb.set)
        tv.pack(padx=12, pady=2, fill='x')
        sb.pack(padx=12, fill='x')

        bf = ttk.Frame(self)
        bf.pack(padx=12, pady=(8, 10), fill='x')
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=self._ok).pack(side='right')

    def _ok(self):
        self.result = self.var.get()
        self.destroy()


# ── Column-mapping dialog ─────────────────────────────────────────────────────

class ColMapDialog(tk.Toplevel):
    _FIELDS = [
        ('internal_pn',  '公司料號 (10-Code) *',   ['material', '料號', 'internal', '10-code', 'pn', 'part']),
        ('vendor_pn',    'Vendor PN / MPN *',       ['vendor', 'mpn', 'component', 'mfr part', '廠商料號']),
        ('description',  '描述 Description',        ['desc', '描述', 'text', 'name', 'short']),
        ('manufacturer', '製造商 Manufacturer',     ['mfr', 'manufacturer', 'vendor name', '廠商']),
        ('qty',          '用量 Qty',                ['qty', 'quantity', '數量', 'amount']),
        ('reference',    '參考件號 Ref/Location',   ['ref', 'location', 'designator', '位置']),
    ]

    def __init__(self, parent, columns, preview_rows, initial_project=''):
        super().__init__(parent)
        self.title('欄位對應 Column Mapping')
        self.resizable(True, False)
        self.grab_set()
        self.result = None
        _none = '(略)'
        choices = [_none] + list(columns)

        ttk.Label(self, text='請對應每個欄位到 SAP BOM 的欄位名稱（* 至少擇一）',
                  font=FONT_UI).grid(row=0, column=0, columnspan=2,
                                     sticky='w', padx=12, pady=(10, 4))

        self._vars = {}
        for r, (key, label, hints) in enumerate(self._FIELDS, start=1):
            ttk.Label(self, text=label, font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=2)
            guess = next((c for c in columns
                          for h in hints if h.lower() in c.lower()), _none)
            var = tk.StringVar(value=guess)
            ttk.Combobox(self, textvariable=var, values=choices,
                         state='readonly', width=30).grid(
                row=r, column=1, padx=8, pady=2, sticky='w')
            self._vars[key] = var

        ttk.Separator(self, orient='horizontal').grid(
            row=len(self._FIELDS)+1, column=0, columnspan=2,
            sticky='ew', padx=8, pady=6)

        nr = len(self._FIELDS) + 2
        ttk.Label(self, text='專案名稱 *', font=FONT_UI).grid(
            row=nr, column=0, sticky='w', padx=12, pady=2)
        self.sv_proj = tk.StringVar(value=initial_project)
        ttk.Entry(self, textvariable=self.sv_proj, width=32).grid(
            row=nr, column=1, padx=8, pady=2, sticky='w')

        nr += 1
        ttk.Label(self, text='Preview (前 3 列)', font=FONT_UI).grid(
            row=nr, column=0, columnspan=2, sticky='w', padx=12, pady=(8, 2))
        nr += 1
        tv = ttk.Treeview(self, columns=list(columns), show='headings', height=3)
        for col in columns:
            tv.heading(col, text=col)
            tv.column(col, width=max(60, min(160, len(str(col)) * 9)))
        for prow in preview_rows[:3]:
            tv.insert('', 'end', values=[str(v) if v is not None else '' for v in prow])
        tv_sb = ttk.Scrollbar(self, orient='horizontal', command=tv.xview)
        tv.configure(xscrollcommand=tv_sb.set)
        tv.grid(row=nr,   column=0, columnspan=2, sticky='ew', padx=12, pady=2)
        tv_sb.grid(row=nr+1, column=0, columnspan=2, sticky='ew', padx=12)

        nr += 2
        bf = ttk.Frame(self)
        bf.grid(row=nr, column=0, columnspan=2, sticky='ew', padx=12, pady=(6, 10))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='Import', command=self._ok).pack(side='right')

    def _ok(self):
        _none = '(略)'
        mapping = {k: (v.get() if v.get() != _none else None)
                   for k, v in self._vars.items()}
        proj = self.sv_proj.get().strip()
        if not proj:
            messagebox.showwarning('Warning', '請輸入專案名稱', parent=self)
            return
        if not mapping.get('internal_pn') and not mapping.get('vendor_pn'):
            messagebox.showwarning('Warning',
                '至少需要對應「公司料號」或「Vendor PN」其中一個', parent=self)
            return
        self.result = {'mapping': mapping, 'project': proj}
        self.destroy()


# ── Rail dialog ──────────────────────────────────────────────────────────────

class RailDialog(tk.Toplevel):
    """Add / Edit power rail.  Currents entered in mA, stored in A internally."""

    def __init__(self, parent, rail_data=None):
        super().__init__(parent)
        self.title('Add Rail' if rail_data is None else 'Edit Rail')
        self.resizable(True, False)
        self.grab_set()
        self.result = None

        def _ma(key):
            """Convert stored A value → mA string for display."""
            if not rail_data:
                return ''
            v = rail_data.get(key, '') or ''
            try:
                return str(round(float(v) * 1000, 4)).rstrip('0').rstrip('.')
            except (ValueError, TypeError):
                return ''

        def _sv(key):
            return str(rail_data.get(key, '') or '') if rail_data else ''

        self._vars = {}
        p = dict(padx=6, pady=3)

        # Row 0: Rail Name | Power Group
        ttk.Label(self, text='Rail Name *', font=FONT_UI).grid(row=0, column=0, sticky='w', **p)
        self._vars['rail_name'] = tk.StringVar(value=_sv('rail_name'))
        ttk.Entry(self, textvariable=self._vars['rail_name'], width=22).grid(
            row=0, column=1, sticky='ew', **p)
        ttk.Label(self, text='Power Group', font=FONT_UI).grid(row=0, column=2, sticky='w', **p)
        self._vars['power_group'] = tk.StringVar(value=_sv('power_group'))
        ttk.Entry(self, textvariable=self._vars['power_group'], width=22).grid(
            row=0, column=3, sticky='ew', **p)

        # Row 1: Input Rail | (empty right side)
        ttk.Label(self, text='Input Rail', font=FONT_UI,
                  foreground='#1a5a9a').grid(row=1, column=0, sticky='w', **p)
        self._vars['input_rail'] = tk.StringVar(value=_sv('input_rail'))
        ttk.Entry(self, textvariable=self._vars['input_rail'], width=22).grid(
            row=1, column=1, sticky='ew', **p)
        ttk.Label(self, text='(e.g. VCC_0V9)', foreground='gray',
                  font=FONT_UI).grid(row=1, column=2, columnspan=2, sticky='w', **p)

        # Row 2: Description | Condition
        ttk.Label(self, text='Description', font=FONT_UI).grid(row=2, column=0, sticky='w', **p)
        self._vars['rail_description'] = tk.StringVar(value=_sv('rail_description'))
        ttk.Entry(self, textvariable=self._vars['rail_description'], width=22).grid(
            row=2, column=1, sticky='ew', **p)
        ttk.Label(self, text='Condition', font=FONT_UI).grid(row=2, column=2, sticky='w', **p)
        self._vars['condition'] = tk.StringVar(value=_sv('condition'))
        ttk.Entry(self, textvariable=self._vars['condition'], width=22).grid(
            row=2, column=3, sticky='ew', **p)

        # Row 3: Voltage | Tolerance
        ttk.Label(self, text='Voltage (V) *', font=FONT_UI).grid(row=3, column=0, sticky='w', **p)
        self._vars['voltage'] = tk.StringVar(value=_sv('voltage'))
        ttk.Entry(self, textvariable=self._vars['voltage'], width=22).grid(
            row=3, column=1, sticky='ew', **p)
        ttk.Label(self, text='Tolerance (%)', font=FONT_UI).grid(row=3, column=2, sticky='w', **p)
        self._vars['tolerance_percent'] = tk.StringVar(value=_sv('tolerance_percent'))
        ttk.Entry(self, textvariable=self._vars['tolerance_percent'], width=22).grid(
            row=3, column=3, sticky='ew', **p)

        # Row 4: Typ Current mA (optional) | Max Current mA (required)
        ttk.Label(self, text='Typ Current (mA)', font=FONT_UI).grid(row=4, column=0, sticky='w', **p)
        self._vars['typ_current_ma'] = tk.StringVar(value=_ma('typ_current_a'))
        ttk.Entry(self, textvariable=self._vars['typ_current_ma'], width=22).grid(
            row=4, column=1, sticky='ew', **p)
        ttk.Label(self, text='Max Current (mA) *', font=FONT_UI).grid(row=4, column=2, sticky='w', **p)
        self._vars['max_current_ma'] = tk.StringVar(value=_ma('max_current_a'))
        ttk.Entry(self, textvariable=self._vars['max_current_ma'], width=22).grid(
            row=4, column=3, sticky='ew', **p)

        # Row 5: Pin Names (full width)
        ttk.Label(self, text='Pin Names (comma-sep)', font=FONT_UI).grid(
            row=5, column=0, sticky='w', **p)
        self._vars['pin_names'] = tk.StringVar(value=_sv('pin_names'))
        ttk.Entry(self, textvariable=self._vars['pin_names']).grid(
            row=5, column=1, columnspan=3, sticky='ew', **p)

        # Row 6: separator + auto-power labels
        ttk.Separator(self, orient='horizontal').grid(
            row=6, column=0, columnspan=4, sticky='ew', padx=8, pady=4)
        ttk.Label(self, text='Typ Power (W)', font=FONT_UI).grid(row=7, column=0, sticky='w', **p)
        self._lbl_typ_p = ttk.Label(self, text='—', font=FONT_UI, foreground='#1a5a9a')
        self._lbl_typ_p.grid(row=7, column=1, sticky='w', **p)
        ttk.Label(self, text='Max Power (W)', font=FONT_UI).grid(row=7, column=2, sticky='w', **p)
        self._lbl_max_p = ttk.Label(self, text='—', font=FONT_UI, foreground='#1a5a9a')
        self._lbl_max_p.grid(row=7, column=3, sticky='w', **p)

        for key in ('voltage', 'typ_current_ma', 'max_current_ma'):
            self._vars[key].trace_add('write', lambda *_: self._update_power())
        self._update_power()

        # Row 8: buttons
        bf = ttk.Frame(self)
        bf.grid(row=8, column=0, columnspan=4, sticky='ew', padx=12, pady=(4, 10))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=self._ok).pack(side='right')

        self.columnconfigure(1, weight=1)
        self.columnconfigure(3, weight=1)

    def _update_power(self):
        try:
            v   = float(self._vars['voltage'].get())
            mx  = float(self._vars['max_current_ma'].get()) / 1000
            typ_s = self._vars['typ_current_ma'].get().strip()
            typ = float(typ_s) / 1000 if typ_s else mx
            self._lbl_typ_p.configure(text=f'{v * typ:.4f}')
            self._lbl_max_p.configure(text=f'{v * mx:.4f}')
        except ValueError:
            self._lbl_typ_p.configure(text='—')
            self._lbl_max_p.configure(text='—')

    def _ok(self):
        if not self._vars['rail_name'].get().strip():
            messagebox.showwarning('Warning', 'Rail Name 為必填', parent=self)
            return
        try:
            v  = float(self._vars['voltage'].get())
            mx = float(self._vars['max_current_ma'].get()) / 1000
        except ValueError:
            messagebox.showwarning('Warning', 'Voltage 與 Max Current 必須為數字', parent=self)
            return
        typ_s = self._vars['typ_current_ma'].get().strip()
        typ = float(typ_s) / 1000 if typ_s else mx

        self.result = {
            'rail_name':         self._vars['rail_name'].get().strip(),
            'power_group':       self._vars['power_group'].get().strip(),
            'input_rail':        self._vars['input_rail'].get().strip(),
            'rail_description':  self._vars['rail_description'].get().strip(),
            'voltage':           v,
            'tolerance_percent': float(self._vars['tolerance_percent'].get() or 0),
            'typ_current_a':     typ,
            'max_current_a':     mx,
            'condition':         self._vars['condition'].get().strip(),
            'pin_names':         self._vars['pin_names'].get().strip(),
        }
        self.destroy()


# ── Source dialog ─────────────────────────────────────────────────────────────

class SourceDialog(tk.Toplevel):
    _REG_TYPES = ['Load Switch', 'LDO', 'Switching', 'Short-Pad']

    def __init__(self, parent, rail_names=None, source_data=None):
        super().__init__(parent)
        self.title('Add Power Source' if source_data is None else 'Edit Power Source')
        self.resizable(True, False)
        self.grab_set()
        self.result = None
        rail_names = rail_names or []

        def _sv(key): return str(source_data.get(key, '') or '') if source_data else ''

        fields_simple = [
            ('source_rail',    'Source Rail'),
            ('input_voltage',  'Input Voltage (V)'),
            ('output_voltage', 'Output Voltage (V)'),
            ('remark',         'Remark'),
        ]
        self._vars = {}

        ttk.Label(self, text='Rail Name *', font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=12, pady=2)
        self._vars['rail_name'] = tk.StringVar(value=_sv('rail_name'))
        ttk.Combobox(self, textvariable=self._vars['rail_name'],
                     values=rail_names, width=30).grid(
            row=0, column=1, padx=8, pady=2, sticky='w')

        ttk.Label(self, text='Regulator Type', font=FONT_UI).grid(
            row=1, column=0, sticky='w', padx=12, pady=2)
        self._vars['regulator_type'] = tk.StringVar(value=_sv('regulator_type'))
        ttk.Combobox(self, textvariable=self._vars['regulator_type'],
                     values=self._REG_TYPES, state='readonly', width=30).grid(
            row=1, column=1, padx=8, pady=2, sticky='w')

        ttk.Label(self, text='Efficiency (0-1)', font=FONT_UI).grid(
            row=2, column=0, sticky='w', padx=12, pady=2)
        self._vars['efficiency'] = tk.StringVar(value=_sv('efficiency') or '1.0')
        ttk.Entry(self, textvariable=self._vars['efficiency'], width=32).grid(
            row=2, column=1, padx=8, pady=2, sticky='w')

        for r, (key, label) in enumerate(fields_simple, start=3):
            ttk.Label(self, text=label, font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=2)
            self._vars[key] = tk.StringVar(value=_sv(key))
            ttk.Entry(self, textvariable=self._vars[key], width=32).grid(
                row=r, column=1, padx=8, pady=2, sticky='w')

        bf = ttk.Frame(self)
        bf.grid(row=len(fields_simple)+3, column=0, columnspan=2,
                sticky='ew', padx=12, pady=(6, 10))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=self._ok).pack(side='right')

    def _ok(self):
        try:
            eff = float(self._vars['efficiency'].get() or 1.0)
        except ValueError:
            messagebox.showwarning('Warning', 'Efficiency 必須為數字 (0~1)', parent=self)
            return
        self.result = {
            'rail_name':      self._vars['rail_name'].get().strip(),
            'regulator_type': self._vars['regulator_type'].get().strip(),
            'efficiency':     eff,
            'source_rail':    self._vars['source_rail'].get().strip(),
            'input_voltage':  float(self._vars['input_voltage'].get() or 0),
            'output_voltage': float(self._vars['output_voltage'].get() or 0),
            'remark':         self._vars['remark'].get().strip(),
        }
        self.destroy()


# ── Qty column picker (for Power BOM import) ──────────────────────────────────

class _QtyColPicker(tk.Toplevel):
    _HINTS = ['qty', 'quantity', '數量', 'amount', 'count']

    def __init__(self, parent, columns, preview_rows):
        super().__init__(parent)
        self.title('選擇 Qty 欄位')
        self.resizable(True, False)
        self.grab_set()
        self.result = None

        ttk.Label(self, text='選擇含數量 (Qty) 的欄位：',
                  font=FONT_UI).pack(padx=12, pady=(10, 4), anchor='w')
        guess = next((c for c in columns
                      for h in self._HINTS if h.lower() in c.lower()), columns[0])
        self.var = tk.StringVar(value=guess)
        ttk.Combobox(self, textvariable=self.var, values=list(columns),
                     state='readonly', width=32).pack(padx=12, pady=4, anchor='w')

        tv = ttk.Treeview(self, columns=list(columns), show='headings', height=3)
        for col in columns:
            tv.heading(col, text=col)
            tv.column(col, width=max(60, min(160, len(str(col)) * 9)))
        for row in preview_rows[:3]:
            tv.insert('', 'end',
                      values=[str(v) if v is not None else '' for v in row])
        sb = ttk.Scrollbar(self, orient='horizontal', command=tv.xview)
        tv.configure(xscrollcommand=sb.set)
        tv.pack(padx=12, pady=2, fill='x')
        sb.pack(padx=12, fill='x')

        bf = ttk.Frame(self)
        bf.pack(padx=12, pady=(8, 10), fill='x')
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=lambda: self._ok()).pack(side='right')

    def _ok(self):
        self.result = self.var.get()
        self.destroy()


# ── Todo List dialog ─────────────────────────────────────────────────────────

class _TodoListDialog(tk.Toplevel):
    def __init__(self, parent, existing: dict = None):
        super().__init__(parent)
        self.title('Rename List' if existing else 'New List')
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        start_color = (existing or {}).get('color', _LIST_PAL[0])
        self._color = start_color if start_color in _LIST_PAL else _LIST_PAL[0]

        tk.Label(self, text='List Name:', font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=12, pady=(12, 4))
        self._sv = tk.StringVar(value=(existing or {}).get('name', ''))
        e = ttk.Entry(self, textvariable=self._sv, width=24)
        e.grid(row=0, column=1, padx=8, pady=(12, 4))
        e.focus_set()
        e.bind('<Return>', lambda ev: self._ok())

        tk.Label(self, text='Color:', font=FONT_UI).grid(
            row=1, column=0, sticky='w', padx=12, pady=4)
        pal_f = tk.Frame(self)
        pal_f.grid(row=1, column=1, sticky='w', padx=8, pady=4)
        self._cc = []
        for i, clr in enumerate(_LIST_PAL):
            c = tk.Canvas(pal_f, width=22, height=22, bd=0,
                          highlightthickness=0, cursor='hand2')
            c.grid(row=0, column=i, padx=2)
            c.bind('<Button-1>', lambda e, col=clr: self._pick(col))
            self._cc.append((c, clr))
        self._pick(self._color)

        bf = ttk.Frame(self)
        bf.grid(row=2, column=0, columnspan=2, sticky='ew', padx=12, pady=(6, 12))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='Save' if existing else 'Create',
                   command=self._ok).pack(side='right')

    def _pick(self, color):
        self._color = color
        for c, clr in self._cc:
            c.delete('all')
            if clr == color:
                c.create_oval(1, 1, 20, 20, fill=clr, outline='white', width=2)
            else:
                c.create_oval(1, 1, 20, 20, fill=clr, outline='')

    def _ok(self):
        name = self._sv.get().strip()
        if not name:
            messagebox.showwarning('Warning', 'Please enter a list name', parent=self)
            return
        self.result = (name, self._color)
        self.destroy()


# ── Todo Edit dialog ──────────────────────────────────────────────────────────

class _TodoEditDialog(tk.Toplevel):
    _PRIOS = [('None', ''), ('Low  !', 'low'),
              ('Medium  !!', 'medium'), ('High  !!!', 'high')]

    def __init__(self, parent, item: dict):
        super().__init__(parent)
        self.title('Edit Reminder')
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        p = dict(padx=12, pady=4)

        tk.Label(self, text='Title:', font=FONT_UI).grid(row=0, column=0, sticky='w', **p)
        self._title_sv = tk.StringVar(value=item.get('title', ''))
        ttk.Entry(self, textvariable=self._title_sv, width=36).grid(
            row=0, column=1, sticky='ew', **p)

        tk.Label(self, text='Notes:', font=FONT_UI).grid(row=1, column=0, sticky='nw', **p)
        self._notes = tk.Text(self, width=36, height=4, font=FONT_UI, relief='solid', bd=1)
        self._notes.insert('1.0', item.get('notes', '') or '')
        self._notes.grid(row=1, column=1, sticky='ew', padx=12, pady=4)

        tk.Label(self, text='Due Date:', font=FONT_UI).grid(row=2, column=0, sticky='w', **p)
        due_f = ttk.Frame(self)
        due_f.grid(row=2, column=1, sticky='w', **p)
        self._due_sv = tk.StringVar(value=item.get('due_date', '') or '')
        ttk.Entry(due_f, textvariable=self._due_sv, width=14).pack(side='left')
        tk.Label(due_f, text='YYYY-MM-DD', fg='gray', font=FONT_UI).pack(side='left', padx=6)

        tk.Label(self, text='Priority:', font=FONT_UI).grid(row=3, column=0, sticky='w', **p)
        cur_prio = item.get('priority', '') or ''
        disp_map = {v: l for l, v in self._PRIOS}
        self._prio_sv = tk.StringVar(value=disp_map.get(cur_prio, 'None'))
        ttk.Combobox(self, textvariable=self._prio_sv,
                     values=[l for l, _ in self._PRIOS],
                     state='readonly', width=14).grid(row=3, column=1, sticky='w', **p)

        bf = ttk.Frame(self)
        bf.grid(row=4, column=0, columnspan=2, sticky='ew', padx=12, pady=(8, 12))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='Save',   command=self._ok).pack(side='right')
        self.columnconfigure(1, weight=1)

    def _ok(self):
        title = self._title_sv.get().strip()
        if not title:
            messagebox.showwarning('Warning', 'Title cannot be empty', parent=self)
            return
        val_map = {l: v for l, v in self._PRIOS}
        due = self._due_sv.get().strip()
        if due:
            try:
                from datetime import date as _d
                _d.fromisoformat(due)
            except ValueError:
                messagebox.showwarning('Warning', 'Invalid date (use YYYY-MM-DD)', parent=self)
                return
        self.result = {
            'title':    title,
            'notes':    self._notes.get('1.0', 'end').strip(),
            'due_date': due,
            'priority': val_map.get(self._prio_sv.get(), ''),
        }
        self.destroy()


class _TodoSubItemDialog(tk.Toplevel):
    """Add/edit a single checklist sub-item (title + optional due date)."""

    def __init__(self, parent, subitem: dict = None):
        super().__init__(parent)
        self.title('Checklist Item')
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        subitem = subitem or {}
        p = dict(padx=12, pady=4)

        tk.Label(self, text='Title:', font=FONT_UI).grid(row=0, column=0, sticky='w', **p)
        self._title_sv = tk.StringVar(value=subitem.get('title', ''))
        e = ttk.Entry(self, textvariable=self._title_sv, width=32)
        e.grid(row=0, column=1, sticky='ew', **p)
        e.focus_set()
        e.bind('<Return>', lambda ev: self._ok())

        tk.Label(self, text='Due Date:', font=FONT_UI).grid(row=1, column=0, sticky='w', **p)
        due_f = ttk.Frame(self)
        due_f.grid(row=1, column=1, sticky='w', **p)
        self._due_sv = tk.StringVar(value=subitem.get('due_date', '') or '')
        ttk.Entry(due_f, textvariable=self._due_sv, width=14).pack(side='left')
        tk.Label(due_f, text='YYYY-MM-DD', fg='gray', font=FONT_UI).pack(side='left', padx=6)

        bf = ttk.Frame(self)
        bf.grid(row=2, column=0, columnspan=2, sticky='ew', padx=12, pady=(8, 12))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='Save',   command=self._ok).pack(side='right')
        self.columnconfigure(1, weight=1)

    def _ok(self):
        title = self._title_sv.get().strip()
        if not title:
            messagebox.showwarning('Warning', 'Title cannot be empty', parent=self)
            return
        due = self._due_sv.get().strip()
        if due:
            try:
                from datetime import date as _d
                _d.fromisoformat(due)
            except ValueError:
                messagebox.showwarning('Warning', 'Invalid date (use YYYY-MM-DD)', parent=self)
                return
        self.result = {'title': title, 'due_date': due}
        self.destroy()


class _TextPromptDialog(tk.Toplevel):
    """Generic single-line text prompt (folder name, rename, label...)."""

    def __init__(self, parent, title, prompt, initial=''):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        tk.Label(self, text=prompt, font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=12, pady=(12, 4))
        self.sv = tk.StringVar(value=initial)
        e = ttk.Entry(self, textvariable=self.sv, width=32)
        e.grid(row=1, column=0, sticky='ew', padx=12, pady=4)
        e.focus_set()
        e.select_range(0, 'end')
        e.bind('<Return>', lambda ev: self._ok())

        bf = ttk.Frame(self)
        bf.grid(row=2, column=0, sticky='e', padx=12, pady=(8, 12))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=self._ok).pack(side='right')
        self.columnconfigure(0, weight=1)

    def _ok(self):
        val = self.sv.get().strip()
        if not val:
            self.destroy()
            return
        self.result = val
        self.destroy()


# ── Smart File Browser ────────────────────────────────────────────────────────

def _load_app_cfg():
    try:
        with open(APP_CFG, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def _save_app_cfg(cfg):
    try:
        with open(APP_CFG, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass



class _AltPartPicker(tk.Toplevel):
    """Searchable picker over PartsDB — returns (vendor_pn, manufacturer)."""

    def __init__(self, parent, db, initial_query=''):
        super().__init__(parent)
        self.title('Browse Internal Parts DB')
        self.geometry('650x400')
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self._db = db
        self._build()
        if initial_query:
            self._sv_q.set(initial_query)
            self._search()
        else:
            self._load_popular()

    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        # Row 0: search bar
        ttk.Label(self, text='Search:', font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=(8, 4), pady=(8, 4))
        self._sv_q = tk.StringVar()
        ent = ttk.Entry(self, textvariable=self._sv_q, width=30)
        ent.grid(row=0, column=1, sticky='ew', padx=4, pady=(8, 4))
        ent.bind('<Return>', lambda e: self._search())
        ttk.Button(self, text='Search',
                   command=self._search).grid(row=0, column=2, padx=4, pady=(8, 4))
        ttk.Button(self, text='Show Popular',
                   command=self._load_popular).grid(row=0, column=3, padx=(0, 8), pady=(8, 4))

        # Row 1: treeview
        tv_f = tk.Frame(self)
        tv_f.grid(row=1, column=0, columnspan=4, sticky='nsew', padx=8, pady=2)
        tv_f.rowconfigure(0, weight=1)
        tv_f.columnconfigure(0, weight=1)
        cols = ('vendor_pn', 'description', 'manufacturer', 'internal_pn', 'proj_count')
        self._tv = ttk.Treeview(tv_f, columns=cols, show='headings', selectmode='browse')
        for cid, hdr, w in [('vendor_pn',    'Vendor PN',    140),
                             ('description',  'Description',  200),
                             ('manufacturer', 'Manufacturer', 120),
                             ('internal_pn',  'Internal PN',  100),
                             ('proj_count',   'Usage',         50)]:
            self._tv.heading(cid, text=hdr)
            self._tv.column(cid, width=w,
                            anchor='w' if cid != 'proj_count' else 'center')
        vsb = ttk.Scrollbar(tv_f, orient='vertical', command=self._tv.yview)
        self._tv.configure(yscrollcommand=vsb.set)
        self._tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        self._tv.bind('<Double-1>', lambda e: self._select())

        # Row 2: status label
        self._lbl = ttk.Label(self, text='', foreground='gray', font=FONT_UI)
        self._lbl.grid(row=2, column=0, columnspan=4, sticky='w', padx=8)

        # Row 3: action buttons
        bf = tk.Frame(self)
        bf.grid(row=3, column=0, columnspan=4, sticky='ew', padx=8, pady=(4, 8))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='Select', command=self._select).pack(side='right')

    def _load_popular(self):
        rows = self._db.popular_parts(keywords=None, limit=50)
        self._populate(rows)
        self._lbl.configure(text=f'{len(rows)} popular parts shown')

    def _search(self):
        q = self._sv_q.get().strip()
        if not q:
            self._load_popular()
            return
        rows = self._db.search_vendor_pn(q)
        if not rows:
            rows = self._db.popular_parts(keywords=q.split(), limit=50)
        self._populate(rows)
        self._lbl.configure(
            text=f'{len(rows)} result(s) for "{q}"' if rows else 'No results found')

    def _populate(self, rows):
        self._tv.delete(*self._tv.get_children())
        for r in rows:
            self._tv.insert('', 'end', values=(
                r['vendor_pn']    or '',
                r['description']  or '',
                r['manufacturer'] or '',
                r['internal_pn']  or '',
                r['proj_count']   or 0))

    def _select(self):
        sel = self._tv.focus() or (self._tv.selection() or (None,))[0]
        if not sel:
            return
        vals = self._tv.item(sel, 'values')
        if vals:
            self.result = (vals[0], vals[2], vals[1], vals[3])  # (vendor_pn, manufacturer, description, internal_pn)
            self.destroy()


class _SmartFileBrowser(tk.Toplevel):
    """File browser with pinned-folder bookmarks for quick navigation."""

    def __init__(self, parent, title='Select File(s)'):
        super().__init__(parent)
        self.title(title)
        self.geometry('820x520')
        self.minsize(600, 380)
        self.grab_set()
        self.result   = []
        self._cur     = None
        self._cfg     = _load_app_cfg()
        self._build()

    def _build(self):
        self.rowconfigure(1, weight=1)
        self.columnconfigure(1, weight=1)

        # ── Left: pinned folders ───────────────────────────────────────────
        left = tk.Frame(self, bg='#EFEFF4', width=190)
        left.grid(row=0, column=0, rowspan=3, sticky='nsew')
        left.grid_propagate(False)
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(left, text='BOOKMARKS', bg='#EFEFF4', fg='#8E8E93',
                 font=('Microsoft JhengHei UI', 8, 'bold')).grid(
            row=0, column=0, sticky='w', padx=10, pady=(10, 4))

        self._pin_lb = tk.Listbox(left, bd=0, highlightthickness=0,
                                   bg='#EFEFF4', font=FONT_UI,
                                   activestyle='dotbox', selectbackground='#D1D1D6')
        self._pin_lb.grid(row=1, column=0, sticky='nsew', padx=4)
        self._pin_lb.bind('<<ListboxSelect>>', self._on_pin_sel)

        bf = tk.Frame(left, bg='#EFEFF4')
        bf.grid(row=2, column=0, sticky='ew', padx=6, pady=6)
        ttk.Button(bf, text='+ Pin',  width=6, command=self._add_pin).pack(side='left')
        ttk.Button(bf, text='✕ Del',  width=6, command=self._del_pin).pack(side='left', padx=4)

        tk.Frame(self, bg='#D1D1D6', width=1).grid(row=0, column=0, rowspan=3, sticky='nse')

        # ── Right top: navigation bar ──────────────────────────────────────
        nav = tk.Frame(self, bg='white')
        nav.grid(row=0, column=1, sticky='ew')
        ttk.Button(nav, text='↑ Up', width=5, command=self._go_up).pack(
            side='left', padx=4, pady=4)
        self._path_var = tk.StringVar(value='← Select a bookmark  or  Pin a folder')
        tk.Label(nav, textvariable=self._path_var, bg='white', fg='#555',
                 font=FONT_UI, anchor='w').pack(side='left', fill='x', expand=True, padx=4)

        # ── Right middle: file treeview ────────────────────────────────────
        tree_f = tk.Frame(self)
        tree_f.grid(row=1, column=1, sticky='nsew', padx=(0, 4), pady=2)
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        cols = ('name', 'size', 'modified')
        self._tv = ttk.Treeview(tree_f, columns=cols, show='headings',
                                 selectmode='extended')
        self._tv.heading('name',     text='Name');     self._tv.column('name',     width=320, anchor='w')
        self._tv.heading('size',     text='Size');     self._tv.column('size',     width=80,  stretch=False, anchor='e')
        self._tv.heading('modified', text='Modified'); self._tv.column('modified', width=130, stretch=False)
        self._tv.tag_configure('dir',  font=(FONT_UI[0], FONT_UI[1], 'bold'))
        self._tv.tag_configure('file', foreground='#3A3A3A')

        vsb = ttk.Scrollbar(tree_f, orient='vertical', command=self._tv.yview)
        self._tv.configure(yscrollcommand=vsb.set)
        self._tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        self._tv.bind('<Double-1>',       self._dbl_click)
        self._tv.bind('<<TreeviewSelect>>', self._on_sel)

        # ── Right bottom: status + buttons ────────────────────────────────
        bot = tk.Frame(self)
        bot.grid(row=2, column=1, sticky='ew', padx=4, pady=(0, 6))
        self._sel_lbl = tk.Label(bot, text='0 file(s) selected',
                                  fg='#8E8E93', font=FONT_UI)
        self._sel_lbl.pack(side='left', padx=4)
        ttk.Button(bot, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bot, text='OK',     command=self._ok).pack(side='right')

        self._refresh_pins()

    def _refresh_pins(self):
        self._pin_lb.delete(0, 'end')
        for p in self._cfg.get('pinned_folders', []):
            self._pin_lb.insert('end', '  📁 ' + (os.path.basename(p) or p))

    def _on_pin_sel(self, _=None):
        sel = self._pin_lb.curselection()
        if not sel:
            return
        pins = self._cfg.get('pinned_folders', [])
        if sel[0] < len(pins):
            self._navigate(pins[sel[0]])

    def _add_pin(self):
        folder = filedialog.askdirectory(title='Select folder to bookmark', parent=self)
        if folder:
            self._cfg.setdefault('pinned_folders', [])
            if folder not in self._cfg['pinned_folders']:
                self._cfg['pinned_folders'].append(folder)
                _save_app_cfg(self._cfg)
            self._refresh_pins()

    def _del_pin(self):
        sel = self._pin_lb.curselection()
        if not sel:
            return
        pins = self._cfg.get('pinned_folders', [])
        if sel[0] < len(pins):
            pins.pop(sel[0])
            _save_app_cfg(self._cfg)
            self._refresh_pins()

    def _navigate(self, path):
        if not os.path.isdir(path):
            messagebox.showwarning('Not Found', f'Folder not found:\n{path}', parent=self)
            return
        self._cur = path
        self._path_var.set(path)
        self._tv.delete(*self._tv.get_children())
        from datetime import datetime as _dt
        try:
            entries = sorted(os.scandir(path),
                             key=lambda e: (not e.is_dir(), e.name.lower()))
        except PermissionError:
            return
        for e in entries:
            if e.name.startswith('.'):
                continue
            try:
                st = e.stat()
                mt = _dt.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M')
                if e.is_dir():
                    self._tv.insert('', 'end', iid=e.path,
                                    values=('📁  ' + e.name, '', mt), tags=('dir',))
                else:
                    sz = st.st_size
                    size_s = (f'{sz/1024:.0f} KB' if sz < 1048576
                              else f'{sz/1048576:.1f} MB')
                    self._tv.insert('', 'end', iid=e.path,
                                    values=('📄  ' + e.name, size_s, mt), tags=('file',))
            except Exception:
                pass

    def _dbl_click(self, event):
        iid = self._tv.identify_row(event.y)
        if iid and os.path.isdir(iid):
            self._navigate(iid)

    def _go_up(self):
        if self._cur:
            parent = os.path.dirname(self._cur)
            if parent != self._cur:
                self._navigate(parent)

    def _on_sel(self, _=None):
        n = sum(1 for iid in self._tv.selection() if os.path.isfile(iid))
        self._sel_lbl.configure(text=f'{n} file(s) selected')

    def _ok(self):
        self.result = [iid for iid in self._tv.selection() if os.path.isfile(iid)]
        self.destroy()

    @classmethod
    def ask_files(cls, parent, title='Select File(s)'):
        dlg = cls(parent, title)
        parent.wait_window(dlg)
        return dlg.result


# ── New Project Setup dialog ──────────────────────────────────────────────────

class _NewProjectDialog(tk.Toplevel):
    _FOLDERS = [
        ('01. BLOCK',        True),
        ('02. SCHEMATIC',    True),
        ('03. LAYOUT',       True),
        ('04. BOM',          True),
        ('05. SIMULATION',   False),
        ('06. DATASHEET',    True),
        ('07. REVIEW',       True),
        ('08. DOCUMENTS',    True),
        ('09. FIRMWARE',     False),
        ('10. TEST_REPORT',  False),
        ('11. PHOTO',        False),
    ]

    def __init__(self, parent):
        super().__init__(parent)
        self.title('New Project Setup')
        self.resizable(False, True)
        self.grab_set()
        self.result = None
        self._build()

    def _build(self):
        p = dict(padx=14, pady=4)

        # Project name
        tk.Label(self, text='Project Name:', font=FONT_UI).grid(
            row=0, column=0, sticky='w', **p)
        self._name = tk.StringVar()
        ttk.Entry(self, textvariable=self._name, width=36).grid(
            row=0, column=1, sticky='ew', padx=(4, 14), pady=4)

        # Location
        tk.Label(self, text='Location:', font=FONT_UI).grid(
            row=1, column=0, sticky='w', **p)
        loc_f = tk.Frame(self)
        loc_f.grid(row=1, column=1, sticky='ew', padx=(4, 14), pady=4)
        self._loc = tk.StringVar()
        ttk.Entry(loc_f, textvariable=self._loc, width=28).pack(side='left')
        ttk.Button(loc_f, text='Browse', width=7,
                   command=self._browse).pack(side='left', padx=4)

        ttk.Separator(self, orient='horizontal').grid(
            row=2, column=0, columnspan=2, sticky='ew', padx=10, pady=6)

        tk.Label(self, text='Folders to create:', font=FONT_UI).grid(
            row=3, column=0, columnspan=2, sticky='w', padx=14)

        # Checkboxes - two columns
        chk_f = tk.Frame(self)
        chk_f.grid(row=4, column=0, columnspan=2, sticky='ew', padx=14, pady=4)
        self._folder_vars = []
        for i, (name, default) in enumerate(self._FOLDERS):
            var = tk.BooleanVar(value=default)
            ttk.Checkbutton(chk_f, text=name, variable=var).grid(
                row=i // 2, column=i % 2, sticky='w', padx=(0, 16), pady=1)
            self._folder_vars.append((name, var))

        # Custom folder
        ttk.Separator(self, orient='horizontal').grid(
            row=5, column=0, columnspan=2, sticky='ew', padx=10, pady=6)
        cust_f = tk.Frame(self)
        cust_f.grid(row=6, column=0, columnspan=2, sticky='ew', padx=14, pady=(0, 4))
        tk.Label(cust_f, text='Add custom folder:', font=FONT_UI).pack(side='left')
        self._custom = tk.StringVar()
        ttk.Entry(cust_f, textvariable=self._custom, width=20).pack(side='left', padx=6)
        ttk.Button(cust_f, text='+ Add', command=self._add_custom).pack(side='left')
        self._custom_lbl = tk.Label(cust_f, text='', fg='#007AFF', font=FONT_UI)
        self._custom_lbl.pack(side='left', padx=6)

        # Options
        self._open_var = tk.BooleanVar(value=True)
        self._pin_var  = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text='Open in Windows Explorer after creation',
                        variable=self._open_var).grid(
            row=7, column=0, columnspan=2, sticky='w', padx=14, pady=2)
        ttk.Checkbutton(self, text='Pin to Smart File Browser bookmarks',
                        variable=self._pin_var).grid(
            row=8, column=0, columnspan=2, sticky='w', padx=14, pady=(2, 8))

        # Buttons
        bf = ttk.Frame(self)
        bf.grid(row=9, column=0, columnspan=2, sticky='ew', padx=14, pady=(4, 12))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='Create Project', command=self._create).pack(side='right')
        self.columnconfigure(1, weight=1)

    def _browse(self):
        path = filedialog.askdirectory(title='Project location', parent=self)
        if path:
            self._loc.set(path)

    def _add_custom(self):
        name = self._custom.get().strip()
        if not name:
            return
        self._folder_vars.append((name, tk.BooleanVar(value=True)))
        existing = self._custom_lbl.cget('text')
        self._custom_lbl.configure(
            text=(existing + ('  ' if existing else '') + name))
        self._custom.set('')

    def _create(self):
        name = self._name.get().strip()
        loc  = self._loc.get().strip()
        if not name:
            messagebox.showwarning('Warning', 'Enter a project name', parent=self)
            return
        if not loc or not os.path.isdir(loc):
            messagebox.showwarning('Warning', 'Select a valid location', parent=self)
            return

        root = os.path.join(loc, name)
        try:
            os.makedirs(root, exist_ok=True)
        except Exception as e:
            messagebox.showerror('Error', str(e), parent=self)
            return

        created = []
        for folder_name, var in self._folder_vars:
            if var.get():
                try:
                    os.makedirs(os.path.join(root, folder_name), exist_ok=True)
                    created.append(folder_name)
                except Exception:
                    pass

        if self._pin_var.get():
            cfg = _load_app_cfg()
            cfg.setdefault('pinned_folders', [])
            if root not in cfg['pinned_folders']:
                cfg['pinned_folders'].append(root)
            _save_app_cfg(cfg)

        messagebox.showinfo(
            'Project Created',
            f'{root}\n\nCreated {len(created)} folders:\n' +
            '\n'.join(f'  {f}' for f in created),
            parent=self)

        if self._open_var.get():
            os.startfile(root)

        self.result = root
        self.destroy()


# ── GUI ───────────────────────────────────────────────────────────────────────

# ═══════════════════════════════════════════════════════════════════════════════
# SAP BOM COMPARE (ported from bom_check.py, standalone tool) — used by the
# "SAP BOM 比對" tab. Reuses the existing OrcadItem class above (identical
# fields); everything else here is new: Orcad-vs-SAP comparison, config-
# driven column mapping, and the NC-aware location classification.
# ═══════════════════════════════════════════════════════════════════════════════

BOMCHK_CFG_PATH = os.path.join(_BASE, 'bom_check.inf')

BOMCHK_REVISION_HISTORY = """\
AVABD BOM CHECK  Revision History
══════════════════════════════════════════════════════════════════════════════

Rev0.1  2026-05-22
  - Initial release
  - Compare Orcad BOM (.BOM) vs SAP BOM (.txt) by 10-Code
  - Filter SAP items by Layer (level column)
  - Result Tab: By Code / By Location two sub-views
  - Config Tab: editable .inf settings

Rev0.2  2026-07-27
  - SAP BOM file can now also be a .xls export (in addition to .txt)

"""

BOMCHK_APPROVED_VENDOR_KEYWORDS = [
    'STMICRO', 'STMICROELECTRONICS',
    'VISHAY',
    'TAIWAN SEMICONDUCTOR',
    'DIODES INC',
    'PAN JIT', 'PANJIT',
    'NEXPERIA',
    'INFINEON',
    'ROHM',
    'ON SEMICONDUCTOR', 'ONSEMI',
    'TOSHIBA',
    'BROADCOM',
    'EVERLIGHT',
    'SKYWORKS',
    'TEXAS INSTRUMENTS',
    'NXP',
    'MONOLITHIC POWER',
    'QUALCOMM',
    'MEDIATEK',
    'REALTEK',
    'MARVELL',
    'NVIDIA',
    'MACRONIX',
    'MICRON',
    'SAMSUNG',
    'WINBOND',
    'INTEGRATED SILICON', 'ISSI',
    'MXIC',
    'WESTERN DIGITAL',
    'KIOXIA',
    'ATP',
    'KINGSTON',
    'LITE-ON',
    'GIGADEVICE', 'GITEK',
    'MICROCHIP',
    'FUTURE TECHNOLOGY',
    'QUECTEL',
    'FIBOCOM',
    'RINGFYRE',
    'MURATA',
    'AVX',
    'TDK',
    'TAIYO YUDEN',
    'YAGEO',
    'WALSIN',
    'CYNTEC',
    'KAMAYA',
    'KOA',
    'PANASONIC', 'MATSUSHITA',
    'THINKING',
    'NICHICON',
    'RUBYCON',
    'NIPPON CHEMI-CON',
    'KEMET',
    'TOKO',
    'DELTA',
    'PULSE',
    'UMEC',
    'BOTHHAND',
    'CHILISIN',
    'TAI-TECH',
    'BOURNS',
    'BUSSMANN',
    'LITTELFUSE',
    'AEM',
    'UPIH',
    'SEIKO',
    'SANYO',
]

class SapItem:
    __slots__ = ('item_no', 'level', 'code', 'part_name', 'alt',
                 'description', 'qty', 'unit', 'locations')

    def __init__(self, item_no, level, code, part_name, alt,
                 description, qty, unit, locations):
        self.item_no     = item_no
        self.level       = level
        self.code        = code
        self.part_name   = part_name
        self.alt         = alt
        self.description = description
        self.qty         = qty
        self.unit        = unit
        self.locations   = locations  # list[str]


class ConfigManager:
    DEFAULTS = {
        'SAP BOM Analyze Config': {
            'Sap DIP':   '55080',
            'Sap SMD':   '39413',
            'Sap PCB':   '29763',
            'Sap SW IC': '26413',
            'Sap Bom Head Start String':       'D ITEM  STD',
            'Sap Bom Code Head String':        'PART NO',
            'Sap Bom Description Head String': 'DESCRIPTION',
            'Sap Bom ALT Head String':         'ALT',
            'Sap Bom Qty Head String':         '    QPA',
            'Sap Bom Location Head String':    '  SORT STRING',
            'Sap Bom Purchasing Head String':  'Purchasing Name',
            'Sap Bom Compare Level':           '2,3',
        },
        'Orcad BOM Analyze Config': {
            'Orcad Identify String':         '________',
            'Orcad Ignore Analyze Value':    'NC_;_NC',
            'Orcad Ignore Analyze Location': 'TP;H',
            'Orcad Ignore 10-Code':          'New10Code',
            'Orcad Bom Code Column':         '2',
            'Orcad Bom Value Column':        '3',
            'Orcad Bom Vendor PN Column':    '4',
            'Orcad Bom Vendor Column':       '5',
            'Orcad Bom Description Column':  '6',
            'Orcad Bom Footprint Column':    '7',
            'Orcad Bom Qty Column':          '8',
            'Orcad Bom Location Column':     '9',
            'Orcad Bom NC Column':           '10',
        },
        'File Path Config': {
            'Orcad Bom Files Path': '',
            'SAP Bom Files Path':   'C:\\TEMP',
            'Report Files Path':    '',
        },
    }

    def __init__(self, path='bom_check.inf'):
        self.path = path
        self.data = {s: dict(kv) for s, kv in self.DEFAULTS.items()}
        if os.path.exists(path):
            self._load()

    def _load(self):
        section = ''
        try:
            with open(self.path, encoding='utf-8', errors='replace') as f:
                for raw in f:
                    line = raw.rstrip()
                    if not line or line.startswith('#'):
                        continue
                    if '[' in line and ']' in line:
                        section = line[line.index('[') + 1:line.rindex(']')].strip()
                        if section not in self.data:
                            self.data[section] = {}
                    elif '=' in line and section:
                        k, _, v = line.partition('=')
                        self.data[section][k.strip()] = v.strip().strip('"')
        except Exception:
            pass

    def get(self, section, key, default=''):
        return self.data.get(section, {}).get(key, default)

    def to_text(self):
        out = []
        for sec, kvs in self.data.items():
            out.append(f'\n[ {sec} ]\n')
            for k, v in kvs.items():
                out.append(f'{k} = {v}')
            out.append('')
        return '\n'.join(out)

    def save(self):
        try:
            with open(self.path, 'w', encoding='utf-8') as f:
                f.write(self.to_text())
            return ''
        except Exception as e:
            return str(e)


class OrcadParser:
    def __init__(self, cfg: ConfigManager):
        s = 'Orcad BOM Analyze Config'
        self.code_col      = int(cfg.get(s, 'Orcad Bom Code Column',        '2'))  - 1
        self.val_col       = int(cfg.get(s, 'Orcad Bom Value Column',       '3'))  - 1
        self.vendor_pn_col = int(cfg.get(s, 'Orcad Bom Vendor PN Column',   '4'))  - 1
        self.vendor_col    = int(cfg.get(s, 'Orcad Bom Vendor Column',      '5'))  - 1
        self.desc_col      = int(cfg.get(s, 'Orcad Bom Description Column', '6'))  - 1
        self.fp_col        = int(cfg.get(s, 'Orcad Bom Footprint Column',   '7'))  - 1
        self.qty_col       = int(cfg.get(s, 'Orcad Bom Qty Column',         '8'))  - 1
        self.loc_col       = int(cfg.get(s, 'Orcad Bom Location Column',    '9'))  - 1
        self.nc_col        = int(cfg.get(s, 'Orcad Bom NC Column',          '10')) - 1

        ig_val = cfg.get(s, 'Orcad Ignore Analyze Value', 'NC_;_NC')
        self.ig_vals = [v.strip() for v in ig_val.split(';') if v.strip()]

        ig_loc = cfg.get(s, 'Orcad Ignore Analyze Location', 'TP;H;G')
        self.ig_locs = [v.strip().upper() for v in ig_loc.split(';') if v.strip()]

        self.ignore_code = cfg.get(s, 'Orcad Ignore 10-Code', 'New10Code')

    def parse(self, path):
        text  = self._read(path)
        lines = text.splitlines()

        # Find separator line (____...)
        start = -1
        for i, ln in enumerate(lines):
            if ln.strip().startswith('____'):
                start = i + 1
                break
        if start < 0:
            return []

        items, cur = [], None
        cur_is_nc  = False
        self.nc_locs = set()

        for ln in lines[start:]:
            if not ln.strip():
                continue
            parts = ln.split('\t')

            # Defensive realignment: some OrCAD exports drop an entirely-empty
            # property cell (observed: Vendor PN) instead of writing a blank
            # tab, shifting every later column left by one for that row. An
            # item line should reach the NC column; a continuation line only
            # ever reaches Location. Detect a one-field shortfall against
            # whichever applies and re-insert a blank at the Vendor PN slot
            # so Description/Footprint/Qty/Location/NC land back correctly.
            is_item_line = parts[0].strip() != ''
            expected = (self.nc_col + 1) if is_item_line else (self.loc_col + 1)
            if len(parts) == expected - 1:
                parts = parts[:self.vendor_pn_col] + [''] + parts[self.vendor_pn_col:]

            # Continuation line: location wraps to next row (starts with tab)
            is_continuation = ln.startswith('\t') or (
                parts[0].strip() == '' and len(parts) > self.loc_col)
            if is_continuation:
                if cur and len(parts) > self.loc_col:
                    extra = parts[self.loc_col].strip()
                    if extra:
                        new_locs = [l for l in self._locs(extra)
                                    if not self._ig_loc(l)]
                        cur.locations += new_locs
                        if cur_is_nc:
                            self.nc_locs.update(new_locs)
                continue

            if len(parts) <= self.code_col:
                continue

            code      = parts[self.code_col     ].strip() if len(parts) > self.code_col      else ''
            value     = parts[self.val_col      ].strip() if len(parts) > self.val_col       else ''
            vendor_pn = parts[self.vendor_pn_col].strip() if len(parts) > self.vendor_pn_col else ''
            vendor    = parts[self.vendor_col   ].strip() if len(parts) > self.vendor_col    else ''
            desc      = parts[self.desc_col     ].strip() if len(parts) > self.desc_col      else ''
            fp        = parts[self.fp_col       ].strip() if len(parts) > self.fp_col        else ''
            loc_s     = parts[self.loc_col      ].strip() if len(parts) > self.loc_col       else ''
            nc        = parts[self.nc_col       ].strip().upper() if len(parts) > self.nc_col else ''

            qty = 0
            try:
                qty = int(parts[self.qty_col].strip()) if len(parts) > self.qty_col else 0
            except (ValueError, IndexError):
                qty = 0

            # Handle malformed lines where location is embedded in Item column
            # e.g. "U5189" instead of item=189 + location=U5
            if not loc_s and len(parts) <= self.qty_col:
                m = re.match(r'^([A-Za-z]+\d+)\d{3}$', parts[0].strip())
                if m:
                    loc_s = m.group(1)
                    if not qty:
                        qty = 1

            if code == self.ignore_code:
                cur, cur_is_nc = None, False
                continue

            locs = [l for l in self._locs(loc_s) if not self._ig_loc(l)]

            if nc:
                self.nc_locs.update(locs)
                # Keep a throwaway holder so wrapped continuation lines still
                # get their locations folded into nc_locs (see is_continuation
                # above) — it's never added to `items`, just tracked via cur.
                cur = OrcadItem(code, value, desc, fp, qty, locs, vendor, vendor_pn)
                cur_is_nc = True
                continue
            if self._ig_value(value):
                cur, cur_is_nc = None, False
                continue

            cur = OrcadItem(code, value, desc, fp, qty, locs, vendor, vendor_pn)
            cur_is_nc = False
            items.append(cur)

        return items

    # ── helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _read(path):
        for enc in ('utf-8', 'big5', 'cp950', 'latin-1'):
            try:
                with open(path, encoding=enc) as f:
                    return f.read()
            except (UnicodeDecodeError, LookupError):
                continue
        with open(path, encoding='utf-8', errors='replace') as f:
            return f.read()

    @staticmethod
    def _locs(s):
        return [x.strip() for x in s.split(',') if x.strip()]

    def _ig_value(self, val):
        for ig in self.ig_vals:
            if ig.endswith('_') and val.upper().startswith(ig[:-1].upper()):
                return True
            if ig.startswith('_') and val.upper().endswith(ig[1:].upper()):
                return True
            if val.upper() == ig.upper():
                return True
        return False

    def _ig_loc(self, loc):
        u = loc.upper()
        return any(u.startswith(p) for p in self.ig_locs)


class SapParser:
    def __init__(self, cfg: ConfigManager):
        s = 'SAP BOM Analyze Config'
        self.top_pfx = [cfg.get(s, k, d) for k, d in [
            ('Sap DIP',   '55080'),
            ('Sap SMD',   '39413'),
            ('Sap PCB',   '29763'),
            ('Sap SW IC', '26413'),
        ]]
        self.head_start = cfg.get(s, 'Sap Bom Head Start String',       'D ITEM  STD')
        self.code_hdr   = cfg.get(s, 'Sap Bom Code Head String',        'PART NO')
        self.desc_hdr   = cfg.get(s, 'Sap Bom Description Head String', 'DESCRIPTION')
        self.alt_hdr    = cfg.get(s, 'Sap Bom ALT Head String',         'ALT')
        self.qty_hdr    = cfg.get(s, 'Sap Bom Qty Head String',         '    QPA')
        self.loc_hdr    = cfg.get(s, 'Sap Bom Location Head String',    '  SORT STRING')
        self.purch_hdr  = cfg.get(s, 'Sap Bom Purchasing Head String',  'Purchasing Name')

    def parse(self, path):
        """Returns (top_items: list[SapItem], components: list[SapItem])"""
        if path.lower().endswith(('.xls', '.xlsx')):
            return self._parse_excel(path)
        return self._parse_text(path)

    def _parse_text(self, path):
        text  = OrcadParser._read(path)
        lines = text.splitlines()

        # Find header line containing head_start
        hdr, hdr_i = None, -1
        for i, ln in enumerate(lines):
            if self.head_start in ln:
                hdr, hdr_i = ln, i
                break
        if hdr is None:
            return [], []

        # Determine column offsets from header line positions
        c_code = hdr.find(self.code_hdr)
        c_desc = hdr.find(self.desc_hdr)
        c_alt  = hdr.find(self.alt_hdr)
        c_qty  = hdr.find(self.qty_hdr)
        c_loc  = hdr.find(self.loc_hdr)
        c_purch = hdr.find(self.purch_hdr)

        top_items, components = [], []
        cur      = None
        loc_buf  = ''   # raw location string buffer — concat before split

        def _flush_loc():
            if cur is not None and loc_buf:
                cur.locations = self._locs(loc_buf)

        for ln in lines[hdr_i + 1:]:
            raw = ln.rstrip('\r\n')
            if not raw.strip():
                continue

            item_str = raw[:7].strip()

            if item_str.isdigit():
                _flush_loc()          # commit previous item's locations
                loc_buf = ''

                item_no  = int(item_str)
                lvl_str  = raw[7:13].strip()
                level    = int(lvl_str) if lvl_str.isdigit() else 0

                span   = self._slice(raw, c_code, c_desc)
                tokens = span.split()
                code      = tokens[0] if tokens else ''
                part_name = ' '.join(tokens[1:])

                desc    = self._slice(raw, c_desc, c_alt).strip()
                alt     = self._slice(raw, c_alt,  c_qty).strip()

                qty_raw = self._slice(raw, c_qty, c_loc).strip()
                m       = re.search(r'[\d.]+', qty_raw)
                qty     = float(m.group()) if m else 0.0

                if c_loc >= 0 and len(raw) > c_loc:
                    loc_buf = self._slice(raw, c_loc, c_purch).rstrip()

                si  = SapItem(item_no, level, code, part_name, alt,
                               desc, qty, 'PCE', [])
                cur = si

                if code and any(code.startswith(p) for p in self.top_pfx):
                    top_items.append(si)
                elif code:
                    components.append(si)

            elif cur is not None and not raw[:7].strip():
                # Continuation: SAP wraps at fixed width — concat raw, split later
                if c_loc >= 0 and len(raw) > c_loc:
                    loc_buf += self._slice(raw, c_loc, c_purch).rstrip()

        _flush_loc()   # commit last item

        return top_items, components

    def _parse_excel(self, path):
        """Parse a legacy .xls SAP BOM export (cell-based, no fixed-width slicing)."""
        if not HAS_XLRD:
            raise RuntimeError('xlrd 未安裝\n請執行：pip install xlrd')

        wb = xlrd.open_workbook(path)
        sheet = None
        for name in ('RawData', 'Format'):
            if name in wb.sheet_names():
                sheet = wb.sheet_by_name(name)
                break
        if sheet is None:
            sheet = wb.sheet_by_index(0)

        code_hdr = self.code_hdr.strip()
        desc_hdr = self.desc_hdr.strip()
        alt_hdr  = self.alt_hdr.strip()
        qty_hdr  = self.qty_hdr.strip()
        loc_hdr  = self.loc_hdr.strip()

        hdr_i, col = -1, {}
        for r in range(min(5, sheet.nrows)):
            vals = [str(v).strip() for v in sheet.row_values(r)]
            if code_hdr in vals and loc_hdr in vals:
                hdr_i = r
                col = {
                    'item':  vals.index('ITEM'),
                    'level': vals.index('L'),
                    'code':  vals.index(code_hdr),
                    'desc':  vals.index(desc_hdr),
                    'alt':   vals.index(alt_hdr),
                    'qty':   vals.index(qty_hdr),
                    'loc':   vals.index(loc_hdr),
                }
                break
        if hdr_i < 0:
            return [], []

        top_items, components = [], []
        cur     = None
        loc_buf = ''

        def _flush_loc():
            if cur is not None and loc_buf:
                cur.locations = self._locs(loc_buf)

        for r in range(hdr_i + 1, sheet.nrows):
            item_cell = sheet.cell(r, col['item'])

            if item_cell.ctype == xlrd.XL_CELL_NUMBER and item_cell.value != '':
                _flush_loc()
                loc_buf = ''

                item_no  = int(item_cell.value)
                lvl_raw  = str(sheet.cell(r, col['level']).value).strip()
                level    = int(lvl_raw) if lvl_raw.isdigit() else 0
                code     = str(sheet.cell(r, col['code']).value).strip()
                desc     = str(sheet.cell(r, col['desc']).value).strip()
                alt      = str(sheet.cell(r, col['alt']).value).strip()

                qty_val  = sheet.cell(r, col['qty']).value
                qty      = float(qty_val) if qty_val not in ('', None) else 0.0

                loc_buf  = str(sheet.cell(r, col['loc']).value).strip()

                si  = SapItem(item_no, level, code, '', alt, desc, qty, 'PCE', [])
                cur = si

                if code and any(code.startswith(p) for p in self.top_pfx):
                    top_items.append(si)
                elif code:
                    components.append(si)

            elif cur is not None:
                # Continuation row: blank ITEM, SORT STRING wraps — concat, split later
                extra = str(sheet.cell(r, col['loc']).value).strip()
                if extra:
                    loc_buf += extra

        _flush_loc()   # commit last item

        return top_items, components

    @staticmethod
    def _slice(line, start, end):
        if start < 0:
            return ''
        end_pos = end if (end > start) else len(line)
        return line[start:end_pos] if len(line) > start else ''

    @staticmethod
    def _locs(s):
        return [x.strip() for x in s.split(',') if x.strip()]


class BomCompare:
    @staticmethod
    def _nat_key(s):
        return [int(t) if t.isdigit() else t.lower()
                for t in re.split(r'(\d+)', s)]

    @staticmethod
    def compare(orcad_items, sap_items):
        sap_codes   = {i.code for i in sap_items   if i.code}
        orcad_codes = {i.code for i in orcad_items if i.code}
        orcad_only  = [i for i in orcad_items if i.code and i.code not in sap_codes]
        sap_only    = [i for i in sap_items   if i.code and i.code not in orcad_codes]
        return orcad_only, sap_only

    @staticmethod
    def _header(orcad_file, sap_file, title):
        ts   = datetime.now().strftime('%Y-%m-%d %H:%M')
        rows = [f'===== {title}  {ts} =====']
        if orcad_file:
            rows.append(f'Orcad : {os.path.basename(orcad_file)}')
        if sap_file:
            rows.append(f'SAP   : {os.path.basename(sap_file)}')
        rows.append('')
        return rows

    @staticmethod
    def format_by_code(orcad_only, sap_only, orcad_file='', sap_file=''):
        rows = BomCompare._header(orcad_file, sap_file, 'BOM Compare (依 Code)')

        qty_hdr = "Q'ty"

        o_qty = sum(i.qty for i in orcad_only)
        rows.append(f'▶ Orcad Only：共 {len(orcad_only)} 筆十碼，{int(o_qty)} 個零件不在 SAP Bom 中')
        rows.append(f'{"Code":<15} {"Value":<20} {qty_hdr:>5}  Location')
        rows.append('─' * 78)
        for item in orcad_only:
            lstr = ','.join(item.locations)
            rows.append(f'{item.code:<15} {item.value[:19]:<20} {item.qty:>5}  {lstr}')
        rows.append('')

        s_qty = sum(int(i.qty) for i in sap_only)
        rows.append(f'▶ SAP Only：共 {len(sap_only)} 筆十碼，{s_qty} 個零件不在 Orcad Bom 中')
        rows.append(f'{"Code":<15} {"Description":<40} {qty_hdr:>5}  Location')
        rows.append('─' * 78)
        for item in sap_only:
            lstr = ','.join(item.locations)
            rows.append(f'{item.code:<15} {item.description[:39]:<40} {int(item.qty):>5}  {lstr}')

        return '\n'.join(rows)

    @staticmethod
    def format_by_location(orcad_only, sap_only, orcad_file='', sap_file=''):
        rows = BomCompare._header(orcad_file, sap_file, 'BOM Compare (依 Location)')
        nat  = BomCompare._nat_key

        # Expand each item into per-location rows
        o_rows = sorted(
            [(loc, item) for item in orcad_only for loc in (item.locations or ['(無)'])],
            key=lambda x: nat(x[0]))
        s_rows = sorted(
            [(loc, item) for item in sap_only   for loc in (item.locations or ['(無)'])],
            key=lambda x: nat(x[0]))

        rows.append(f'▶ Orcad Only：共 {len(o_rows)} 個位置不在 SAP Bom 中')
        rows.append(f'{"Location":<12} {"Code":<15} {"Value":<20} Description')
        rows.append('─' * 78)
        for loc, item in o_rows:
            rows.append(f'{loc:<12} {item.code:<15} {item.value[:19]:<20} {item.description[:30]}')
        rows.append('')

        qty_hdr = "Q'ty"

        rows.append(f'▶ SAP Only：共 {len(s_rows)} 個位置不在 Orcad Bom 中')
        rows.append(f'{"Location":<12} {"Code":<15} {"Description":<40} {qty_hdr:>5}')
        rows.append('─' * 78)
        for loc, item in s_rows:
            rows.append(f'{loc:<12} {item.code:<15} {item.description[:39]:<40} {int(item.qty):>5}')

        return '\n'.join(rows)

    @staticmethod
    def loc_table_data(orcad_items, sap_items, nc_locs=None, ig_loc_fn=None):
        """Returns list of (loc, sap_code, orcad_code, match: bool)."""
        o_map = {}
        for item in orcad_items:
            for loc in item.locations:
                o_map[loc] = item
        s_map = {}
        for item in sap_items:
            for loc in item.locations:
                if ig_loc_fn and ig_loc_fn(loc):
                    continue  # skip TP/H etc. from SAP side
                s_map[loc] = item

        all_locs = sorted(set(o_map) | set(s_map), key=BomCompare._nat_key)
        result = []
        for loc in all_locs:
            o_item = o_map.get(loc)
            s_item = s_map.get(loc)
            # SAP-only location that is NC in Orcad → expected, skip
            if not o_item and nc_locs and loc in nc_locs:
                continue
            o_code = o_item.code if o_item else ''
            s_code = s_item.code if s_item else ''
            match  = bool(o_code) and bool(s_code) and o_code == s_code
            result.append((loc, s_code, o_code, match))
        return result

    @staticmethod
    def loc_note_data(orcad_items, sap_items, nc_locs=None, ig_loc_fn=None):
        """Returns list of (loc, sap_code, orcad_code, note), location-first.

        Rule: an NC-flagged Orcad location means "do not populate". SAP must
        not have anything there — if it does, that entry is stale and
        flagged 'SAP位置移除' (Orcad side shown blank, since nothing should
        be populated). If SAP also has nothing there, both sides already
        agree and it's dropped — nothing to report.

        For a normal (non-NC, actively populated) Orcad location, a mismatch
        against SAP is classified as a brand-new part+location, an existing
        SAP part number just missing this location, or a straight swap.
        """
        o_map = {}
        for item in orcad_items:
            for loc in item.locations:
                o_map[loc] = item
        s_map = {}
        for item in sap_items:
            for loc in item.locations:
                if ig_loc_fn and ig_loc_fn(loc):
                    continue  # skip TP/H etc. from SAP side
                s_map[loc] = item
        sap_codes = {item.code for item in sap_items if item.code}
        nc_locs   = nc_locs or set()

        all_locs = sorted(set(o_map) | set(s_map), key=BomCompare._nat_key)
        result = []
        for loc in all_locs:
            if loc in nc_locs:
                s_item = s_map.get(loc)
                s_code = s_item.code if s_item else ''
                if s_code:
                    result.append((loc, s_code, '', 'SAP位置移除'))
                continue  # NC + nothing in SAP -> both agree, nothing to report

            o_item = o_map.get(loc)
            s_item = s_map.get(loc)
            o_code = o_item.code if o_item else ''
            s_code = s_item.code if s_item else ''

            if o_code and s_code and o_code == s_code:
                continue  # matched, nothing to report

            if s_code and not o_code:
                note = 'SAP位置移除'
            elif o_code and not s_code:
                note = 'SAP 既有料號' if o_code in sap_codes else '新增 Part number & 位置'
            else:
                note = '換料'
            result.append((loc, s_code, o_code, note))
        return result

    @staticmethod
    def format_loc_note(rows, orcad_file='', sap_file=''):
        out = BomCompare._header(orcad_file, sap_file, 'BOM Compare (Location 分類)')
        out.append(f'▶ 共 {len(rows)} 筆位置差異')
        out.append(f'{"Location":<12} {"SAP PN":<16} {"Orcad PN":<16} Note')
        out.append('─' * 78)
        for loc, s_code, o_code, note in rows:
            out.append(f'{loc:<12} {s_code:<16} {o_code:<16} {note}')
        return '\n'.join(out)



class MouserApp:
    TITLE = 'Mouser BOM Search  Rev0.2'

    def __init__(self, root: tk.Tk):
        self.root      = root
        self.client    = MouserClient()
        self.dk_client = DigiKeyClient()
        self.db        = PartsDB(DB_PATH)
        self._all_items: list[OrcadItem] = []
        self._valid:     list[OrcadItem] = []
        self._bom_path    = ''
        self._batch_rows  = []
        self._searching   = False
        self._dk_rows     = []
        self._dk_searching = False
        self._lib_projects: list = []
        self.pwr_db        = PowerDB(POWER_DB_PATH)
        self._pwr_part_id  = None
        self._pwr_rail_id  = None
        self._pwr_part_sel: dict[int, dict] = {}  # {part_id: {'checked': bool, 'qty': int}}

        self.todo_db          = TodoDB(TODO_DB_PATH)
        self._todo_list_id    = None
        self._todo_list_col   = '#007AFF'
        self._todo_alert_ids  = set()   # items already alerted this session
        self._todo_cal_year   = 0
        self._todo_cal_month  = 0
        self._todo_cal_woff   = 0       # week offset from current week
        self._proj_cur_path   = None    # right-panel current folder
        self.risk_db          = RiskDB(RISK_DB_PATH)
        self._risk_scan_id    = None
        self._risk_scanning   = False
        self._risk_items_buf  = []      # [(item_id, vpn, qty), ...]
        self._bom_list_data:     list[tuple] = []
        self._bom_no_code_items: list[OrcadItem] = []
        self._note_npp_path: str = ''
        self._exp_bom_file:  str  = ''
        self._exp_exp_file:  str  = ''
        self._exp_design:    str  = ''
        self._exp_headers:   list = []
        self._exp_rows:      list = []
        self._exp_items:     list = []
        self._exp_changes:   list = []
        self._exp_parser          = ExpParser()
        self._diff_file_a:       str  = ''
        self._diff_file_b:       str  = ''
        self._diff_results:      list = []
        self._diff_code_results: list = []
        self.bomchk_cfg     = ConfigManager(BOMCHK_CFG_PATH)
        self.bomchk_orcad_p = OrcadParser(self.bomchk_cfg)
        self.bomchk_sap_p   = SapParser(self.bomchk_cfg)
        self.bomchk_cmp     = BomCompare()
        self.bomchk_orcad_file = ''
        self.bomchk_sap_file   = ''
        self.bomchk_orcad_items: list[OrcadItem] = []
        self.bomchk_sap_top:   list = []
        self.bomchk_sap_comps: list = []
        self._bomchk_last_orcad_only:        list = []
        self._bomchk_last_sap_only:          list = []
        self._bomchk_last_tbl_data:          list = []
        self._bomchk_last_note_tbl_data:     list = []
        self._bomchk_last_orcad_check_items: list = []
        self._bomchk_last_vendor_items:      list = []
        self._bomchk_last_vendor_keywords:   list = []

        root.title(self.TITLE)
        root.geometry('1060x720')
        root.minsize(860, 540)
        self._build_ui()

    # ── UI Build ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True, padx=4, pady=4)
        self._main_nb = nb
        self.tab_batch  = ttk.Frame(nb)
        self.tab_single = ttk.Frame(nb)
        self.tab_dk     = ttk.Frame(nb)
        self.tab_lib    = ttk.Frame(nb)
        nb.add(self.tab_batch,  text='Mouser 查價')
        nb.add(self.tab_single, text='Single Part Search')
        nb.add(self.tab_dk,     text='DigiKey 查價')
        nb.add(self.tab_lib,    text='料號庫')
        self.tab_pwr  = ttk.Frame(nb)
        self.tab_todo = ttk.Frame(nb)
        self.tab_proj = ttk.Frame(nb)
        self.tab_risk = ttk.Frame(nb)
        self.tab_bom  = ttk.Frame(nb)
        self.tab_note = ttk.Frame(nb)
        self.tab_exp  = ttk.Frame(nb)
        self.tab_diff = ttk.Frame(nb)
        nb.add(self.tab_pwr,  text='Power DB Builder')
        nb.add(self.tab_todo, text='Reminders')
        nb.add(self.tab_proj, text='Projects')
        nb.add(self.tab_risk, text='BOM Risk Scanner')
        nb.add(self.tab_bom,  text='備料清單')
        nb.add(self.tab_note, text='Note')
        nb.add(self.tab_exp,  text='EXP 料源替換')
        nb.add(self.tab_diff, text='BOM Diff')
        self.tab_bomchk = ttk.Frame(nb)
        nb.add(self.tab_bomchk, text='SAP BOM 比對')
        self._build_batch_tab()
        self._build_single_tab()
        self._build_digikey_tab()
        self._build_library_tab()
        self._build_power_tab()
        self._build_todo_tab()
        self._build_proj_tab()
        self._build_risk_tab()
        self._build_bom_tab()
        self._build_note_tab()
        self._build_exp_tab()
        self._build_diff_tab()
        self._build_sapbom_tab()

    # ── Tab 1: Batch ──────────────────────────────────────────────────────────

    def _build_batch_tab(self):
        f = self.tab_batch
        f.rowconfigure(2, weight=1)
        f.columnconfigure(0, weight=1)
        p = dict(padx=6, pady=3)

        # Row 0: file picker
        r0 = ttk.Frame(f)
        r0.grid(row=0, column=0, sticky='ew', **p)
        r0.columnconfigure(1, weight=1)
        ttk.Button(r0, text='Open .BOM File', command=self._open_bom,
                   width=16).grid(row=0, column=0, sticky='w')
        self.sv_bom = tk.StringVar()
        ttk.Entry(r0, textvariable=self.sv_bom,
                  state='readonly').grid(row=0, column=1, sticky='ew', padx=6)

        # Row 1: action bar
        r1 = ttk.Frame(f)
        r1.grid(row=1, column=0, sticky='ew', **p)
        ttk.Label(r1, text='Filter:').pack(side='left')
        self.sv_filter = tk.StringVar()
        ttk.Entry(r1, textvariable=self.sv_filter, width=18).pack(side='left', padx=4)
        ttk.Button(r1, text='Apply', command=self._apply_filter).pack(side='left')
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=8)
        self.btn_sel = ttk.Button(r1, text='Search Selected',
                                  command=self._search_selected)
        self.btn_sel.pack(side='left')
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Label(r1, text='Build Qty:').pack(side='left')
        self.sv_build_qty = tk.IntVar(value=1)
        ttk.Spinbox(r1, from_=1, to=99999, textvariable=self.sv_build_qty,
                    width=8).pack(side='left', padx=4)
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=8)
        self.btn_batch = ttk.Button(r1, text='批量查全部 (算總價)',
                                    command=self._start_batch)
        self.btn_batch.pack(side='left')
        self.lbl_status = ttk.Label(r1, text='', foreground='gray', font=FONT_UI)
        self.lbl_status.pack(side='left', padx=10)

        # Row 2: PanedWindow
        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=2, column=0, sticky='nsew', padx=4, pady=4)

        # Left pane: BOM list
        frm_l = ttk.Frame(pw)
        frm_l.rowconfigure(1, weight=1)
        frm_l.columnconfigure(0, weight=1)
        self.lbl_count = ttk.Label(frm_l, text='', font=FONT_UI)
        self.lbl_count.grid(row=0, column=0, columnspan=2, sticky='w', padx=4, pady=2)
        self.lb = tk.Listbox(frm_l, font=FONT_TX, activestyle='dotbox',
                             selectmode='single', width=44)
        sb = ttk.Scrollbar(frm_l, orient='vertical', command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.grid(row=1, column=0, sticky='nsew')
        sb.grid(row=1, column=1, sticky='ns')
        self.lb.bind('<<ListboxSelect>>', self._on_select)
        self.lb.bind('<Double-1>', lambda e: self._search_selected())

        # Right pane
        frm_r = ttk.Frame(pw)
        frm_r.rowconfigure(1, weight=1)
        frm_r.columnconfigure(0, weight=1)
        self.txt_detail = scrolledtext.ScrolledText(
            frm_r, font=FONT_TX, height=6, bg='#f0f8ff', wrap='none')
        self.txt_detail.grid(row=0, column=0, sticky='ew', padx=4, pady=(4, 2))
        self.txt_result = scrolledtext.ScrolledText(
            frm_r, font=FONT_TX, bg='#f8f8f8', wrap='none')
        self.txt_result.grid(row=1, column=0, sticky='nsew', padx=4, pady=2)
        self.txt_result.tag_configure('no_stock',  background='#FFA500', foreground='#000000')
        self.txt_result.tag_configure('not_found', foreground='#999999')
        self.txt_result.tag_configure('error_row', foreground='#cc0000')
        self.txt_result.tag_configure('hdr',       foreground='#0055aa')
        self.txt_result.tag_configure('separator', foreground='#aaaaaa')
        self.txt_result.tag_configure('tbl_hdr',   foreground='#555555')
        r_btn = ttk.Frame(frm_r)
        r_btn.grid(row=2, column=0, sticky='ew', padx=4, pady=4)
        ttk.Button(r_btn, text='Clear',
                   command=self._clear_result).pack(side='left')
        self.btn_save = ttk.Button(r_btn, text='Save Results as File ...',
                                   command=self._save_result)
        self.btn_save.pack(side='left', padx=6)
        self.btn_excel = ttk.Button(r_btn, text='Save as Excel ...',
                                    command=self._save_excel)
        self.btn_excel.pack(side='left', padx=6)

        pw.add(frm_l, weight=1)
        pw.add(frm_r, weight=3)

    # ── Tab 2: Single ─────────────────────────────────────────────────────────

    def _build_single_tab(self):
        f = self.tab_single
        f.rowconfigure(2, weight=1)
        f.columnconfigure(0, weight=1)
        p = dict(padx=6, pady=3)

        # Row 0: mode selector
        r0 = ttk.Frame(f)
        r0.grid(row=0, column=0, sticky='w', **p)
        ttk.Label(r0, text='Search mode:').pack(side='left')
        self.sv_mode = tk.StringVar(value='part')
        ttk.Radiobutton(r0, text='By MPN (exact)',
                        variable=self.sv_mode, value='part').pack(side='left', padx=6)
        ttk.Radiobutton(r0, text='By Keyword',
                        variable=self.sv_mode, value='kw').pack(side='left')

        # Row 1: search bar (one entry → searches both)
        r1 = ttk.Frame(f)
        r1.grid(row=1, column=0, sticky='ew', **p)
        r1.columnconfigure(1, weight=1)
        ttk.Label(r1, text='Search:').grid(row=0, column=0, sticky='w')
        self.sv_term = tk.StringVar()
        ent = ttk.Entry(r1, textvariable=self.sv_term)
        ent.grid(row=0, column=1, sticky='ew', padx=6)
        ent.bind('<Return>', lambda _: self._do_single())
        self.btn_single = ttk.Button(r1, text='Search Mouser + DigiKey',
                                     command=self._do_single)
        self.btn_single.grid(row=0, column=2, sticky='w')

        # Row 2: side-by-side result panels
        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=2, column=0, sticky='nsew', padx=4, pady=4)

        def _result_panel(parent, title, title_fg, bg):
            frm = ttk.Frame(parent)
            frm.rowconfigure(1, weight=1)
            frm.columnconfigure(0, weight=1)
            ttk.Label(frm, text=title, foreground=title_fg,
                      font=(FONT_UI[0], FONT_UI[1], 'bold')).grid(
                row=0, column=0, sticky='w', padx=4, pady=(4, 0))
            txt = scrolledtext.ScrolledText(frm, font=FONT_TX, bg=bg, wrap='none')
            txt.grid(row=1, column=0, sticky='nsew', padx=4, pady=(2, 4))
            txt.tag_configure('no_stock',  background='#FFA500', foreground='#000000')
            txt.tag_configure('hdr',       foreground=title_fg)
            txt.tag_configure('separator', foreground='#aaaaaa')
            return frm, txt

        frm_m,  self.txt_single_m  = _result_panel(pw, 'Mouser',  '#0055cc', '#f8f8f8')
        frm_dk, self.txt_single_dk = _result_panel(pw, 'DigiKey', '#cc0000', '#fff8f8')
        pw.add(frm_m,  weight=1)
        pw.add(frm_dk, weight=1)

        # Alias for any existing code that references txt_single
        self.txt_single = self.txt_single_m

        # Row 3: status bar
        r3 = ttk.Frame(f)
        r3.grid(row=3, column=0, sticky='ew', padx=6, pady=4)
        self.lbl_single_st = ttk.Label(r3, text='', foreground='gray', font=FONT_UI)
        self.lbl_single_st.pack(side='left')
        ttk.Button(r3, text='Clear', command=self._single_clear).pack(side='right')

    # ── BOM open / populate ───────────────────────────────────────────────────

    def _open_bom(self):
        path = filedialog.askopenfilename(
            title='Select Orcad BOM',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if not path:
            return
        try:
            all_items = BomParser.parse(path)
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return
        self._bom_path  = path
        self._all_items = all_items
        self._valid     = [i for i in all_items if BomParser.valid_code(i.code)]
        self.sv_bom.set(path)
        self.sv_filter.set('')
        self._populate(self._valid)
        self.sv_dk_bom.set(path)
        self.sv_dk_filter.set('')
        self._populate_dk(self._valid)
        skipped = len(all_items) - len(self._valid)
        self.lbl_status.configure(
            text=f'Loaded  {len(all_items)} total  |  {len(self._valid)} valid  |  {skipped} skipped (incomplete code)',
            foreground='gray')

    def _apply_filter(self):
        kw = self.sv_filter.get().strip().upper()
        src = self._valid
        if kw:
            src = [i for i in src if kw in (i.vendor_pn + i.value + i.code + i.description).upper()]
        self._populate(src)

    def _populate(self, items):
        self.lb.delete(0, 'end')
        for idx, item in enumerate(items):
            mpn = item.vendor_pn or item.value
            self.lb.insert('end', f'{idx+1:>4}  {mpn:<30} {item.code}')
        self.lbl_count.configure(
            text=f'{len(items)} items  (valid code only)')

    def _on_select(self, _=None):
        sel = self.lb.curselection()
        if not sel:
            return
        kw  = self.sv_filter.get().strip().upper()
        src = ([i for i in self._valid
                if kw in (i.vendor_pn + i.value + i.code + i.description).upper()]
               if kw else self._valid)
        if sel[0] >= len(src):
            return
        item = src[sel[0]]
        mpn  = item.vendor_pn or item.value
        lstr = ', '.join(item.locations[:8])
        if len(item.locations) > 8:
            lstr += f' ... (+{len(item.locations)-8})'
        txt = (f'VENDOR_PN  : {mpn}\n'
               f'Value      : {item.value}\n'
               f'10-Code    : {item.code}\n'
               f'Desc       : {item.description}\n'
               f'Vendor     : {item.vendor}\n'
               f'Qty / Loc  : {item.qty}  {lstr}')
        self.txt_detail.delete('1.0', 'end')
        self.txt_detail.insert('end', txt)

    # ── Single search ─────────────────────────────────────────────────────────

    def _search_selected(self):
        sel = self.lb.curselection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇料件')
            return
        kw  = self.sv_filter.get().strip().upper()
        src = ([i for i in self._valid
                if kw in (i.vendor_pn + i.value + i.code + i.description).upper()]
               if kw else self._valid)
        if sel[0] >= len(src):
            return
        mpn = src[sel[0]].vendor_pn or src[sel[0]].value
        self._run_search(mpn, self.txt_result, self.btn_sel, self.lbl_status)

    def _single_clear(self):
        self.txt_single_m.delete('1.0', 'end')
        self.txt_single_dk.delete('1.0', 'end')
        self.lbl_single_st.configure(text='')

    def _do_single(self):
        term = self.sv_term.get().strip()
        if not term:
            return
        mode = self.sv_mode.get()
        self.btn_single.configure(state='disabled')
        self.lbl_single_st.configure(
            text='Searching Mouser + DigiKey ...', foreground='gray')
        self.txt_single_m.delete('1.0', 'end')
        self.txt_single_dk.delete('1.0', 'end')
        self._single_pending  = 2
        self._single_m_count  = 0
        self._single_dk_count = 0
        threading.Thread(target=self._single_mouser_worker,
                         args=(term, mode), daemon=True).start()
        threading.Thread(target=self._single_dk_worker,
                         args=(term, mode), daemon=True).start()

    def _single_one_done(self):
        self._single_pending -= 1
        if self._single_pending <= 0:
            self.btn_single.configure(state='normal')
            self.lbl_single_st.configure(
                text=(f'Mouser: {self._single_m_count} result(s)  |  '
                      f'DigiKey: {self._single_dk_count} result(s)'),
                foreground='gray')

    # ── Single: Mouser worker ─────────────────────────────────────────────────

    def _single_mouser_worker(self, term, mode):
        try:
            parts = (self.client.search_by_keyword(term)
                     if mode == 'kw' else self.client.search_by_part(term))
            self.root.after(0, self._single_mouser_done, parts, term)
        except MouserAPIError as e:
            self.root.after(0, self._single_mouser_err, str(e))
        except Exception as e:
            self.root.after(0, self._single_mouser_err, repr(e))

    def _single_mouser_done(self, parts, term):
        self._single_m_count = len(parts)
        w = self.txt_single_m
        if parts:
            w.insert('end', f'── {term}  ({len(parts)} result(s)) ──\n')
            for p in parts:
                self._insert_detail(w, p, term)
            w.see('end')
        else:
            w.insert('end', f'No results for: {term}\n')
        self._single_one_done()

    def _single_mouser_err(self, msg):
        self.txt_single_m.insert('end', f'Error: {msg}\n')
        self._single_one_done()

    # ── Single: DigiKey worker ────────────────────────────────────────────────

    def _single_dk_worker(self, term, mode):
        try:
            products = (self.dk_client.search_by_keyword(term)
                        if mode == 'kw' else self.dk_client.search_by_mpn(term))
            self.root.after(0, self._single_dk_done, products, term)
        except DigiKeyError as e:
            self.root.after(0, self._single_dk_err, str(e))
        except Exception as e:
            self.root.after(0, self._single_dk_err, repr(e))

    def _single_dk_done(self, products, term):
        self._single_dk_count = len(products)
        w = self.txt_single_dk
        if products:
            w.insert('end', f'── {term}  ({len(products)} result(s)) ──\n')
            for p in products:
                self._dk_insert_detail(w, p, term)
            w.see('end')
        else:
            w.insert('end', f'No results for: {term}\n')
        self._single_one_done()

    def _single_dk_err(self, msg):
        self.txt_single_dk.insert('end', f'Error: {msg}\n')
        self._single_one_done()

    def _run_search(self, term, widget, btn, status, mode='part'):
        btn.configure(state='disabled')
        status.configure(text='Searching ...', foreground='gray')

        def worker():
            try:
                parts = (self.client.search_by_keyword(term)
                         if mode == 'kw' else self.client.search_by_part(term))
                self.root.after(0, done, parts)
            except MouserAPIError as e:
                self.root.after(0, err, str(e))
            except Exception as e:
                self.root.after(0, err, repr(e))

        def done(parts):
            btn.configure(state='normal')
            if not parts:
                status.configure(text=f'No results: {term}', foreground='orange')
                return
            status.configure(text=f'{len(parts)} result(s)  [{term}]', foreground='gray')
            widget.insert('end', f'── {term}  ({len(parts)} result(s)) ──\n')
            for p in parts:
                self._insert_detail(widget, p, term)
            widget.see('end')

        def err(msg):
            btn.configure(state='normal')
            status.configure(text=f'Error: {msg}', foreground='red')

        threading.Thread(target=worker, daemon=True).start()

    def _insert_detail(self, w, part, search_term=''):
        text  = MouserClient.fmt_detail(part)
        avail = part.get('Availability', '')
        mpn   = part.get('ManufacturerPartNumber', '')
        for line in text.splitlines(keepends=True):
            start = w.index('end - 1c linestart')
            w.insert('end', line)
            end = w.index('end - 1c')
            s = line.strip()
            if s.startswith('┌') or s.startswith('└'):
                w.tag_add('separator', start, end)
            elif s.startswith('MPN') and search_term.upper() in mpn.upper():
                w.tag_add('hdr', start, end)
            elif s.startswith('Availability'):
                val = avail.strip()
                if not val or val == '0' or 'Non-Stock' in val:
                    w.tag_add('no_stock', start, end)

    # ── Batch query ───────────────────────────────────────────────────────────

    def _start_batch(self):
        if not self._valid:
            messagebox.showwarning('Warning', '請先載入 BOM 檔案')
            return
        if self._searching:
            messagebox.showwarning('Warning', '查詢進行中，請稍候')
            return

        try:
            build_qty = max(1, int(self.sv_build_qty.get()))
        except Exception:
            build_qty = 1
        self.sv_build_qty.set(build_qty)

        if not messagebox.askyesno('Confirm',
                f'將查詢 {len(self._valid)} 筆有效料件\n'
                f'生產數量 = {build_qty} 片\n\n是否開始？'):
            return

        self._searching  = True
        self._batch_rows = []
        self.btn_batch.configure(state='disabled')
        self.btn_sel.configure(state='disabled')
        self.txt_result.delete('1.0', 'end')

        ts  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        hdr = (f'Mouser 批量查價  {ts}  (Build Qty = {build_qty})\n'
               f'BOM: {os.path.basename(self._bom_path)}\n'
               f'Valid items: {len(self._valid)}  (code 不完整者已略過)\n'
               f'{"─"*126}\n'
               f'{"No.":<5} {"10-Code":<15} {"VENDOR_PN":<30} {"Manufacturer":<22} '
               f'{"BOM":>5}  {"TotalQty":>9}  {"Avail":>9}  '
               f'{"Best Pr":>9}  {"Unit Pr":>9}  {"Line Total":>12}  Cur\n'
               f'{"─"*126}\n')
        self.txt_result.insert('end', hdr)
        self.txt_result.tag_add('tbl_hdr', '1.0', 'end')

        threading.Thread(
            target=self._batch_worker,
            args=(list(self._valid), build_qty),
            daemon=True).start()

    def _batch_worker(self, items, build_qty=1):
        found = 0
        not_found = 0
        for i, item in enumerate(items):
            self.root.after(0, self._update_progress, i + 1, len(items))
            mpn       = item.vendor_pn or item.value
            total_qty = item.qty * build_qty
            try:
                parts = self.client.search_by_part(mpn)
                if parts:
                    p              = parts[0]
                    unit_price, cur = MouserClient.price_at(p, total_qty)
                    price_best, _   = MouserClient.price_best(p)
                    avail           = p.get('Availability', '')
                    avail_num       = ''.join(c for c in avail.split()[0] if c.isdigit() or c == ',') if avail else ''
                    mfr             = p.get('Manufacturer', '')
                    line_total      = total_qty * MouserClient._parse_price(unit_price)
                    found          += 1
                    warn = not unit_price or avail_num in ('', '0')
                    row = (i + 1, item.code, mpn, mfr,
                           item.qty, total_qty, avail_num,
                           price_best, unit_price, line_total, cur,
                           'warn' if warn else 'ok')
                    self._batch_rows.append(row)
                    self.root.after(0, self._append_batch_row, row)
                else:
                    not_found += 1
                    row = (i + 1, item.code, mpn, '-- Not Found --',
                           item.qty, total_qty, '', '', '', 0.0, '', 'notfound')
                    self._batch_rows.append(row)
                    self.root.after(0, self._append_batch_row, row)
            except Exception as e:
                not_found += 1
                row = (i + 1, item.code, mpn, f'-- Error: {e} --',
                       item.qty, total_qty, '', '', '', 0.0, '', 'error')
                self._batch_rows.append(row)
                self.root.after(0, self._append_batch_row, row)
            time.sleep(0.5)

        grand_total = sum(r[9] for r in self._batch_rows if r[-1] in ('ok', 'warn'))
        self.root.after(0, self._batch_done, found, not_found, len(items), grand_total)

    def _update_progress(self, current, total):
        self.lbl_status.configure(
            text=f'Searching {current}/{total} ...', foreground='gray')

    def _append_batch_row(self, row):
        no, code, mpn, mfr, bom_qty, total_qty, avail, price_best, unit_price, line_total, cur, row_type = row
        lt_str = f'${line_total:,.3f}' if line_total else ''
        line = (f'{no:<5} {code:<15} {mpn:<30} {mfr:<22} '
                f'{bom_qty:>5}  {total_qty:>9}  {avail:>9}  '
                f'{price_best:>9}  {unit_price:>9}  {lt_str:>12}  {cur}\n')
        start = self.txt_result.index('end - 1c linestart')
        self.txt_result.insert('end', line)
        end = self.txt_result.index('end - 1c')
        if row_type == 'warn':
            self.txt_result.tag_add('no_stock', start, end)
        elif row_type == 'notfound':
            self.txt_result.tag_add('not_found', start, end)
        elif row_type == 'error':
            self.txt_result.tag_add('error_row', start, end)
        self.txt_result.see('end')

    def _batch_done(self, found, not_found, total, grand_total=0.0):
        self._searching = False
        self.btn_batch.configure(state='normal')
        self.btn_sel.configure(state='normal')
        summary = (f'{"─"*126}\n'
                   f'完成  查詢 {total} 筆  |  找到 {found} 筆  |  '
                   f'未找到/略過 {not_found} 筆\n'
                   f'Grand Total = ${grand_total:,.2f} USD\n')
        self.txt_result.insert('end', summary)
        self.txt_result.see('end')
        self.lbl_status.configure(
            text=f'Done  {found}/{total} found  |  Total = ${grand_total:,.2f}',
            foreground='gray')

    # ── Result area ───────────────────────────────────────────────────────────

    def _clear_result(self):
        self.txt_result.delete('1.0', 'end')
        self._batch_rows = []

    def _save_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        if not self._batch_rows:
            messagebox.showinfo('Info', '沒有資料可匯出')
            return

        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = f'Mouser_Batch_{ts}.xlsx'
        path = filedialog.asksaveasfilename(
            initialfile=name,
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Mouser BOM Price'

        # ── styles ────────────────────────────────────────────────────────────
        hdr_fill  = PatternFill('solid', fgColor='4472C4')
        hdr_font  = Font(bold=True, color='FFFFFF')
        total_fill = PatternFill('solid', fgColor='FFE699')
        total_font = Font(bold=True)
        fill_warn  = PatternFill('solid', fgColor='FFA500')
        fill_nf    = PatternFill('solid', fgColor='DDDDDD')
        fill_err   = PatternFill('solid', fgColor='FFB3B3')
        center     = Alignment(horizontal='center')
        right      = Alignment(horizontal='right')
        thin       = Side(style='thin', color='AAAAAA')
        border     = Border(bottom=thin)

        # ── header row ────────────────────────────────────────────────────────
        headers = ['No.', '10-Code', 'VENDOR_PN', 'Manufacturer',
                   'BOM Qty', 'Total Qty', 'Avail', 'Best Price',
                   'Unit Price', 'Line Total', 'Currency', 'Status']
        col_widths = [5, 16, 32, 24, 8, 10, 12, 11, 11, 13, 8, 10]
        ws.append(headers)
        for cell, w in zip(ws[1], col_widths):
            cell.fill      = hdr_fill
            cell.font      = hdr_font
            cell.alignment = center
            ws.column_dimensions[cell.column_letter].width = w

        # ── data rows ─────────────────────────────────────────────────────────
        status_label = {'ok': 'Found', 'warn': 'No Stock',
                        'notfound': 'Not Found', 'error': 'Error'}
        for row in self._batch_rows:
            no, code, mpn, mfr, bom_qty, total_qty, avail, \
                price_best, unit_price, line_total, cur, row_type = row
            ws.append([no, code, mpn, mfr, bom_qty, total_qty, avail,
                       price_best, unit_price,
                       line_total if line_total else None,
                       cur, status_label.get(row_type, row_type)])
            xr = ws.max_row
            fill = {'warn': fill_warn, 'notfound': fill_nf,
                    'error': fill_err}.get(row_type)
            for cell in ws[xr]:
                if fill:
                    cell.fill = fill
            ws.cell(xr, 10).number_format = '$#,##0.000'
            ws.cell(xr, 5).alignment  = center
            ws.cell(xr, 6).alignment  = center
            ws.cell(xr, 10).alignment = right

        # ── grand total row ───────────────────────────────────────────────────
        grand = sum(r[9] for r in self._batch_rows if r[-1] in ('ok', 'warn'))
        ws.append(['', '', '', 'Grand Total', '', '', '', '', '', grand, 'USD', ''])
        tr = ws.max_row
        for cell in ws[tr]:
            cell.fill   = total_fill
            cell.font   = total_font
            cell.border = border
        ws.cell(tr, 10).number_format = '$#,##0.00'
        ws.cell(tr, 10).alignment     = right

        # ── BOM info in row above header ──────────────────────────────────────
        ws.insert_rows(1)
        ws['A1'] = (f'BOM: {os.path.basename(self._bom_path)}  |  '
                    f'Build Qty: {self.sv_build_qty.get()}  |  '
                    f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        ws['A1'].font = Font(italic=True, color='555555')
        ws.merge_cells('A1:L1')
        ws.freeze_panes = 'A3'

        wb.save(path)
        messagebox.showinfo('Done', f'已儲存至\n{path}')

    def _save_result(self):
        text = self.txt_result.get('1.0', 'end').strip()
        if not text:
            messagebox.showinfo('Info', '沒有內容可儲存')
            return
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = f'Mouser_Batch_{ts}.txt'
        path = filedialog.asksaveasfilename(
            initialfile=name,
            filetypes=[('Text', '*.txt'), ('All', '*.*')])
        if path:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write(text)
            messagebox.showinfo('Done', f'已儲存至\n{path}')


    # ── Tab 3: DigiKey ────────────────────────────────────────────────────────

    def _build_digikey_tab(self):
        f = self.tab_dk
        f.rowconfigure(2, weight=1)
        f.columnconfigure(0, weight=1)
        p = dict(padx=6, pady=3)

        # Row 0: file picker
        r0 = ttk.Frame(f)
        r0.grid(row=0, column=0, sticky='ew', **p)
        r0.columnconfigure(1, weight=1)
        ttk.Button(r0, text='Open .BOM File', command=self._open_bom,
                   width=16).grid(row=0, column=0, sticky='w')
        self.sv_dk_bom = tk.StringVar()
        ttk.Entry(r0, textvariable=self.sv_dk_bom,
                  state='readonly').grid(row=0, column=1, sticky='ew', padx=6)

        # Row 1: action bar
        r1 = ttk.Frame(f)
        r1.grid(row=1, column=0, sticky='ew', **p)
        ttk.Label(r1, text='Filter:').pack(side='left')
        self.sv_dk_filter = tk.StringVar()
        ttk.Entry(r1, textvariable=self.sv_dk_filter, width=18).pack(side='left', padx=4)
        ttk.Button(r1, text='Apply', command=self._apply_filter_dk).pack(side='left')
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=8)
        self.btn_dk_sel = ttk.Button(r1, text='Search Selected',
                                     command=self._search_selected_dk)
        self.btn_dk_sel.pack(side='left')
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Label(r1, text='Build Qty:').pack(side='left')
        self.sv_dk_qty = tk.IntVar(value=1)
        ttk.Spinbox(r1, from_=1, to=99999, textvariable=self.sv_dk_qty,
                    width=8).pack(side='left', padx=4)
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=8)
        self.btn_dk_batch = ttk.Button(r1, text='批量查全部 (算總價)',
                                       command=self._dk_start_batch)
        self.btn_dk_batch.pack(side='left')
        self.lbl_dk_status = ttk.Label(r1, text='', foreground='gray', font=FONT_UI)
        self.lbl_dk_status.pack(side='left', padx=10)

        # Row 2: PanedWindow
        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=2, column=0, sticky='nsew', padx=4, pady=4)

        # Left pane: BOM list
        frm_l = ttk.Frame(pw)
        frm_l.rowconfigure(1, weight=1)
        frm_l.columnconfigure(0, weight=1)
        self.lbl_dk_count = ttk.Label(frm_l, text='', font=FONT_UI)
        self.lbl_dk_count.grid(row=0, column=0, columnspan=2, sticky='w', padx=4, pady=2)
        self.lb_dk = tk.Listbox(frm_l, font=FONT_TX, activestyle='dotbox',
                                selectmode='single', width=44)
        sb_dk = ttk.Scrollbar(frm_l, orient='vertical', command=self.lb_dk.yview)
        self.lb_dk.configure(yscrollcommand=sb_dk.set)
        self.lb_dk.grid(row=1, column=0, sticky='nsew')
        sb_dk.grid(row=1, column=1, sticky='ns')
        self.lb_dk.bind('<<ListboxSelect>>', self._on_select_dk)
        self.lb_dk.bind('<Double-1>', lambda e: self._search_selected_dk())

        # Right pane
        frm_r = ttk.Frame(pw)
        frm_r.rowconfigure(1, weight=1)
        frm_r.columnconfigure(0, weight=1)
        self.txt_dk_detail = scrolledtext.ScrolledText(
            frm_r, font=FONT_TX, height=6, bg='#f0f8ff', wrap='none')
        self.txt_dk_detail.grid(row=0, column=0, sticky='ew', padx=4, pady=(4, 2))
        self.txt_dk_result = scrolledtext.ScrolledText(
            frm_r, font=FONT_TX, bg='#f8f8f8', wrap='none')
        self.txt_dk_result.grid(row=1, column=0, sticky='nsew', padx=4, pady=2)
        self.txt_dk_result.tag_configure('no_stock',  background='#FFA500', foreground='#000000')
        self.txt_dk_result.tag_configure('not_found', foreground='#999999')
        self.txt_dk_result.tag_configure('error_row', foreground='#cc0000')
        self.txt_dk_result.tag_configure('hdr',       foreground='#0055aa')
        self.txt_dk_result.tag_configure('separator', foreground='#aaaaaa')
        self.txt_dk_result.tag_configure('tbl_hdr',   foreground='#555555')
        r_btn = ttk.Frame(frm_r)
        r_btn.grid(row=2, column=0, sticky='ew', padx=4, pady=4)
        ttk.Button(r_btn, text='Clear',
                   command=self._dk_clear).pack(side='left')
        ttk.Button(r_btn, text='Save Results as File ...',
                   command=self._dk_save_txt).pack(side='left', padx=6)
        ttk.Button(r_btn, text='Save as Excel ...',
                   command=self._dk_save_excel).pack(side='left', padx=6)

        pw.add(frm_l, weight=1)
        pw.add(frm_r, weight=3)

    # ── DigiKey handlers ──────────────────────────────────────────────────────

    def _dk_run_search(self, term, widget, btn, status, mode='part'):
        btn.configure(state='disabled')
        status.configure(text='Searching ...', foreground='gray')

        def worker():
            try:
                parts = (self.dk_client.search_by_keyword(term)
                         if mode == 'kw' else self.dk_client.search_by_mpn(term))
                self.root.after(0, done, parts)
            except DigiKeyError as e:
                self.root.after(0, err, str(e))
            except Exception as e:
                self.root.after(0, err, repr(e))

        def done(parts):
            btn.configure(state='normal')
            if not parts:
                status.configure(text=f'No results: {term}', foreground='orange')
                return
            status.configure(text=f'{len(parts)} result(s)  [{term}]', foreground='gray')
            widget.insert('end', f'── {term}  ({len(parts)} result(s)) ──\n')
            for p in parts:
                self._dk_insert_detail(widget, p, term)
            widget.see('end')

        def err(msg):
            btn.configure(state='normal')
            status.configure(text=f'Error: {msg}', foreground='red')

        threading.Thread(target=worker, daemon=True).start()

    def _dk_insert_detail(self, w, product, search_term=''):
        text = DigiKeyClient.fmt_detail(product)
        mpn  = product.get('ManufacturerProductNumber', '')
        qty  = product.get('QuantityAvailable', 0)
        ns   = product.get('NonStock', False)
        for line in text.splitlines(keepends=True):
            start = w.index('end - 1c linestart')
            w.insert('end', line)
            end = w.index('end - 1c')
            s = line.strip()
            if s.startswith('┌') or s.startswith('└'):
                w.tag_add('separator', start, end)
            elif s.startswith('MPN') and search_term.upper() in mpn.upper():
                w.tag_add('hdr', start, end)
            elif s.startswith('Availability') and (not qty or ns):
                w.tag_add('no_stock', start, end)

    def _populate_dk(self, items):
        self.lb_dk.delete(0, 'end')
        for idx, item in enumerate(items):
            mpn = item.vendor_pn or item.value
            self.lb_dk.insert('end', f'{idx+1:>4}  {mpn:<30} {item.code}')
        self.lbl_dk_count.configure(text=f'{len(items)} items  (valid code only)')

    def _apply_filter_dk(self):
        kw  = self.sv_dk_filter.get().strip().upper()
        src = self._valid
        if kw:
            src = [i for i in src
                   if kw in (i.vendor_pn + i.value + i.code + i.description).upper()]
        self._populate_dk(src)

    def _on_select_dk(self, _=None):
        sel = self.lb_dk.curselection()
        if not sel:
            return
        kw  = self.sv_dk_filter.get().strip().upper()
        src = ([i for i in self._valid
                if kw in (i.vendor_pn + i.value + i.code + i.description).upper()]
               if kw else self._valid)
        if sel[0] >= len(src):
            return
        item = src[sel[0]]
        mpn  = item.vendor_pn or item.value
        lstr = ', '.join(item.locations[:8])
        if len(item.locations) > 8:
            lstr += f' ... (+{len(item.locations)-8})'
        txt = (f'VENDOR_PN  : {mpn}\n'
               f'Value      : {item.value}\n'
               f'10-Code    : {item.code}\n'
               f'Desc       : {item.description}\n'
               f'Vendor     : {item.vendor}\n'
               f'Qty / Loc  : {item.qty}  {lstr}')
        self.txt_dk_detail.delete('1.0', 'end')
        self.txt_dk_detail.insert('end', txt)

    def _search_selected_dk(self):
        sel = self.lb_dk.curselection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇料件')
            return
        kw  = self.sv_dk_filter.get().strip().upper()
        src = ([i for i in self._valid
                if kw in (i.vendor_pn + i.value + i.code + i.description).upper()]
               if kw else self._valid)
        if sel[0] >= len(src):
            return
        mpn = src[sel[0]].vendor_pn or src[sel[0]].value
        self._dk_run_search(mpn, self.txt_dk_result, self.btn_dk_sel, self.lbl_dk_status)

    def _dk_start_batch(self):
        if not self._valid:
            messagebox.showwarning('Warning', '請先在「BOM 批量查價」頁籤載入 BOM 檔案')
            return
        if self._dk_searching:
            messagebox.showwarning('Warning', '查詢進行中，請稍候')
            return
        try:
            build_qty = max(1, int(self.sv_dk_qty.get()))
        except Exception:
            build_qty = 1
        self.sv_dk_qty.set(build_qty)
        if not messagebox.askyesno('Confirm',
                f'將查詢 {len(self._valid)} 筆有效料件 (DigiKey)\n'
                f'生產數量 = {build_qty} 片\n\n是否開始？'):
            return

        self._dk_searching = True
        self._dk_rows      = []
        self.btn_dk_batch.configure(state='disabled')
        self.btn_dk_sel.configure(state='disabled')
        self.txt_dk_result.delete('1.0', 'end')

        ts  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        hdr = (f'DigiKey 批量查價  {ts}  (Build Qty = {build_qty})\n'
               f'BOM: {os.path.basename(self._bom_path)}\n'
               f'Valid items: {len(self._valid)}\n'
               f'{"─"*126}\n'
               f'{"No.":<5} {"10-Code":<15} {"VENDOR_PN":<30} {"Manufacturer":<22} '
               f'{"BOM":>5}  {"TotalQty":>9}  {"Avail":>9}  '
               f'{"Best Pr":>9}  {"Unit Pr":>9}  {"Line Total":>12}  Cur\n'
               f'{"─"*126}\n')
        self.txt_dk_result.insert('end', hdr)
        self.txt_dk_result.tag_add('tbl_hdr', '1.0', 'end')

        threading.Thread(
            target=self._dk_batch_worker,
            args=(list(self._valid), build_qty),
            daemon=True).start()

    def _dk_batch_worker(self, items, build_qty=1):
        found = not_found = 0
        for i, item in enumerate(items):
            self.root.after(0, self._dk_update_progress, i + 1, len(items))
            mpn       = item.vendor_pn or item.value
            total_qty = item.qty * build_qty
            try:
                parts = self.dk_client.search_by_mpn(mpn, limit=5)
                if parts:
                    p              = parts[0]
                    unit_price, _  = DigiKeyClient.price_at(p, total_qty)
                    price_best, _  = DigiKeyClient.price_best(p)
                    avail          = p.get('QuantityAvailable', 0)
                    mfr            = (p.get('Manufacturer') or {}).get('Name', '')
                    line_total     = total_qty * unit_price
                    found         += 1
                    warn           = not unit_price or avail == 0
                    row = (i + 1, item.code, mpn, mfr,
                           item.qty, total_qty, avail,
                           price_best, unit_price, line_total, 'USD',
                           'warn' if warn else 'ok')
                    self._dk_rows.append(row)
                    self.root.after(0, self._dk_append_batch_row, row)
                else:
                    not_found += 1
                    row = (i + 1, item.code, mpn, '-- Not Found --',
                           item.qty, total_qty, 0, 0.0, 0.0, 0.0, '', 'notfound')
                    self._dk_rows.append(row)
                    self.root.after(0, self._dk_append_batch_row, row)
            except Exception as e:
                not_found += 1
                row = (i + 1, item.code, mpn, f'-- Error: {e} --',
                       item.qty, total_qty, 0, 0.0, 0.0, 0.0, '', 'error')
                self._dk_rows.append(row)
                self.root.after(0, self._dk_append_batch_row, row)
            time.sleep(0.3)

        grand = sum(r[9] for r in self._dk_rows if r[-1] in ('ok', 'warn'))
        self.root.after(0, self._dk_batch_done, found, not_found, len(items), grand)

    def _dk_update_progress(self, current, total):
        self.lbl_dk_status.configure(
            text=f'Searching {current}/{total} ...', foreground='gray')

    def _dk_append_batch_row(self, row):
        no, code, mpn, mfr, bom_qty, total_qty, avail, \
            price_best, unit_price, line_total, cur, row_type = row
        pb_s  = f'${price_best:.4f}' if price_best else ''
        up_s  = f'${unit_price:.4f}' if unit_price else ''
        lt_s  = f'${line_total:,.3f}' if line_total else ''
        av_s  = str(avail) if avail else ''
        line  = (f'{no:<5} {code:<15} {mpn:<30} {mfr:<22} '
                 f'{bom_qty:>5}  {total_qty:>9}  {av_s:>9}  '
                 f'{pb_s:>9}  {up_s:>9}  {lt_s:>12}  {cur}\n')
        start = self.txt_dk_result.index('end - 1c linestart')
        self.txt_dk_result.insert('end', line)
        end = self.txt_dk_result.index('end - 1c')
        if row_type == 'warn':
            self.txt_dk_result.tag_add('no_stock', start, end)
        elif row_type == 'notfound':
            self.txt_dk_result.tag_add('not_found', start, end)
        elif row_type == 'error':
            self.txt_dk_result.tag_add('error_row', start, end)
        self.txt_dk_result.see('end')

    def _dk_batch_done(self, found, not_found, total, grand_total=0.0):
        self._dk_searching = False
        self.btn_dk_batch.configure(state='normal')
        self.btn_dk_sel.configure(state='normal')
        summary = (f'{"─"*126}\n'
                   f'完成  查詢 {total} 筆  |  找到 {found} 筆  |  '
                   f'未找到/略過 {not_found} 筆\n'
                   f'Grand Total = ${grand_total:,.2f} USD\n')
        self.txt_dk_result.insert('end', summary)
        self.txt_dk_result.see('end')
        self.lbl_dk_status.configure(
            text=f'Done  {found}/{total} found  |  Total = ${grand_total:,.2f}',
            foreground='gray')

    def _dk_clear(self):
        self.txt_dk_detail.delete('1.0', 'end')
        self.txt_dk_result.delete('1.0', 'end')
        self._dk_rows = []
        self.lbl_dk_status.configure(text='')

    def _dk_save_txt(self):
        text = self.txt_dk_result.get('1.0', 'end').strip()
        if not text:
            messagebox.showinfo('Info', '沒有批量查詢結果可儲存')
            return
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'DigiKey_Batch_{ts}.txt',
            filetypes=[('Text', '*.txt'), ('All', '*.*')])
        if path:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write(text)
            messagebox.showinfo('Done', f'已儲存至\n{path}')

    def _dk_save_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        if not self._dk_rows:
            messagebox.showinfo('Info', '沒有批量查詢結果可匯出')
            return
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'DigiKey_Batch_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'DigiKey BOM Price'

        hdr_fill   = PatternFill('solid', fgColor='1F6CBF')
        hdr_font   = Font(bold=True, color='FFFFFF')
        total_fill = PatternFill('solid', fgColor='FFE699')
        total_font = Font(bold=True)
        fill_warn  = PatternFill('solid', fgColor='FFA500')
        fill_nf    = PatternFill('solid', fgColor='DDDDDD')
        fill_err   = PatternFill('solid', fgColor='FFB3B3')
        center     = Alignment(horizontal='center')
        right      = Alignment(horizontal='right')
        thin       = Side(style='thin', color='AAAAAA')
        border     = Border(bottom=thin)

        headers    = ['No.', '10-Code', 'VENDOR_PN', 'Manufacturer',
                      'BOM Qty', 'Total Qty', 'Avail', 'Best Price (USD)',
                      'Unit Price (USD)', 'Line Total (USD)', 'Currency', 'Status']
        col_widths = [5, 16, 32, 24, 8, 10, 12, 16, 16, 16, 8, 10]
        ws.append(headers)
        for cell, w in zip(ws[1], col_widths):
            cell.fill      = hdr_fill
            cell.font      = hdr_font
            cell.alignment = center
            ws.column_dimensions[cell.column_letter].width = w

        status_label = {'ok': 'Found', 'warn': 'No Stock',
                        'notfound': 'Not Found', 'error': 'Error'}
        for row in self._dk_rows:
            no, code, mpn, mfr, bom_qty, total_qty, avail, \
                price_best, unit_price, line_total, cur, row_type = row
            ws.append([no, code, mpn, mfr, bom_qty, total_qty,
                       avail if avail else None,
                       price_best if price_best else None,
                       unit_price if unit_price else None,
                       line_total if line_total else None,
                       cur, status_label.get(row_type, row_type)])
            xr   = ws.max_row
            fill = {'warn': fill_warn, 'notfound': fill_nf,
                    'error': fill_err}.get(row_type)
            for cell in ws[xr]:
                if fill:
                    cell.fill = fill
            for col in (8, 9, 10):
                ws.cell(xr, col).number_format = '$#,##0.0000'
                ws.cell(xr, col).alignment = right
            for col in (5, 6):
                ws.cell(xr, col).alignment = center

        grand = sum(r[9] for r in self._dk_rows if r[-1] in ('ok', 'warn'))
        ws.append(['', '', '', 'Grand Total', '', '', '', '', '', grand, 'USD', ''])
        tr = ws.max_row
        for cell in ws[tr]:
            cell.fill   = total_fill
            cell.font   = total_font
            cell.border = border
        ws.cell(tr, 10).number_format = '$#,##0.00'
        ws.cell(tr, 10).alignment     = right

        ws.insert_rows(1)
        ws['A1'] = (f'BOM: {os.path.basename(self._bom_path)}  |  '
                    f'Build Qty: {self.sv_dk_qty.get()}  |  '
                    f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        ws['A1'].font = Font(italic=True, color='555555')
        ws.merge_cells('A1:L1')
        ws.freeze_panes = 'A3'

        wb.save(path)
        messagebox.showinfo('Done', f'已儲存至\n{path}')


    # ── Tab 4: Parts Library ──────────────────────────────────────────────────

    def _build_library_tab(self):
        f = self.tab_lib
        f.rowconfigure(2, weight=1)
        f.columnconfigure(0, weight=1)
        p = dict(padx=6, pady=3)

        # Row 0: import + DB status
        r0 = ttk.Frame(f)
        r0.grid(row=0, column=0, sticky='ew', **p)
        ttk.Button(r0, text='Import SAP BOM ...', command=self._lib_import,
                   width=20).pack(side='left')
        ttk.Button(r0, text='Import AVABD Excel ...', command=self._lib_import_avabd,
                   width=22).pack(side='left', padx=4)
        ttk.Separator(r0, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Button(r0, text='Customer BOM 查詢 ...', command=self._lib_cust_bom,
                   width=22).pack(side='left')
        ttk.Separator(r0, orient='vertical').pack(side='left', fill='y', padx=8)
        self.lbl_lib_stats = ttk.Label(r0, text='', foreground='gray', font=FONT_UI)
        self.lbl_lib_stats.pack(side='left')

        # Row 1: query bar
        r1 = ttk.Frame(f)
        r1.grid(row=1, column=0, sticky='ew', **p)
        ttk.Label(r1, text='查詢 (料號/Vendor PN):').pack(side='left')
        self.sv_lib_vpn = tk.StringVar()
        vpn_e = ttk.Entry(r1, textvariable=self.sv_lib_vpn, width=22)
        vpn_e.pack(side='left', padx=4)
        vpn_e.bind('<Return>', lambda _: self._lib_search_vpn())
        ttk.Button(r1, text='Search', command=self._lib_search_vpn).pack(side='left')
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=10)
        ttk.Label(r1, text='關鍵字1:').pack(side='left')
        self.sv_lib_kw = tk.StringVar()
        kw_e = ttk.Entry(r1, textvariable=self.sv_lib_kw, width=14)
        kw_e.pack(side='left', padx=4)
        kw_e.bind('<Return>', lambda _: self._lib_popular())
        ttk.Label(r1, text='關鍵字2:').pack(side='left')
        self.sv_lib_kw2 = tk.StringVar()
        kw_e2 = ttk.Entry(r1, textvariable=self.sv_lib_kw2, width=14)
        kw_e2.pack(side='left', padx=4)
        kw_e2.bind('<Return>', lambda _: self._lib_popular())
        ttk.Label(r1, text='熱門料件 Top:').pack(side='left', padx=(6, 2))
        self.sv_lib_top = tk.IntVar(value=50)
        ttk.Spinbox(r1, from_=10, to=500, textvariable=self.sv_lib_top,
                    width=6).pack(side='left', padx=4)
        ttk.Button(r1, text='查詢', command=self._lib_popular).pack(side='left')
        ttk.Separator(r1, orient='vertical').pack(side='left', fill='y', padx=10)
        ttk.Button(r1, text='Clear', command=self._lib_clear).pack(side='left')

        # Row 2: PanedWindow
        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=2, column=0, sticky='nsew', padx=4, pady=4)

        # Left pane: project list
        frm_l = ttk.Frame(pw)
        frm_l.rowconfigure(1, weight=1)
        frm_l.columnconfigure(0, weight=1)
        ttk.Label(frm_l, text='Projects', font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=4, pady=2)
        self.lb_proj = tk.Listbox(frm_l, font=FONT_TX, activestyle='dotbox',
                                   selectmode='single', width=26)
        sb_proj = ttk.Scrollbar(frm_l, orient='vertical', command=self.lb_proj.yview)
        self.lb_proj.configure(yscrollcommand=sb_proj.set)
        self.lb_proj.grid(row=1, column=0, sticky='nsew')
        sb_proj.grid(row=1, column=1, sticky='ns')
        self.lb_proj.bind('<<ListboxSelect>>', self._lib_on_proj_select)
        ttk.Button(frm_l, text='Delete Project',
                   command=self._lib_delete_project).grid(
            row=2, column=0, columnspan=2, sticky='ew', padx=4, pady=4)

        # Right pane: Treeview
        frm_r = ttk.Frame(pw)
        frm_r.rowconfigure(0, weight=1)
        frm_r.columnconfigure(0, weight=1)
        cols = ('internal_pn', 'vendor_pn', 'description', 'manufacturer',
                'priority', 'proj_count', 'total_qty', 'projects')
        self.tv_lib = ttk.Treeview(frm_r, columns=cols, show='headings',
                                    selectmode='browse')
        col_cfg = [
            ('internal_pn',  '公司料號',    130, 'w'),
            ('vendor_pn',    'Vendor PN',   150, 'w'),
            ('description',  '描述',        210, 'w'),
            ('manufacturer', '製造商',      120, 'w'),
            ('priority',     'Priority',     60, 'center'),
            ('proj_count',   '使用專案數',   80, 'center'),
            ('total_qty',    '總用量',       70, 'center'),
            ('projects',     '專案列表',    200, 'w'),
        ]
        for cid, hdr, w, anchor in col_cfg:
            self.tv_lib.heading(cid, text=hdr,
                                command=lambda c=cid: self._lib_sort(c))
            self.tv_lib.column(cid, width=w, anchor=anchor, minwidth=50)
        self.tv_lib.tag_configure('no_ipn',       background='#FFFACD')
        self.tv_lib.tag_configure('popular',      background='#E8F5E9')
        self.tv_lib.tag_configure('found_row',    background='#DCFFE4')
        self.tv_lib.tag_configure('not_found_row', background='#FFEBEE')
        vsb = ttk.Scrollbar(frm_r, orient='vertical',   command=self.tv_lib.yview)
        hsb = ttk.Scrollbar(frm_r, orient='horizontal', command=self.tv_lib.xview)
        self.tv_lib.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tv_lib.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.tv_lib.bind('<ButtonRelease-1>', self._lib_tv_copy_cell)
        self.tv_lib.bind('<Control-c>',       self._lib_tv_copy_row)
        self._lib_ctx = tk.Menu(self.root, tearoff=0)
        self._lib_ctx.add_command(label='複製此欄', command=self._lib_ctx_copy_cell)
        self._lib_ctx.add_command(label='複製整列 (Tab 分隔)',
                                  command=self._lib_tv_copy_row)
        self.tv_lib.bind('<Button-3>', self._lib_tv_show_ctx)
        self._lib_ctx_col = ''

        r_btn = ttk.Frame(frm_r)
        r_btn.grid(row=2, column=0, columnspan=2, sticky='ew', padx=4, pady=4)
        self.lbl_lib_result = ttk.Label(r_btn, text='', foreground='gray', font=FONT_UI)
        self.lbl_lib_result.pack(side='left')
        ttk.Button(r_btn, text='Export Excel ...',
                   command=self._lib_export_excel).pack(side='right', padx=4)

        pw.add(frm_l, weight=1)
        pw.add(frm_r, weight=3)

        self._lib_refresh_stats()
        self._lib_refresh_projects()

    # ── Library handlers ──────────────────────────────────────────────────────

    def _lib_refresh_stats(self):
        try:
            parts, projects = self.db.stats()
            self.lbl_lib_stats.configure(
                text=f'DB: {os.path.basename(self.db.path)}  '
                     f'|  {parts:,} parts  |  {projects} projects',
                foreground='gray')
        except Exception as e:
            self.lbl_lib_stats.configure(text=f'DB error: {e}', foreground='red')

    def _lib_refresh_projects(self):
        self._lib_projects = list(self.db.get_projects())
        self.lb_proj.delete(0, 'end')
        self.lb_proj.insert('end', f'{"(全部 All)":<24}')
        for row in self._lib_projects:
            name = row['name'][:20]
            self.lb_proj.insert('end', f'{name:<20} ({row["part_count"]})')

    def _lib_on_proj_select(self, _=None):
        sel = self.lb_proj.curselection()
        if not sel:
            return
        proj_id = (None if sel[0] == 0
                   else self._lib_projects[sel[0] - 1]['id'])
        try:
            limit = max(10, int(self.sv_lib_top.get()))
        except (ValueError, TypeError):
            limit = 50
        rows = self.db.popular_parts(project_id=proj_id, limit=limit)
        self._lib_show_rows(rows)
        label = '全部專案' if proj_id is None else self._lib_projects[sel[0]-1]['name']
        self.lbl_lib_result.configure(
            text=f'{label}: {len(rows)} 筆', foreground='gray')

    def _lib_import(self):
        path = filedialog.askopenfilename(
            title='選擇 BOM 檔案 (SAP .txt / Excel / CSV)',
            filetypes=[('SAP BOM txt', '*.txt'),
                       ('Excel', '*.xlsx *.xls'),
                       ('CSV', '*.csv'),
                       ('All', '*.*')])
        if not path:
            return

        # ── SAP BOM .txt — auto-parse, no column mapping needed ───────────────
        if SapBomParser.detect(path):
            try:
                project, items = SapBomParser.parse(path)
            except Exception as e:
                messagebox.showerror('SAP BOM 解析失敗', str(e))
                return
            if not messagebox.askyesno('確認匯入',
                    f'偵測到 SAP BOM 報表\n'
                    f'專案：{project}\n'
                    f'料件數：{len(items)} 筆\n\n是否匯入？'):
                return
            try:
                n = self.db.import_rows(items, project, path)
                self._lib_refresh_stats()
                self._lib_refresh_projects()
                messagebox.showinfo('Done', f'匯入完成\n專案：{project}\n有效筆數：{n}')
            except Exception as e:
                messagebox.showerror('Error', f'匯入失敗:\n{e}')
            return

        # ── Excel / CSV — show column-mapping dialog ──────────────────────────
        try:
            columns, preview, all_rows = self._lib_read_file(path)
        except Exception as e:
            messagebox.showerror('Error', f'讀取檔案失敗:\n{e}')
            return

        initial = os.path.splitext(os.path.basename(path))[0]
        dlg = ColMapDialog(self.root, columns, preview, initial)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return

        mapping = dlg.result['mapping']
        project = dlg.result['project']
        rows = [{k: (raw.get(col, '') if col else '')
                 for k, col in mapping.items()} for raw in all_rows]

        try:
            n = self.db.import_rows(rows, project, path)
            self._lib_refresh_stats()
            self._lib_refresh_projects()
            messagebox.showinfo('Done', f'匯入完成\n專案：{project}\n有效筆數：{n}')
        except Exception as e:
            messagebox.showerror('Error', f'匯入失敗:\n{e}')

    def _lib_read_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in ('.xlsx', '.xls'):
            sheets = _load_excel_sheets(path)
            raw = next(iter(sheets.values()), [])
            if not raw:
                raise ValueError('檔案是空的')
            header = [str(c) if c is not None else f'Col{i}'
                      for i, c in enumerate(raw[0])]
            data = [dict(zip(header, [str(v) if v is not None else '' for v in r]))
                    for r in raw[1:] if any(v is not None for v in r)]
            preview = [list(r) for r in raw[1:4]]
        elif ext in ('.csv', '.txt', '.tsv'):
            import csv
            data = preview = header = None
            # Detect delimiter: tab for .tsv/.txt, comma for .csv
            delimiters = ['\t', ','] if ext in ('.txt', '.tsv') else [',', '\t']
            for enc in ('utf-8-sig', 'utf-8', 'big5', 'cp950', 'latin-1'):
                for delim in delimiters:
                    try:
                        with open(path, encoding=enc, newline='') as fh:
                            reader = csv.DictReader(fh, delimiter=delim)
                            hdr = list(reader.fieldnames or [])
                            if len(hdr) < 2:
                                continue        # wrong delimiter, try next
                            rows = list(reader)
                        header  = hdr
                        data    = rows
                        preview = [list(r.values()) for r in data[:3]]
                        break
                    except (UnicodeDecodeError, LookupError):
                        continue
                if data is not None:
                    break
            if data is None:
                raise ValueError('無法解碼檔案或偵測欄位分隔符號失敗')
        else:
            raise ValueError(f'不支援的格式: {ext}')
        return header, preview, data

    def _lib_search_vpn(self):
        vpn = self.sv_lib_vpn.get().strip()
        if not vpn:
            return
        rows = self.db.search_vendor_pn(vpn)
        self._lib_show_rows(rows)
        self.lbl_lib_result.configure(
            text=f'查詢「{vpn}」: {len(rows)} 筆', foreground='gray')

    def _lib_popular(self):
        kw1 = self.sv_lib_kw.get().strip()
        kw2 = self.sv_lib_kw2.get().strip()
        keywords = [k for k in [kw1, kw2] if k]
        try:
            limit = max(10, int(self.sv_lib_top.get()))
        except (ValueError, TypeError):
            limit = 50
        rows = self.db.popular_parts(keywords=keywords, limit=limit)
        self._lib_show_rows(rows)
        kw_label = '  (關鍵字: ' + ' + '.join(keywords) + ')' if keywords else ''
        label = f'熱門料件 Top {limit}' + kw_label
        self.lbl_lib_result.configure(text=f'{label}: {len(rows)} 筆',
                                       foreground='gray')

    def _lib_show_rows(self, rows):
        self.tv_lib.delete(*self.tv_lib.get_children())
        for i, r in enumerate(rows):
            prio = r['priority'] if r['priority'] is not None else ''
            vals = (r['internal_pn'] or '', r['vendor_pn'] or '',
                    r['description'] or '', r['manufacturer'] or '',
                    prio, r['proj_count'], r['total_qty'] or 0,
                    r['projects'] or '')
            if not r['internal_pn']:
                tag = 'no_ipn'
            elif i < 3:
                tag = 'popular'
            else:
                tag = ''
            self.tv_lib.insert('', 'end', values=vals, tags=(tag,))

    def _lib_sort(self, col):
        items = [(self.tv_lib.set(iid, col), iid)
                 for iid in self.tv_lib.get_children()]
        try:
            items.sort(key=lambda x: float(x[0]) if x[0] else 0, reverse=True)
        except ValueError:
            items.sort(key=lambda x: x[0].lower())
        for idx, (_, iid) in enumerate(items):
            self.tv_lib.move(iid, '', idx)

    def _lib_import_avabd(self):
        path = filedialog.askopenfilename(
            title='選擇 AVABD Excel 檔案',
            filetypes=[('Excel', '*.xlsx *.xls'), ('All', '*.*')])
        if not path:
            return

        try:
            sheets = _load_excel_sheets(path)
        except Exception as e:
            messagebox.showerror('Error', f'無法開啟 Excel：{e}')
            return

        total_inserted = 0
        summary_lines  = []
        KEY_COLS = ('PART_NUMBER', 'VENDOR_PN', 'DESCRIPTION', 'VENDOR', 'Priority')

        for sh_name, rows in sheets.items():
            if not rows:
                continue

            # Find header row — first row with PART_NUMBER
            hdr_idx = None
            for ri, row in enumerate(rows):
                if any(str(c or '').strip() == 'PART_NUMBER' for c in row):
                    hdr_idx = ri
                    break
            if hdr_idx is None:
                continue

            hdr = [str(c or '').strip() for c in rows[hdr_idx]]
            col = {name: hdr.index(name) for name in KEY_COLS if name in hdr}
            if 'PART_NUMBER' not in col or 'VENDOR_PN' not in col:
                continue

            import_rows = []
            for row in rows[hdr_idx + 1:]:
                ipn = str(row[col['PART_NUMBER']] or '').strip() if 'PART_NUMBER' in col else ''
                vpn = str(row[col['VENDOR_PN']]   or '').strip() if 'VENDOR_PN'   in col else ''
                if not ipn and not vpn:
                    continue
                import_rows.append({
                    'internal_pn': ipn,
                    'vendor_pn':   vpn,
                    'description': str(row[col['DESCRIPTION']] or '').strip() if 'DESCRIPTION' in col else '',
                    'manufacturer':str(row[col['VENDOR']]       or '').strip() if 'VENDOR'       in col else '',
                    'priority':    row[col['Priority']] if 'Priority' in col else None,
                    'qty': 0, 'reference': '',
                })

            if import_rows:
                n = self.db.import_rows(import_rows, project_name=sh_name,
                                        source_file=os.path.basename(path))
                total_inserted += n
                summary_lines.append(f'  {sh_name}: {n} 筆')

        self._lib_refresh_projects()
        if summary_lines:
            messagebox.showinfo('Import 完成',
                f'已匯入 {total_inserted} 筆料號：\n' + '\n'.join(summary_lines))
        else:
            messagebox.showwarning('Import', '找不到可匯入的工作表（需有 PART_NUMBER 欄位）')

    def _lib_cust_bom(self):
        """Load customer BOM → batch-lookup vendor PNs against the parts DB."""
        if self.db.stats()[1] == 0:
            messagebox.showwarning('Warning', '料號庫尚無資料，請先 Import SAP BOM')
            return
        path = filedialog.askopenfilename(
            title='選擇客戶 BOM 檔案',
            filetypes=[('Excel', '*.xlsx *.xls'), ('CSV', '*.csv'),
                       ('Text/TSV', '*.txt *.tsv'), ('All', '*.*')])
        if not path:
            return
        try:
            columns, preview, all_rows = self._lib_read_file(path)
        except Exception as e:
            messagebox.showerror('Error', f'讀取失敗:\n{e}')
            return

        dlg = _VpnColPicker(self.root, columns, preview)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return
        vpn_col = dlg.result

        # Gather unique vendor PNs (preserve order)
        seen = {}
        for row in all_rows:
            vpn = str(row.get(vpn_col, '') or '').strip()
            if vpn and vpn not in seen:
                seen[vpn] = None

        results = []
        for vpn in seen:
            hits = self.db.lookup_vendor_pn_exact(vpn)
            if hits:
                for h in hits:
                    results.append({
                        'status':       'found' if h['internal_pn'] else 'no_ipn',
                        'internal_pn':  h['internal_pn'] or '',
                        'vendor_pn':    vpn,
                        'db_vpn':       h['vendor_pn'] or '',
                        'description':  h['description'] or '',
                        'manufacturer': h['manufacturer'] or '',
                        'proj_count':   h['proj_count'],
                        'total_qty':    h['total_qty'] or 0,
                        'projects':     h['projects'] or '',
                    })
            else:
                results.append({
                    'status':       'not_found',
                    'internal_pn':  '',
                    'vendor_pn':    vpn,
                    'db_vpn':       '',
                    'description':  '',
                    'manufacturer': '',
                    'proj_count':   0,
                    'total_qty':    0,
                    'projects':     '— 未登錄 —',
                })

        self._lib_show_cust_results(results, os.path.basename(path))

    def _lib_show_cust_results(self, results, filename=''):
        self.tv_lib.delete(*self.tv_lib.get_children())
        # Restore standard column headings in case they were changed
        self.tv_lib.heading('projects', text='專案列表 / 狀態')

        found = not_found = no_ipn = 0
        for r in results:
            status = r['status']
            if status == 'found':
                found += 1
                tag = 'found_row'
            elif status == 'no_ipn':
                no_ipn += 1
                tag = 'no_ipn'
            else:
                not_found += 1
                tag = 'not_found_row'

            proj_val = r['proj_count'] if r['proj_count'] else ''
            qty_val  = r['total_qty']  if r['total_qty']  else ''
            self.tv_lib.insert('', 'end', tags=(tag,), values=(
                r['internal_pn'],
                r['vendor_pn'],
                r['description'],
                r['manufacturer'],
                r['priority'] if r['priority'] is not None else '',
                proj_val,
                qty_val,
                r['projects'],
            ))

        total = len(results)
        msg = (f'客戶 BOM [{filename}]  共 {total} 筆：'
               f'  ✓ 找到 {found}')
        if no_ipn:
            msg += f'  (其中 {no_ipn} 筆有料但無公司料號)'
        msg += f'  ✗ 未登錄 {not_found}'
        self.lbl_lib_result.configure(text=msg, foreground='gray')

    _TV_LIB_COLS = ('internal_pn', 'vendor_pn', 'description', 'manufacturer',
                    'proj_count', 'total_qty', 'projects')

    def _lib_tv_copy_cell(self, event):
        region = self.tv_lib.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self.tv_lib.identify_column(event.x)   # '#1', '#2', …
        iid = self.tv_lib.identify_row(event.y)
        if not iid:
            return
        self._lib_ctx_col = col
        col_idx = int(col[1:]) - 1
        value = self.tv_lib.set(iid, self._TV_LIB_COLS[col_idx])
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.lbl_lib_result.configure(text=f'已複製: {value}', foreground='#1a7a1a')

    def _lib_tv_copy_row(self, _=None):
        iid = self.tv_lib.focus()
        if not iid:
            sel = self.tv_lib.selection()
            iid = sel[0] if sel else ''
        if not iid:
            return
        values = [str(self.tv_lib.set(iid, c)) for c in self._TV_LIB_COLS]
        self.root.clipboard_clear()
        self.root.clipboard_append('\t'.join(values))
        self.lbl_lib_result.configure(text='已複製整列 (Tab 分隔)', foreground='#1a7a1a')

    def _lib_get_selected_part_id(self):
        iid = self.tv_lib.focus() or (self.tv_lib.selection() or (None,))[0]
        if not iid:
            return None
        vals = self.tv_lib.item(iid, 'values')
        if not vals:
            return None
        with self.db._conn() as c:
            row = c.execute(
                'SELECT id FROM parts WHERE internal_pn=? AND vendor_pn=?',
                (vals[0], vals[1])).fetchone()
        return row[0] if row else None

    def _lib_attach_file(self):
        part_id = self._lib_get_selected_part_id()
        if part_id is None:
            messagebox.showwarning('Attachments', '請先選擇一個料件')
            return
        paths = _SmartFileBrowser.ask_files(self.root, 'Attach File(s) to Part')
        for p in paths:
            self.db.add_attachment(part_id, p)
        if paths:
            messagebox.showinfo('Done', f'{len(paths)} file(s) attached.')

    def _lib_tv_show_ctx(self, event):
        col = self.tv_lib.identify_column(event.x)
        iid = self.tv_lib.identify_row(event.y)
        if iid:
            self.tv_lib.selection_set(iid)
            self.tv_lib.focus(iid)
        self._lib_ctx_col = col

        # Rebuild attachment section dynamically
        try:
            self._lib_ctx.delete('Attach File(s)...')
        except Exception:
            pass
        # Remove old dynamic entries (index 2 onwards)
        while True:
            try:
                self._lib_ctx.delete(2)
            except Exception:
                break
        self._lib_ctx.add_separator()
        self._lib_ctx.add_command(label='Attach File(s)...',
                                   command=self._lib_attach_file)
        part_id = self._lib_get_selected_part_id()
        if part_id:
            atts = self.db.get_attachments(part_id)
            if atts:
                for att in atts:
                    fname = os.path.basename(att['file_path'])
                    self._lib_ctx.add_command(
                        label=f'📄 {fname}',
                        command=lambda p=att['file_path']: self._open_file(p))
                for att in atts:
                    fname = os.path.basename(att['file_path'])
                    self._lib_ctx.add_command(
                        label=f'✕ Remove  {fname}',
                        foreground='#FF3B30',
                        command=lambda aid=att['id']: self.db.delete_attachment(aid))

        self._lib_ctx.tk_popup(event.x_root, event.y_root)

    def _lib_ctx_copy_cell(self):
        iid = self.tv_lib.focus()
        if not iid or not self._lib_ctx_col:
            return
        col_idx = int(self._lib_ctx_col[1:]) - 1
        value = self.tv_lib.set(iid, self._TV_LIB_COLS[col_idx])
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.lbl_lib_result.configure(text=f'已複製: {value}', foreground='#1a7a1a')

    def _lib_clear(self):
        self.tv_lib.delete(*self.tv_lib.get_children())
        self.tv_lib.heading('projects', text='專案列表')
        self.sv_lib_vpn.set('')
        self.sv_lib_kw.set(''); self.sv_lib_kw2.set('')
        self.lbl_lib_result.configure(text='')
        self.lb_proj.selection_clear(0, 'end')

    def _lib_delete_project(self):
        sel = self.lb_proj.curselection()
        if not sel or sel[0] == 0:
            messagebox.showwarning('Warning', '請先選擇要刪除的專案')
            return
        proj = self._lib_projects[sel[0] - 1]
        if not messagebox.askyesno('Confirm',
                f'刪除專案「{proj["name"]}」及其所有 BOM 資料？\n此操作不可復原'):
            return
        self.db.delete_project(proj['id'])
        self._lib_refresh_stats()
        self._lib_refresh_projects()
        self._lib_clear()

    def _lib_export_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        items = self.tv_lib.get_children()
        if not items:
            messagebox.showinfo('Info', '沒有資料可匯出')
            return
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'PartsLib_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Parts Library'
        hdr_fill = PatternFill('solid', fgColor='1F6CBF')
        hdr_font = Font(bold=True, color='FFFFFF')
        headers  = ['公司料號', 'Vendor PN', '描述', '製造商',
                    '使用專案數', '總用量', '專案列表']
        col_wids = [18, 22, 32, 18, 12, 10, 45]
        ws.append(headers)
        for cell, w in zip(ws[1], col_wids):
            cell.fill = hdr_fill
            cell.font = hdr_font
            ws.column_dimensions[cell.column_letter].width = w

        no_ipn_fill = PatternFill('solid', fgColor='FFFACD')
        pop_fill    = PatternFill('solid', fgColor='E8F5E9')
        for iid in items:
            vals = list(self.tv_lib.item(iid, 'values'))
            ws.append(vals)
            tags = self.tv_lib.item(iid, 'tags')
            if 'no_ipn' in tags:
                for cell in ws[ws.max_row]:
                    cell.fill = no_ipn_fill
            elif 'popular' in tags:
                for cell in ws[ws.max_row]:
                    cell.fill = pop_fill

        ws.insert_rows(1)
        ws['A1'] = f'匯出時間: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        ws['A1'].font = Font(italic=True, color='555555')
        ws.merge_cells('A1:G1')
        ws.freeze_panes = 'A3'

        wb.save(path)
        messagebox.showinfo('Done', f'已儲存至\n{path}')

    # ── Tab 5: Power DB Builder ───────────────────────────────────────────────

    def _build_power_tab(self):
        f = self.tab_pwr
        f.rowconfigure(1, weight=1)
        f.columnconfigure(0, weight=1)
        p = dict(padx=6, pady=3)

        # Row 0: action bar
        r0 = ttk.Frame(f)
        r0.grid(row=0, column=0, sticky='ew', **p)
        ttk.Button(r0, text='New Part',    command=self._pwr_new_part,    width=10).pack(side='left')
        ttk.Button(r0, text='Save Part',   command=self._pwr_save_part,   width=10).pack(side='left', padx=2)
        ttk.Button(r0, text='Delete Part', command=self._pwr_delete_part, width=10).pack(side='left')
        ttk.Separator(r0, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Label(r0, text='Search:').pack(side='left')
        self.sv_pwr_search = tk.StringVar()
        se = ttk.Entry(r0, textvariable=self.sv_pwr_search, width=20)
        se.pack(side='left', padx=4)
        se.bind('<Return>', lambda _: self._pwr_search())
        ttk.Button(r0, text='Go', command=self._pwr_search, width=4).pack(side='left')
        ttk.Separator(r0, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Button(r0, text='Import BOM',   command=self._pwr_import_bom,   width=12).pack(side='left')
        ttk.Button(r0, text='Export DB Excel', command=self._pwr_export_db_excel, width=14).pack(side='left', padx=2)
        ttk.Button(r0, text='Power Tree', command=self._pwr_show_power_tree, width=12).pack(side='left', padx=2)
        ttk.Separator(r0, orient='vertical').pack(side='left', fill='y', padx=8)
        self.lbl_pwr_stats = ttk.Label(r0, text='', foreground='gray', font=FONT_UI)
        self.lbl_pwr_stats.pack(side='left', padx=4)

        # Row 1: nested Notebook — "Parts & Rails" vs. the independent "Power Sources" library
        pdb_nb = ttk.Notebook(f)
        pdb_nb.grid(row=1, column=0, sticky='nsew', padx=4, pady=4)
        tab_pdb_main = ttk.Frame(pdb_nb)
        tab_pdb_src  = ttk.Frame(pdb_nb)
        tab_pdb_main.rowconfigure(0, weight=1)
        tab_pdb_main.columnconfigure(0, weight=1)
        tab_pdb_src.rowconfigure(1, weight=1)
        tab_pdb_src.columnconfigure(0, weight=1)
        pdb_nb.add(tab_pdb_main, text='Parts & Rails')
        pdb_nb.add(tab_pdb_src,  text='Power Sources (DC/DC Library)')

        # main PanedWindow (horizontal)
        pw = ttk.PanedWindow(tab_pdb_main, orient='horizontal')
        pw.grid(row=0, column=0, sticky='nsew')

        # ── Left pane: parts list ─────────────────────────────────────────────
        frm_l = ttk.Frame(pw)
        frm_l.rowconfigure(1, weight=1)
        frm_l.columnconfigure(0, weight=1)
        ttk.Label(frm_l, text='Parts', font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=4, pady=2)
        pcols = ('check', 'qty', 'vendor_pn', 'company_pn', 'manufacturer', 'description')
        self.tv_pwr_parts = ttk.Treeview(frm_l, columns=pcols, show='headings',
                                          selectmode='browse', height=30)
        for cid, hdr, w, anchor in [
            ('check',        '✓',           28, 'center'),
            ('qty',          'Qty',          38, 'center'),
            ('vendor_pn',    'Vendor PN',   120, 'w'),
            ('company_pn',   'Company PN',   90, 'w'),
            ('manufacturer', 'Mfr',          80, 'w'),
            ('description',  'Description', 150, 'w'),
        ]:
            self.tv_pwr_parts.heading(cid, text=hdr)
            self.tv_pwr_parts.column(cid, width=w, minwidth=20, anchor=anchor)
        vsb_p = ttk.Scrollbar(frm_l, orient='vertical', command=self.tv_pwr_parts.yview)
        self.tv_pwr_parts.configure(yscrollcommand=vsb_p.set)
        self.tv_pwr_parts.grid(row=1, column=0, sticky='nsew')
        vsb_p.grid(row=1, column=1, sticky='ns')
        self.tv_pwr_parts.bind('<<TreeviewSelect>>', self._pwr_on_part_select)
        self.tv_pwr_parts.bind('<ButtonRelease-1>',  self._pwr_parts_click)
        pw.add(frm_l, weight=1)

        # ── Right pane: vertical PanedWindow ──────────────────────────────────
        pw_r = ttk.PanedWindow(pw, orient='vertical')

        # TOP: part form
        frm_form = ttk.LabelFrame(pw_r, text='Part Information', padding=4)
        frm_form.columnconfigure(1, weight=1)
        frm_form.columnconfigure(3, weight=1)
        self._pwr_form_vars = {}
        form_fields = [
            ('vendor_pn',    'Vendor PN *',   0, 0),
            ('company_pn',   'Company PN',    0, 2),
            ('manufacturer', 'Manufacturer',  1, 0),
            ('description',  'Description',   1, 2),
        ]
        for key, label, row, col in form_fields:
            ttk.Label(frm_form, text=label, font=FONT_UI).grid(
                row=row, column=col, sticky='w', padx=4, pady=2)
            var = tk.StringVar()
            ttk.Entry(frm_form, textvariable=var).grid(
                row=row, column=col+1, sticky='ew', padx=(0, 8), pady=2)
            self._pwr_form_vars[key] = var

        ttk.Label(frm_form, text='Datasheet', font=FONT_UI).grid(
            row=2, column=0, sticky='w', padx=4, pady=2)
        self._pwr_form_vars['datasheet_path'] = tk.StringVar()
        ds_frame = ttk.Frame(frm_form)
        ds_frame.grid(row=2, column=1, columnspan=3, sticky='ew', pady=2)
        ds_frame.columnconfigure(0, weight=1)
        ttk.Entry(ds_frame, textvariable=self._pwr_form_vars['datasheet_path']).grid(
            row=0, column=0, sticky='ew', padx=(0, 4))
        ttk.Button(ds_frame, text='Browse...', width=9,
                   command=self._pwr_browse_datasheet).grid(row=0, column=1)

        ttk.Label(frm_form, text='Remark', font=FONT_UI).grid(
            row=3, column=0, sticky='nw', padx=4, pady=2)
        self._pwr_remark = tk.Text(frm_form, height=2, font=FONT_UI, wrap='word')
        self._pwr_remark.grid(row=3, column=1, columnspan=3, sticky='ew', padx=(0, 8), pady=2)
        pw_r.add(frm_form, weight=0)

        # MID: rails treeview
        frm_rails = ttk.LabelFrame(pw_r, text='Power Rails', padding=4)
        frm_rails.rowconfigure(1, weight=1)
        frm_rails.columnconfigure(0, weight=1)
        rb = ttk.Frame(frm_rails)
        rb.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 4))
        ttk.Button(rb, text='Add Rail',    command=self._pwr_add_rail,    width=10).pack(side='left')
        ttk.Button(rb, text='Edit Rail',   command=self._pwr_edit_rail,   width=10).pack(side='left', padx=2)
        ttk.Button(rb, text='Delete Rail', command=self._pwr_delete_rail, width=10).pack(side='left')
        ttk.Separator(rb, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Button(rb, text='Paste JSON / TSV', command=self._pwr_paste_rails, width=16).pack(side='left')
        rcols = ('rail_name', 'power_group', 'input_rail', 'voltage', 'typ_i', 'max_i',
                 'typ_p', 'max_p', 'condition')
        self.tv_pwr_rails = ttk.Treeview(frm_rails, columns=rcols, show='headings',
                                          selectmode='browse', height=6)
        for cid, hdr, w in [
            ('rail_name',   'Rail Name',   110), ('power_group', 'Power Group', 100),
            ('input_rail',  'Input Rail',  100), ('voltage',     'V',            50),
            ('typ_i',  'Typ I(mA)', 75),
            ('max_i',       'Max I(mA)',    75), ('typ_p',  'Typ P(W)',  75),
            ('max_p',       'Max P(W)',     75), ('condition', 'Condition', 100),
        ]:
            self.tv_pwr_rails.heading(cid, text=hdr)
            self.tv_pwr_rails.column(cid, width=w, anchor='center', minwidth=40)
        vsb_r = ttk.Scrollbar(frm_rails, orient='vertical', command=self.tv_pwr_rails.yview)
        self.tv_pwr_rails.configure(yscrollcommand=vsb_r.set)
        self.tv_pwr_rails.grid(row=1, column=0, sticky='nsew')
        vsb_r.grid(row=1, column=1, sticky='ns')
        pw_r.add(frm_rails, weight=1)

        pw.add(pw_r, weight=3)

        # ── Independent "Power Sources / DC-DC Library" tab ─────────────────────
        # Not gated by which Part is selected on the left — a standalone catalog of
        # reusable DC/DC designs, plus any legacy per-IC sources (shown with their owner).
        sb2 = ttk.Frame(tab_pdb_src)
        sb2.grid(row=0, column=0, sticky='ew', padx=4, pady=(4, 0))
        ttk.Button(sb2, text='Add',    command=self._pwr_add_source,    width=8).pack(side='left')
        ttk.Button(sb2, text='Edit',   command=self._pwr_edit_source,   width=8).pack(side='left', padx=2)
        ttk.Button(sb2, text='Delete', command=self._pwr_delete_source, width=8).pack(side='left')
        src_cols = ('owner', 'rail_name', 'source_rail', 'regulator_type', 'efficiency',
                    'input_voltage', 'output_voltage', 'remark')
        self.tv_pwr_src = ttk.Treeview(tab_pdb_src, columns=src_cols, show='headings',
                                        selectmode='browse')
        for cid, hdr, w in [
            ('owner',          'Owner',       100), ('rail_name',      'Rail Name',   100),
            ('source_rail',    'Source Rail', 100), ('regulator_type', 'Type',         90),
            ('efficiency',     'Eff',          50), ('input_voltage',  'Vin(V)',       60),
            ('output_voltage', 'Vout(V)',      60), ('remark',         'Remark',      160),
        ]:
            self.tv_pwr_src.heading(cid, text=hdr)
            self.tv_pwr_src.column(cid, width=w, anchor='center', minwidth=40)
        vsb_s2 = ttk.Scrollbar(tab_pdb_src, orient='vertical', command=self.tv_pwr_src.yview)
        self.tv_pwr_src.configure(yscrollcommand=vsb_s2.set)
        self.tv_pwr_src.grid(row=1, column=0, sticky='nsew', padx=(4, 0), pady=4)
        vsb_s2.grid(row=1, column=1, sticky='ns', pady=4)

        self._pwr_refresh_stats()
        self._pwr_search()
        self._pwr_load_all_sources()

    # ── Power DB handlers ────────────────────────────────────────────────────

    def _pwr_refresh_stats(self):
        try:
            parts, rails = self.pwr_db.stats()
            self.lbl_pwr_stats.configure(
                text=f'DB: {os.path.basename(self.pwr_db.path)}  |  '
                     f'{parts} parts  |  {rails} rails',
                foreground='gray')
        except Exception as e:
            self.lbl_pwr_stats.configure(text=f'DB error: {e}', foreground='red')

    def _pwr_search(self):
        kw = self.sv_pwr_search.get().strip()
        rows = self.pwr_db.search_parts(kw)
        self.tv_pwr_parts.delete(*self.tv_pwr_parts.get_children())
        for r in rows:
            pid = r['id']
            sel = self._pwr_part_sel.setdefault(pid, {'checked': False, 'qty': 1})
            chk = '☑' if sel['checked'] else '☐'
            self.tv_pwr_parts.insert('', 'end', iid=str(pid),
                                      values=(chk, sel['qty'],
                                              r['vendor_pn'], r['company_pn'] or '',
                                              r['manufacturer'] or '', r['description'] or ''))

    def _pwr_parts_click(self, event):
        """Single-click: toggle ✓ on col #1, open inline qty editor on col #2."""
        region = self.tv_pwr_parts.identify_region(event.x, event.y)
        col    = self.tv_pwr_parts.identify_column(event.x)
        iid    = self.tv_pwr_parts.identify_row(event.y)
        if region != 'cell' or not iid:
            return
        pid = int(iid)

        if col == '#1':  # checkbox column
            sel = self._pwr_part_sel.setdefault(pid, {'checked': False, 'qty': 1})
            sel['checked'] = not sel['checked']
            self.tv_pwr_parts.set(iid, 'check', '☑' if sel['checked'] else '☐')

        elif col == '#2':  # qty column — show inline Entry
            bbox = self.tv_pwr_parts.bbox(iid, '#2')
            if not bbox:
                return
            x, y, w, h = bbox
            var = tk.StringVar(value=str(
                self._pwr_part_sel.get(pid, {'qty': 1})['qty']))
            ent = ttk.Entry(self.tv_pwr_parts, textvariable=var,
                            justify='center', width=5)
            ent.place(x=x, y=y, width=w, height=h)
            ent.focus_set()
            ent.select_range(0, 'end')

            def _save(_=None):
                try:
                    q = max(1, int(float(var.get())))
                except ValueError:
                    q = 1
                self._pwr_part_sel.setdefault(
                    pid, {'checked': False, 'qty': 1})['qty'] = q
                self.tv_pwr_parts.set(iid, 'qty', q)
                ent.destroy()

            ent.bind('<Return>',   _save)
            ent.bind('<Tab>',      _save)
            ent.bind('<FocusOut>', _save)
            ent.bind('<Escape>',   lambda _: ent.destroy())

    def _pwr_on_part_select(self, _=None):
        sel = self.tv_pwr_parts.selection()
        if not sel:
            return
        part_id = int(sel[0])
        self._pwr_part_id = part_id
        self._pwr_load_part_to_form(part_id)
        self._pwr_load_rails(part_id)

    def _pwr_load_part_to_form(self, part_id):
        p = self.pwr_db.get_part(part_id)
        if not p:
            return
        for key in ('vendor_pn', 'company_pn', 'manufacturer', 'description', 'datasheet_path'):
            self._pwr_form_vars[key].set(p[key] or '')
        self._pwr_remark.delete('1.0', 'end')
        self._pwr_remark.insert('1.0', p['remark'] or '')

    def _pwr_new_part(self):
        self._pwr_part_id = None
        self._pwr_clear_form()

    def _pwr_clear_form(self):
        for var in self._pwr_form_vars.values():
            var.set('')
        self._pwr_remark.delete('1.0', 'end')
        self.tv_pwr_rails.delete(*self.tv_pwr_rails.get_children())
        self.tv_pwr_parts.selection_remove(*self.tv_pwr_parts.selection())

    def _pwr_save_part(self):
        vpn = self._pwr_form_vars['vendor_pn'].get().strip()
        if not vpn:
            messagebox.showwarning('Warning', 'Vendor PN 為必填')
            return
        data = {k: v.get().strip() for k, v in self._pwr_form_vars.items()}
        data['remark'] = self._pwr_remark.get('1.0', 'end').strip()
        if self._pwr_part_id:
            data['id'] = self._pwr_part_id   # pass id so upsert can UPDATE by id
        try:
            part_id = self.pwr_db.upsert_part(data)
            self._pwr_part_id = part_id
            self._pwr_search()
            self.tv_pwr_parts.selection_set(str(part_id))
            self._pwr_refresh_stats()
        except Exception as e:
            messagebox.showerror('Error', f'儲存失敗:\n{e}')

    def _pwr_delete_part(self):
        if not self._pwr_part_id:
            messagebox.showwarning('Warning', '請先選擇一個 Part')
            return
        p = self.pwr_db.get_part(self._pwr_part_id)
        if not messagebox.askyesno('Confirm',
                f'刪除 {p["vendor_pn"]} 及其所有 Rails / Sources？\n此操作不可復原'):
            return
        self.pwr_db.delete_part(self._pwr_part_id)
        self._pwr_part_id = None
        self._pwr_clear_form()
        self._pwr_search()
        self._pwr_refresh_stats()

    def _pwr_browse_datasheet(self):
        path = filedialog.askopenfilename(
            title='選擇 Datasheet',
            filetypes=[('PDF', '*.pdf'), ('All', '*.*')])
        if path:
            self._pwr_form_vars['datasheet_path'].set(path)

    def _pwr_load_rails(self, part_id):
        self.tv_pwr_rails.delete(*self.tv_pwr_rails.get_children())
        sum_typ_i = sum_max_i = sum_typ_p = sum_max_p = 0.0
        for r in self.pwr_db.get_rails(part_id):
            typ_a = r['typ_current_a']
            max_a = r['max_current_a']
            typ_p = r['voltage'] * typ_a
            max_p = r['voltage'] * max_a
            self.tv_pwr_rails.insert('', 'end', iid=str(r['id']), values=(
                r['rail_name'], r['power_group'] or '', r['input_rail'] or '',
                f"{r['voltage']:.2f}", f"{typ_a * 1000:.2f}",
                f"{max_a * 1000:.2f}", f"{typ_p:.4f}", f"{max_p:.4f}",
                r['condition'] or ''))
            sum_typ_i += typ_a; sum_max_i += max_a
            sum_typ_p += typ_p; sum_max_p += max_p
        self.tv_pwr_rails.tag_configure(
            'total', background='#D6EAF8', font=('Microsoft JhengHei UI', 9, 'bold'))
        self.tv_pwr_rails.insert('', 'end', iid='__total__', tags=('total',), values=(
            'TOTAL', '', '', '', f'{sum_typ_i * 1000:.2f}', f'{sum_max_i * 1000:.2f}',
            f'{sum_typ_p:.4f}', f'{sum_max_p:.4f}', ''))

    def _pwr_add_rail(self):
        if not self._pwr_part_id:
            messagebox.showwarning('Warning', '請先選擇或儲存一個 Part')
            return
        dlg = RailDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return
        dlg.result['part_id'] = self._pwr_part_id
        self.pwr_db.upsert_rail(dlg.result)
        self._pwr_load_rails(self._pwr_part_id)
        self._pwr_refresh_stats()

    def _pwr_edit_rail(self):
        sel = self.tv_pwr_rails.selection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇要編輯的 Rail')
            return
        if sel[0] == '__total__':
            messagebox.showwarning('Warning', 'TOTAL 列不可編輯')
            return
        rail_id = int(sel[0])
        with self.pwr_db._conn() as c:
            r = c.execute('SELECT * FROM power_rails WHERE id=?', (rail_id,)).fetchone()
        if not r:
            return
        dlg = RailDialog(self.root, dict(r))
        self.root.wait_window(dlg)
        if dlg.result is None:
            return
        dlg.result['id'] = rail_id
        dlg.result['part_id'] = self._pwr_part_id
        self.pwr_db.upsert_rail(dlg.result)
        self._pwr_load_rails(self._pwr_part_id)

    def _pwr_delete_rail(self):
        sel = self.tv_pwr_rails.selection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇要刪除的 Rail')
            return
        if sel[0] == '__total__':
            messagebox.showwarning('Warning', 'TOTAL 列不可刪除')
            return
        if not messagebox.askyesno('Confirm', '刪除這條 Rail？'):
            return
        self.pwr_db.delete_rail(int(sel[0]))
        self._pwr_load_rails(self._pwr_part_id)
        self._pwr_refresh_stats()

    def _pwr_load_all_sources(self):
        """Power Sources / DC-DC Library — independent of the Parts selection."""
        self.tv_pwr_src.delete(*self.tv_pwr_src.get_children())
        for s in self.pwr_db.get_all_sources():
            self.tv_pwr_src.insert('', 'end', iid=str(s['id']), values=(
                s['owner_vpn'] or '(library)',
                s['rail_name'], s['source_rail'] or '',
                s['regulator_type'] or '', f"{s['efficiency']:.2f}",
                f"{s['input_voltage']:.2f}", f"{s['output_voltage']:.2f}",
                s['remark'] or ''))

    def _pwr_add_source(self):
        rail_names = self.pwr_db.get_all_rail_names()
        dlg = SourceDialog(self.root, rail_names)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return
        dlg.result['part_id'] = None  # new entries are always independent library rows
        self.pwr_db.upsert_source(dlg.result)
        self._pwr_load_all_sources()

    def _pwr_edit_source(self):
        sel = self.tv_pwr_src.selection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇要編輯的 Source')
            return
        src_id = int(sel[0])
        with self.pwr_db._conn() as c:
            s = c.execute('SELECT * FROM power_sources WHERE id=?', (src_id,)).fetchone()
        if not s:
            return
        rail_names = self.pwr_db.get_all_rail_names()
        dlg = SourceDialog(self.root, rail_names, dict(s))
        self.root.wait_window(dlg)
        if dlg.result is None:
            return
        dlg.result['id'] = src_id
        # part_id intentionally omitted: upsert_source's UPDATE branch never writes
        # part_id, so a legacy IC-owned row keeps its owner untouched.
        self.pwr_db.upsert_source(dlg.result)
        self._pwr_load_all_sources()

    def _pwr_delete_source(self):
        sel = self.tv_pwr_src.selection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇要刪除的 Source')
            return
        if not messagebox.askyesno('Confirm', '刪除這筆 Power Source？'):
            return
        self.pwr_db.delete_source(int(sel[0]))
        self._pwr_load_all_sources()

    # ── Power BOM Import ──────────────────────────────────────────────────────

    def _pwr_import_bom(self):
        if self.pwr_db.stats()[0] == 0:
            messagebox.showwarning('Warning', 'Power DB 尚無資料，請先新增 IC Parts')
            return
        path = filedialog.askopenfilename(
            title='選擇 BOM 檔案',
            filetypes=[('BOM (OrCAD)', '*.bom *.BOM'),
                       ('Excel', '*.xlsx *.xls'), ('CSV', '*.csv'),
                       ('All', '*.*')])
        if not path:
            return

        seen: dict[str, int] = {}

        if path.lower().endswith('.bom'):
            # OrCAD .BOM format — use BomParser directly
            try:
                items = BomParser.parse(path)
            except Exception as e:
                messagebox.showerror('Error', f'讀取 .BOM 失敗:\n{e}')
                return
            if not items:
                messagebox.showwarning('Warning', '.BOM 檔案無有效資料')
                return
            for item in items:
                vpn = (item.vendor_pn or '').strip()
                if not vpn:
                    continue
                seen[vpn] = seen.get(vpn, 0) + (item.qty or 1)
        else:
            # Excel / CSV
            try:
                columns, preview, all_rows = self._lib_read_file(path)
            except Exception as e:
                messagebox.showerror('Error', f'讀取失敗:\n{e}')
                return

            dlg_vpn = _VpnColPicker(self.root, columns, preview)
            self.root.wait_window(dlg_vpn)
            if dlg_vpn.result is None:
                return

            dlg_qty = _QtyColPicker(self.root, columns, preview)
            self.root.wait_window(dlg_qty)
            if dlg_qty.result is None:
                return

            vpn_col = dlg_vpn.result
            qty_col = dlg_qty.result

            for row in all_rows:
                vpn = str(row.get(vpn_col, '') or '').strip()
                if not vpn:
                    continue
                try:
                    qty = int(float(str(row.get(qty_col, 1) or 1)))
                except (ValueError, TypeError):
                    qty = 1
                seen[vpn] = seen.get(vpn, 0) + qty

        # Match against power DB
        budget_rows: list[dict] = []
        missing: list[tuple] = []
        for vpn, qty in seen.items():
            part = self.pwr_db.get_part_by_vpn(vpn)
            if part is None:
                missing.append((vpn, qty))
                continue
            rails = self.pwr_db.get_rails(part['id'])
            if not rails:
                missing.append((vpn, qty))
                continue
            for r in rails:
                budget_rows.append({
                    'vendor_pn':    vpn,
                    'rail_name':    r['rail_name'],
                    'power_group':  r['power_group'] or r['rail_name'],
                    'input_rail':   r['input_rail'] or '',
                    'voltage':      r['voltage'],
                    'qty':          qty,
                    'typ_i':        r['typ_current_a'] * qty,
                    'max_i':        r['max_current_a'] * qty,
                    'typ_p':        r['voltage'] * r['typ_current_a'] * qty,
                    'max_p':        r['voltage'] * r['max_current_a'] * qty,
                    'condition':    r['condition'] or '',
                })

        _PwrBudgetWindow(self.root, budget_rows, missing,
                         os.path.basename(path), self._pwr_export_excel)

    # ── Export DB Excel (no BOM needed) ──────────────────────────────────────

    def _pwr_export_db_excel(self):
        """Export checked parts from the power DB to Excel, using per-part qty."""
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return

        # Gather checked parts from the parts treeview
        checked: list[tuple[int, int]] = []  # [(part_id, qty), ...]
        for iid in self.tv_pwr_parts.get_children():
            pid = int(iid)
            sel = self._pwr_part_sel.get(pid, {'checked': False, 'qty': 1})
            if sel['checked']:
                checked.append((pid, sel['qty']))

        if not checked:
            messagebox.showwarning(
                'No Selection',
                '請先在 Parts 清單的 ✓ 欄打勾選擇要 Export 的零件\n'
                '（點一下 ☐ 即可勾選，雙擊 Qty 欄可修改數量）')
            return

        budget_rows: list[dict] = []
        for pid, qty in checked:
            p = self.pwr_db.get_part(pid)
            if p is None:
                continue
            for r in self.pwr_db.get_rails(pid):
                budget_rows.append({
                    'vendor_pn':   p['vendor_pn'],
                    'rail_name':   r['rail_name'],
                    'power_group': r['power_group'] or r['rail_name'],
                    'input_rail':  r['input_rail'] or '',
                    'voltage':     r['voltage'],
                    'qty':         qty,
                    'typ_i':       r['typ_current_a'] * qty,
                    'max_i':       r['max_current_a'] * qty,
                    'typ_p':       r['voltage'] * r['typ_current_a'] * qty,
                    'max_p':       r['voltage'] * r['max_current_a'] * qty,
                    'condition':   r['condition'] or '',
                })

        if not budget_rows:
            messagebox.showwarning('Warning', '選取的零件沒有 Rail 資料')
            return

        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'PowerDB_Export_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        self._pwr_export_excel(budget_rows, [],
                               f'DB Export ({len(checked)} parts)', save_path=path)

    # ── Power Tree (hierarchical current/power roll-up) ────────────────────────

    def _pwr_build_power_tree_data(self, checked):
        """Build a supply tree from checked [(part_id, qty), ...] pairs.

        Nets are matched purely by string equality on rail_name / source_rail
        (there is no real FK between power_rails/power_sources), mirroring the
        matching already used by _pwr_export_excel's Input Rail grouping.

        Returns a list of root net nodes (nets nobody in the selection
        produces, e.g. VBAT). Each node:
            {'name', 'typ_i', 'max_i', 'typ_p', 'max_p',   # this net's total demand
             'consumers': [ {vendor_pn, power_group, qty, voltage,
                              typ_i, max_i, typ_p, max_p, condition}, ... ],
             'children': [ {'source': dict, 'cycle': bool, 'out_node': node|None,
                             'in_typ_i', 'in_max_i', 'in_typ_p', 'in_max_p'}, ... ]}
        A child's out_node['typ_i']/['max_i'] IS "this Power Source's output
        current = sum of all its downstream rail currents". in_typ_i/in_max_i
        is that same current reflected back onto the parent net after dividing
        by efficiency (what this source itself draws upstream).
        """
        from collections import defaultdict

        rail_demand = defaultdict(lambda: {'typ_i': 0.0, 'max_i': 0.0, 'typ_p': 0.0,
                                            'max_p': 0.0, 'consumers': []})
        sources_by_input = defaultdict(list)
        produced_nets = set()

        for pid, qty in checked:
            p = self.pwr_db.get_part(pid)
            if p is None:
                continue
            vendor_pn = p['vendor_pn']
            for r in self.pwr_db.get_rails(pid):
                net = r['rail_name']
                typ_i = r['typ_current_a'] * qty
                max_i = r['max_current_a'] * qty
                typ_p = r['voltage'] * r['typ_current_a'] * qty
                max_p = r['voltage'] * r['max_current_a'] * qty
                d = rail_demand[net]
                d['typ_i'] += typ_i; d['max_i'] += max_i
                d['typ_p'] += typ_p; d['max_p'] += max_p
                d['consumers'].append({
                    'rail_id': r['id'],
                    'vendor_pn': vendor_pn, 'power_group': r['power_group'] or net,
                    'qty': qty, 'voltage': r['voltage'],
                    'typ_i': typ_i, 'max_i': max_i, 'typ_p': typ_p, 'max_p': max_p,
                    'condition': r['condition'] or '',
                })
            for s in self.pwr_db.get_sources(pid):
                if not s['rail_name'] or not s['source_rail']:
                    continue  # incomplete row — not enough info to place in the tree
                src = dict(s)
                src['vendor_pn'] = vendor_pn
                sources_by_input[s['source_rail']].append(src)
                produced_nets.add(s['rail_name'])

        # Independent DC/DC library entries (part_id IS NULL) — available regardless
        # of which parts are checked, same string-matching rules as per-part sources.
        for s in self.pwr_db.get_library_sources():
            if not s['rail_name'] or not s['source_rail']:
                continue
            src = dict(s)
            src['vendor_pn'] = '(library)'
            sources_by_input[s['source_rail']].append(src)
            produced_nets.add(s['rail_name'])

        def build_net_node(net_name, visited):
            d = rail_demand.get(net_name)
            typ_i = d['typ_i'] if d else 0.0
            max_i = d['max_i'] if d else 0.0
            typ_p = d['typ_p'] if d else 0.0
            max_p = d['max_p'] if d else 0.0
            consumers = d['consumers'] if d else []
            children = []
            for src in sources_by_input.get(net_name, []):
                out_net = src['rail_name']
                if out_net in visited:
                    children.append({'source': src, 'out_node': None, 'cycle': True,
                                      'in_typ_i': 0.0, 'in_max_i': 0.0,
                                      'in_typ_p': 0.0, 'in_max_p': 0.0})
                    continue
                out_node = build_net_node(out_net, visited | {net_name})
                eff = src['efficiency'] or 1.0
                in_typ_i = out_node['typ_i'] / eff
                in_max_i = out_node['max_i'] / eff
                in_typ_p = out_node['typ_p'] / eff
                in_max_p = out_node['max_p'] / eff
                typ_i += in_typ_i; max_i += in_max_i
                typ_p += in_typ_p; max_p += in_max_p
                children.append({'source': src, 'out_node': out_node, 'cycle': False,
                                  'in_typ_i': in_typ_i, 'in_max_i': in_max_i,
                                  'in_typ_p': in_typ_p, 'in_max_p': in_max_p})
            return {'name': net_name, 'typ_i': typ_i, 'max_i': max_i,
                    'typ_p': typ_p, 'max_p': max_p,
                    'consumers': consumers, 'children': children}

        all_nets = set(rail_demand) | set(sources_by_input)
        roots = sorted(n for n in all_nets if n not in produced_nets)
        return [build_net_node(root, {root}) for root in roots]

    def _pwr_show_power_tree(self):
        checked: list[tuple[int, int]] = []
        for iid in self.tv_pwr_parts.get_children():
            pid = int(iid)
            sel = self._pwr_part_sel.get(pid, {'checked': False, 'qty': 1})
            if sel['checked']:
                checked.append((pid, sel['qty']))

        if not checked:
            messagebox.showwarning(
                'No Selection',
                '請先在 Parts 清單的 ✓ 欄打勾選擇要建立 Power Tree 的零件\n'
                '（點一下 ☐ 即可勾選，雙擊 Qty 欄可修改數量）')
            return

        roots = self._pwr_build_power_tree_data(checked)
        if not roots:
            messagebox.showwarning('Warning', '選取的零件沒有 Rail 資料')
            return

        _PwrTreeWindow(self.root, self, checked, len(checked))

    # ── Paste JSON / TSV for batch rail input ─────────────────────────────────

    def _pwr_paste_rails(self):
        if self._pwr_part_id is None:
            messagebox.showwarning('Warning', '請先選擇或儲存一個 Part')
            return
        dlg = _PasteRailsDialog(self.root)
        self.root.wait_window(dlg)
        if not dlg.result:
            return
        for rail_data in dlg.result:
            rail_data['part_id'] = self._pwr_part_id
            self.pwr_db.upsert_rail(rail_data)
        self._pwr_load_rails(self._pwr_part_id)
        self._pwr_refresh_stats()
        messagebox.showinfo('Done', f'已新增 {len(dlg.result)} 條 Rail')

    def _pwr_export_excel(self, budget_rows, missing, filename, save_path=None):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        if save_path:
            path = save_path
        else:
            ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
            path = filedialog.asksaveasfilename(
                initialfile=f'Power_Budget_{ts}.xlsx',
                filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
            if not path:
                return

        wb = openpyxl.Workbook()
        hdr_fill  = PatternFill('solid', fgColor='1F6CBF')
        hdr_font  = Font(bold=True, color='FFFFFF')
        right     = Alignment(horizontal='right')
        center    = Alignment(horizontal='center')

        def _make_hdr(ws, headers, col_widths):
            ws.append(headers)
            for cell, w in zip(ws[1], col_widths):
                cell.fill = hdr_fill
                cell.font = hdr_font
                ws.column_dimensions[cell.column_letter].width = w
            ws.freeze_panes = 'A2'

        # ── Summary sheet (cross-tab matrix) ────────────────────────────────────
        from collections import defaultdict

        ws_sum = wb.active
        ws_sum.title = 'Summary'

        # Build column list — grouped by rail_name (PCB net).
        # Multiple power_groups that share the same rail_name are summed into one column.
        power_groups = list(dict.fromkeys(r['rail_name'] for r in budget_rows))
        n_pg = len(power_groups)

        # Look up power_sources for each rail_name (efficiency / type / Vin net)
        src_map: dict[str, dict] = {}
        with self.pwr_db._conn() as _c:
            for _pg in power_groups:
                _s = _c.execute(
                    'SELECT * FROM power_sources WHERE rail_name=? LIMIT 1',
                    (_pg,)).fetchone()
                if _s:
                    src_map[_pg] = {
                        'eff':  _s['efficiency'] or 1.0,
                        'type': _s['regulator_type'] or '',
                        'vin':  _s['source_rail'] or '',
                    }
                else:
                    src_map[_pg] = {'eff': 1.0, 'type': '', 'vin': ''}

        pg_voltage: dict[str, float] = {}
        for r in budget_rows:
            pg_voltage.setdefault(r['rail_name'], r['voltage'])

        # Per-IC aggregations (keyed by rail_name)
        ics_order = list(dict.fromkeys(r['vendor_pn'] for r in budget_rows))
        ic_qty: dict[str, int] = {}
        for r in budget_rows:
            ic_qty[r['vendor_pn']] = r['qty']

        ic_pg_i: dict = defaultdict(lambda: defaultdict(float))
        ic_pg_p: dict = defaultdict(lambda: defaultdict(float))
        for r in budget_rows:
            ic_pg_i[r['vendor_pn']][r['rail_name']] += r['typ_i']
            ic_pg_p[r['vendor_pn']][r['rail_name']] += r['typ_p']

        pg_total_i: dict = defaultdict(float)
        pg_total_p: dict = defaultdict(float)
        for r in budget_rows:
            pg_total_i[r['rail_name']] += r['typ_i']
            pg_total_p[r['rail_name']] += r['typ_p']

        # Per-pg input power and loss
        pg_pin:  dict[str, float] = {}
        pg_loss: dict[str, float] = {}
        for _pg in power_groups:
            _eff   = src_map[_pg]['eff']
            _p_out = pg_total_p[_pg]
            _p_in  = _p_out / _eff if _eff > 0 else _p_out
            pg_pin[_pg]  = _p_in
            pg_loss[_pg] = _p_in - _p_out

        total_pout = sum(pg_total_p[p] for p in power_groups)
        total_pin  = sum(pg_pin[p] for p in power_groups)

        # ── Custom fills / fonts for the header band ───────────────────────────
        fill_eff   = PatternFill('solid', fgColor='00B050')  # green
        fill_vin   = PatternFill('solid', fgColor='FF99CC')  # pink
        fill_vout  = PatternFill('solid', fgColor='7B9CD0')  # blue-purple
        fill_col   = PatternFill('solid', fgColor='1F6CBF')  # blue (column header)
        fill_foot  = PatternFill('solid', fgColor='D9EAD3')  # light-green footer
        fill_grand = PatternFill('solid', fgColor='BDD7EE')  # light-blue grand total
        bold_wht   = Font(bold=True, color='FFFFFF')
        bold       = Font(bold=True)

        def _apply_pg_row(rn, fill, font=None):
            for ci in range(3, n_pg + 3):
                cell = ws_sum.cell(rn, ci)
                cell.fill = fill
                cell.alignment = center
                if font:
                    cell.font = font

        def _num_cell(rn, ci, fmt='0.000'):
            cell = ws_sum.cell(rn, ci)
            if isinstance(cell.value, (int, float)):
                cell.number_format = fmt
                cell.alignment = right

        # Row 1 — PWR Efficiency
        ws_sum.append(
            ['PWR Efficiency', ''] +
            [f"{int(round(src_map[p]['eff'] * 100))}%" for p in power_groups] + [''])
        rn = ws_sum.max_row
        _apply_pg_row(rn, fill_eff, bold_wht)
        ws_sum.cell(rn, 1).font = bold

        # Row 2 — Type
        ws_sum.append(['Type', ''] + [src_map[p]['type'] for p in power_groups] + [''])
        rn = ws_sum.max_row
        ws_sum.cell(rn, 1).font = bold
        for ci in range(3, n_pg + 3):
            ws_sum.cell(rn, ci).alignment = center

        # Row 3 — Vin (V)
        ws_sum.append(['Vin (V)', ''] + [src_map[p]['vin'] for p in power_groups] + [''])
        rn = ws_sum.max_row
        ws_sum.cell(rn, 1).font = bold
        _apply_pg_row(rn, fill_vin)

        # Row 4 — Vout (V)
        ws_sum.append(
            ['Vout (V)', ''] + [pg_voltage.get(p, 0) for p in power_groups] + [''])
        rn = ws_sum.max_row
        ws_sum.cell(rn, 1).font = bold
        _apply_pg_row(rn, fill_vout, bold_wht)
        for ci in range(3, n_pg + 3):
            _num_cell(rn, ci, '0.00')

        # Row 5 — Column headers (Vout Net names)
        ws_sum.append(["Devices", "Q'ty"] + power_groups + ['TDP (W)'])
        rn = ws_sum.max_row
        for ci in range(1, n_pg + 4):
            cell = ws_sum.cell(rn, ci)
            cell.fill = fill_col
            cell.font = bold_wht
            cell.alignment = center
        ws_sum.freeze_panes = ws_sum.cell(rn + 1, 3)

        # ── Data rows ─────────────────────────────────────────────────────────
        for vpn in ics_order:
            row_vals = [vpn, ic_qty[vpn]]
            for p in power_groups:
                iv = ic_pg_i[vpn].get(p)
                row_vals.append(iv if iv else None)
            tdp = sum(ic_pg_p[vpn][p] for p in power_groups)
            row_vals.append(tdp if tdp else None)
            ws_sum.append(row_vals)
            rn = ws_sum.max_row
            ws_sum.cell(rn, 2).alignment = center
            for ci in range(3, n_pg + 4):
                _num_cell(rn, ci, '0.000')

        # ── Footer rows ───────────────────────────────────────────────────────
        def _footer(label, per_pg_vals, tdp_val, fill=fill_foot):
            ws_sum.append([label, ''] + list(per_pg_vals) + [tdp_val])
            rn = ws_sum.max_row
            ws_sum.cell(rn, 1).font = bold
            ws_sum.cell(rn, 1).fill = fill
            ws_sum.cell(rn, 2).fill = fill
            for ci in range(3, n_pg + 4):
                cell = ws_sum.cell(rn, ci)
                cell.fill = fill
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0.000'
                    cell.alignment = right

        _footer('Total current per Voltage [A]',
                [pg_total_i[p] for p in power_groups], '')
        _footer('Total power per Voltage [W]',
                [pg_total_p[p] for p in power_groups], total_pout)
        _footer('Reg. Loss [W]',
                [pg_loss[p] for p in power_groups], sum(pg_loss.values()))
        _footer('Total Power Input[W]',
                [pg_pin[p] for p in power_groups], total_pin)
        _footer('12Vin current [A]',
                [pg_pin[p] / 12.0 for p in power_groups], total_pin / 12.0)

        def _grand(label, tdp_val):
            ws_sum.append([label] + [''] * (n_pg + 1) + [tdp_val])
            rn = ws_sum.max_row
            for ci in range(1, n_pg + 4):
                cell = ws_sum.cell(rn, ci)
                cell.fill = fill_grand
                cell.font = bold
            cell = ws_sum.cell(rn, n_pg + 3)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '0.000'
                cell.alignment = right

        _grand('Total power consumption [W]', total_pout)
        _grand('Total power consumption (Power optimization 90% included) [W]', total_pin)

        # ── Input Rail Subtotal section ────────────────────────────────────────
        # Group by input_rail (PCB supply net). If empty, fall back to voltage label.
        from collections import defaultdict as _dd
        v_grp: dict = _dd(lambda: {'typ_i': 0.0, 'max_i': 0.0,
                                    'typ_p': 0.0, 'max_p': 0.0, 'volt': 0.0})
        for r in budget_rows:
            vk = r['input_rail'] if r['input_rail'] else f"{r['voltage']:.3f}V"
            v_grp[vk]['typ_i'] += r['typ_i']
            v_grp[vk]['max_i'] += r['max_i']
            v_grp[vk]['typ_p'] += r['typ_p']
            v_grp[vk]['max_p'] += r['max_p']
            if not v_grp[vk]['volt']:
                v_grp[vk]['volt'] = r['voltage']

        fill_vsub  = PatternFill('solid', fgColor='FFF2CC')   # light yellow
        fill_vhdr  = PatternFill('solid', fgColor='F4B942')   # orange header
        font_vsub  = Font(bold=True)

        # blank separator row
        ws_sum.append([])

        # header row for input-rail subtable
        vsub_cols = ['Input Rail', 'Total Typ I (A)', 'Total Max I (A)',
                     'Total Typ P (W)', 'Total Max P (W)']
        ws_sum.append(vsub_cols + [''] * (n_pg - 2))
        rn = ws_sum.max_row
        for ci in range(1, len(vsub_cols) + 1):
            cell = ws_sum.cell(rn, ci)
            cell.fill = fill_vhdr
            cell.font = Font(bold=True, color='FFFFFF')
            cell.alignment = center

        for vk in sorted(v_grp):
            d = v_grp[vk]
            ws_sum.append([vk, d['typ_i'], d['max_i'], d['typ_p'], d['max_p']])
            rn = ws_sum.max_row
            for ci in range(1, 6):
                cell = ws_sum.cell(rn, ci)
                cell.fill = fill_vsub
                cell.font = font_vsub
                if ci > 1:
                    cell.number_format = '0.000'
                    cell.alignment = right
                else:
                    cell.number_format = '0.000'
                    cell.alignment = center

        # grand-total voltage row
        ws_sum.append(['TOTAL',
                       sum(v_grp[v]['typ_i'] for v in v_grp),
                       sum(v_grp[v]['max_i'] for v in v_grp),
                       sum(v_grp[v]['typ_p'] for v in v_grp),
                       sum(v_grp[v]['max_p'] for v in v_grp)])
        rn = ws_sum.max_row
        for ci in range(1, 6):
            cell = ws_sum.cell(rn, ci)
            cell.fill = fill_grand
            cell.font = Font(bold=True)
            if ci > 1:
                cell.number_format = '0.000'
                cell.alignment = right

        # ── Column widths ─────────────────────────────────────────────────────
        ws_sum.column_dimensions['A'].width = 46
        ws_sum.column_dimensions['B'].width = 6
        for ci in range(3, n_pg + 3):
            ws_sum.column_dimensions[
                ws_sum.cell(1, ci).column_letter].width = 15
        ws_sum.column_dimensions[
            ws_sum.cell(1, n_pg + 3).column_letter].width = 12

        # ── Per-IC sheets ─────────────────────────────────────────────────────
        ic_groups: dict[str, list] = defaultdict(list)
        for r in budget_rows:
            ic_groups[r['vendor_pn']].append(r)

        for vpn, rows in ic_groups.items():
            ws_ic = wb.create_sheet(title=vpn[:31])
            _make_hdr(ws_ic,
                      ['Vendor PN', 'Rail Name', 'Power Group', 'Voltage (V)', 'Qty',
                       'Typ I (A)', 'Max I (A)', 'Typ P (W)', 'Max P (W)', 'Condition'],
                      [18, 14, 16, 12, 6, 12, 12, 12, 12, 18])
            for r in rows:
                ws_ic.append([r['vendor_pn'], r['rail_name'], r['power_group'],
                               r['voltage'], r['qty'], r['typ_i'], r['max_i'],
                               r['typ_p'], r['max_p'], r['condition']])
                xr = ws_ic.max_row
                for col in range(6, 10):
                    ws_ic.cell(xr, col).number_format = '0.0000'
                    ws_ic.cell(xr, col).alignment = right

        # ── Missing Parts sheet ───────────────────────────────────────────────
        if missing:
            ws_miss = wb.create_sheet(title='Missing Parts')
            _make_hdr(ws_miss, ['Vendor PN', 'Qty'], [24, 8])
            for vpn, qty in missing:
                ws_miss.append([vpn, qty])

        # ── Title row ─────────────────────────────────────────────────────────
        ws_sum.insert_rows(1)
        ws_sum['A1'] = (f'BOM: {filename}  |  '
                        f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        ws_sum['A1'].font = Font(italic=True, color='555555')
        last_col_letter = ws_sum.cell(1, n_pg + 3).column_letter
        ws_sum.merge_cells(f'A1:{last_col_letter}1')

        wb.save(path)
        messagebox.showinfo('Done', f'已儲存至\n{path}')


    # ── Tab: Reminders ────────────────────────────────────────────────────────

    def _build_todo_tab(self):
        f = self.tab_todo
        f.rowconfigure(0, weight=1)
        f.columnconfigure(1, weight=1)

        # ── Left sidebar ──────────────────────────────────────────────────────
        sidebar = tk.Frame(f, bg='#EFEFF4', width=220)
        sidebar.grid(row=0, column=0, sticky='nsew')
        sidebar.grid_propagate(False)
        sidebar.rowconfigure(1, weight=1)
        sidebar.columnconfigure(0, weight=1)

        search_f = tk.Frame(sidebar, bg='#EFEFF4')
        search_f.pack(fill='x', padx=8, pady=(10, 6))
        self._todo_search_sv = tk.StringVar()
        se = ttk.Entry(search_f, textvariable=self._todo_search_sv)
        se.pack(side='left', fill='x', expand=True)
        se.bind('<Return>', lambda e: self._todo_search())
        ttk.Button(search_f, text='🔍', width=3,
                   command=self._todo_search).pack(side='left', padx=(2, 0))

        self._todo_today_row = tk.Frame(sidebar, bg='#EFEFF4', cursor='hand2')
        self._todo_today_row.pack(fill='x', padx=6, pady=(10, 4))
        today_dot = tk.Canvas(self._todo_today_row, width=16, height=16,
                              bg='#EFEFF4', bd=0, highlightthickness=0)
        today_dot.create_oval(1, 1, 15, 15, fill='#FF9500', outline='')
        today_dot.pack(side='left', padx=(8, 6), pady=6)
        self._todo_today_lbl = tk.Label(self._todo_today_row, text='⏰ Today',
                                        bg='#EFEFF4', fg='#1C1C1E', font=FONT_UI, anchor='w')
        self._todo_today_lbl.pack(side='left', fill='x', expand=True)
        for w in (self._todo_today_row, today_dot, self._todo_today_lbl):
            w.bind('<Button-1>', lambda e: self._todo_select_today())

        tk.Label(sidebar, text='MY LISTS', bg='#EFEFF4',
                 fg='#8E8E93',
                 font=('Microsoft JhengHei UI', 8, 'bold')).pack(
            anchor='w', padx=14, pady=(14, 4))

        self._todo_list_frame = tk.Frame(sidebar, bg='#EFEFF4')
        self._todo_list_frame.pack(fill='both', expand=True, padx=6)

        tk.Frame(sidebar, bg='#D1D1D6', height=1).pack(fill='x', padx=10, pady=(0, 2))

        add_btn = tk.Label(sidebar, text='+ Add List', bg='#EFEFF4',
                           fg='#007AFF', font=FONT_UI, cursor='hand2')
        add_btn.pack(anchor='w', padx=14, pady=(8, 2))
        add_btn.bind('<Button-1>', lambda e: self._todo_add_list())

        proj_btn = tk.Label(sidebar, text='🗁 New Project...', bg='#EFEFF4',
                            fg='#34C759', font=FONT_UI, cursor='hand2')
        proj_btn.pack(anchor='w', padx=14, pady=(2, 8))
        proj_btn.bind('<Button-1>', lambda e: self._todo_new_project())

        # 1-px divider
        tk.Frame(f, bg='#D1D1D6', width=1).grid(row=0, column=0, sticky='nse')

        # ── Right main area ───────────────────────────────────────────────────
        main = tk.Frame(f, bg=_TODO_BG)
        main.grid(row=0, column=1, sticky='nsew')
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        # Header
        hdr = tk.Frame(main, bg=_TODO_BG)
        hdr.grid(row=0, column=0, sticky='ew', padx=20, pady=(16, 8))

        self._todo_hdr_lbl = tk.Label(
            hdr, text='Select a list', bg=_TODO_BG,
            fg='#1C1C1E',
            font=('Microsoft JhengHei UI', 20, 'bold'))
        self._todo_hdr_lbl.pack(side='left')

        self._todo_count_lbl = tk.Label(
            hdr, text='', bg=_TODO_BG, fg='#8E8E93',
            font=('Microsoft JhengHei UI', 11))
        self._todo_count_lbl.pack(side='left', padx=8)

        self._todo_export_btn = ttk.Button(hdr, text='📊 Export Report',
                                           command=self._todo_export_report)
        self._todo_export_btn.pack(side='left', padx=(4, 0))

        self._todo_comp_var = tk.BooleanVar(value=True)
        self._todo_comp_cb = ttk.Checkbutton(hdr, text='Show Completed',
                        variable=self._todo_comp_var,
                        command=self._todo_refresh_items)
        self._todo_comp_cb.pack(side='right')

        # View toggle (List / Month / Week) — hidden in the Today rollup
        self._todo_view_var = tk.StringVar(value='list')
        self._todo_view_radios = []
        for txt, val in [('List', 'list'), ('Month', 'month'), ('Week', 'week')]:
            rb = ttk.Radiobutton(hdr, text=txt, variable=self._todo_view_var,
                                  value=val, command=self._todo_switch_view)
            rb.pack(side='right')
            self._todo_view_radios.append(rb)

        # ── List view: Treeview ───────────────────────────────────────────────
        main.columnconfigure(0, weight=1)
        cols = ('status', 'title', 'list', 'attach', 'progress', 'due', 'priority', 'created')
        self._todo_tv = ttk.Treeview(main, columns=cols, show='tree headings',
                                      selectmode='browse')
        self._todo_tv.heading('#0',      text='')
        self._todo_tv.heading('status',  text='')
        self._todo_tv.heading('title',   text='Reminder')
        self._todo_tv.heading('list',    text='List')
        self._todo_tv.heading('attach',  text='')
        self._todo_tv.heading('progress', text='Progress')
        self._todo_tv.heading('due',     text='Due')
        self._todo_tv.heading('priority', text='')
        self._todo_tv.heading('created', text='Created')
        self._todo_tv.column('#0',       width=32, stretch=False, anchor='w')
        self._todo_tv.column('status',   width=28, stretch=False, anchor='center')
        self._todo_tv.column('title',    width=280, stretch=True,  anchor='w')
        self._todo_tv.column('list',     width=0, stretch=False, anchor='w')
        self._todo_tv.column('attach',   width=24, stretch=False, anchor='center')
        self._todo_tv.column('progress', width=76, stretch=False, anchor='center')
        self._todo_tv.column('due',      width=80, stretch=False, anchor='center')
        self._todo_tv.column('priority', width=36, stretch=False, anchor='center')
        self._todo_tv.column('created',  width=100, stretch=False, anchor='center')
        self._todo_tv.tag_configure('done', foreground='#A0A0A5')
        self._todo_tv.tag_configure('sub', foreground='#6E6E73')
        self._todo_tv.tag_configure('subdone', foreground='#C7C7CC')
        self._todo_tv_vsb = ttk.Scrollbar(main, orient='vertical',
                                           command=self._todo_tv.yview)
        self._todo_tv.configure(yscrollcommand=self._todo_tv_vsb.set)
        self._todo_tv.grid(row=1, column=0, sticky='nsew', padx=(16, 0), pady=(0, 6))
        self._todo_tv_vsb.grid(row=1, column=1, sticky='ns', padx=(0, 16), pady=(0, 6))
        self._todo_tv.bind('<Button-1>', self._todo_tv_click)
        self._todo_tv.bind('<Double-1>', self._todo_tv_dbl_click)
        self._todo_tv.bind('<Button-3>', self._todo_tv_ctx)

        # ── Calendar view (hidden initially) ─────────────────────────────────
        self._todo_cal_outer = tk.Frame(main, bg=_TODO_BG)
        self._todo_cal_outer.grid(row=1, column=0, columnspan=2,
                                   sticky='nsew', padx=16, pady=(0, 6))
        self._todo_cal_outer.rowconfigure(1, weight=1)
        self._todo_cal_outer.columnconfigure(0, weight=1)
        self._build_todo_cal()
        self._todo_cal_outer.grid_remove()   # hidden until user selects Month/Week

        # ── New Reminder entry bar ────────────────────────────────────────────
        entry_bar = tk.Frame(main, bg=_TODO_BG)
        entry_bar.grid(row=2, column=0, columnspan=2, sticky='ew',
                       padx=16, pady=(0, 16))
        entry_bar.columnconfigure(0, weight=1)
        self._todo_entry = ttk.Entry(entry_bar, font=FONT_UI)
        self._todo_entry.grid(row=0, column=0, sticky='ew', ipady=4)
        self._todo_entry.bind('<Return>', lambda e: self._todo_add_item())
        self._todo_add_btn = ttk.Button(entry_bar, text='Add', command=self._todo_add_item,
                                        width=6)
        self._todo_add_btn.grid(row=0, column=1, padx=(6, 0))

        self._todo_refresh_lists()
        self._todo_start_alert()

    def _todo_tv_click(self, event):
        """Single click on the status column toggles complete (item or sub-item)."""
        if self._todo_tv.identify_region(event.x, event.y) != 'cell':
            return
        if self._todo_tv.identify_column(event.x) != '#1':
            return
        iid = self._todo_tv.identify_row(event.y)
        if not iid:
            return
        if iid.startswith('s'):
            self.todo_db.toggle_subitem(int(iid[1:]))
        else:
            self.todo_db.toggle(int(iid))
        self._todo_refresh_current()

    def _todo_tv_dbl_click(self, event):
        iid = self._todo_tv.identify_row(event.y)
        if not iid:
            return
        if iid.startswith('s'):
            self._todo_sub_edit_dialog(int(iid[1:]))
        else:
            self._todo_edit_dialog(int(iid))

    def _todo_tv_ctx(self, event):
        iid = self._todo_tv.identify_row(event.y)
        if not iid:
            return
        self._todo_tv.selection_set(iid)
        if iid.startswith('s'):
            self._todo_sub_ctx_menu(event, int(iid[1:]))
        else:
            self._todo_ctx_menu(event, int(iid))

    # ── Sidebar ───────────────────────────────────────────────────────────────

    def _todo_refresh_lists(self):
        for w in self._todo_list_frame.winfo_children():
            w.destroy()
        for lst in self.todo_db.get_lists():
            self._todo_make_list_row(lst)

    def _todo_make_list_row(self, lst):
        is_sel = (lst['id'] == self._todo_list_id)
        bg     = '#D1D1D6' if is_sel else '#EFEFF4'

        row = tk.Frame(self._todo_list_frame, bg=bg, cursor='hand2')
        row.pack(fill='x', pady=1)

        dot = tk.Canvas(row, width=16, height=16, bg=bg, bd=0, highlightthickness=0)
        dot.create_oval(1, 1, 15, 15, fill=lst['color'], outline='')
        dot.pack(side='left', padx=(8, 6), pady=8)

        tk.Label(row, text=lst['name'], bg=bg, fg='#1C1C1E',
                 font=FONT_UI, anchor='w').pack(side='left', fill='x', expand=True)

        count = lst['active_count'] or 0
        if count:
            tk.Label(row, text=str(count), bg='#8E8E93', fg='white',
                     font=('Microsoft JhengHei UI', 8, 'bold'),
                     padx=5, pady=1).pack(side='right', padx=(4, 8), pady=6)

        frac, total = self.todo_db.get_list_progress(lst['id'])
        if total:
            filled = round(frac * 5)
            bar = '■' * filled + '□' * (5 - filled)
            tk.Label(row, text=bar, bg=bg, fg='#8E8E93',
                     font=('Courier New', 8)).pack(side='right', padx=(4, 0))

        cb = lambda e, lid=lst['id'], col=lst['color'], nm=lst['name']: \
            self._todo_select_list(lid, col, nm)
        row.bind('<Button-1>', cb)
        for child in row.winfo_children():
            child.bind('<Button-1>', cb)

        row.bind('<Button-3>', lambda e, lid=lst['id'], nm=lst['name'], col=lst['color']:
                 self._todo_list_ctx(e, lid, nm, col))

    def _todo_list_ctx(self, event, list_id, name, color):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label='Rename...',
                      command=lambda: self._todo_rename_list_dialog(list_id, name, color))
        m.add_separator()
        m.add_command(label=f'Delete "{name}"', foreground='#FF3B30',
                      command=lambda: self._todo_delete_list(list_id))
        m.tk_popup(event.x_root, event.y_root)

    def _todo_rename_list_dialog(self, list_id, name, color):
        dlg = _TodoListDialog(self.root, existing={'name': name, 'color': color})
        self.root.wait_window(dlg)
        if not dlg.result:
            return
        new_name, new_color = dlg.result
        try:
            self.todo_db.rename_list(list_id, new_name, new_color)
        except sqlite3.IntegrityError:
            messagebox.showwarning('Rename List',
                                   f'A list named "{new_name}" already exists.')
            return
        if self._todo_list_id == list_id:
            self._todo_list_col = new_color
            self._todo_hdr_lbl.configure(text=new_name)
        self._todo_refresh_lists()

    def _todo_delete_list(self, list_id):
        if not messagebox.askyesno('Delete List',
                'Delete this list and all its reminders?'):
            return
        self.todo_db.delete_list(list_id)
        if self._todo_list_id == list_id:
            self._todo_list_id  = None
            self._todo_list_col = '#007AFF'
            self._todo_hdr_lbl.configure(text='Select a list')
            self._todo_count_lbl.configure(text='')
            for w in self._todo_items_frm.winfo_children():
                w.destroy()
        self._todo_refresh_lists()

    def _todo_select_list(self, list_id, color, name=''):
        self._todo_list_id  = list_id
        self._todo_list_col = color
        self._todo_hdr_lbl.configure(text=name or 'Select a list')
        self._todo_set_today_highlight(False)
        self._todo_export_btn.pack(side='left', padx=(4, 0))
        self._todo_comp_cb.pack(side='right')
        for rb in self._todo_view_radios:
            rb.pack(side='right')
        self._todo_entry.configure(state='normal')
        self._todo_add_btn.configure(state='normal')
        self._todo_tv.column('list', width=0)
        self._todo_refresh_lists()
        self._todo_refresh_items()

    def _todo_select_today(self):
        self._todo_list_id  = 'TODAY'
        self._todo_list_col = '#FF9500'
        self._todo_set_today_highlight(True)
        self._todo_enter_flat_mode('Today')
        self._todo_refresh_lists()
        self._todo_refresh_today()

    def _todo_enter_flat_mode(self, header_text):
        """Shared setup for read-only cross-list rollup views (Today, Search):
        disables the add-bar and per-list view controls, widens the List column."""
        self._todo_hdr_lbl.configure(text=header_text)
        self._todo_export_btn.pack_forget()
        self._todo_comp_cb.pack_forget()
        for rb in self._todo_view_radios:
            rb.pack_forget()
        self._todo_entry.configure(state='disabled')
        self._todo_add_btn.configure(state='disabled')
        if self._todo_view_var.get() != 'list':
            self._todo_view_var.set('list')
            self._todo_cal_outer.grid_remove()
            self._todo_tv.grid()
            self._todo_tv_vsb.grid()
        self._todo_tv.column('list', width=90)

    def _todo_set_today_highlight(self, selected):
        bg = '#D1D1D6' if selected else '#EFEFF4'
        self._todo_today_row.configure(bg=bg)
        for w in self._todo_today_row.winfo_children():
            w.configure(bg=bg)

    # ── Items area ────────────────────────────────────────────────────────────

    def _todo_refresh_current(self):
        """Dispatch to the per-list view or the cross-list Today rollup,
        whichever is currently selected — shared by every mutation handler
        (toggle, edit, delete, due-date, priority, attachments...) so they
        refresh the right view regardless of which one triggered them."""
        if self._todo_list_id == 'TODAY':
            self._todo_refresh_today()
        elif self._todo_list_id == 'SEARCH':
            self._todo_refresh_search(self._todo_search_sv.get().strip())
        else:
            self._todo_refresh_items()

    @staticmethod
    def _todo_created(ts):
        try:
            return ts[:16] if ts else ''
        except Exception:
            return ''

    def _todo_insert_subitems(self, item_id, list_col=''):
        subs = self.todo_db.get_subitems(item_id)
        for i, s in enumerate(subs, 1):
            sdue   = self._todo_fmt_due(s['due_date'])[0] if s['due_date'] else ''
            attach = '📎' if self.todo_db.get_attachments(subitem_id=s['id']) else ''
            self._todo_tv.insert(str(item_id), 'end', iid=f"s{s['id']}",
                                  values=('✓' if s['completed'] else '○',
                                          f"      {i}. {s['title'] or ''}",
                                          list_col, attach, '', sdue, '', ''),
                                  tags=('subdone',) if s['completed'] else ('sub',))
        return subs

    def _todo_refresh_items(self):
        if self._todo_list_id is None or self._todo_list_id in ('TODAY', 'SEARCH'):
            return
        for iid in self._todo_tv.get_children():
            self._todo_tv.delete(iid)

        show_comp = self._todo_comp_var.get()
        items     = self.todo_db.get_items(self._todo_list_id, show_comp)
        active    = [it for it in items if not it['completed']]
        done_list = [it for it in items if     it['completed']]

        count = len(active)
        self._todo_count_lbl.configure(
            text=f'{count} remaining' if count else '')

        marks = {'high': '!!!', 'medium': '!!', 'low': '!'}

        def _insert_reminder(it, status, extra_tags):
            due    = self._todo_fmt_due(it['due_date'])[0] if it['due_date'] else ''
            title  = it['title'] or ''
            attach = '📎' if self.todo_db.get_attachments(it['id']) else ''
            self._todo_tv.insert('', 'end', iid=str(it['id']),
                                  values=(status, title, '', attach,
                                          self._todo_progress_bar(1 if it['completed'] else 0, 1),
                                          due, marks.get(it['priority'] or '', ''),
                                          self._todo_created(it['created_at'])),
                                  tags=extra_tags)
            subs = self._todo_insert_subitems(it['id'])
            if subs:
                sdone = sum(1 for s in subs if s['completed'])
                self._todo_tv.item(str(it['id']),
                                    values=(status, f'{title}  ({sdone}/{len(subs)})', '', attach,
                                            self._todo_progress_bar(sdone, len(subs)),
                                            due, marks.get(it['priority'] or '', ''),
                                            self._todo_created(it['created_at'])),
                                    open=True, tags=extra_tags)

        for it in active:
            _insert_reminder(it, '○', ())

        if done_list and show_comp:
            for it in done_list:
                _insert_reminder(it, '✓', ('done',))

        self._todo_refresh_lists()

    def _todo_refresh_today(self):
        for iid in self._todo_tv.get_children():
            self._todo_tv.delete(iid)

        from datetime import date
        today_str = date.today().isoformat()
        item_rows, sub_rows = self.todo_db.get_due_rollup(today_str)

        n_due = len(item_rows) + sum(1 for s in sub_rows if s['item_id'] not in
                                      {it['id'] for it in item_rows})
        self._todo_count_lbl.configure(text=f'{n_due} due' if n_due else 'Nothing due')

        marks = {'high': '!!!', 'medium': '!!', 'low': '!'}
        seen_item_ids = set()
        for it in item_rows:
            due    = self._todo_fmt_due(it['due_date'])[0]
            attach = '📎' if self.todo_db.get_attachments(it['id']) else ''
            title  = it['title'] or ''
            subs   = self._todo_insert_subitems(it['id'], list_col=it['list_name'])
            progress = self._todo_progress_bar(
                sum(1 for s in subs if s['completed']), len(subs)) if subs \
                else self._todo_progress_bar(0, 1)
            if subs:
                sdone = sum(1 for s in subs if s['completed'])
                title = f'{title}  ({sdone}/{len(subs)})'
            self._todo_tv.insert('', 'end', iid=str(it['id']),
                                  values=('○', title, it['list_name'], attach, progress,
                                          due, marks.get(it['priority'] or '', ''), ''),
                                  open=bool(subs))
            seen_item_ids.add(it['id'])

        for s in sub_rows:
            if s['item_id'] in seen_item_ids:
                continue   # already nested under its parent item above
            sdue = self._todo_fmt_due(s['due_date'])[0]
            self._todo_tv.insert('', 'end', iid=f"s{s['id']}",
                                  values=('○', f"{s['item_title']} › {s['title']}",
                                          s['list_name'], '', '', sdue, '', ''))

        self._todo_refresh_lists()

    def _todo_search(self):
        query = self._todo_search_sv.get().strip()
        if not query:
            return
        self._todo_list_id  = 'SEARCH'
        self._todo_list_col = '#8E8E93'
        self._todo_set_today_highlight(False)
        self._todo_enter_flat_mode(f'Search: "{query}"')
        self._todo_refresh_lists()
        self._todo_refresh_search(query)

    def _todo_refresh_search(self, query):
        for iid in self._todo_tv.get_children():
            self._todo_tv.delete(iid)
        if not query:
            self._todo_count_lbl.configure(text='')
            return

        items, subs = self.todo_db.search_items(query)
        n = len(items) + sum(1 for s in subs if s['item_id'] not in
                              {it['id'] for it in items})
        self._todo_count_lbl.configure(text=f'{n} result{"s" if n != 1 else ""}' if n else 'No results')

        marks = {'high': '!!!', 'medium': '!!', 'low': '!'}
        seen_item_ids = set()
        for it in items:
            due    = self._todo_fmt_due(it['due_date'])[0] if it['due_date'] else ''
            attach = '📎' if self.todo_db.get_attachments(it['id']) else ''
            title  = it['title'] or ''
            item_subs = self._todo_insert_subitems(it['id'], list_col=it['list_name'])
            progress = self._todo_progress_bar(
                sum(1 for s in item_subs if s['completed']), len(item_subs)) if item_subs \
                else self._todo_progress_bar(1 if it['completed'] else 0, 1)
            if item_subs:
                sdone = sum(1 for s in item_subs if s['completed'])
                title = f'{title}  ({sdone}/{len(item_subs)})'
            status = '✓' if it['completed'] else '○'
            self._todo_tv.insert('', 'end', iid=str(it['id']),
                                  values=(status, title, it['list_name'], attach, progress,
                                          due, marks.get(it['priority'] or '', ''), ''),
                                  open=bool(item_subs),
                                  tags=('done',) if it['completed'] else ())
            seen_item_ids.add(it['id'])

        for s in subs:
            if s['item_id'] in seen_item_ids:
                continue   # already nested under its parent item above
            sdue = self._todo_fmt_due(s['due_date'])[0] if s['due_date'] else ''
            self._todo_tv.insert('', 'end', iid=f"s{s['id']}",
                                  values=('✓' if s['completed'] else '○',
                                          f"{s['item_title']} › {s['title']}",
                                          s['list_name'], '', '', sdue, '', ''))

        self._todo_refresh_lists()

    def _todo_export_report(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        if not isinstance(self._todo_list_id, int):
            messagebox.showwarning('Export Report', '請先選擇一個清單')
            return

        list_id   = self._todo_list_id
        list_name = self._todo_hdr_lbl.cget('text')
        items     = self.todo_db.get_items(list_id, show_completed=True)
        frac, total = self.todo_db.get_list_progress(list_id)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Progress Report'

        ws.append([f'{list_name}  進度報告'])
        ws['A1'].font = Font(bold=True, size=14)
        ws.append([f'匯出時間: {datetime.now().strftime("%Y-%m-%d %H:%M")}'])
        ws.append([f'整體完成度: {round(frac * 100)}%  （{total} 個項目）'])
        ws.append([])

        headers = ['No', 'Status', 'Reminder', 'Due', 'Priority', 'Progress', 'Checklist']
        widths  = [5, 8, 40, 12, 10, 12, 60]
        hdr_row = ws.max_row + 1
        ws.append(headers)
        hdr_fill = PatternFill('solid', fgColor='1F6CBF')
        hdr_font = Font(bold=True, color='FFFFFF')
        for cell, w in zip(ws[hdr_row], widths):
            cell.fill = hdr_fill
            cell.font = hdr_font
            ws.column_dimensions[cell.column_letter].width = w
        ws.freeze_panes = f'A{hdr_row + 1}'

        for i, it in enumerate(items, 1):
            subs = self.todo_db.get_subitems(it['id'])
            if subs:
                sdone   = sum(1 for s in subs if s['completed'])
                pct     = f'{round(sdone / len(subs) * 100)}%'
                checklist = '; '.join(
                    f"{'✓' if s['completed'] else '○'} {s['title']}" for s in subs)
            else:
                pct = '100%' if it['completed'] else '0%'
                checklist = ''
            ws.append([i, '✓' if it['completed'] else '○', it['title'] or '',
                      it['due_date'] or '', it['priority'] or '', pct, checklist])

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_name = re.sub(r'[\\/:*?"<>|]', '_', list_name)
        path = filedialog.asksaveasfilename(
            title='Export Progress Report',
            initialfile=f'{safe_name}_progress_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        try:
            wb.save(path)
            messagebox.showinfo('Done', f'已匯出至\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _todo_make_item_row(self, item):
        done    = bool(item['completed'])
        color   = self._todo_list_col
        item_id = item['id']
        CB      = 22

        row = tk.Frame(self._todo_items_frm, bg='white',
                        highlightbackground='#EBEBEB',
                        highlightthickness=1)
        row.pack(fill='x', pady=2)
        row.columnconfigure(1, weight=1)

        # Circle checkbox
        cb_cv = tk.Canvas(row, width=CB, height=CB,
                           bg='white', bd=0, highlightthickness=0, cursor='hand2')
        cb_cv.grid(row=0, column=0, rowspan=2, padx=(12, 8), pady=10, sticky='n')

        def _draw(cv, is_done):
            cv.delete('all')
            if is_done:
                cv.create_oval(1, 1, CB-1, CB-1, fill=color, outline=color)
                cv.create_line(5, CB//2, 9, CB-6, fill='white', width=2,
                               capstyle='round', joinstyle='round')
                cv.create_line(9, CB-6, CB-4, 5, fill='white', width=2,
                               capstyle='round', joinstyle='round')
            else:
                cv.create_oval(1, 1, CB-1, CB-1, fill='', outline='#C7C7CC', width=1.5)

        _draw(cb_cv, done)
        cb_cv.bind('<Button-1>', lambda e, iid=item_id:
                   (self.todo_db.toggle(iid), self._todo_refresh_items()))

        # Title
        t_font = (FONT_UI[0], FONT_UI[1], 'overstrike') if done else FONT_UI
        t_fg   = '#A0A0A5' if done else '#1C1C1E'
        title_lbl = tk.Label(row, text=item['title'] or '',
                              bg='white', fg=t_fg, font=t_font, anchor='w')
        title_lbl.grid(row=0, column=1, sticky='ew', pady=(9, 1))

        # Sub-row tags
        tags = tk.Frame(row, bg='white')
        tags.grid(row=1, column=1, sticky='ew', pady=(0, 8))

        notes = (item['notes'] or '').strip()
        if notes:
            preview = notes.splitlines()[0][:60]
            tk.Label(tags, text=preview, bg='white', fg='#A0A0A5',
                     font=(FONT_UI[0], 8)).pack(side='left', padx=(0, 6))

        due = item['due_date'] or ''
        if due:
            lbl, fg = self._todo_fmt_due(due)
            tk.Label(tags, text=lbl, bg='#F2F2F7', fg=fg,
                     font=(FONT_UI[0], 8), padx=5, pady=1).pack(side='left', padx=(0, 4))

        prio = item['priority'] or ''
        if prio in _PRIO_CLR:
            marks = {'high': '!!!', 'medium': '!!', 'low': '!'}[prio]
            tk.Label(tags, text=marks, bg='white', fg=_PRIO_CLR[prio],
                     font=(FONT_UI[0], FONT_UI[1], 'bold')).pack(side='right', padx=(0, 12))

        # Hover
        row.bind('<Enter>', lambda e, r=row: r.configure(highlightbackground='#D1D1D6'))
        row.bind('<Leave>', lambda e, r=row: r.configure(highlightbackground='#EBEBEB'))

        # Double-click to edit
        title_lbl.bind('<Double-1>', lambda e, iid=item_id: self._todo_edit_dialog(iid))

        # Right-click context
        def _ctx(e, iid=item_id):
            self._todo_ctx_menu(e, iid)
        for w in (row, cb_cv, title_lbl, tags):
            w.bind('<Button-3>', _ctx)
            w.bind('<MouseWheel>',
                   lambda e: self._todo_canvas.yview_scroll(-1*(e.delta//120), 'units'))

    @staticmethod
    def _todo_fmt_due(due_str):
        from datetime import date as _date
        try:
            due   = _date.fromisoformat(due_str)
            delta = (due - _date.today()).days
            if delta < 0:
                return f'Overdue {due.strftime("%m/%d")}', '#FF3B30'
            if delta == 0:
                return 'Today',    '#FF9500'
            if delta == 1:
                return 'Tomorrow', '#FF9500'
            if delta < 7:
                return due.strftime('%a'), '#8E8E93'
            return due.strftime('%m/%d'), '#8E8E93'
        except (ValueError, TypeError):
            return due_str, '#8E8E93'

    @staticmethod
    def _todo_progress_bar(done_n, total_n):
        filled = round(done_n / total_n * 5) if total_n else 0
        return '■' * filled + '□' * (5 - filled)

    # ── Actions ───────────────────────────────────────────────────────────────

    def _todo_add_item(self):
        title = self._todo_entry.get().strip()
        if not title:
            self._todo_entry.focus_set()
            return
        if self._todo_list_id is None or self._todo_list_id in ('TODAY', 'SEARCH'):
            messagebox.showwarning('Reminders', 'Please select a list first')
            return
        self.todo_db.add_item(self._todo_list_id, title)
        self._todo_entry.delete(0, 'end')
        self._todo_entry.focus_set()
        self._todo_refresh_items()

    def _todo_add_list(self):
        dlg = _TodoListDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            name, color = dlg.result
            list_id = self.todo_db.add_list(name, color)
            self._todo_select_list(list_id, color, name)

    def _todo_new_project(self):
        dlg = _NewProjectDialog(self.root)
        self.root.wait_window(dlg)

    def _todo_edit_dialog(self, item_id):
        with self.todo_db._conn() as c:
            item = c.execute('SELECT * FROM todo_items WHERE id=?',
                             (item_id,)).fetchone()
        if not item:
            return
        dlg = _TodoEditDialog(self.root, dict(item))
        self.root.wait_window(dlg)
        if dlg.result:
            self.todo_db.update(item_id, **dlg.result)
            self._todo_refresh_current()

    def _open_file(self, path):
        import os
        if not os.path.exists(path):
            messagebox.showwarning('File Not Found',
                                   f'File not found:\n{path}\n\n'
                                   'It may have been moved or deleted.')
            return
        os.startfile(path)

    def _todo_attach_file(self, item_id):
        paths = _SmartFileBrowser.ask_files(self.root, 'Attach File(s) to Reminder')
        for p in paths:
            self.todo_db.add_attachment(item_id, p)
        if paths:
            self._todo_refresh_current()

    def _todo_ctx_menu(self, event, item_id):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label='Toggle Complete ○/✓',
                      command=lambda: (self.todo_db.toggle(item_id),
                                       self._todo_refresh_current()))
        m.add_command(label='Edit...', command=lambda: self._todo_edit_dialog(item_id))
        m.add_separator()
        m.add_command(label='☑ Add Sub-item to Checklist...',
                      command=lambda: self._todo_add_subitem(item_id))
        m.add_separator()

        # Quick Due Date submenu
        due_sub = tk.Menu(m, tearoff=0)
        for lbl, days in [('Today', 0), ('+3 Days', 3), ('+1 Week', 7)]:
            due_sub.add_command(
                label=lbl,
                command=lambda d=days, iid=item_id: self._todo_set_due(iid, d))
        due_sub.add_separator()
        due_sub.add_command(label='Clear',
                            command=lambda iid=item_id: self._todo_set_due(iid, None))
        m.add_cascade(label='Due Date ▶', menu=due_sub)

        # Priority submenu
        prio_sub = tk.Menu(m, tearoff=0)
        for label, val in [('High  !!!', 'high'), ('Medium  !!', 'medium'),
                            ('Low  !', 'low'), ('None', '')]:
            prio_sub.add_command(
                label=label,
                command=lambda v=val, iid=item_id: (
                    self.todo_db.update(iid, priority=v),
                    self._todo_refresh_current()))
        m.add_cascade(label='Priority ▶', menu=prio_sub)
        # Attachments submenu
        atts = self.todo_db.get_attachments(item_id)
        att_sub = tk.Menu(m, tearoff=0)
        att_sub.add_command(label='Attach File(s)...',
                            command=lambda: self._todo_attach_file(item_id))
        if atts:
            att_sub.add_separator()
            for att in atts:
                fname = os.path.basename(att['file_path'])
                att_sub.add_command(
                    label=f'📄 {fname}',
                    command=lambda p=att['file_path']: self._open_file(p))
            att_sub.add_separator()
            for att in atts:
                fname = os.path.basename(att['file_path'])
                att_sub.add_command(
                    label=f'✕ Remove  {fname}',
                    foreground='#FF3B30',
                    command=lambda aid=att['id']: (
                        self.todo_db.delete_attachment(aid),
                        self._todo_refresh_current()))
        m.add_cascade(label=f'📎 Attachments ({len(atts)})', menu=att_sub)

        # Note link submenu
        with self.todo_db._conn() as c:
            row = c.execute('SELECT linked_note_path FROM todo_items WHERE id=?',
                            (item_id,)).fetchone()
        linked = row['linked_note_path'] if row else ''
        note_sub = tk.Menu(m, tearoff=0)
        if linked:
            note_sub.add_command(label=f'📄 Open "{os.path.basename(linked)}"',
                                 command=lambda: self._todo_open_linked_note(item_id))
            note_sub.add_separator()
        note_sub.add_command(label='Link Existing Note...',
                             command=lambda: self._todo_link_note(item_id))
        note_sub.add_command(label='Create New Linked Note...',
                             command=lambda: self._todo_create_linked_note(item_id))
        if linked:
            note_sub.add_separator()
            note_sub.add_command(label='Unlink', foreground='#FF3B30',
                                 command=lambda: (self.todo_db.set_linked_note(item_id, ''),
                                                  self._todo_refresh_current()))
        m.add_cascade(label='🔗 Note' + (' •' if linked else ''), menu=note_sub)

        m.add_separator()
        m.add_command(label='Delete', foreground='#FF3B30',
                      command=lambda: (self.todo_db.delete_item(item_id),
                                       self._todo_refresh_current()))
        m.tk_popup(event.x_root, event.y_root)

    def _todo_link_note(self, item_id):
        path = filedialog.askopenfilename(
            title='Select Note to Link', initialdir=NOTE_FOLDER,
            filetypes=[('Text', '*.txt'), ('All', '*.*')])
        if not path:
            return
        self.todo_db.set_linked_note(item_id, path)
        self._todo_refresh_current()

    def _todo_create_linked_note(self, item_id):
        with self.todo_db._conn() as c:
            title = c.execute('SELECT title FROM todo_items WHERE id=?',
                              (item_id,)).fetchone()['title']
        os.makedirs(NOTE_FOLDER, exist_ok=True)
        safe = re.sub(r'[\\/:*?"<>|]', '_', title)[:60] or 'note'
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = os.path.join(NOTE_FOLDER, f'{safe}_{ts}.txt')
        try:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write('')
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return
        self.todo_db.set_linked_note(item_id, path)
        self._todo_refresh_current()
        if hasattr(self, '_note_refresh'):
            self._note_refresh()

    def _todo_open_linked_note(self, item_id):
        with self.todo_db._conn() as c:
            row = c.execute('SELECT linked_note_path FROM todo_items WHERE id=?',
                            (item_id,)).fetchone()
        path = row['linked_note_path'] if row else ''
        if not path or not os.path.isfile(path):
            messagebox.showerror('Note', '找不到連結的筆記檔案')
            return
        self._main_nb.select(self.tab_note)
        iid = self._note_path_iid(path)
        if iid is None:
            self._note_refresh()
            iid = self._note_path_iid(path)
        if iid is not None:
            self._note_tv.selection_set(iid)
            self._note_tv.focus(iid)
            self._note_tv.see(iid)
            self._note_load_selected()

    def _todo_set_due(self, item_id, days):
        from datetime import date, timedelta
        due = '' if days is None else (date.today() + timedelta(days=days)).isoformat()
        self.todo_db.update(item_id, due_date=due)
        self._todo_refresh_current()
        if hasattr(self, '_todo_cal_cv'):
            self._todo_render_cal()

    # ── Checklist sub-items ──────────────────────────────────────────────────

    def _todo_add_subitem(self, item_id):
        dlg = _TodoSubItemDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            self.todo_db.add_subitem(item_id, dlg.result['title'], dlg.result['due_date'])
            self._todo_refresh_current()

    def _todo_sub_edit_dialog(self, subitem_id):
        with self.todo_db._conn() as c:
            row = c.execute('SELECT * FROM todo_subitems WHERE id=?',
                            (subitem_id,)).fetchone()
        if not row:
            return
        dlg = _TodoSubItemDialog(self.root, dict(row))
        self.root.wait_window(dlg)
        if dlg.result:
            self.todo_db.update_subitem(subitem_id, **dlg.result)
            self._todo_refresh_current()

    def _todo_sub_set_due(self, subitem_id, days):
        from datetime import date, timedelta
        due = '' if days is None else (date.today() + timedelta(days=days)).isoformat()
        self.todo_db.update_subitem(subitem_id, due_date=due)
        self._todo_refresh_current()

    def _todo_sub_ctx_menu(self, event, subitem_id):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label='Toggle Complete ○/✓',
                      command=lambda: (self.todo_db.toggle_subitem(subitem_id),
                                       self._todo_refresh_current()))
        m.add_command(label='Edit...',
                      command=lambda: self._todo_sub_edit_dialog(subitem_id))
        m.add_separator()

        due_sub = tk.Menu(m, tearoff=0)
        for lbl, days in [('Today', 0), ('+3 Days', 3), ('+1 Week', 7)]:
            due_sub.add_command(
                label=lbl,
                command=lambda d=days, sid=subitem_id: self._todo_sub_set_due(sid, d))
        due_sub.add_separator()
        due_sub.add_command(label='Clear',
                            command=lambda sid=subitem_id: self._todo_sub_set_due(sid, None))
        m.add_cascade(label='Due Date ▶', menu=due_sub)

        atts = self.todo_db.get_attachments(subitem_id=subitem_id)
        att_sub = tk.Menu(m, tearoff=0)
        att_sub.add_command(label='Attach File(s)...',
                            command=lambda: self._todo_sub_attach_file(subitem_id))
        if atts:
            att_sub.add_separator()
            for att in atts:
                fname = os.path.basename(att['file_path'])
                att_sub.add_command(
                    label=f'📄 {fname}',
                    command=lambda p=att['file_path']: self._open_file(p))
            att_sub.add_separator()
            for att in atts:
                fname = os.path.basename(att['file_path'])
                att_sub.add_command(
                    label=f'✕ Remove  {fname}',
                    foreground='#FF3B30',
                    command=lambda aid=att['id']: (
                        self.todo_db.delete_attachment(aid),
                        self._todo_refresh_current()))
        m.add_cascade(label=f'📎 Attachments ({len(atts)})', menu=att_sub)

        m.add_separator()
        m.add_command(label='Delete', foreground='#FF3B30',
                      command=lambda: (self.todo_db.delete_subitem(subitem_id),
                                       self._todo_refresh_current()))
        m.tk_popup(event.x_root, event.y_root)

    def _todo_sub_attach_file(self, subitem_id):
        with self.todo_db._conn() as c:
            item_id = c.execute('SELECT item_id FROM todo_subitems WHERE id=?',
                                (subitem_id,)).fetchone()['item_id']
        paths = _SmartFileBrowser.ask_files(self.root, 'Attach File(s) to Checklist Item')
        for p in paths:
            self.todo_db.add_attachment(item_id, p, subitem_id=subitem_id)
        if paths:
            self._todo_refresh_current()

    # ── View switching ────────────────────────────────────────────────────────

    def _todo_switch_view(self):
        v = self._todo_view_var.get()
        if v == 'list':
            self._todo_cal_outer.grid_remove()
            self._todo_tv.grid()
            self._todo_tv_vsb.grid()
        else:
            self._todo_tv.grid_remove()
            self._todo_tv_vsb.grid_remove()
            self._todo_cal_outer.grid()
            self._todo_cal_mode = v
            from datetime import date
            today = date.today()
            if self._todo_cal_year == 0:
                self._todo_cal_year  = today.year
                self._todo_cal_month = today.month
            self._todo_render_cal()

    # ── Alert system ──────────────────────────────────────────────────────────

    def _todo_start_alert(self):
        self._todo_check_alerts()

    def _todo_check_alerts(self):
        from datetime import date
        today = date.today().isoformat()
        with self.todo_db._conn() as c:
            rows = c.execute('''
                SELECT ti.id, ti.title, ti.due_date, tl.name AS list_name
                FROM todo_items ti
                JOIN todo_lists tl ON tl.id = ti.list_id
                WHERE ti.completed=0 AND ti.due_date!='' AND ti.due_date<=?
                ORDER BY ti.due_date
            ''', (today,)).fetchall()
        new = [r for r in rows if r['id'] not in self._todo_alert_ids]
        if new:
            for r in new:
                self._todo_alert_ids.add(r['id'])
            self._show_alert_popup(new)
        self.root.after(300_000, self._todo_check_alerts)   # re-check every 5 min

    def _show_alert_popup(self, items):
        pop = tk.Toplevel(self.root)
        pop.title('Reminders')
        pop.resizable(False, False)
        pop.attributes('-topmost', True)
        tk.Label(pop, text=f'  {len(items)} reminder(s) need attention  ',
                 font=('Microsoft JhengHei UI', 11, 'bold'),
                 fg='#FF3B30').pack(padx=20, pady=(14, 6))
        from datetime import date
        today = date.today()
        for it in items[:8]:
            try:
                delta = (date.fromisoformat(it['due_date']) - today).days
                when  = 'Due Today' if delta == 0 else f'Overdue {abs(delta)}d'
            except Exception:
                when = it['due_date']
            row = tk.Frame(pop)
            row.pack(fill='x', padx=20, pady=1)
            tk.Label(row, text='●', fg='#FF3B30', font=FONT_UI).pack(side='left')
            tk.Label(row, text=f"  {it['title']}  [{it['list_name']}]",
                     font=FONT_UI, anchor='w').pack(side='left', fill='x', expand=True)
            tk.Label(row, text=when, fg='#FF9500',
                     font=(FONT_UI[0], 8)).pack(side='right')
        if len(items) > 8:
            tk.Label(pop, text=f'  ...and {len(items)-8} more',
                     fg='#8E8E93', font=FONT_UI).pack(anchor='w', padx=20)
        ttk.Button(pop, text='OK', command=pop.destroy).pack(pady=(10, 14))

    # ── Calendar view ─────────────────────────────────────────────────────────

    def _build_todo_cal(self):
        import calendar as _cm
        self._todo_cal_mod  = _cm
        self._todo_cal_mode = 'month'
        self._todo_cell_map = {}

        nav = tk.Frame(self._todo_cal_outer, bg=_TODO_BG)
        nav.grid(row=0, column=0, sticky='ew', pady=(0, 4))
        ttk.Button(nav, text='◀', width=3,
                   command=self._todo_cal_prev).pack(side='left')
        self._todo_cal_title = tk.Label(
            nav, text='', bg=_TODO_BG,
            font=('Microsoft JhengHei UI', 11, 'bold'), fg='#1C1C1E')
        self._todo_cal_title.pack(side='left', padx=10)
        ttk.Button(nav, text='▶', width=3,
                   command=self._todo_cal_next).pack(side='left')

        self._todo_cal_cv = tk.Canvas(
            self._todo_cal_outer, bg='white', bd=0,
            highlightbackground='#E5E5EA', highlightthickness=1)
        self._todo_cal_cv.grid(row=1, column=0, sticky='nsew')
        self._todo_cal_cv.bind('<Configure>', lambda e: self._todo_render_cal())
        self._todo_cal_cv.bind('<Button-1>',  self._todo_cal_click)

    def _todo_cal_prev(self):
        if self._todo_cal_mode == 'month':
            if self._todo_cal_month == 1:
                self._todo_cal_month = 12; self._todo_cal_year -= 1
            else:
                self._todo_cal_month -= 1
        else:
            self._todo_cal_woff -= 1
        self._todo_render_cal()

    def _todo_cal_next(self):
        if self._todo_cal_mode == 'month':
            if self._todo_cal_month == 12:
                self._todo_cal_month = 1; self._todo_cal_year += 1
            else:
                self._todo_cal_month += 1
        else:
            self._todo_cal_woff += 1
        self._todo_render_cal()

    def _todo_cal_items_on(self, date_str):
        with self.todo_db._conn() as c:
            return c.execute('''
                SELECT ti.id, ti.title, ti.completed, tl.color
                FROM todo_items ti JOIN todo_lists tl ON tl.id=ti.list_id
                WHERE ti.due_date=?
                ORDER BY ti.completed, ti.created_at
            ''', (date_str,)).fetchall()

    def _todo_render_cal(self):
        if self._todo_cal_mode == 'month':
            self._todo_render_month()
        else:
            self._todo_render_week()

    def _todo_render_month(self):
        from datetime import date
        cv = self._todo_cal_cv
        cv.delete('all')
        self._todo_cell_map = {}
        W, H = cv.winfo_width(), cv.winfo_height()
        if W < 20 or H < 20:
            return
        yr, mo = self._todo_cal_year, self._todo_cal_month
        self._todo_cal_title.configure(text=f'{yr} / {mo:02d}  (Month)')
        today = date.today()
        cal   = self._todo_cal_mod.monthcalendar(yr, mo)
        col_w = W / 7
        hdr_h = 22
        cell_h = (H - hdr_h) / len(cal)
        for i, d in enumerate(['Mon','Tue','Wed','Thu','Fri','Sat','Sun']):
            cv.create_text(i*col_w + col_w/2, hdr_h/2, text=d,
                           font=('Microsoft JhengHei UI', 8, 'bold'), fill='#8E8E93')
        for ri, week in enumerate(cal):
            for ci, day in enumerate(week):
                if day == 0:
                    continue
                x0, y0 = ci*col_w, hdr_h + ri*cell_h
                x1, y1 = x0+col_w, y0+cell_h
                date_str  = f'{yr:04d}-{mo:02d}-{day:02d}'
                is_today  = (date(yr, mo, day) == today)
                rect = cv.create_rectangle(x0+1, y0+1, x1, y1,
                                           fill='#EBF5FF' if is_today else 'white',
                                           outline='#E5E5EA')
                self._todo_cell_map[rect] = date_str
                cv.create_text(x0+5, y0+4, anchor='nw', text=str(day),
                               font=('Microsoft JhengHei UI', 8,
                                     'bold' if is_today else 'normal'),
                               fill='#007AFF' if is_today else '#1C1C1E')
                for j, it in enumerate(self._todo_cal_items_on(date_str)):
                    iy = y0 + 18 + j*14
                    if iy + 10 > y1:
                        cv.create_text(x0+5, iy, anchor='nw',
                                       text=f'+{len(self._todo_cal_items_on(date_str))-j}',
                                       font=('Microsoft JhengHei UI', 7),
                                       fill='#8E8E93')
                        break
                    fg = '#A0A0A5' if it['completed'] else '#1C1C1E'
                    cv.create_oval(x0+4, iy+3, x0+8, iy+7,
                                   fill=it['color'] or '#007AFF', outline='')
                    cv.create_text(x0+11, iy, anchor='nw',
                                   text=(it['title'] or '')[:int(col_w//8)],
                                   font=('Microsoft JhengHei UI', 7), fill=fg)

    def _todo_render_week(self):
        from datetime import date, timedelta
        cv = self._todo_cal_cv
        cv.delete('all')
        self._todo_cell_map = {}
        W, H = cv.winfo_width(), cv.winfo_height()
        if W < 20 or H < 20:
            return
        today  = date.today()
        wstart = today - timedelta(days=today.weekday()) + timedelta(weeks=self._todo_cal_woff)
        self._todo_cal_title.configure(
            text=f'Week  {wstart.strftime("%Y/%m/%d")} – '
                 f'{(wstart+timedelta(days=6)).strftime("%m/%d")}  (Week)')
        col_w, hdr_h = W/7, 38
        for ci in range(7):
            day = wstart + timedelta(days=ci)
            x0  = ci * col_w
            if ci > 0:
                cv.create_line(x0, 0, x0, H, fill='#E5E5EA')
            is_today = (day == today)
            cv.create_text(x0+col_w/2, 12, anchor='center',
                           text=day.strftime('%a'),
                           font=('Microsoft JhengHei UI', 8, 'bold'), fill='#8E8E93')
            cv.create_text(x0+col_w/2, 26, anchor='center',
                           text=day.strftime('%m/%d'),
                           font=('Microsoft JhengHei UI', 9,
                                 'bold' if is_today else 'normal'),
                           fill='#007AFF' if is_today else '#1C1C1E')
            date_str = day.isoformat()
            items = self._todo_cal_items_on(date_str)
            for j, it in enumerate(items):
                iy = hdr_h + j*20
                if iy+14 > H:
                    cv.create_text(x0+5, iy, anchor='nw',
                                   text=f'+{len(items)-j}',
                                   font=('Microsoft JhengHei UI', 7), fill='#8E8E93')
                    break
                fg = '#A0A0A5' if it['completed'] else '#1C1C1E'
                cv.create_oval(x0+4, iy+4, x0+8, iy+8,
                               fill=it['color'] or '#007AFF', outline='')
                cv.create_text(x0+12, iy+1, anchor='nw',
                               text=(it['title'] or '')[:int(col_w//8)],
                               font=('Microsoft JhengHei UI', 8), fill=fg)

    def _todo_cal_click(self, event):
        hits = self._todo_cal_cv.find_overlapping(
            event.x-1, event.y-1, event.x+1, event.y+1)
        for h in hits:
            if h in self._todo_cell_map:
                date_str = self._todo_cell_map[h]
                items = self._todo_cal_items_on(date_str)
                if not items:
                    return
                m = tk.Menu(self.root, tearoff=0)
                for it in items:
                    lbl = ('✓ ' if it['completed'] else '○ ') + (it['title'] or '')
                    m.add_command(label=lbl,
                                  command=lambda iid=it['id']: self._todo_edit_dialog(iid))
                m.tk_popup(event.x_root, event.y_root)
                return


    # ── Tab: Projects ─────────────────────────────────────────────────────────

    def _build_proj_tab(self):
        f = self.tab_proj
        f.rowconfigure(1, weight=1)
        f.columnconfigure(0, weight=1)

        # ── Top toolbar ───────────────────────────────────────────────────────
        tb = tk.Frame(f, bg='#F2F2F7')
        tb.grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 4))

        ttk.Button(tb, text='+ New Project',
                   command=self._proj_new).pack(side='left', padx=(0, 4))
        ttk.Button(tb, text='+ Add Existing',
                   command=self._proj_add_existing).pack(side='left')
        ttk.Button(tb, text='↺ Refresh',
                   command=self._proj_refresh_tree).pack(side='left', padx=(8, 0))

        tk.Label(tb, text='🔍', bg='#F2F2F7', font=FONT_UI).pack(side='left', padx=(16, 2))
        self._proj_search_sv = tk.StringVar()
        srch = ttk.Entry(tb, textvariable=self._proj_search_sv, width=20)
        srch.pack(side='left')
        srch.bind('<KeyRelease>', lambda e: self._proj_search())

        # ── Main paned area ───────────────────────────────────────────────────
        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=1, column=0, sticky='nsew', padx=8, pady=(0, 8))

        # ── Left: project tree ────────────────────────────────────────────────
        left = tk.Frame(pw)
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(left, text='MY PROJECTS', bg='#EFEFF4', fg='#8E8E93',
                 font=('Microsoft JhengHei UI', 8, 'bold'),
                 anchor='w').grid(row=0, column=0, columnspan=2,
                                  sticky='ew', padx=8, pady=(6, 2))

        self._proj_tree = ttk.Treeview(left, show='tree', selectmode='browse')
        self._proj_tree.column('#0', width=200)
        proj_vsb = ttk.Scrollbar(left, orient='vertical',
                                  command=self._proj_tree.yview)
        self._proj_tree.configure(yscrollcommand=proj_vsb.set)
        self._proj_tree.grid(row=1, column=0, sticky='nsew')
        proj_vsb.grid(row=1, column=1, sticky='ns')
        self._proj_tree.tag_configure('missing', foreground='#FF9500')
        self._proj_tree.tag_configure('subfolder', foreground='#555555')
        self._proj_tree.bind('<<TreeviewSelect>>', self._proj_tree_select)
        self._proj_tree.bind('<<TreeviewOpen>>',   self._proj_tree_expand)
        self._proj_tree.bind('<Button-3>',          self._proj_left_ctx)

        # ── Right: file browser ───────────────────────────────────────────────
        right = tk.Frame(pw)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        nav = tk.Frame(right)
        nav.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(4, 2))
        ttk.Button(nav, text='↑ Up', width=5,
                   command=self._proj_go_up).pack(side='left', padx=(4, 6))
        self._proj_path_var = tk.StringVar(value='← Select a project')
        tk.Label(nav, textvariable=self._proj_path_var, fg='#555',
                 font=FONT_UI, anchor='w').pack(side='left', fill='x', expand=True)
        ttk.Button(nav, text='🗁 Explorer',
                   command=self._proj_open_in_explorer).pack(side='right', padx=4)

        cols = ('name', 'size', 'modified')
        self._proj_tv = ttk.Treeview(right, columns=cols, show='headings',
                                      selectmode='extended')
        self._proj_tv.heading('name',     text='Name')
        self._proj_tv.heading('size',     text='Size')
        self._proj_tv.heading('modified', text='Modified')
        self._proj_tv.column('name',     width=340, stretch=True,  anchor='w')
        self._proj_tv.column('size',     width=80,  stretch=False, anchor='e')
        self._proj_tv.column('modified', width=130, stretch=False)
        self._proj_tv.tag_configure('dir',  font=(FONT_UI[0], FONT_UI[1], 'bold'))
        self._proj_tv.tag_configure('file', foreground='#3A3A3A')
        for _k, (_bg, _) in _PROJ_TAGS.items():
            self._proj_tv.tag_configure(f'clr_{_k}', background=_bg)
        file_vsb = ttk.Scrollbar(right, orient='vertical',
                                  command=self._proj_tv.yview)
        self._proj_tv.configure(yscrollcommand=file_vsb.set)
        self._proj_tv.grid(row=1, column=0, sticky='nsew')
        file_vsb.grid(row=1, column=1, sticky='ns')
        self._proj_tv.bind('<Double-1>',          self._proj_dbl_click)
        self._proj_tv.bind('<<TreeviewSelect>>', self._proj_file_sel)
        self._proj_tv.bind('<Button-3>',          self._proj_right_ctx)

        bot = tk.Frame(right)
        bot.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(2, 4))
        self._proj_sel_lbl = tk.Label(bot, text='', fg='#8E8E93', font=FONT_UI)
        self._proj_sel_lbl.pack(side='left', padx=4)
        ttk.Button(bot, text='Copy Path',
                   command=self._proj_copy_path).pack(side='right', padx=4)
        ttk.Button(bot, text='Open',
                   command=self._proj_open_selected).pack(side='right')

        pw.add(left,  weight=1)
        pw.add(right, weight=3)

        self._proj_refresh_tree()

    # ── Projects: left tree ───────────────────────────────────────────────────

    def _proj_refresh_tree(self):
        # Store expanded state
        expanded = {iid for iid in self._proj_tree.get_children()
                    if self._proj_tree.item(iid, 'open')}
        self._proj_tree.delete(*self._proj_tree.get_children())
        cfg  = _load_app_cfg()
        lbls = cfg.get('project_labels', {})
        for path in cfg.get('pinned_folders', []):
            exists  = os.path.isdir(path)
            name    = lbls.get(path, os.path.basename(path) or path)
            icon    = '📁' if exists else '⚠'
            tags    = () if exists else ('missing',)
            iid     = self._proj_tree.insert(
                '', 'end', iid=path,
                text=f'{icon}  {name}', tags=tags, open=(path in expanded))
            # Placeholder child for expand arrow
            if exists:
                try:
                    has_sub = any(e.is_dir() for e in os.scandir(path))
                except Exception:
                    has_sub = False
                if has_sub:
                    self._proj_tree.insert(iid, 'end', iid=path + '/__dummy__',
                                           text='')

    def _proj_tree_expand(self, event):
        iid = self._proj_tree.focus()
        if not iid:
            return
        path = iid
        # Remove placeholder
        for child in self._proj_tree.get_children(iid):
            if child.endswith('/__dummy__'):
                self._proj_tree.delete(child)
        # Load real subfolders (one level)
        if not os.path.isdir(path):
            return
        existing_children = set(self._proj_tree.get_children(iid))
        try:
            subdirs = sorted(
                (e for e in os.scandir(path) if e.is_dir() and not e.name.startswith('.')),
                key=lambda e: e.name.lower())
        except Exception:
            return
        for e in subdirs:
            if e.path not in existing_children:
                child_iid = self._proj_tree.insert(
                    iid, 'end', iid=e.path,
                    text=f'📁  {e.name}', tags=('subfolder',))
                # Add dummy child if this subfolder has sub-subfolders
                try:
                    if any(s.is_dir() for s in os.scandir(e.path)):
                        self._proj_tree.insert(child_iid, 'end',
                                               iid=e.path + '/__dummy__', text='')
                except Exception:
                    pass

    def _proj_tree_select(self, _=None):
        sel = self._proj_tree.selection()
        if not sel:
            return
        path = sel[0]
        if path.endswith('/__dummy__'):
            return
        if os.path.isdir(path):
            self._proj_navigate(path)

    def _proj_left_ctx(self, event):
        iid = self._proj_tree.identify_row(event.y)
        if not iid or iid.endswith('/__dummy__'):
            return
        self._proj_tree.selection_set(iid)
        # Only allow rename/remove for top-level project roots
        cfg    = _load_app_cfg()
        is_top = iid in cfg.get('pinned_folders', [])
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label='Open in Explorer',
                      command=lambda: os.startfile(iid) if os.path.isdir(iid) else None)
        if is_top:
            m.add_separator()
            m.add_command(label='Rename Label...',
                          command=lambda p=iid: self._proj_rename(p))
            m.add_command(label='Remove from list', foreground='#FF3B30',
                          command=lambda p=iid: self._proj_remove(p))
        m.tk_popup(event.x_root, event.y_root)

    # ── Projects: right file panel ────────────────────────────────────────────

    def _proj_navigate(self, path):
        if not os.path.isdir(path):
            messagebox.showwarning('Not Found', f'Folder not found:\n{path}')
            return
        self._proj_cur_path = path
        self._proj_path_var.set(path)
        self._proj_search_sv.set('')
        self._proj_populate(path)

    def _proj_populate(self, path, filter_str=''):
        self._proj_tv.delete(*self._proj_tv.get_children())
        from datetime import datetime as _dt
        file_tags = _load_app_cfg().get('file_tags', {})
        try:
            entries = sorted(os.scandir(path),
                             key=lambda e: (not e.is_dir(), e.name.lower()))
        except PermissionError:
            return
        q = filter_str.lower()
        for e in entries:
            if e.name.startswith('.'):
                continue
            if q and q not in e.name.lower():
                continue
            try:
                st   = e.stat()
                mt   = _dt.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M')
                if e.is_dir():
                    self._proj_tv.insert('', 'end', iid=e.path,
                                         values=('📁  ' + e.name, '', mt),
                                         tags=('dir',))
                else:
                    sz   = st.st_size
                    sz_s = (f'{sz/1024:.0f} KB' if sz < 1_048_576
                            else f'{sz/1_048_576:.1f} MB')
                    clr  = file_tags.get(e.path)
                    tags = (f'clr_{clr}', 'file') if clr in _PROJ_TAGS else ('file',)
                    self._proj_tv.insert('', 'end', iid=e.path,
                                         values=('📄  ' + e.name, sz_s, mt),
                                         tags=tags)
            except Exception:
                pass

    def _proj_go_up(self):
        if not self._proj_cur_path:
            return
        parent = os.path.dirname(self._proj_cur_path)
        if parent != self._proj_cur_path:
            self._proj_navigate(parent)

    def _proj_dbl_click(self, event):
        iid = self._proj_tv.identify_row(event.y)
        if not iid:
            return
        if os.path.isdir(iid):
            self._proj_navigate(iid)
        else:
            self._open_file(iid)

    def _proj_file_sel(self, _=None):
        n = len(self._proj_tv.selection())
        self._proj_sel_lbl.configure(
            text=f'{n} item(s) selected' if n else '')

    def _proj_open_selected(self):
        for iid in self._proj_tv.selection():
            self._open_file(iid)

    def _proj_copy_path(self):
        sel = self._proj_tv.selection()
        if not sel:
            if self._proj_cur_path:
                self.root.clipboard_clear()
                self.root.clipboard_append(self._proj_cur_path)
            return
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(sel))

    def _proj_open_in_explorer(self):
        path = self._proj_cur_path
        if path and os.path.isdir(path):
            os.startfile(path)

    def _proj_right_ctx(self, event):
        iid = self._proj_tv.identify_row(event.y)
        if not iid:
            return
        self._proj_tv.selection_set(iid)
        m = tk.Menu(self.root, tearoff=0)
        if os.path.isfile(iid):
            m.add_command(label='Open',
                          command=lambda: self._open_file(iid))
        else:
            m.add_command(label='Open in Explorer',
                          command=lambda: os.startfile(iid))
        m.add_command(label='Copy Path',
                      command=lambda: (self.root.clipboard_clear(),
                                       self.root.clipboard_append(iid)))

        # Tag / highlight submenu (files only)
        if os.path.isfile(iid):
            m.add_separator()
            tag_sub = tk.Menu(m, tearoff=0)
            for key, (_, lbl) in _PROJ_TAGS.items():
                tag_sub.add_command(
                    label=lbl,
                    command=lambda k=key, p=iid: self._proj_set_tag(p, k))
            tag_sub.add_separator()
            tag_sub.add_command(label='✕  Clear Tag',
                                command=lambda p=iid: self._proj_set_tag(p, None))
            m.add_cascade(label='🏷  Tag Color', menu=tag_sub)

        m.tk_popup(event.x_root, event.y_root)

    def _proj_set_tag(self, path, color_key):
        cfg = _load_app_cfg()
        cfg.setdefault('file_tags', {})
        if color_key is None:
            cfg['file_tags'].pop(path, None)
        else:
            cfg['file_tags'][path] = color_key
        _save_app_cfg(cfg)
        self._proj_search()   # refresh current view

    def _proj_search(self):
        if self._proj_cur_path:
            self._proj_populate(self._proj_cur_path,
                                self._proj_search_sv.get())

    # ── Projects: project management ──────────────────────────────────────────

    def _proj_add_existing(self):
        path = filedialog.askdirectory(title='Select project folder to add')
        if not path:
            return
        cfg = _load_app_cfg()
        cfg.setdefault('pinned_folders', [])
        if path not in cfg['pinned_folders']:
            cfg['pinned_folders'].append(path)
            _save_app_cfg(cfg)
        self._proj_refresh_tree()

    def _proj_remove(self, path):
        if not messagebox.askyesno('Remove', f'Remove "{os.path.basename(path)}" from list?\n(Folder is NOT deleted)'):
            return
        cfg = _load_app_cfg()
        cfg.get('pinned_folders', []).discard if False else None
        try:
            cfg['pinned_folders'].remove(path)
        except (ValueError, KeyError):
            pass
        cfg.get('project_labels', {}).pop(path, None)
        _save_app_cfg(cfg)
        self._proj_refresh_tree()
        if self._proj_cur_path and self._proj_cur_path.startswith(path):
            self._proj_cur_path = None
            self._proj_path_var.set('← Select a project')
            self._proj_tv.delete(*self._proj_tv.get_children())

    def _proj_rename(self, path):
        cfg  = _load_app_cfg()
        cur  = cfg.get('project_labels', {}).get(path, os.path.basename(path))
        dlg  = tk.Toplevel(self.root)
        dlg.title('Rename Label')
        dlg.resizable(False, False)
        dlg.grab_set()
        tk.Label(dlg, text='Display name:', font=FONT_UI).pack(padx=16, pady=(12, 4))
        sv = tk.StringVar(value=cur)
        ttk.Entry(dlg, textvariable=sv, width=28).pack(padx=16, pady=4)
        def _ok():
            lbl = sv.get().strip()
            if lbl:
                cfg.setdefault('project_labels', {})[path] = lbl
                _save_app_cfg(cfg)
            dlg.destroy()
            self._proj_refresh_tree()
        bf = ttk.Frame(dlg)
        bf.pack(padx=16, pady=(4, 12), fill='x')
        ttk.Button(bf, text='Cancel', command=dlg.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=_ok).pack(side='right')

    def _proj_new(self):
        dlg = _NewProjectDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            cfg = _load_app_cfg()
            if dlg.result not in cfg.get('pinned_folders', []):
                cfg.setdefault('pinned_folders', []).append(dlg.result)
                _save_app_cfg(cfg)
            self._proj_refresh_tree()
            self._proj_navigate(dlg.result)



    # ── Tab: BOM Risk Scanner ─────────────────────────────────────────────────

    def _build_risk_tab(self):
        f = self.tab_risk
        f.rowconfigure(1, weight=1)
        f.columnconfigure(0, weight=1)

        # Progress bar (row 0)
        prog_f = ttk.Frame(f)
        prog_f.grid(row=0, column=0, sticky='ew', padx=8, pady=(6, 0))
        self._risk_prog = ttk.Progressbar(prog_f, mode='determinate', length=400)
        self._risk_prog.pack(side='left', fill='x', expand=True)
        self._risk_prog_lbl = ttk.Label(prog_f, text='', foreground='gray', font=FONT_UI)
        self._risk_prog_lbl.pack(side='left', padx=8)

        # Main paned window (row 1)
        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=1, column=0, sticky='nsew', padx=8, pady=6)

        # ── Left panel ────────────────────────────────────────────────────────
        left = tk.Frame(pw, width=220)
        left.grid_propagate(False)
        left.rowconfigure(4, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(left, text='BOM RISK SCANNER', fg='#8E8E93',
                 font=('Microsoft JhengHei UI', 8, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(10, 4))

        ttk.Button(left, text='📂 Import BOM...',
                   command=self._risk_import_bom).grid(
            row=1, column=0, columnspan=2, sticky='ew', padx=8, pady=2)

        ttk.Separator(left, orient='horizontal').grid(
            row=2, column=0, columnspan=2, sticky='ew', padx=8, pady=6)

        tk.Label(left, text='Sessions', font=FONT_UI).grid(
            row=3, column=0, sticky='w', padx=8)

        self._risk_sess_lb = tk.Listbox(left, font=FONT_UI, bd=0,
                                         highlightthickness=0,
                                         selectbackground='#D1D1D6')
        self._risk_sess_lb.grid(row=4, column=0, columnspan=2, sticky='nsew', padx=4)
        self._risk_sess_lb.bind('<Double-1>', lambda e: self._risk_load_session())

        sess_btn_f = tk.Frame(left)
        sess_btn_f.grid(row=5, column=0, columnspan=2, sticky='ew', padx=4, pady=2)
        ttk.Button(sess_btn_f, text='Load',   width=6,
                   command=self._risk_load_session).pack(side='left')
        ttk.Button(sess_btn_f, text='Delete', width=6,
                   command=self._risk_delete_session).pack(side='left', padx=4)

        ttk.Separator(left, orient='horizontal').grid(
            row=6, column=0, columnspan=2, sticky='ew', padx=8, pady=6)

        self._risk_run_btn = ttk.Button(left, text='▶ Run Risk Scan',
                                         command=self._risk_run_scan, state='disabled')
        self._risk_run_btn.grid(row=7, column=0, columnspan=2, sticky='ew', padx=8, pady=2)

        self._risk_stop_btn = ttk.Button(left, text='■ Stop',
                                          command=self._risk_stop_scan)
        self._risk_stop_btn.grid(row=8, column=0, columnspan=2, sticky='ew', padx=8, pady=2)
        self._risk_stop_btn.grid_remove()

        ttk.Button(left, text='📊 Export Excel...',
                   command=self._risk_export_excel).grid(
            row=9, column=0, columnspan=2, sticky='ew', padx=8, pady=2)

        ttk.Separator(left, orient='horizontal').grid(
            row=10, column=0, columnspan=2, sticky='ew', padx=8, pady=6)

        ttk.Button(left, text='🔄 Alt Parts DB...',
                   command=self._risk_alt_dialog).grid(
            row=11, column=0, columnspan=2, sticky='ew', padx=8, pady=(2, 10))

        # ── Center: BOM table ─────────────────────────────────────────────────
        center = tk.Frame(pw)
        center.rowconfigure(0, weight=1)
        center.columnconfigure(0, weight=1)

        cols = ('no', 'vpn', 'desc', 'qty', 'm_stk', 'dk_stk',
                'lt_wk', 'lifecycle', 'score', 'level')
        self._risk_tv = ttk.Treeview(center, columns=cols, show='headings',
                                      selectmode='browse')
        hdrs = [('no',   'No.',        40,  'center'),
                ('vpn',  'Vendor PN', 160,  'w'),
                ('desc', 'Description',180, 'w'),
                ('qty',  'Qty',        50,  'center'),
                ('m_stk','Mouser',     80,  'center'),
                ('dk_stk','DigiKey',   80,  'center'),
                ('lt_wk','LT(wk)',     55,  'center'),
                ('lifecycle','Lifecycle',80,'center'),
                ('score','Score',       55, 'center'),
                ('level','Risk',        75, 'center')]
        for cid, hdr, w, anch in hdrs:
            self._risk_tv.heading(cid, text=hdr)
            self._risk_tv.column(cid, width=w, stretch=(cid == 'desc'),
                                 anchor=anch, minwidth=30)
        for lvl, bg in _RISK_CLR.items():
            self._risk_tv.tag_configure(lvl, background=bg)
        self._risk_tv.tag_configure('pending', foreground='#999999')

        tv_vsb = ttk.Scrollbar(center, orient='vertical',   command=self._risk_tv.yview)
        tv_hsb = ttk.Scrollbar(center, orient='horizontal', command=self._risk_tv.xview)
        self._risk_tv.configure(yscrollcommand=tv_vsb.set,
                                 xscrollcommand=tv_hsb.set)
        self._risk_tv.grid(row=0, column=0, sticky='nsew')
        tv_vsb.grid(row=0, column=1, sticky='ns')
        tv_hsb.grid(row=1, column=0, sticky='ew')
        self._risk_tv.bind('<Button-3>', self._risk_tv_ctx)
        self._risk_tv.bind('<Control-c>',
                           lambda e: self._risk_copy_vpn())

        # ── Right: summary dashboard ──────────────────────────────────────────
        right = tk.Frame(pw, width=185, bg='#F8F8F8')
        right.grid_propagate(False)
        right.columnconfigure(0, weight=1)

        tk.Label(right, text='Risk Summary', bg='#F8F8F8', fg='#8E8E93',
                 font=('Microsoft JhengHei UI', 8, 'bold')).grid(
            row=0, column=0, padx=10, pady=(10, 4))

        self._risk_score_lbl = tk.Label(
            right, text='--', bg='#F8F8F8',
            font=('Microsoft JhengHei UI', 32, 'bold'), fg='#1C1C1E')
        self._risk_score_lbl.grid(row=1, column=0, pady=(4, 0))

        self._risk_level_lbl = tk.Label(
            right, text='', bg='#F8F8F8',
            font=('Microsoft JhengHei UI', 10, 'bold'))
        self._risk_level_lbl.grid(row=2, column=0, pady=(0, 8))

        ttk.Separator(right, orient='horizontal').grid(
            row=3, column=0, sticky='ew', padx=10)

        self._risk_summary_vars = {}
        for i, (key, lbl) in enumerate([
                ('total',    'Total Parts'),
                ('critical', '🔴 Critical'),
                ('high',     '🟠 High'),
                ('medium',   '🟡 Medium'),
                ('low',      '🟢 Low')]):
            tk.Label(right, text=lbl, bg='#F8F8F8', font=FONT_UI,
                     anchor='w').grid(row=4+i*2, column=0, sticky='w',
                                      padx=14, pady=(6, 0))
            sv = tk.StringVar(value='0')
            self._risk_summary_vars[key] = sv
            tk.Label(right, textvariable=sv, bg='#F8F8F8',
                     font=(FONT_UI[0], FONT_UI[1], 'bold'),
                     fg='#1C1C1E', anchor='e').grid(
                row=4+i*2, column=0, sticky='e', padx=14)

        pw.add(left,   weight=0)
        pw.add(center, weight=3)
        pw.add(right,  weight=0)

        self._risk_refresh_sessions()

    # ── Risk: sessions ────────────────────────────────────────────────────────

    def _risk_refresh_sessions(self):
        self._risk_sess_lb.delete(0, 'end')
        self._risk_sessions = list(self.risk_db.get_scans())
        for s in self._risk_sessions:
            self._risk_sess_lb.insert('end', f"  {s['name']}  ({s['created_at'][:10]})")

    def _risk_load_session(self):
        sel = self._risk_sess_lb.curselection()
        if not sel:
            return
        scan = self._risk_sessions[sel[0]]
        self._risk_scan_id = scan['id']
        items = self.risk_db.get_items(self._risk_scan_id)
        self._risk_populate_tv(items)
        self._risk_update_summary(items)
        self._risk_run_btn.configure(state='normal')

    def _risk_delete_session(self):
        sel = self._risk_sess_lb.curselection()
        if not sel:
            return
        scan = self._risk_sessions[sel[0]]
        if not messagebox.askyesno('Delete',
                f"Delete scan '{scan['name']}' ?\n(Cannot undo)"):
            return
        self.risk_db.delete_scan(scan['id'])
        if self._risk_scan_id == scan['id']:
            self._risk_scan_id = None
            self._risk_tv.delete(*self._risk_tv.get_children())
            self._risk_run_btn.configure(state='disabled')
        self._risk_refresh_sessions()

    # ── Risk: import BOM ──────────────────────────────────────────────────────

    def _risk_import_bom(self):
        path = filedialog.askopenfilename(
            title='Import BOM for Risk Scan',
            filetypes=[('BOM Files', '*.bom *.BOM'),
                       ('Excel', '*.xlsx *.xls'),
                       ('CSV', '*.csv'),
                       ('All', '*.*')])
        if not path:
            return

        ext          = os.path.splitext(path)[1].lower()
        default_name = os.path.splitext(os.path.basename(path))[0]
        parsed_rows  = []   # list of (vpn, desc, qty)

        # .bom path (OrCAD / SAP fixed-width) — no column picker needed
        if ext == '.bom':
            try:
                if SapBomParser.detect(path):
                    _, raw = SapBomParser.parse(path)
                    rows = [(r['vendor_pn'], r['description'], r['qty'])
                            for r in raw if (r.get('vendor_pn') or '').strip()]
                else:
                    items = BomParser.parse(path)
                    rows  = [(i.vendor_pn or i.value, i.description, i.qty)
                             for i in items if (i.vendor_pn or i.value).strip()]
            except Exception as e:
                messagebox.showerror('Error', str(e))
                return
            seen = set()
            for vpn, desc, qty in rows:
                vpn = vpn.strip()
                if vpn and vpn not in seen:
                    seen.add(vpn)
                    parsed_rows.append((vpn, desc, qty))

        # Excel / CSV path — use column picker
        else:
            try:
                columns, preview, all_rows = self._lib_read_file(path)
            except Exception as e:
                messagebox.showerror('Error', str(e))
                return
            dlg_vpn = _VpnColPicker(self.root, columns, preview)
            self.root.wait_window(dlg_vpn)
            if dlg_vpn.result is None:
                return
            vpn_col  = dlg_vpn.result
            desc_col = next((c for c in columns
                             if any(h in c.lower()
                                    for h in ['desc', 'description', 'text', 'name'])), None)
            qty_col  = next((c for c in columns
                             if any(h in c.lower()
                                    for h in ['qty', 'quantity', 'amount'])), None)
            seen = set()
            for row in all_rows:
                vpn = str(row.get(vpn_col, '') or '').strip()
                if not vpn or vpn in seen:
                    continue
                seen.add(vpn)
                desc = str(row.get(desc_col, '') or '').strip() if desc_col else ''
                try:
                    qty = int(str(row.get(qty_col, 0) or 0).replace(',', '')) if qty_col else 0
                except Exception:
                    qty = 0
                parsed_rows.append((vpn, desc, qty))

        if not parsed_rows:
            messagebox.showwarning('Import BOM', 'No valid parts found in file.')
            return

        # Scan name dialog
        name_dlg = tk.Toplevel(self.root)
        name_dlg.title('Scan Name')
        name_dlg.resizable(False, False)
        name_dlg.grab_set()
        sv_name = tk.StringVar(value=default_name)
        ttk.Label(name_dlg, text='Scan name:', font=FONT_UI).pack(padx=14, pady=(12, 4))
        ttk.Entry(name_dlg, textvariable=sv_name, width=32).pack(padx=14, pady=4)
        confirmed = [False]
        def _ok():
            confirmed[0] = True
            name_dlg.destroy()
        bf = ttk.Frame(name_dlg)
        bf.pack(padx=14, pady=(4, 12))
        ttk.Button(bf, text='Cancel', command=name_dlg.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK',     command=_ok).pack(side='right')
        self.root.wait_window(name_dlg)
        if not confirmed[0]:
            return

        # Save to DB
        scan_id = self.risk_db.new_scan(sv_name.get().strip() or default_name, path)
        self._risk_scan_id   = scan_id
        self._risk_items_buf = []
        for vpn, desc, qty in parsed_rows:
            item_id = self.risk_db.add_item(scan_id, vpn, desc, qty)
            self._risk_items_buf.append((item_id, vpn, qty))

        items = self.risk_db.get_items(scan_id)
        self._risk_populate_tv(items)
        self._risk_update_summary(items)
        self._risk_run_btn.configure(state='normal')
        self._risk_refresh_sessions()
        messagebox.showinfo('Imported',
                            f'{len(self._risk_items_buf)} unique parts loaded.\n'
                            'Click ▶ Run Risk Scan to start.')

    def _risk_populate_tv(self, items):
        self._risk_tv.delete(*self._risk_tv.get_children())
        for i, it in enumerate(items, 1):
            lvl = it['risk_level'] or ''
            m   = str(it['mouser_stock']) if it['mouser_stock'] >= 0 else '--'
            dk  = str(it['dk_stock'])     if it['dk_stock']     >= 0 else '--'
            lt  = str(it['lead_time_wk']) if it['lead_time_wk'] >= 0 else '--'
            sc  = str(it['risk_total'])   if lvl else '--'
            tag = lvl if lvl in _RISK_CLR else 'pending'
            self._risk_tv.insert('', 'end', iid=str(it['id']),
                                  values=(i, it['vendor_pn'],
                                          (it['description'] or '')[:40],
                                          it['qty'], m, dk, lt,
                                          (it['lifecycle'] or '')[:12] or '--',
                                          sc, lvl or '--'),
                                  tags=(tag,))

    def _risk_update_row(self, item_id, m_stk, dk_stk, lt_wk, lifecycle, score, level):
        iid = str(item_id)
        if not self._risk_tv.exists(iid):
            return
        vals = list(self._risk_tv.item(iid, 'values'))
        vals[4] = str(m_stk)  if m_stk  >= 0 else '--'
        vals[5] = str(dk_stk) if dk_stk >= 0 else '--'
        vals[6] = str(lt_wk)  if lt_wk  >= 0 else '--'
        vals[7] = (lifecycle or '')[:12] or '--'
        vals[8] = str(score)
        vals[9] = level
        self._risk_tv.item(iid, values=vals,
                           tags=(level if level in _RISK_CLR else 'pending',))

    def _risk_update_summary(self, items=None):
        if items is None:
            items = (self.risk_db.get_items(self._risk_scan_id)
                     if self._risk_scan_id else [])
        counted = [it for it in items if it['risk_level']]
        lvl_cnt = {k: 0 for k in ('CRITICAL', 'HIGH', 'MEDIUM', 'LOW')}
        total_sc = 0
        for it in counted:
            lv = it['risk_level']
            if lv in lvl_cnt:
                lvl_cnt[lv] += 1
            total_sc += it['risk_total']
        avg = total_sc // len(counted) if counted else 0
        if avg > 100:  proj_lv = 'CRITICAL'
        elif avg > 60: proj_lv = 'HIGH'
        elif avg > 30: proj_lv = 'MEDIUM'
        else:          proj_lv = 'LOW'
        clrs = {'CRITICAL': '#CC0000', 'HIGH': '#FF6600',
                'MEDIUM': '#CC9900', 'LOW': '#228B22'}
        self._risk_score_lbl.configure(text=str(avg) if counted else '--')
        self._risk_level_lbl.configure(
            text=proj_lv if counted else '',
            fg=clrs.get(proj_lv, '#1C1C1E'))
        self._risk_summary_vars['total'].set(str(len(items)))
        self._risk_summary_vars['critical'].set(str(lvl_cnt['CRITICAL']))
        self._risk_summary_vars['high'].set(str(lvl_cnt['HIGH']))
        self._risk_summary_vars['medium'].set(str(lvl_cnt['MEDIUM']))
        self._risk_summary_vars['low'].set(str(lvl_cnt['LOW']))

    # ── Risk: scan worker ─────────────────────────────────────────────────────

    def _risk_run_scan(self):
        if self._risk_scan_id is None:
            messagebox.showwarning('Risk Scan', 'Import a BOM first.')
            return
        if self._risk_scanning:
            return
        items = self.risk_db.get_items(self._risk_scan_id)
        n = len(items)
        if not messagebox.askyesno('Run Scan',
                f'Scan {n} parts via Mouser?\n'
                f'Approx. {max(1, n // 10)} min.'):
            return
        self._risk_items_buf = [(it['id'], it['vendor_pn'], it['qty']) for it in items]
        self._risk_scanning = True
        self._risk_run_btn.grid_remove()
        self._risk_stop_btn.grid()
        self._risk_prog.configure(maximum=n, value=0)
        self._risk_prog_lbl.configure(text='Starting...', foreground='gray')
        threading.Thread(target=self._risk_scan_worker,
                         args=(list(self._risk_items_buf),), daemon=True).start()

    def _risk_stop_scan(self):
        self._risk_scanning = False

    def _risk_scan_worker(self, items):
        dk_disabled    = False
        dk_fail_streak = 0

        for i, (item_id, vpn, qty) in enumerate(items):
            if not self._risk_scanning:
                break
            self.root.after(0, self._risk_prog_update, i + 1, len(items), vpn)

            m_stk = lt_wk = dk_stk = -1
            lifecycle = ''

            # ── Mouser (sequential — confirmed fast) ────────────────────
            try:
                parts = self.client.search_by_part(vpn, _retries=2)
                if parts:
                    p         = parts[0]
                    m_stk     = self._parse_mouser_stock(p.get('Availability', ''))
                    lt_wk     = self._parse_lead_time_wk(p.get('LeadTime', '') or '')
                    lifecycle = p.get('LifecycleStatus', '') or ''
            except Exception:
                pass

            # ── DigiKey (thread + hard 5s timeout) ──────────────────────
            if not dk_disabled:
                dk_result = [None]
                dk_exc    = [False]

                def _qdk(r=dk_result, e=dk_exc, v=vpn):
                    try:
                        prods = self.dk_client.search_by_mpn(v, limit=3)
                        r[0]  = prods[0] if prods else None
                    except Exception:
                        e[0] = True

                t = threading.Thread(target=_qdk, daemon=True)
                t.start()
                t.join(timeout=5)

                if t.is_alive():
                    # SSL / network hang — count as failure
                    dk_fail_streak += 1
                    if dk_fail_streak >= 3:
                        dk_disabled = True
                        self.root.after(0, self._risk_prog_lbl.configure,
                                        {'text': 'DigiKey timeout — Mouser only',
                                         'foreground': 'orange'})
                elif dk_result[0] is not None:
                    dk_fail_streak = 0
                    p      = dk_result[0]
                    dk_stk = p.get('QuantityAvailable', 0) or 0
                    st     = (p.get('ProductStatus') or {}).get('Status', '')
                    if st and not lifecycle:
                        lifecycle = st
                    if lt_wk < 0:
                        dk_lt = p.get('ManufacturerLeadWeeks')
                        if dk_lt is not None:
                            try:
                                lt_wk = int(float(str(dk_lt)))
                            except (ValueError, TypeError):
                                pass
                else:
                    # API error or part not found — reset streak
                    if dk_exc[0]:
                        dk_fail_streak += 1
                        if dk_fail_streak >= 3:
                            dk_disabled = True
                            self.root.after(0, self._risk_prog_lbl.configure,
                                            {'text': 'DigiKey error — Mouser only',
                                             'foreground': 'orange'})
                    else:
                        dk_fail_streak = 0   # not found is OK, not an error

            # ── Score & persist ─────────────────────────────────────────
            try:
                ri, rl, rs, rlife, total, level = self._calc_risk(
                    m_stk, dk_stk, lt_wk, lifecycle)
                self.risk_db.update_item(item_id,
                    mouser_stock=m_stk, dk_stock=dk_stk,
                    lead_time_wk=lt_wk, lifecycle=lifecycle,
                    risk_inv=ri, risk_lt=rl, risk_src=rs,
                    risk_life=rlife, risk_total=total, risk_level=level)
                self.root.after(0, self._risk_update_row, item_id,
                                m_stk, dk_stk, lt_wk, lifecycle, total, level)
            except Exception:
                pass
            time.sleep(0.1)

        self.root.after(0, self._risk_scan_done)

    def _risk_prog_update(self, current, total, vpn):
        self._risk_prog.configure(value=current)
        self._risk_prog_lbl.configure(
            text=f'Scanning {current}/{total}  {vpn[:30]}',
            foreground='gray')

    def _risk_scan_done(self):
        self._risk_scanning = False
        self._risk_stop_btn.grid_remove()
        self._risk_run_btn.grid()
        self._risk_prog_lbl.configure(text='Scan complete.', foreground='green')
        items = self.risk_db.get_items(self._risk_scan_id)
        self._risk_populate_tv(items)
        self._risk_update_summary(items)

    # ── Risk: scoring ─────────────────────────────────────────────────────────

    @staticmethod
    def _parse_mouser_stock(avail):
        if not avail:
            return -1
        # Primary: "3,200 In Stock"
        m = re.search(r'([\d,]+)\s+In\s+Stock', avail, re.IGNORECASE)
        if m:
            try:
                return int(m.group(1).replace(',', ''))
            except Exception:
                pass
        # Fallback: leading number e.g. "0", "1234 units", non-standard formats
        m2 = re.match(r'([\d,]+)', avail.strip())
        if m2:
            try:
                return int(m2.group(1).replace(',', ''))
            except Exception:
                pass
        # Non-numeric strings ("Non-Stock", "Factory Stock") → unknown
        return -1

    @staticmethod
    def _parse_lead_time_wk(avail):
        import re as _re
        if not avail:
            return -1
        m = _re.search(r'(\d+)\s*[Ww]eek', avail)
        if m:
            return int(m.group(1))
        return -1

    @staticmethod
    @staticmethod
    def _extract_rlc_query(desc):
        """Extract searchable value+package string from an RLC component description.
        e.g. 'RES SMD 1/10W 10kohm F 0402' -> '10KOHM 0402'
             'CAP CER 100nF 25V X5R 0402'  -> '100NF 0402'
        Returns '' for non-RLC or unparseable descriptions.
        """
        if not desc:
            return ''
        d = desc.strip().upper()
        rlc_pfx = ('RES', 'CAP', 'IND', 'RESISTOR', 'CAPACITOR', 'INDUCTOR',
                   'FERRITE', 'MLCC', 'BEAD', 'CHOKE', 'VARISTOR')
        if not any(d.startswith(p) for p in rlc_pfx):
            return ''
        parts = []
        pkg_re  = re.compile(r'^(0[12]\d\d|1[02]\d\d|20\d\d|25\d\d|01005)$')
        # Component value: digits + optional SI prefix + optional unit
        val_re  = re.compile(r'^\d+(?:\.\d+)?[PNUMKGR]?(?:OHM|F|H)?$')
        # Fractional resistor value: e.g. 4R7, 0R01
        frac_re = re.compile(r'^\d+R\d+$')
        for tok in d.split():
            if pkg_re.match(tok):
                parts.append(tok)
            elif frac_re.match(tok):
                parts.append(tok)
            elif (val_re.match(tok)
                  and not tok.endswith('V')   # skip voltage e.g. 25V
                  and not tok.endswith('W')   # skip power e.g. 100W
                  and tok not in ('F', 'H')):  # skip bare unit letters
                parts.append(tok)
        return ' '.join(parts)

    @staticmethod
    def _calc_risk(m_stk, dk_stk, lt_wk, lifecycle):
        total_stk = max(0, m_stk) + max(0, dk_stk)
        # Inventory
        if total_stk < 100:   risk_inv = 40
        elif total_stk < 1000: risk_inv = 20
        else:                  risk_inv = 5
        # Lead time
        if lt_wk < 0:     risk_lt = 15   # unknown → medium
        elif lt_wk > 52:  risk_lt = 40
        elif lt_wk > 26:  risk_lt = 20
        else:             risk_lt = 5
        # Source
        suppliers = sum(1 for s in (m_stk, dk_stk) if s > 0)
        if suppliers == 0:   risk_src = 20
        elif suppliers == 1: risk_src = 10
        else:                risk_src = 5
        # Lifecycle
        lc = (lifecycle or '').upper()
        if any(k in lc for k in ('EOL', 'OBSOLETE', 'DISCONTINUED')):
            risk_life = 40
        elif any(k in lc for k in ('NRND', 'NOT RECOMMENDED')):
            risk_life = 20
        else:
            risk_life = 0
        total = risk_inv + risk_lt + risk_src + risk_life
        if total > 100:  level = 'CRITICAL'
        elif total > 60: level = 'HIGH'
        elif total > 30: level = 'MEDIUM'
        else:            level = 'LOW'
        return risk_inv, risk_lt, risk_src, risk_life, total, level

    # ── Risk: right-click context ─────────────────────────────────────────────

    def _risk_tv_ctx(self, event):
        iid = self._risk_tv.identify_row(event.y)
        if not iid:
            return
        self._risk_tv.selection_set(iid)
        vals = self._risk_tv.item(iid, 'values')
        vpn  = vals[1] if vals else ''
        desc = vals[2] if vals else ''
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label=f'Copy PN:  {vpn}',
                      command=lambda v=vpn: self._risk_copy_vpn(v))
        m.add_separator()
        m.add_command(label=f'Single Search: {vpn}',
                      command=lambda: (self.sv_term.set(vpn), self._do_single()))
        m.add_separator()
        m.add_command(label='View Alternatives...',
                      command=lambda v=vpn, d=desc: self._risk_alt_dialog(v, initial_desc=d))
        m.tk_popup(event.x_root, event.y_root)

    def _risk_copy_vpn(self, vpn=''):
        if not vpn:
            sel = self._risk_tv.selection()
            if sel:
                vpn = self._risk_tv.item(sel[0], 'values')[1]
        if vpn:
            self.root.clipboard_clear()
            self.root.clipboard_append(vpn)

    def _risk_export_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl not installed')
            return
        if not self._risk_scan_id:
            messagebox.showinfo('Export', 'No scan loaded.')
            return
        items = self.risk_db.get_items(self._risk_scan_id)
        if not items:
            messagebox.showinfo('Export', 'No data to export.')
            return
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'BOM_Risk_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'BOM Risk Scan'
        from openpyxl.styles import PatternFill, Font, Alignment
        hdr_fill = PatternFill('solid', fgColor='2E4057')
        hdr_font = Font(bold=True, color='FFFFFF')
        fill_map = {
            'CRITICAL': PatternFill('solid', fgColor='FFB3B3'),
            'HIGH':     PatternFill('solid', fgColor='FFD9B3'),
            'MEDIUM':   PatternFill('solid', fgColor='FFFACC'),
            'LOW':      PatternFill('solid', fgColor='D4FFCC'),
        }

        # ── Sheet 1: BOM Risk ────────────────────────────────────────────
        headers = ['No.', 'Vendor PN', 'Description', 'Qty',
                   'Mouser Stock', 'DigiKey Stock', 'Lead Time(wk)',
                   'Lifecycle', 'Inv Risk', 'LT Risk', 'Src Risk',
                   'Life Risk', 'Total Score', 'Risk Level']
        col_w = [5, 24, 30, 6, 13, 13, 14, 12, 10, 10, 10, 10, 12, 12]
        ws.append(headers)
        for cell, w in zip(ws[1], col_w):
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[cell.column_letter].width = w
        for i, it in enumerate(items, 1):
            row = [i, it['vendor_pn'], it['description'], it['qty'],
                   it['mouser_stock'] if it['mouser_stock'] >= 0 else '',
                   it['dk_stock']     if it['dk_stock']     >= 0 else '',
                   it['lead_time_wk'] if it['lead_time_wk'] >= 0 else '',
                   it['lifecycle'],
                   it['risk_inv'], it['risk_lt'], it['risk_src'], it['risk_life'],
                   it['risk_total'], it['risk_level']]
            ws.append(row)
            lvl = it['risk_level']
            if lvl in fill_map:
                for cell in ws[ws.max_row]:
                    cell.fill = fill_map[lvl]
        ws.freeze_panes = 'A2'

        # ── Sheet 2: Alt Parts ───────────────────────────────────────────
        ws2 = wb.create_sheet('Alt Parts')
        alt_hdrs = ['Primary VPN', 'Primary Description', 'Risk Level',
                    'Alt VPN', 'Alt Manufacturer', 'Compatibility', 'Note']
        alt_w    = [24, 34, 12, 24, 20, 20, 24]
        ws2.append(alt_hdrs)
        for cell, w in zip(ws2[1], alt_w):
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center')
            ws2.column_dimensions[cell.column_letter].width = w
        for it in items:
            alts = self.risk_db.get_alts(it['vendor_pn'])
            if not alts:
                continue
            lvl  = it['risk_level']
            fill = fill_map.get(lvl)
            for a in alts:
                ws2.append([it['vendor_pn'], it['description'], lvl,
                             a['alt_vpn'], a['alt_mfr'], a['compat'], a['note']])
                if fill:
                    for cell in ws2[ws2.max_row]:
                        cell.fill = fill
        ws2.freeze_panes = 'A2'

        wb.save(path)
        messagebox.showinfo('Done', f'Exported to:\n{path}')

    def _risk_alt_dialog(self, initial_vpn='', initial_desc=''):
        COMPAT = ['Pin-to-Pin', 'Function Compatible',
                  'PCB Change Required', 'Firmware Change Required']
        dlg = tk.Toplevel(self.root)
        dlg.title('Alternative Parts Database')
        dlg.geometry('700x460')
        dlg.resizable(True, True)

        # ── Top: search bar ──────────────────────────────────────────────
        top = tk.Frame(dlg)
        top.pack(fill='x', padx=10, pady=(10, 4))
        ttk.Label(top, text='Primary VPN:', font=FONT_UI).pack(side='left')
        sv_pri = tk.StringVar(value=initial_vpn)
        ttk.Entry(top, textvariable=sv_pri, width=22).pack(side='left', padx=6)
        ttk.Button(top, text='Search',
                   command=lambda: _refresh(sv_pri.get())).pack(side='left')
        ttk.Button(top, text='Show All',
                   command=lambda: _refresh('')).pack(side='left', padx=4)

        # ── Bottom: action buttons ───────────────────────────────────────
        bot = tk.Frame(dlg)
        bot.pack(side='bottom', fill='x', padx=10, pady=(4, 10))

        # ── Center: treeview ─────────────────────────────────────────────
        tv_frame = tk.Frame(dlg)
        tv_frame.pack(fill='both', expand=True, padx=10, pady=(0, 4))
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        cols = ('primary', 'alt', 'mfr', 'compat', 'note')
        tv = ttk.Treeview(tv_frame, columns=cols, show='headings', selectmode='browse')
        for cid, hdr, w in [('primary', 'Primary VPN',  130),
                             ('alt',     'Alt VPN',      130),
                             ('mfr',     'Manufacturer', 110),
                             ('compat',  'Compatibility',130),
                             ('note',    'Note',         150)]:
            tv.heading(cid, text=hdr)
            tv.column(cid, width=w, anchor='w')
        vsb = ttk.Scrollbar(tv_frame, orient='vertical', command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')

        def _refresh(vpn=''):
            tv.delete(*tv.get_children())
            for r in self.risk_db.get_alts(vpn):
                tv.insert('', 'end', iid=str(r['id']),
                          values=(r['primary_vpn'], r['alt_vpn'],
                                  r['alt_mfr'], r['compat'], r['note']))

        _refresh(initial_vpn)

        # ── Add dialog ───────────────────────────────────────────────────
        def _add():
            add_dlg = tk.Toplevel(dlg)
            add_dlg.title('Add Alternative')
            add_dlg.resizable(False, False)
            add_dlg.grab_set()
            add_dlg.columnconfigure(1, weight=1)
            flds = {}

            def _browse_db():
                rlc_q = MouserApp._extract_rlc_query(initial_desc)
                picker = _AltPartPicker(add_dlg, self.db, initial_query=rlc_q)
                add_dlg.wait_window(picker)
                if picker.result:
                    flds['alt_vpn'].set(picker.result[0])
                    flds['alt_mfr'].set(picker.result[1])

            row_offset = 0

            # Row 0: primary part description (read-only, shown when available)
            if initial_desc:
                ttk.Label(add_dlg, text='Primary Part:', font=FONT_UI).grid(
                    row=0, column=0, sticky='nw', padx=12, pady=(10, 2))
                tk.Label(add_dlg, text=initial_desc, fg='#555555',
                         font=FONT_UI, wraplength=300, justify='left',
                         anchor='w').grid(
                    row=0, column=1, columnspan=2, sticky='ew',
                    padx=8, pady=(10, 2))
                row_offset = 1

            # Primary VPN
            r = row_offset
            ttk.Label(add_dlg, text='Primary VPN:', font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=3)
            sv = tk.StringVar(value=sv_pri.get())
            ttk.Entry(add_dlg, textvariable=sv, width=28).grid(
                row=r, column=1, columnspan=2, sticky='ew', padx=8, pady=3)
            flds['primary_vpn'] = sv

            # Alt VPN + Browse button
            r = row_offset + 1
            ttk.Label(add_dlg, text='Alt VPN:', font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=3)
            sv = tk.StringVar()
            ttk.Entry(add_dlg, textvariable=sv, width=22).grid(
                row=r, column=1, sticky='ew', padx=(8, 2), pady=3)
            ttk.Button(add_dlg, text='Browse...', width=9,
                       command=_browse_db).grid(
                row=r, column=2, sticky='w', padx=(0, 8), pady=3)
            flds['alt_vpn'] = sv

            # Manufacturer
            r = row_offset + 2
            ttk.Label(add_dlg, text='Manufacturer:', font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=3)
            sv = tk.StringVar()
            ttk.Entry(add_dlg, textvariable=sv, width=28).grid(
                row=r, column=1, columnspan=2, sticky='ew', padx=8, pady=3)
            flds['alt_mfr'] = sv

            # Note
            r = row_offset + 3
            ttk.Label(add_dlg, text='Note:', font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=3)
            sv = tk.StringVar()
            ttk.Entry(add_dlg, textvariable=sv, width=28).grid(
                row=r, column=1, columnspan=2, sticky='ew', padx=8, pady=3)
            flds['note'] = sv

            # Compatibility
            r = row_offset + 4
            ttk.Label(add_dlg, text='Compatibility:', font=FONT_UI).grid(
                row=r, column=0, sticky='w', padx=12, pady=3)
            sv_c = tk.StringVar(value=COMPAT[0])
            ttk.Combobox(add_dlg, textvariable=sv_c, values=COMPAT,
                         state='readonly', width=26).grid(
                row=r, column=1, columnspan=2, sticky='ew', padx=8, pady=3)

            def _ok2():
                pri = flds['primary_vpn'].get().strip()
                alt = flds['alt_vpn'].get().strip()
                if not pri or not alt:
                    messagebox.showwarning('Add Alternative',
                                           'Primary VPN and Alt VPN are required.',
                                           parent=add_dlg)
                    return
                self.risk_db.add_alt(pri, alt,
                                     flds['alt_mfr'].get().strip(),
                                     sv_c.get(),
                                     flds['note'].get().strip())
                add_dlg.destroy()
                sv_pri.set(pri)
                _refresh(pri)

            r = row_offset + 5
            bf2 = ttk.Frame(add_dlg)
            bf2.grid(row=r, column=0, columnspan=3, padx=12, pady=(4, 12))
            ttk.Button(bf2, text='Cancel', command=add_dlg.destroy).pack(side='right', padx=4)
            ttk.Button(bf2, text='Add',    command=_ok2).pack(side='right')

        def _del():
            iid = tv.focus()
            if not iid:
                messagebox.showinfo('Delete', 'Select a row first.', parent=dlg)
                return
            if messagebox.askyesno('Delete', 'Remove this alternative?', parent=dlg):
                self.risk_db.delete_alt(int(iid))
                _refresh(sv_pri.get())

        ttk.Button(bot, text='+ Add',    command=_add).pack(side='left')
        ttk.Button(bot, text='✕ Delete', command=_del).pack(side='left', padx=6)
        ttk.Button(bot, text='Close',    command=dlg.destroy).pack(side='right')

    # ── Tab: 備料清單 ─────────────────────────────────────────────────────────

    def _build_bom_tab(self):
        f = self.tab_bom
        f.columnconfigure(1, weight=1)
        f.rowconfigure(2, weight=1)
        p = dict(padx=6, pady=3)

        ttk.Button(f, text='Open Orcad Bom File', command=self._bom_open,
                   width=20).grid(row=0, column=0, sticky='w', **p)
        self._sv_bom = tk.StringVar()
        ttk.Entry(f, textvariable=self._sv_bom,
                  state='readonly').grid(row=0, column=1, sticky='ew', **p)

        ttk.Button(f, text='產生備料清單',
                   command=self._bom_generate).grid(row=1, column=0, sticky='w', **p)

        self._txt_bom = scrolledtext.ScrolledText(
            f, font=FONT_TX, wrap='none', bg='#f8f8f8')
        self._txt_bom.grid(row=2, column=0, columnspan=2,
                           sticky='nsew', padx=6, pady=3)
        self._txt_bom.tag_configure('warn_dup', background='#FFA500', foreground='#000000')
        self._txt_bom.tag_configure('no_code',  background='#FFA500', foreground='#000000')
        self._txt_bom.tag_configure('non_std',  background='#FFFACD', foreground='#000000')
        self._txt_bom.tag_configure('warn_qty', background='#FF6B6B', foreground='#000000')

        row3 = ttk.Frame(f)
        row3.grid(row=3, column=0, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Button(row3, text='Export Excel',
                   command=self._bom_export_excel).pack(side='left')
        ttk.Button(row3, text='Copy All',
                   command=self._bom_copy_all).pack(side='left', padx=6)

    def _bom_open(self):
        path = filedialog.askopenfilename(
            title='Select Orcad BOM File',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if path:
            self._sv_bom.set(path)

    @staticmethod
    def _bom_group_by_code(items):
        groups, order = {}, []
        for item in items:
            if item.code not in groups:
                groups[item.code] = {'values': [], 'descs': [], 'qty': 0, 'locations': []}
                order.append(item.code)
            g = groups[item.code]
            if item.value not in g['values']:
                g['values'].append(item.value)
            if item.description not in g['descs']:
                g['descs'].append(item.description)
            g['qty']       += item.qty
            g['locations'] += item.locations
        return order, groups

    def _bom_generate(self):
        path = self._sv_bom.get()
        if not path:
            messagebox.showwarning('Warning', '請先選擇 Orcad BOM 檔案')
            return
        try:
            all_items = BomParser.parse(path)
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return

        def _is_no_code(code):
            if not code or not code.strip():
                return True
            c = code.strip().upper()
            return c in ('TBD', 'NEW10CODE') or c.startswith('COM')

        no_code_items = [i for i in all_items if _is_no_code(i.code)]
        coded_items   = [i for i in all_items if not _is_no_code(i.code)]

        order, groups = self._bom_group_by_code(coded_items)
        total_qty     = sum(g['qty'] for g in groups.values())
        self._bom_list_data     = [(code, groups[code]) for code in order]
        self._bom_no_code_items = no_code_items

        w = self._txt_bom
        w.delete('1.0', 'end')
        w.insert('end',
            f'備料清單  BOM：{os.path.basename(path)}\n'
            f'共 {len(order)} 種料號，總計 {total_qty:,} pcs\n'
            f'{"═"*80}\n'
            f'{"No.":<5} {"10-Code":<16} {"Qty":>6}  Value / Description\n'
            f'{"─"*80}\n')

        for i, code in enumerate(order):
            g    = groups[code]
            val  = g['values'][0]  if g['values'] else ''
            desc = g['descs'][0]   if g['descs']  else ''
            line = (f'{i+1:<5} {code:<16} {g["qty"]:>6}  '
                    f'{val[:28]:<28} / {desc[:32]}\n')
            line_start = w.index('end - 1c linestart')
            w.insert('end', line)
            if not code.strip().upper().endswith('M'):
                w.tag_add('non_std', line_start, w.index('end - 1c'))

            if len(g['values']) > 1:
                warn = f'{"":5} ⚠ Value 不一致: {" | ".join(g["values"])}\n'
                start = w.index('end - 1c linestart')
                w.insert('end', warn)
                end = w.index('end - 1c')
                w.tag_add('warn_dup', start, end)

            if len(g['descs']) > 1:
                warn = f'{"":5} ⚠ Desc 不一致:  {" | ".join(g["descs"])}\n'
                start = w.index('end - 1c linestart')
                w.insert('end', warn)
                end = w.index('end - 1c')
                w.tag_add('warn_dup', start, end)

            loc_count = len(g['locations'])
            if loc_count != g['qty']:
                warn = (f'{"":5} ⚠ Qty/Location 不符: '
                        f'BOM qty={g["qty"]}  實際 location={loc_count}\n')
                start = w.index('end - 1c linestart')
                w.insert('end', warn)
                end = w.index('end - 1c')
                w.tag_add('warn_qty', start, end)

        if no_code_items:
            w.insert('end',
                f'\n{"═"*80}\n'
                f'★ 沒有料號（共 {len(no_code_items)} 筆，不計入總量）\n'
                f'{"─"*80}\n'
                f'{"Value":<30} {"Qty":>6}  {"Description":<36}  Location\n'
                f'{"─"*80}\n')
            for item in no_code_items:
                lstr = ','.join(item.locations[:6])
                if len(item.locations) > 6:
                    lstr += f',... (+{len(item.locations)-6})'
                line = (f'{item.value[:29]:<30} {item.qty:>6}  '
                        f'{item.description[:35]:<36}  {lstr}\n')
                start = w.index('end - 1c linestart')
                w.insert('end', line)
                end = w.index('end - 1c')
                w.tag_add('no_code', start, end)

    def _bom_copy_all(self):
        text = self._txt_bom.get('1.0', 'end').strip()
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)

    def _bom_export_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        if not self._bom_list_data and not self._bom_no_code_items:
            messagebox.showwarning('Warning', '請先執行產生備料清單')
            return

        hdr_fill  = PatternFill('solid', fgColor='1F6CBF')
        hdr_font  = Font(bold=True, color='FFFFFF')
        fill_warn = PatternFill('solid', fgColor='FFA500')

        def xl_hdr(ws, headers, col_widths):
            ws.append(headers)
            for cell, w in zip(ws[1], col_widths):
                cell.fill = hdr_fill
                cell.font = hdr_font
                ws.column_dimensions[cell.column_letter].width = w
            ws.freeze_panes = 'A2'

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '備料清單'
        xl_hdr(ws,
            ['No', '10-Code', 'Qty', 'Value', 'Description', 'Locations', 'Remark'],
            [5, 16, 6, 28, 40, 80, 40])

        for i, (code, g) in enumerate(self._bom_list_data):
            val    = g['values'][0]  if g['values'] else ''
            desc   = g['descs'][0]   if g['descs']  else ''
            locs   = ','.join(g['locations'])
            remark_parts = []
            if len(g['values']) > 1:
                remark_parts.append('⚠ Value 不一致: ' + ' | '.join(g['values']))
            if len(g['descs']) > 1:
                remark_parts.append('⚠ Desc 不一致: '  + ' | '.join(g['descs']))
            loc_count = len(g['locations'])
            if loc_count != g['qty']:
                remark_parts.append(
                    f'⚠ Qty/Location 不符: BOM qty={g["qty"]} 實際 location={loc_count}')
            remark = '  '.join(remark_parts)
            ws.append([i + 1, code, g['qty'], val, desc, locs, remark])
            if remark:
                row = ws.max_row
                for col in range(1, 8):
                    ws.cell(row, col).fill = fill_warn

        if self._bom_no_code_items:
            ws2 = wb.create_sheet('沒有料號')
            xl_hdr(ws2,
                ['Value', 'Qty', 'Description', 'Locations'],
                [28, 6, 40, 80])
            for item in self._bom_no_code_items:
                ws2.append([item.value, item.qty, item.description,
                             ','.join(item.locations)])
                row = ws2.max_row
                for col in range(1, 5):
                    ws2.cell(row, col).fill = fill_warn

        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'BOM_備料清單_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        try:
            wb.save(path)
            messagebox.showinfo('Done', f'已儲存至\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    # ── Tab: Note ─────────────────────────────────────────────────────────────

    def _build_note_tab(self):
        os.makedirs(NOTE_FOLDER, exist_ok=True)
        f = self.tab_note
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)

        # ── Left: file list ───────────────────────────────────────
        left = ttk.Frame(pw, width=260)
        left.grid_propagate(False)
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        # Search bar
        search_f = ttk.Frame(left)
        search_f.grid(row=0, column=0, columnspan=2, sticky='ew', padx=4, pady=(4, 2))
        search_f.columnconfigure(0, weight=1)
        self._note_search_sv = tk.StringVar()
        se = ttk.Entry(search_f, textvariable=self._note_search_sv)
        se.grid(row=0, column=0, sticky='ew')
        se.bind('<Return>', lambda e: self._note_search())
        ttk.Button(search_f, text='🔍', width=3,
                   command=self._note_search).grid(row=0, column=1, padx=(2, 0))
        ttk.Button(search_f, text='✕', width=3,
                   command=self._note_clear_search).grid(row=0, column=2, padx=(2, 0))

        self._note_tv = ttk.Treeview(left, show='tree', selectmode='browse')
        self._note_tv.column('#0', width=240, minwidth=120)
        self._note_tv.grid(row=1, column=0, sticky='nsew')
        self._note_tv.bind('<Double-1>', self._note_dbl_click)

        vsb = ttk.Scrollbar(left, orient='vertical', command=self._note_tv.yview)
        vsb.grid(row=1, column=1, sticky='ns')
        self._note_tv.configure(yscrollcommand=vsb.set)

        # Category parent nodes
        self._note_cat_folder = self._note_tv.insert(
            '', 'end', text='📁 Note Folder', open=True, tags=('cat',))
        self._note_cat_extern = self._note_tv.insert(
            '', 'end', text='🔗 External Links', open=True, tags=('cat',))
        self._note_tv.tag_configure(
            'cat', font=('Microsoft JhengHei UI', 9, 'bold'))
        self._note_cat_search = None   # created on demand by _note_search

        self._note_iid_path: dict[str, str] = {}
        self._note_iid_kind: dict[str, str] = {}   # iid -> 'folder' | 'file' | 'extern'
        self._note_current_path: str = ''
        self._note_search_active = False
        self._note_md_after_id = None
        self._note_tv.bind('<<TreeviewSelect>>', lambda e: self._note_load_selected())

        btn_lf = ttk.Frame(left)
        btn_lf.grid(row=2, column=0, columnspan=2, sticky='ew', padx=4, pady=4)
        ttk.Button(btn_lf, text='Add File',
                   command=self._note_add_file).pack(side='left')
        ttk.Button(btn_lf, text='New Folder',
                   command=self._note_new_folder).pack(side='left', padx=2)
        ttk.Button(btn_lf, text='Rename',
                   command=self._note_rename).pack(side='left')
        ttk.Button(btn_lf, text='Delete',
                   command=self._note_delete).pack(side='left', padx=2)

        btn_lf2 = ttk.Frame(left)
        btn_lf2.grid(row=3, column=0, columnspan=2, sticky='ew', padx=4, pady=(0, 4))
        ttk.Button(btn_lf2, text='Refresh',
                   command=self._note_refresh).pack(side='left')
        ttk.Button(btn_lf2, text='⚙ Npp',
                   command=self._note_set_npp).pack(side='right')

        pw.add(left, weight=1)

        # ── Right: paste & save ───────────────────────────────────
        right = ttk.Frame(pw)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        top_rf = ttk.Frame(right)
        top_rf.grid(row=0, column=0, sticky='ew', padx=4, pady=(4, 0))
        top_rf.columnconfigure(0, weight=1)
        self._note_link_lbl = tk.Label(top_rf, text='', fg='#007AFF', cursor='',
                                       font=FONT_UI, anchor='w')
        self._note_link_lbl.grid(row=0, column=0, sticky='w')
        self._note_md_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top_rf, text='Markdown 標記', variable=self._note_md_var,
                        command=self._note_toggle_md).grid(row=0, column=1, sticky='e')

        self._note_txt = scrolledtext.ScrolledText(
            right, font=FONT_TX, wrap='word', undo=True)
        self._note_txt.grid(row=1, column=0, sticky='nsew', padx=4, pady=(4, 0))
        self._note_txt.bind('<KeyRelease>', self._note_on_keyrelease)
        self._note_md_tags_ready = False

        btn_rf = ttk.Frame(right)
        btn_rf.grid(row=2, column=0, sticky='ew', padx=4, pady=4)
        ttk.Button(btn_rf, text='Save',
                   command=self._note_save_current).pack(side='left')
        ttk.Button(btn_rf, text='Save as New Note',
                   command=self._note_save_text).pack(side='left', padx=4)
        ttk.Button(btn_rf, text='Clear',
                   command=lambda: self._note_txt.delete('1.0', 'end')).pack(
                   side='left')
        ttk.Button(btn_rf, text='Open in Notepad++',
                   command=self._note_open_npp).pack(side='right')

        pw.add(right, weight=2)

        self._note_refresh()

    # ── Tree population ──────────────────────────────────────────────────────

    def _note_populate_folder(self, parent_iid, dir_path):
        try:
            entries = sorted(os.listdir(dir_path))
        except OSError:
            return
        # folders first, then files, each alphabetical
        dirs  = [n for n in entries if os.path.isdir(os.path.join(dir_path, n))]
        files = [n for n in entries if os.path.isfile(os.path.join(dir_path, n))]
        for name in dirs:
            full = os.path.join(dir_path, name)
            iid  = self._note_tv.insert(parent_iid, 'end', text=f'📁 {name}', open=False)
            self._note_iid_path[iid] = full
            self._note_iid_kind[iid] = 'folder'
            self._note_populate_folder(iid, full)
        for name in files:
            full = os.path.join(dir_path, name)
            iid  = self._note_tv.insert(parent_iid, 'end', text=name)
            self._note_iid_path[iid] = full
            self._note_iid_kind[iid] = 'file'

    def _note_refresh(self):
        for iid in self._note_tv.get_children(self._note_cat_folder):
            self._note_tv.delete(iid)
        for iid in self._note_tv.get_children(self._note_cat_extern):
            self._note_tv.delete(iid)
        if self._note_cat_search is not None:
            self._note_tv.delete(self._note_cat_search)
            self._note_cat_search = None
        self._note_search_active = False
        self._note_iid_path = {}
        self._note_iid_kind = {}

        os.makedirs(NOTE_FOLDER, exist_ok=True)
        self._note_populate_folder(self._note_cat_folder, NOTE_FOLDER)

        cfg = _load_app_cfg()
        self._note_npp_path = cfg.get('npp_path', '')
        for entry in cfg.get('notes', []):
            iid = self._note_tv.insert(
                self._note_cat_extern, 'end', text=entry.get('label', ''))
            self._note_iid_path[iid] = entry.get('path', '')
            self._note_iid_kind[iid] = 'extern'

    def _note_save_extern(self):
        cfg = _load_app_cfg()
        cfg['npp_path'] = self._note_npp_path
        cfg['notes'] = [
            {'label': self._note_tv.item(iid)['text'],
             'path':  self._note_iid_path.get(iid, '')}
            for iid in self._note_tv.get_children(self._note_cat_extern)
        ]
        _save_app_cfg(cfg)

    def _note_add_file(self):
        path = filedialog.askopenfilename(
            title='Select File to Link',
            filetypes=[('All files', '*.*')])
        if not path:
            return
        label = os.path.basename(path)
        iid   = self._note_tv.insert(self._note_cat_extern, 'end', text=label)
        self._note_iid_path[iid] = path
        self._note_iid_kind[iid] = 'extern'
        self._note_save_extern()

    def _note_target_dir(self):
        """Directory a new file/folder should be created in: the currently
        selected folder (or the parent of a selected file), else NOTE_FOLDER."""
        iid = self._note_tv.focus()
        if iid and iid in self._note_iid_kind:
            path = self._note_iid_path.get(iid, '')
            kind = self._note_iid_kind[iid]
            if kind == 'folder' and os.path.isdir(path):
                return path
            if kind == 'file' and os.path.isfile(path):
                return os.path.dirname(path)
        return NOTE_FOLDER

    def _note_new_folder(self):
        dlg = _TextPromptDialog(self.root, 'New Folder', 'Folder name:')
        self.root.wait_window(dlg)
        if not dlg.result:
            return
        target = os.path.join(self._note_target_dir(), dlg.result)
        try:
            os.makedirs(target, exist_ok=False)
        except FileExistsError:
            messagebox.showwarning('New Folder', '這個資料夾已經存在')
            return
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return
        self._note_refresh()

    # ── Loading / editing ────────────────────────────────────────────────────

    def _note_path_iid(self, path):
        for iid, p in self._note_iid_path.items():
            if p == path and self._note_iid_kind.get(iid) == 'file':
                return iid
        return None

    def _note_dbl_click(self, event):
        iid = self._note_tv.identify_row(event.y)
        if not iid or self._note_iid_kind.get(iid) != 'file':
            return   # folders just expand/collapse via the normal tree behaviour
        self._note_open_npp()

    def _note_load_selected(self):
        iid = self._note_tv.focus()
        if not iid or self._note_iid_kind.get(iid) not in ('file', 'extern'):
            return
        path = self._note_iid_path.get(iid, '')
        if not path or not os.path.isfile(path):
            return

        if self._note_txt.edit_modified():
            resp = messagebox.askyesnocancel(
                'Unsaved Changes', '目前筆記有未儲存的修改，要先儲存嗎？')
            if resp is None:
                prev_iid = self._note_path_iid(self._note_current_path)
                if prev_iid:
                    self._note_tv.selection_set(prev_iid)
                    self._note_tv.focus(prev_iid)
                return
            if resp:
                self._note_save_current()

        try:
            with open(path, encoding='utf-8', errors='replace') as fh:
                content = fh.read()
        except Exception:
            return
        self._note_txt.delete('1.0', 'end')
        self._note_txt.insert('1.0', content)
        self._note_txt.edit_reset()
        self._note_txt.edit_modified(False)
        self._note_current_path = path
        if self._note_md_var.get():
            self._note_apply_md_highlight()
        self._note_refresh_linked_label()

    def _note_save_current(self):
        path = self._note_current_path
        if not path:
            messagebox.showinfo('Note', '尚未載入任何檔案，請使用「Save as New Note」')
            return
        try:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write(self._note_txt.get('1.0', 'end'))
            self._note_txt.edit_modified(False)
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _note_open_npp(self):
        iid = self._note_tv.focus()
        if not iid or self._note_iid_kind.get(iid) not in ('file', 'extern'):
            messagebox.showinfo('Note', '請先選取一個檔案')
            return
        file_path = self._note_iid_path.get(iid, '')
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror('Error', f'檔案不存在：\n{file_path}')
            return
        npp = self._note_npp_path
        if not npp or not os.path.exists(npp):
            for candidate in [
                r'C:\Program Files\Notepad++\notepad++.exe',
                r'C:\Program Files (x86)\Notepad++\notepad++.exe',
            ]:
                if os.path.exists(candidate):
                    npp = candidate
                    self._note_npp_path = npp
                    self._note_save_extern()
                    break
            else:
                messagebox.showerror(
                    'Error',
                    'Notepad++ 路徑未設定或找不到\n請按「⚙ Npp」設定')
                return
        try:
            subprocess.Popen([npp, file_path])
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _note_delete(self):
        iid = self._note_tv.focus()
        kind = self._note_iid_kind.get(iid)
        if not iid or kind is None:
            return
        parent = self._note_tv.parent(iid)
        fname  = self._note_tv.item(iid)['text']
        fpath  = self._note_iid_path.get(iid, '')
        if kind == 'file':
            if not messagebox.askyesno('Delete', f'刪除檔案？\n{fname}'):
                return
            try:
                os.remove(fpath)
            except Exception as e:
                messagebox.showerror('Error', str(e))
                return
        elif kind == 'folder':
            if not messagebox.askyesno(
                    'Delete Folder', f'刪除資料夾「{fname}」及裡面所有檔案？此動作無法復原。'):
                return
            try:
                import shutil
                shutil.rmtree(fpath)
            except Exception as e:
                messagebox.showerror('Error', str(e))
                return
        # elif kind == 'extern': just unlink, the original file is untouched
        del self._note_iid_path[iid]
        del self._note_iid_kind[iid]
        self._note_tv.delete(iid)
        if parent == self._note_cat_extern:
            self._note_save_extern()

    def _note_rename(self):
        iid = self._note_tv.focus()
        kind = self._note_iid_kind.get(iid)
        if not iid or kind is None:
            return
        old_name = self._note_tv.item(iid)['text']
        if kind == 'folder':
            old_name = old_name.replace('📁 ', '', 1)
        dlg = _TextPromptDialog(self.root, 'Rename', 'New name:', initial=old_name)
        self.root.wait_window(dlg)
        if not dlg.result or dlg.result == old_name:
            return
        new_name = dlg.result

        if kind == 'extern':
            self._note_tv.item(iid, text=new_name)
            self._note_save_extern()
            return

        old_path = self._note_iid_path[iid]
        new_path = os.path.join(os.path.dirname(old_path), new_name)
        if os.path.exists(new_path):
            messagebox.showwarning('Rename', '目標名稱已存在')
            return
        try:
            os.rename(old_path, new_path)
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return
        if self._note_current_path == old_path:
            self._note_current_path = new_path
        elif self._note_current_path.startswith(old_path + os.sep):
            self._note_current_path = new_path + self._note_current_path[len(old_path):]
        self._note_refresh()

    def _note_save_text(self):
        text = self._note_txt.get('1.0', 'end').strip()
        if not text:
            messagebox.showinfo('Note', '文字區域為空，無法儲存')
            return
        target_dir = self._note_target_dir()
        os.makedirs(target_dir, exist_ok=True)
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            title='Save Note',
            initialdir=target_dir,
            initialfile=f'note_{ts}.txt',
            filetypes=[('Text', '*.txt'), ('All', '*.*')])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write(self._note_txt.get('1.0', 'end'))
            self._note_refresh()
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _note_set_npp(self):
        path = filedialog.askopenfilename(
            title='Select Notepad++.exe',
            initialdir=r'C:\Program Files',
            filetypes=[('Executable', '*.exe'), ('All', '*.*')])
        if path:
            self._note_npp_path = path
            self._note_save_extern()
            messagebox.showinfo('Done', f'Notepad++ 路徑已設定：\n{path}')

    # ── Search ────────────────────────────────────────────────────────────────

    def _note_search(self):
        query = self._note_search_sv.get().strip()
        if not query:
            self._note_clear_search()
            return
        matches = []
        for root_dir, _dirs, files in os.walk(NOTE_FOLDER):
            for fname in files:
                full = os.path.join(root_dir, fname)
                hit_name = query.lower() in fname.lower()
                snippet = ''
                if not hit_name:
                    try:
                        with open(full, encoding='utf-8', errors='replace') as fh:
                            content = fh.read()
                    except Exception:
                        continue
                    idx = content.lower().find(query.lower())
                    if idx < 0:
                        continue
                    start = max(0, idx - 20)
                    snippet = content[start:idx + len(query) + 20].replace('\n', ' ')
                matches.append((os.path.relpath(full, NOTE_FOLDER), full, snippet))

        for iid in self._note_tv.get_children(self._note_cat_folder):
            self._note_tv.detach(iid)
        for iid in self._note_tv.get_children(self._note_cat_extern):
            self._note_tv.detach(iid)
        if self._note_cat_search is not None:
            self._note_tv.delete(self._note_cat_search)
        self._note_cat_search = self._note_tv.insert(
            '', 0, text=f'🔍 Results ({len(matches)})', open=True, tags=('cat',))
        self._note_search_active = True
        for relpath, full, snippet in matches:
            label = relpath if not snippet else f'{relpath}  —  …{snippet}…'
            iid = self._note_tv.insert(self._note_cat_search, 'end', text=label)
            self._note_iid_path[iid] = full
            self._note_iid_kind[iid] = 'file'

    def _note_clear_search(self):
        self._note_search_sv.set('')
        if not self._note_search_active:
            return
        self._note_refresh()

    # ── Markdown highlighting ────────────────────────────────────────────────

    def _note_ensure_md_tags(self):
        if self._note_md_tags_ready:
            return
        txt = self._note_txt
        base_family = FONT_TX[0]
        txt.tag_configure('md_h1', font=(base_family, 16, 'bold'))
        txt.tag_configure('md_h2', font=(base_family, 13, 'bold'))
        txt.tag_configure('md_h3', font=(base_family, 11, 'bold'))
        txt.tag_configure('md_bullet', foreground='#007AFF')
        txt.tag_configure('md_code', font=('Courier New', FONT_TX[1]),
                          background='#F2F2F7', foreground='#AF52DE')
        txt.tag_configure('md_bold', font=(base_family, FONT_TX[1], 'bold'))
        self._note_md_tags_ready = True

    def _note_apply_md_highlight(self):
        self._note_ensure_md_tags()
        txt = self._note_txt
        for tag in ('md_h1', 'md_h2', 'md_h3', 'md_bullet', 'md_code', 'md_bold'):
            txt.tag_remove(tag, '1.0', 'end')

        content = txt.get('1.0', 'end-1c')
        for i, line in enumerate(content.split('\n'), start=1):
            stripped = line.lstrip()
            if stripped.startswith('### '):
                txt.tag_add('md_h3', f'{i}.0', f'{i}.end')
            elif stripped.startswith('## '):
                txt.tag_add('md_h2', f'{i}.0', f'{i}.end')
            elif stripped.startswith('# '):
                txt.tag_add('md_h1', f'{i}.0', f'{i}.end')
            elif stripped.startswith(('- ', '* ')):
                indent = len(line) - len(stripped)
                txt.tag_add('md_bullet', f'{i}.{indent}', f'{i}.{indent + 1}')
            for m in re.finditer(r'`([^`]+)`', line):
                txt.tag_add('md_code', f'{i}.{m.start()}', f'{i}.{m.end()}')
            for m in re.finditer(r'\*\*([^*]+)\*\*', line):
                txt.tag_add('md_bold', f'{i}.{m.start()}', f'{i}.{m.end()}')

    def _note_toggle_md(self):
        if self._note_md_var.get():
            self._note_apply_md_highlight()
        else:
            self._note_ensure_md_tags()
            for tag in ('md_h1', 'md_h2', 'md_h3', 'md_bullet', 'md_code', 'md_bold'):
                self._note_txt.tag_remove(tag, '1.0', 'end')

    def _note_on_keyrelease(self, event=None):
        if not self._note_md_var.get():
            return
        if self._note_md_after_id:
            self.root.after_cancel(self._note_md_after_id)
        self._note_md_after_id = self.root.after(400, self._note_apply_md_highlight)

    # ── Reminder linking (display side — links are created from Reminders) ────

    def _note_refresh_linked_label(self):
        if not self._note_current_path:
            self._note_link_lbl.configure(text='', cursor='')
            self._note_link_ids = []
            return
        rows = self.todo_db.get_items_by_note(self._note_current_path)
        if not rows:
            self._note_link_lbl.configure(text='', cursor='')
            self._note_link_ids = []
            return
        names = ', '.join(f"{r['title']} ({r['list_name']})" for r in rows)
        self._note_link_lbl.configure(text=f'🔗 Linked from: {names}', cursor='hand2')
        self._note_link_ids = [(r['id'], r['list_id'], r['list_name']) for r in rows]
        self._note_link_lbl.bind('<Button-1>', lambda e: self._note_jump_to_reminder())

    def _note_jump_to_reminder(self):
        if not getattr(self, '_note_link_ids', None):
            return
        item_id, list_id, list_name = self._note_link_ids[0]
        lst = next((l for l in self.todo_db.get_lists() if l['id'] == list_id), None)
        color = lst['color'] if lst else '#007AFF'
        self._main_nb.select(self.tab_todo)
        self._todo_select_list(list_id, color, list_name)
        if self._todo_tv.exists(str(item_id)):
            self._todo_tv.see(str(item_id))
            self._todo_tv.selection_set(str(item_id))
            self._todo_tv.focus(str(item_id))

    # ── Tab: EXP 料源替換 ─────────────────────────────────────────────────────

    def _build_exp_tab(self):
        f = self.tab_exp
        f.columnconfigure(1, weight=1)
        f.rowconfigure(3, weight=1)
        p = dict(padx=6, pady=3)

        # Row 0: BOM file
        ttk.Button(f, text='Open BOM File', command=self._exp_open_bom,
                   width=18).grid(row=0, column=0, sticky='w', **p)
        self._sv_exp_bom = tk.StringVar()
        ttk.Entry(f, textvariable=self._sv_exp_bom,
                  state='readonly').grid(row=0, column=1, sticky='ew', **p)

        # Row 1: EXP file
        ttk.Button(f, text='Open EXP File', command=self._exp_open_exp,
                   width=18).grid(row=1, column=0, sticky='w', **p)
        self._sv_exp_exp = tk.StringVar()
        ttk.Entry(f, textvariable=self._sv_exp_exp,
                  state='readonly').grid(row=1, column=1, sticky='ew', **p)

        # Row 2: Analyze button + status
        row2 = ttk.Frame(f)
        row2.grid(row=2, column=0, columnspan=2, sticky='ew', padx=6, pady=3)
        ttk.Button(row2, text='Load & Analyze',
                   command=self._exp_analyze).pack(side='left')
        self._exp_status = ttk.Label(row2, text='', foreground='gray',
                                     font=FONT_UI)
        self._exp_status.pack(side='left', padx=10)

        # Row 3: Treeview
        tv_f = ttk.Frame(f)
        tv_f.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=6, pady=3)
        tv_f.rowconfigure(0, weight=1)
        tv_f.columnconfigure(0, weight=1)

        cols = ('code', 'orig_vpn', 'alt_vpn', 'alt_mfr', 'db', 'refs')
        self._exp_tv = ttk.Treeview(tv_f, columns=cols, show='tree headings',
                                     selectmode='browse')
        self._exp_tv.heading('#0',       text='☑')
        self._exp_tv.heading('code',     text='10-Code')
        self._exp_tv.heading('orig_vpn', text='Orig VPN')
        self._exp_tv.heading('alt_vpn',  text='Alt VPN')
        self._exp_tv.heading('alt_mfr',  text='Alt Mfr')
        self._exp_tv.heading('db',       text='DB')
        self._exp_tv.heading('refs',     text='Part Refs')
        self._exp_tv.column('#0',       width=28,  stretch=False)
        self._exp_tv.column('code',     width=130, stretch=False)
        self._exp_tv.column('orig_vpn', width=180, stretch=False)
        self._exp_tv.column('alt_vpn',  width=150, stretch=False)
        self._exp_tv.column('alt_mfr',  width=160, stretch=False)
        self._exp_tv.column('db',       width=40,  stretch=False)
        self._exp_tv.column('refs',     width=300)
        self._exp_tv.grid(row=0, column=0, sticky='nsew')
        self._exp_tv.tag_configure('checked',  background='#D4EDDA')
        self._exp_tv.tag_configure('alt_sel',  background='#CCE5FF')
        self._exp_tv.tag_configure('no_db',    foreground='#CC4400')
        self._exp_tv.bind('<Button-1>', self._exp_tv_click)
        self._exp_tv.bind('<Button-3>', self._exp_right_click)

        self._exp_ctx = tk.Menu(self._exp_tv, tearoff=0)
        self._exp_ctx.add_command(label='🔍 Search DB for replacement...',
                                   command=self._exp_search_db)

        vsb = ttk.Scrollbar(tv_f, orient='vertical', command=self._exp_tv.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        self._exp_tv.configure(yscrollcommand=vsb.set)

        # Row 4: Apply + Save buttons
        row4 = ttk.Frame(f)
        row4.grid(row=4, column=0, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Button(row4, text='Apply Selected',
                   command=self._exp_apply).pack(side='left')
        ttk.Button(row4, text='Save EXP as...',
                   command=self._exp_save).pack(side='left', padx=6)
        ttk.Button(row4, text='Export Excel',
                   command=self._exp_export_excel).pack(side='left')

        # Row 5: Result log
        self._exp_log = scrolledtext.ScrolledText(
            f, font=FONT_TX, wrap='none', bg='#f8f8f8', height=8)
        self._exp_log.grid(row=5, column=0, columnspan=2,
                           sticky='ew', padx=6, pady=3)

    def _exp_open_bom(self):
        path = filedialog.askopenfilename(
            title='Select Orcad BOM File',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if path:
            self._exp_bom_file = path
            self._sv_exp_bom.set(path)

    def _exp_open_exp(self):
        path = filedialog.askopenfilename(
            title='Select Orcad EXP File',
            filetypes=[('EXP files', '*.exp *.EXP'), ('All', '*.*')])
        if path:
            self._exp_exp_file = path
            self._sv_exp_exp.set(path)

    def _exp_analyze(self):
        if not self._exp_bom_file:
            messagebox.showwarning('Warning', '請先選擇 BOM 檔案')
            return
        if not self._exp_exp_file:
            messagebox.showwarning('Warning', '請先選擇 EXP 檔案')
            return
        try:
            bom_items = BomParser.parse(self._exp_bom_file)
        except Exception as e:
            messagebox.showerror('BOM Error', str(e))
            return
        try:
            self._exp_design, self._exp_headers, self._exp_rows = \
                self._exp_parser.parse(self._exp_exp_file)
        except Exception as e:
            messagebox.showerror('EXP Error', str(e))
            return

        def _clean_vpn(vpn: str) -> str:
            """Strip leading single-letter prefix e.g. 'M      PMEG3005EL' → 'PMEG3005EL'."""
            m = re.match(r'^[A-Za-z]\s+(.*)', vpn.strip())
            return m.group(1).strip() if m else vpn.strip()

        self._exp_items = []
        for item in bom_items:
            clean_vpn = _clean_vpn(item.vendor_pn)
            try:
                ss_rows = self.db.get_second_sources(item.code, clean_vpn)
            except Exception:
                ss_rows = []
            alts = []
            for row in ss_rows:
                alts.append({
                    'new_code':   row['internal_pn'] or item.code,
                    'new_vpn':    row['vendor_pn'],
                    'new_desc':   row['description'] or '',
                    'new_vendor': row['manufacturer'] or '',
                })
            self._exp_items.append({
                'code':      item.code,
                'vendor_pn': item.vendor_pn,
                'desc':      item.description,
                'qty':       item.qty,
                'locations': list(item.locations),
                'checked':   False,
                'sel_alt':   0,
                'alts':      alts,
            })

        # Populate treeview
        for iid in self._exp_tv.get_children():
            self._exp_tv.delete(iid)
        self._exp_iid_item: dict[str, int] = {}   # iid → index in _exp_items
        self._exp_iid_alt:  dict[str, tuple] = {} # iid → (item_idx, alt_idx)

        for idx, item in enumerate(self._exp_items):
            refs_str = ','.join(item['locations'][:8])
            if len(item['locations']) > 8:
                refs_str += f',... (+{len(item["locations"])-8})'
            has_alt = bool(item['alts'])
            if has_alt:
                first_alt = item['alts'][0]
                alt_vpn   = first_alt['new_vpn']
                alt_mfr   = first_alt['new_vendor']
                p_tags    = ()
            else:
                alt_vpn   = ''
                alt_mfr   = ''
                p_tags    = ('no_db',)
            parent = self._exp_tv.insert(
                '', 'end',
                text='☐',
                values=(item['code'], item['vendor_pn'],
                        alt_vpn, alt_mfr,
                        '✓' if has_alt else '', refs_str),
                tags=p_tags)
            self._exp_iid_item[parent] = idx

            if len(item['alts']) > 1:
                for aidx, alt in enumerate(item['alts']):
                    child = self._exp_tv.insert(
                        parent, 'end',
                        text=' ',
                        values=('', '', alt['new_vpn'], alt['new_vendor'],
                                '✓', ''))
                    self._exp_iid_alt[child] = (idx, aidx)
                    if aidx == 0:
                        self._exp_tv.item(child, tags=('alt_sel',))

        total    = len(self._exp_items)
        with_alt = sum(1 for i in self._exp_items if i['alts'])
        self._exp_status.config(
            text=f'共 {total} 個 BOM item，其中 {with_alt} 個有替代料',
            foreground='green' if with_alt else 'gray')
        self._exp_log.delete('1.0', 'end')

    def _exp_tv_click(self, event):
        region = self._exp_tv.identify_region(event.x, event.y)
        iid    = self._exp_tv.identify_row(event.y)
        if not iid:
            return

        if iid in self._exp_iid_item:
            # Click on parent → toggle checkbox if clicking col #0
            col = self._exp_tv.identify_column(event.x)
            if col == '#0':
                idx  = self._exp_iid_item[iid]
                item = self._exp_items[idx]
                item['checked'] = not item['checked']
                chk  = '☑' if item['checked'] else '☐'
                tags = ('checked',) if item['checked'] else ()
                self._exp_tv.item(iid, text=chk, tags=tags)

        elif iid in self._exp_iid_alt:
            # Click on child → mark as selected alt
            item_idx, alt_idx = self._exp_iid_alt[iid]
            item = self._exp_items[item_idx]
            item['sel_alt'] = alt_idx
            alt  = item['alts'][alt_idx]
            # Update parent row display
            parent = self._exp_tv.parent(iid)
            refs_str = ','.join(item['locations'][:8])
            if len(item['locations']) > 8:
                refs_str += f',... (+{len(item["locations"])-8})'
            chk = '☑' if item['checked'] else '☐'
            tags = ('checked',) if item['checked'] else ()
            self._exp_tv.item(parent,
                text=chk,
                values=(item['code'], item['vendor_pn'],
                        alt['new_vpn'], alt['new_vendor'], '✓', refs_str),
                tags=tags)
            # Highlight selected child
            for sibling in self._exp_tv.get_children(parent):
                self._exp_tv.item(sibling, tags=())
            self._exp_tv.item(iid, tags=('alt_sel',))

    def _exp_apply(self):
        if not self._exp_items:
            messagebox.showinfo('Info', '請先執行 Load & Analyze')
            return
        if not self._exp_rows:
            messagebox.showinfo('Info', 'EXP 資料未載入')
            return

        # Build a quick lookup: part_ref → row index in _exp_rows
        ref_idx: dict[str, int] = {}
        for i, row in enumerate(self._exp_rows):
            if len(row) > ExpParser.COL_PART_REF:
                ref_idx[row[ExpParser.COL_PART_REF]] = i

        changes: list[tuple] = []  # (ref, orig_code, new_code, new_vpn)
        for item in self._exp_items:
            if not item['checked']:
                continue
            alt = item['alts'][item['sel_alt']]
            for ref in item['locations']:
                ri = ref_idx.get(ref)
                if ri is None:
                    continue
                row = self._exp_rows[ri]
                # Pad row if needed
                while len(row) <= max(ExpParser.COL_PART_NO,
                                      ExpParser.COL_DESC2,
                                      ExpParser.COL_VALUE1,
                                      ExpParser.COL_VENDOR2,
                                      ExpParser.COL_VPN):
                    row.append('')
                orig_code = row[ExpParser.COL_PART_NO]
                new_desc   = alt['new_desc'] or row[ExpParser.COL_DESC]
                new_vendor = alt['new_vendor']
                row[ExpParser.COL_PART_NO] = alt['new_code'] or orig_code
                row[ExpParser.COL_DESC]    = new_desc   # DESCRIPTION (col 17)
                row[ExpParser.COL_DESC2]   = new_desc   # Description (col 21)
                row[ExpParser.COL_VALUE]   = alt['new_vpn']
                row[ExpParser.COL_VALUE1]  = alt['new_vpn']
                row[ExpParser.COL_VENDOR]  = new_vendor  # VENDOR (col 69)
                row[ExpParser.COL_VENDOR2] = new_vendor  # Vendor  (col 72)
                row[ExpParser.COL_VPN]     = alt['new_vpn']
                changes.append((ref, orig_code, alt['new_code'], alt['new_vpn'],
                                alt['new_vendor'], alt['new_desc']))

        self._exp_changes = changes
        self._exp_log.delete('1.0', 'end')
        if not changes:
            self._exp_log.insert('end', '沒有勾選任何 item，無變更。\n')
            return
        self._exp_log.insert('end',
            f'套用完成，共修改 {len(changes)} 個 Part Reference：\n'
            f'{"─"*100}\n'
            f'{"Part Ref":<12} {"Orig 10-Code":<18} {"New 10-Code":<18} '
            f'{"New VPN":<22} {"Vendor":<24} Description\n'
            f'{"─"*100}\n')
        for ref, orig, new_c, new_v, vendor, desc in changes:
            self._exp_log.insert('end',
                f'{ref:<12} {orig:<18} {new_c:<18} '
                f'{new_v[:21]:<22} {vendor[:23]:<24} {desc[:40]}\n')

    def _exp_export_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        if not self._exp_changes:
            messagebox.showwarning('Warning', '請先執行 Apply Selected')
            return

        wb  = openpyxl.Workbook()
        ws  = wb.active
        ws.title = 'EXP 料源替換'

        hdr_fill = PatternFill('solid', fgColor='1F6CBF')
        hdr_font = Font(bold=True, color='FFFFFF')
        headers  = ['Part Ref', 'Orig 10-Code', 'New 10-Code',
                    'New VPN', 'Vendor', 'Description']
        widths   = [12, 18, 18, 24, 28, 50]
        ws.append(headers)
        for cell, w in zip(ws[1], widths):
            cell.fill = hdr_fill
            cell.font = hdr_font
            ws.column_dimensions[cell.column_letter].width = w
        ws.freeze_panes = 'A2'

        for ref, orig, new_c, new_v, vendor, desc in self._exp_changes:
            ws.append([ref, orig, new_c, new_v, vendor, desc])

        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'EXP_替換_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        try:
            wb.save(path)
            messagebox.showinfo('Done', f'已儲存至\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _exp_save(self):
        if not self._exp_rows:
            messagebox.showinfo('Info', 'EXP 資料未載入，請先 Load & Analyze')
            return
        base = os.path.splitext(os.path.basename(self._exp_exp_file))[0]
        path = filedialog.asksaveasfilename(
            title='Save Modified EXP',
            initialdir=os.path.dirname(self._exp_exp_file) or None,
            initialfile=f'{base}_ss.EXP',
            filetypes=[('EXP files', '*.EXP *.exp'), ('All', '*.*')])
        if not path:
            return
        try:
            self._exp_parser.write(path, self._exp_design,
                                   self._exp_headers, self._exp_rows)
            messagebox.showinfo('Done', f'已儲存：\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _exp_right_click(self, event):
        iid = self._exp_tv.identify_row(event.y)
        if not iid:
            return
        self._exp_tv.focus(iid)
        self._exp_tv.selection_set(iid)
        ok = iid in self._exp_iid_item or iid in self._exp_iid_alt
        self._exp_ctx.entryconfigure(0, state='normal' if ok else 'disabled')
        self._exp_ctx.tk_popup(event.x_root, event.y_root)

    def _exp_search_db(self):
        iid = self._exp_tv.focus()
        if not iid:
            return
        if iid in self._exp_iid_item:
            item_idx   = self._exp_iid_item[iid]
            parent_iid = iid
        elif iid in self._exp_iid_alt:
            item_idx, _ = self._exp_iid_alt[iid]
            parent_iid  = self._exp_tv.parent(iid)
        else:
            return

        item     = self._exp_items[item_idx]
        picker   = _AltPartPicker(self.root, self.db, initial_query=item['code'])
        self.root.wait_window(picker)
        if not picker.result:
            return

        vendor_pn, manufacturer, description, internal_pn = picker.result
        new_alt = {
            'new_code':   internal_pn or item['code'],
            'new_vpn':    vendor_pn,
            'new_desc':   description or '',
            'new_vendor': manufacturer or '',
        }
        item['alts'].append(new_alt)
        new_alt_idx = len(item['alts']) - 1
        item['sel_alt'] = new_alt_idx

        # Add child row
        child = self._exp_tv.insert(parent_iid, 'end',
            text=' ',
            values=('', '', vendor_pn, manufacturer, '✓', ''))
        self._exp_iid_alt[child] = (item_idx, new_alt_idx)

        # Clear previous alt_sel tags on siblings
        for sibling in self._exp_tv.get_children(parent_iid):
            self._exp_tv.item(sibling, tags=())
        self._exp_tv.item(child, tags=('alt_sel',))

        # Update parent row display to reflect chosen alt
        refs_str = ','.join(item['locations'][:8])
        if len(item['locations']) > 8:
            refs_str += f',... (+{len(item["locations"])-8})'
        chk  = '☑' if item['checked'] else '☐'
        tags = ('checked',) if item['checked'] else ()
        self._exp_tv.item(parent_iid,
            text=chk,
            values=(item['code'], item['vendor_pn'],
                    vendor_pn, manufacturer, '✓', refs_str),
            tags=tags)
        self._exp_tv.item(parent_iid, open=True)

    # ── Tab: BOM Diff ─────────────────────────────────────────────────────────

    def _build_diff_tab(self):
        f = self.tab_diff
        f.columnconfigure(1, weight=1)
        f.rowconfigure(3, weight=1)
        p = dict(padx=6, pady=3)

        ttk.Button(f, text='Open BOM A', command=self._diff_open_a,
                   width=14).grid(row=0, column=0, sticky='w', **p)
        self._sv_diff_a = tk.StringVar()
        ttk.Entry(f, textvariable=self._sv_diff_a,
                  state='readonly').grid(row=0, column=1, sticky='ew', **p)

        ttk.Button(f, text='Open BOM B', command=self._diff_open_b,
                   width=14).grid(row=1, column=0, sticky='w', **p)
        self._sv_diff_b = tk.StringVar()
        ttk.Entry(f, textvariable=self._sv_diff_b,
                  state='readonly').grid(row=1, column=1, sticky='ew', **p)

        row2 = ttk.Frame(f)
        row2.grid(row=2, column=0, columnspan=2, sticky='ew', padx=6, pady=3)
        ttk.Button(row2, text='Compare', command=self._diff_analyze).pack(side='left')
        self._diff_hide_unch = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text='Hide Unchanged',
                        variable=self._diff_hide_unch,
                        command=self._diff_refresh_all).pack(side='left', padx=10)
        self._diff_status = ttk.Label(row2, text='', foreground='gray', font=FONT_UI)
        self._diff_status.pack(side='left', padx=10)

        nb_v = ttk.Notebook(f)
        nb_v.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=6, pady=3)

        # ── Tab 1: By Location ────────────────────────────────────────────────
        tv_f = ttk.Frame(nb_v)
        nb_v.add(tv_f, text='By Location')
        tv_f.rowconfigure(0, weight=1)
        tv_f.columnconfigure(0, weight=1)

        cols = ('status', 'location', 'code_a', 'code_b',
                'value_a', 'value_b', 'vpn_a', 'vpn_b')
        self._diff_tv = ttk.Treeview(tv_f, columns=cols, show='headings',
                                      selectmode='browse')
        self._diff_tv.heading('status',   text='Status')
        self._diff_tv.heading('location', text='Location')
        self._diff_tv.heading('code_a',   text='10-Code (A)')
        self._diff_tv.heading('code_b',   text='10-Code (B)')
        self._diff_tv.heading('value_a',  text='Value (A)')
        self._diff_tv.heading('value_b',  text='Value (B)')
        self._diff_tv.heading('vpn_a',    text='VPN (A)')
        self._diff_tv.heading('vpn_b',    text='VPN (B)')
        self._diff_tv.column('status',   width=75,  stretch=False)
        self._diff_tv.column('location', width=75,  stretch=False)
        self._diff_tv.column('code_a',   width=120, stretch=False)
        self._diff_tv.column('code_b',   width=120, stretch=False)
        self._diff_tv.column('value_a',  width=130, stretch=True)
        self._diff_tv.column('value_b',  width=130, stretch=True)
        self._diff_tv.column('vpn_a',    width=150, stretch=True)
        self._diff_tv.column('vpn_b',    width=150, stretch=True)
        self._diff_tv.tag_configure('added',     background='#D4EDDA')
        self._diff_tv.tag_configure('removed',   background='#F8D7DA')
        self._diff_tv.tag_configure('changed',   background='#FFF3CD')
        self._diff_tv.tag_configure('unchanged', background='#FFFFFF')

        sb_y = ttk.Scrollbar(tv_f, orient='vertical',   command=self._diff_tv.yview)
        sb_x = ttk.Scrollbar(tv_f, orient='horizontal', command=self._diff_tv.xview)
        self._diff_tv.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        self._diff_tv.grid(row=0, column=0, sticky='nsew')
        sb_y.grid(row=0, column=1, sticky='ns')
        sb_x.grid(row=1, column=0, sticky='ew')

        # ── Tab 2: By 料號 (Qty 變化) ─────────────────────────────────────────
        tv_f2 = ttk.Frame(nb_v)
        nb_v.add(tv_f2, text='By 料號 (Qty 變化)')
        tv_f2.rowconfigure(0, weight=1)
        tv_f2.columnconfigure(0, weight=1)

        cols2 = ('status', 'code', 'value', 'vpn', 'qty_a', 'qty_b', 'delta')
        self._diff_code_tv = ttk.Treeview(tv_f2, columns=cols2, show='headings',
                                           selectmode='browse')
        self._diff_code_tv.heading('status', text='Status')
        self._diff_code_tv.heading('code',   text='10-Code')
        self._diff_code_tv.heading('value',  text='Value')
        self._diff_code_tv.heading('vpn',    text='VPN')
        self._diff_code_tv.heading('qty_a',  text='Qty (A)')
        self._diff_code_tv.heading('qty_b',  text='Qty (B)')
        self._diff_code_tv.heading('delta',  text='差異 (B-A)')
        self._diff_code_tv.column('status', width=80,  stretch=False)
        self._diff_code_tv.column('code',   width=140, stretch=False)
        self._diff_code_tv.column('value',  width=130, stretch=True)
        self._diff_code_tv.column('vpn',    width=150, stretch=True)
        self._diff_code_tv.column('qty_a',  width=70,  stretch=False, anchor='center')
        self._diff_code_tv.column('qty_b',  width=70,  stretch=False, anchor='center')
        self._diff_code_tv.column('delta',  width=90,  stretch=False, anchor='center')
        self._diff_code_tv.tag_configure('added',     background='#D4EDDA')
        self._diff_code_tv.tag_configure('removed',   background='#F8D7DA')
        self._diff_code_tv.tag_configure('changed',   background='#FFF3CD')
        self._diff_code_tv.tag_configure('unchanged', background='#FFFFFF')

        sb2_y = ttk.Scrollbar(tv_f2, orient='vertical',   command=self._diff_code_tv.yview)
        sb2_x = ttk.Scrollbar(tv_f2, orient='horizontal', command=self._diff_code_tv.xview)
        self._diff_code_tv.configure(yscrollcommand=sb2_y.set, xscrollcommand=sb2_x.set)
        self._diff_code_tv.grid(row=0, column=0, sticky='nsew')
        sb2_y.grid(row=0, column=1, sticky='ns')
        sb2_x.grid(row=1, column=0, sticky='ew')

        self._diff_summary = ttk.Label(f, text='', font=FONT_UI, foreground='#333333')
        self._diff_summary.grid(row=4, column=0, columnspan=2, sticky='w', padx=6, pady=2)

        row5 = ttk.Frame(f)
        row5.grid(row=5, column=0, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Button(row5, text='Export TXT',
                   command=self._diff_export_txt).pack(side='left')
        ttk.Button(row5, text='Export Excel',
                   command=self._diff_export_excel).pack(side='left', padx=6)

    def _diff_open_a(self):
        path = filedialog.askopenfilename(
            title='Select BOM A',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if path:
            self._diff_file_a = path
            self._sv_diff_a.set(path)

    def _diff_open_b(self):
        path = filedialog.askopenfilename(
            title='Select BOM B',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if path:
            self._diff_file_b = path
            self._sv_diff_b.set(path)

    @staticmethod
    def _diff_build_loc_map(items):
        m = {}
        for item in items:
            for loc in item.locations:
                m[loc.strip()] = item
        return m

    def _diff_analyze(self):
        if not self._diff_file_a:
            messagebox.showwarning('Warning', '請先選擇 BOM A 檔案')
            return
        if not self._diff_file_b:
            messagebox.showwarning('Warning', '請先選擇 BOM B 檔案')
            return
        try:
            items_a = BomParser.parse(self._diff_file_a)
        except Exception as e:
            messagebox.showerror('BOM A Error', str(e))
            return
        try:
            items_b = BomParser.parse(self._diff_file_b)
        except Exception as e:
            messagebox.showerror('BOM B Error', str(e))
            return

        map_a  = self._diff_build_loc_map(items_a)
        map_b  = self._diff_build_loc_map(items_b)
        locs_a = set(map_a.keys())
        locs_b = set(map_b.keys())

        def _loc_sort_key(x):
            prefix = re.sub(r'\d.*', '', x).upper()
            m = re.search(r'\d+', x)
            return (prefix, int(m.group()) if m else 0)

        all_locs = sorted(locs_a | locs_b, key=_loc_sort_key)

        results = []
        for loc in all_locs:
            in_a, in_b = loc in locs_a, loc in locs_b
            if in_a and not in_b:
                ia = map_a[loc]
                results.append({'status': 'Removed', 'location': loc,
                                 'code_a': ia.code,       'code_b': '',
                                 'value_a': ia.value,     'value_b': '',
                                 'vpn_a': ia.vendor_pn,   'vpn_b': ''})
            elif in_b and not in_a:
                ib = map_b[loc]
                results.append({'status': 'Added', 'location': loc,
                                 'code_a': '',             'code_b': ib.code,
                                 'value_a': '',            'value_b': ib.value,
                                 'vpn_a': '',              'vpn_b': ib.vendor_pn})
            else:
                ia, ib = map_a[loc], map_b[loc]
                if (ia.code != ib.code or ia.value != ib.value
                        or ia.vendor_pn != ib.vendor_pn):
                    status = 'Changed'
                else:
                    status = 'Unchanged'
                results.append({'status': status, 'location': loc,
                                 'code_a': ia.code,      'code_b': ib.code,
                                 'value_a': ia.value,    'value_b': ib.value,
                                 'vpn_a': ia.vendor_pn,  'vpn_b': ib.vendor_pn})

        self._diff_results = results
        self._diff_refresh_tv()

        # ── By-code aggregation ───────────────────────────────────────────────
        all_codes: dict = {}
        for item in items_a:
            key = item.code or item.vendor_pn
            if not key:
                continue
            if key not in all_codes:
                all_codes[key] = {'value': item.value, 'vpn': item.vendor_pn,
                                  'qty_a': 0, 'qty_b': 0}
            all_codes[key]['qty_a'] += item.qty

        for item in items_b:
            key = item.code or item.vendor_pn
            if not key:
                continue
            if key not in all_codes:
                all_codes[key] = {'value': item.value, 'vpn': item.vendor_pn,
                                  'qty_a': 0, 'qty_b': 0}
            all_codes[key]['qty_b'] += item.qty

        code_results = []
        for key in sorted(all_codes):
            d = all_codes[key]
            delta = d['qty_b'] - d['qty_a']
            if d['qty_a'] == 0:
                c_status = 'Added'
            elif d['qty_b'] == 0:
                c_status = 'Removed'
            elif delta != 0:
                c_status = 'Changed'
            else:
                c_status = 'Unchanged'
            code_results.append({
                'status': c_status, 'code': key,
                'value': d['value'], 'vpn': d['vpn'],
                'qty_a': d['qty_a'], 'qty_b': d['qty_b'], 'delta': delta,
            })
        self._diff_code_results = code_results
        self._diff_refresh_code_tv()

        n_added   = sum(1 for r in results if r['status'] == 'Added')
        n_removed = sum(1 for r in results if r['status'] == 'Removed')
        n_changed = sum(1 for r in results if r['status'] == 'Changed')
        n_unch    = sum(1 for r in results if r['status'] == 'Unchanged')
        self._diff_summary.config(
            text=(f'Total A: {len(locs_a)}  Total B: {len(locs_b)}   |   '
                  f'Added: {n_added}   Removed: {n_removed}   '
                  f'Changed: {n_changed}   Unchanged: {n_unch}'))
        self._diff_status.config(
            text=f'{len(results)} locations compared', foreground='gray')

    def _diff_refresh_tv(self):
        tv = self._diff_tv
        tv.delete(*tv.get_children())
        hide_unch = self._diff_hide_unch.get()
        tag_map = {'Added': 'added', 'Removed': 'removed',
                   'Changed': 'changed', 'Unchanged': 'unchanged'}
        for r in self._diff_results:
            if hide_unch and r['status'] == 'Unchanged':
                continue
            tv.insert('', 'end', tags=(tag_map.get(r['status'], ''),), values=(
                r['status'], r['location'],
                r['code_a'],  r['code_b'],
                r['value_a'], r['value_b'],
                r['vpn_a'],   r['vpn_b'],
            ))

    def _diff_refresh_code_tv(self):
        tv = self._diff_code_tv
        tv.delete(*tv.get_children())
        hide_unch = self._diff_hide_unch.get()
        tag_map = {'Added': 'added', 'Removed': 'removed',
                   'Changed': 'changed', 'Unchanged': 'unchanged'}
        for r in self._diff_code_results:
            if hide_unch and r['status'] == 'Unchanged':
                continue
            delta = r['delta']
            delta_str = f'+{delta}' if delta > 0 else str(delta)
            tv.insert('', 'end', tags=(tag_map.get(r['status'], ''),), values=(
                r['status'], r['code'], r['value'], r['vpn'],
                r['qty_a'], r['qty_b'], delta_str,
            ))

    def _diff_refresh_all(self):
        self._diff_refresh_tv()
        self._diff_refresh_code_tv()

    def _diff_export_txt(self):
        if not self._diff_results:
            messagebox.showwarning('Warning', '請先執行 Compare')
            return
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'BOM_Diff_{ts}.txt',
            filetypes=[('Text', '*.txt'), ('All', '*.*')])
        if not path:
            return
        n_added   = sum(1 for r in self._diff_results if r['status'] == 'Added')
        n_removed = sum(1 for r in self._diff_results if r['status'] == 'Removed')
        n_changed = sum(1 for r in self._diff_results if r['status'] == 'Changed')
        n_unch    = sum(1 for r in self._diff_results if r['status'] == 'Unchanged')
        lines = [
            'BOM Diff Report',
            f"Generated : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"BOM A     : {self._diff_file_a}",
            f"BOM B     : {self._diff_file_b}",
            '',
            'Summary',
            f"  Added    : {n_added}",
            f"  Removed  : {n_removed}",
            f"  Changed  : {n_changed}",
            f"  Unchanged: {n_unch}",
            '',
        ]
        col = '{:<10} {:<18} {:<18} {:<28} {:<28} {:<22} {:<22}'

        def _section(title, status):
            rows = [r for r in self._diff_results if r['status'] == status]
            if not rows:
                return
            lines.append(f'=== {title} ({len(rows)}) ===')
            lines.append(col.format('Location', '10-Code (A)', '10-Code (B)',
                                    'Value (A)', 'Value (B)', 'VPN (A)', 'VPN (B)'))
            lines.append('-' * 128)
            for r in rows:
                lines.append(col.format(
                    r['location'],
                    r['code_a'][:18],  r['code_b'][:18],
                    r['value_a'][:28], r['value_b'][:28],
                    r['vpn_a'][:22],   r['vpn_b'][:22]))
            lines.append('')

        _section('ADDED   (in B, not in A)', 'Added')
        _section('REMOVED (in A, not in B)', 'Removed')
        _section('CHANGED', 'Changed')
        _section('UNCHANGED', 'Unchanged')
        try:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write('\n'.join(lines))
            messagebox.showinfo('Done', f'已儲存：\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _diff_export_excel(self):
        if not self._diff_results:
            messagebox.showwarning('Warning', '請先執行 Compare')
            return
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        fills = {
            'Added':     PatternFill('solid', fgColor='D4EDDA'),
            'Removed':   PatternFill('solid', fgColor='F8D7DA'),
            'Changed':   PatternFill('solid', fgColor='FFF3CD'),
            'Unchanged': None,
        }
        hdr_fill = PatternFill('solid', fgColor='1F6CBF')
        hdr_font = Font(bold=True, color='FFFFFF')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'BOM Diff'
        headers    = ['Status', 'Location', '10-Code (A)', '10-Code (B)',
                      'Value (A)', 'Value (B)', 'VPN (A)', 'VPN (B)']
        col_widths = [12, 10, 18, 18, 30, 30, 24, 24]
        ws.append(headers)
        for cell, w in zip(ws[1], col_widths):
            cell.fill = hdr_fill
            cell.font = hdr_font
            ws.column_dimensions[cell.column_letter].width = w
        ws.freeze_panes = 'A2'
        for r in self._diff_results:
            ws.append([r['status'], r['location'],
                       r['code_a'],  r['code_b'],
                       r['value_a'], r['value_b'],
                       r['vpn_a'],   r['vpn_b']])
            fill = fills.get(r['status'])
            if fill:
                row_idx = ws.max_row
                for col_idx in range(1, 9):
                    ws.cell(row_idx, col_idx).fill = fill
        # ── By 料號 sheet ─────────────────────────────────────────────────────
        ws_code = wb.create_sheet('By 料號')
        code_hdrs   = ['Status', '10-Code', 'Value', 'VPN', 'Qty (A)', 'Qty (B)', '差異 (B-A)']
        code_widths = [12, 20, 30, 24, 10, 10, 12]
        ws_code.append(code_hdrs)
        for cell, w in zip(ws_code[1], code_widths):
            cell.fill = hdr_fill
            cell.font = hdr_font
            ws_code.column_dimensions[cell.column_letter].width = w
        ws_code.freeze_panes = 'A2'
        for r in self._diff_code_results:
            delta = r['delta']
            delta_str = f'+{delta}' if delta > 0 else str(delta)
            ws_code.append([r['status'], r['code'], r['value'], r['vpn'],
                            r['qty_a'], r['qty_b'], delta_str])
            fill = fills.get(r['status'])
            if fill:
                row_idx = ws_code.max_row
                for col_idx in range(1, 8):
                    ws_code.cell(row_idx, col_idx).fill = fill

        ws2 = wb.create_sheet('Summary')
        n_added   = sum(1 for r in self._diff_results if r['status'] == 'Added')
        n_removed = sum(1 for r in self._diff_results if r['status'] == 'Removed')
        n_changed = sum(1 for r in self._diff_results if r['status'] == 'Changed')
        n_unch    = sum(1 for r in self._diff_results if r['status'] == 'Unchanged')
        for row in [['BOM A',     self._diff_file_a],
                    ['BOM B',     self._diff_file_b],
                    ['Added',     n_added],
                    ['Removed',   n_removed],
                    ['Changed',   n_changed],
                    ['Unchanged', n_unch]]:
            ws2.append(row)
        ws2.column_dimensions['A'].width = 14
        ws2.column_dimensions['B'].width = 60
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'BOM_Diff_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        try:
            wb.save(path)
            messagebox.showinfo('Done', f'已儲存至\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))



    # ═══════════════════════════════════════════════════════════════════
    # SAP BOM 比對 tab (ported from bom_check.py)
    # ═══════════════════════════════════════════════════════════════════

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_sapbom_tab(self):
        f = self.tab_bomchk
        nb = ttk.Notebook(f)
        nb.pack(fill='both', expand=True, padx=4, pady=4)

        self.tab_bomchk_cmp    = ttk.Frame(nb)
        self.tab_bomchk_orcad  = ttk.Frame(nb)
        self.tab_bomchk_vendor = ttk.Frame(nb)
        self.tab_bomchk_cfg    = ttk.Frame(nb)
        self.tab_bomchk_result = ttk.Frame(nb)
        self.tab_bomchk_help   = ttk.Frame(nb)
        self.tab_bomchk_rev    = ttk.Frame(nb)

        nb.add(self.tab_bomchk_cmp,    text='Bom Compare')
        nb.add(self.tab_bomchk_orcad,  text='Orcad Bom Check')
        nb.add(self.tab_bomchk_vendor, text='Vendor 篩選')
        nb.add(self.tab_bomchk_result, text='Compare Result')
        nb.add(self.tab_bomchk_help,   text='使用說明')
        nb.add(self.tab_bomchk_rev,    text='Revision')

        self._bomchk_build_compare_tab()
        self._bomchk_build_orcad_tab()
        self._bomchk_build_vendor_tab()
        self._bomchk_build_config_tab()
        self._bomchk_build_result_tab()
        self._bomchk_build_help_tab()
        self._bomchk_build_revision_tab()

    # ── Tab 1: BOM Compare ────────────────────────────────────────────────────

    def _bomchk_build_compare_tab(self):
        f = self.tab_bomchk_cmp
        f.columnconfigure(1, weight=1)
        f.rowconfigure(3, weight=1)
        p = dict(padx=6, pady=3)

        self.sv_bomchk_cmp_level = tk.StringVar(value='2,3')

        # Row 0: Orcad file picker
        ttk.Button(f, text='Open Orcad Bom File', command=self._bomchk_open_orcad,
                   width=20).grid(row=0, column=0, sticky='w', **p)
        self.sv_bomchk_orcad = tk.StringVar()
        ttk.Entry(f, textvariable=self.sv_bomchk_orcad,
                  state='readonly').grid(row=0, column=1, sticky='ew', **p)

        # Row 1: SAP file picker
        ttk.Button(f, text='Open SAP Bom File', command=self._bomchk_open_sap,
                   width=20).grid(row=1, column=0, sticky='w', **p)
        self.sv_bomchk_sap = tk.StringVar()
        ttk.Entry(f, textvariable=self.sv_bomchk_sap,
                  state='readonly').grid(row=1, column=1, sticky='ew', **p)

        # Row 2: action buttons + centre label
        row2 = ttk.Frame(f)
        row2.grid(row=2, column=0, columnspan=2, sticky='ew', padx=6, pady=3)
        ttk.Button(row2, text='BOM Compare',
                   command=self._bomchk_do_compare).pack(side='left')
        ttk.Label(row2, text='Compare Result').pack(side='left', expand=True)

        # Row 3: scrolled result area
        self.txt_bomchk_result = scrolledtext.ScrolledText(
            f, font=FONT_TX, wrap='none', bg='#f8f8f8')
        self.txt_bomchk_result.grid(row=3, column=0, columnspan=2,
                             sticky='nsew', padx=6, pady=3)

        # Row 4: save / clear
        row4 = ttk.Frame(f)
        row4.grid(row=4, column=0, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Button(row4, text='Export Excel',
                   command=self._bomchk_save_excel_compare).pack(side='left')
        ttk.Button(row4, text='Clear Log',
                   command=self._bomchk_clear_log).pack(side='right')

    # ── Tab 2: Orcad BOM Check ────────────────────────────────────────────────

    def _bomchk_build_orcad_tab(self):
        f = self.tab_bomchk_orcad
        f.columnconfigure(1, weight=1)
        f.rowconfigure(2, weight=1)
        p = dict(padx=6, pady=3)

        ttk.Button(f, text='Open Orcad Bom File', command=self._bomchk_open_orcad_check,
                   width=20).grid(row=0, column=0, sticky='w', **p)
        self.sv_bomchk_orcad2 = tk.StringVar()
        ttk.Entry(f, textvariable=self.sv_bomchk_orcad2,
                  state='readonly').grid(row=0, column=1, sticky='ew', **p)

        ttk.Button(f, text='Analyze Orcad BOM',
                   command=self._bomchk_do_orcad_check).grid(row=1, column=0, sticky='w', **p)

        self.txt_bomchk_orcad = scrolledtext.ScrolledText(
            f, font=FONT_TX, wrap='none', bg='#f8f8f8')
        self.txt_bomchk_orcad.grid(row=2, column=0, columnspan=2,
                            sticky='nsew', padx=6, pady=3)

        row3 = ttk.Frame(f)
        row3.grid(row=3, column=0, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Button(row3, text='Export Excel',
                   command=self._bomchk_save_excel_orcad).pack(side='left')

    # ── Tab 3: Vendor 篩選 ────────────────────────────────────────────────────

    def _bomchk_build_vendor_tab(self):
        f = self.tab_bomchk_vendor
        f.columnconfigure(1, weight=1)
        f.rowconfigure(2, weight=1)
        p = dict(padx=6, pady=3)

        ttk.Button(f, text='Open Orcad Bom File', command=self._bomchk_open_vendor_check,
                   width=20).grid(row=0, column=0, sticky='w', **p)
        self.sv_bomchk_vendor = tk.StringVar()
        ttk.Entry(f, textvariable=self.sv_bomchk_vendor,
                  state='readonly').grid(row=0, column=1, sticky='ew', **p)

        ttk.Button(f, text='Analyze Vendor',
                   command=self._bomchk_do_vendor_check).grid(row=1, column=0, sticky='w', **p)

        pw = ttk.PanedWindow(f, orient='horizontal')
        pw.grid(row=2, column=0, columnspan=2, sticky='nsew', padx=4, pady=4)

        # Left pane: editable approved vendor keyword list
        frm_list = ttk.Frame(pw)
        frm_list.rowconfigure(1, weight=1)
        frm_list.columnconfigure(0, weight=1)
        ttk.Label(frm_list, text='核可 Vendor 關鍵字（每行一個）').grid(
            row=0, column=0, columnspan=2, sticky='w', padx=4, pady=2)
        self.txt_bomchk_vendor_list = scrolledtext.ScrolledText(
            frm_list, font=FONT_TX, wrap='none', width=28)
        self.txt_bomchk_vendor_list.grid(row=1, column=0, columnspan=2,
                                  sticky='nsew', padx=4, pady=2)
        row_btn = ttk.Frame(frm_list)
        row_btn.grid(row=2, column=0, columnspan=2, sticky='ew', padx=4, pady=4)
        ttk.Button(row_btn, text='儲存', command=self._bomchk_save_vendor_list).pack(side='left')
        ttk.Button(row_btn, text='載入', command=self._bomchk_load_vendor_list_to_widget).pack(side='left', padx=4)
        ttk.Button(row_btn, text='還原預設', command=self._bomchk_reset_vendor_list).pack(side='left')

        self.lbl_bomchk_vendor_path = ttk.Label(frm_list, foreground='gray',
                                         font=('Courier New', 7), wraplength=220, justify='left')
        self.lbl_bomchk_vendor_path.grid(row=3, column=0, columnspan=2, sticky='w', padx=4, pady=0)

        # Right pane: results
        frm_result = ttk.Frame(pw)
        frm_result.rowconfigure(0, weight=1)
        frm_result.columnconfigure(0, weight=1)
        self.txt_bomchk_vendor = scrolledtext.ScrolledText(
            frm_result, font=FONT_TX, wrap='none', bg='#f8f8f8')
        self.txt_bomchk_vendor.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)
        self.txt_bomchk_vendor.tag_configure('vendor_warn', background='#FFA500', foreground='#000000')

        pw.add(frm_list,   weight=1)
        pw.add(frm_result, weight=3)

        row3 = ttk.Frame(f)
        row3.grid(row=3, column=0, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Button(row3, text='Export Excel',
                   command=self._bomchk_save_excel_vendor).pack(side='left')

        self._bomchk_load_vendor_list_to_widget()

    # ── Tab 4: Config ─────────────────────────────────────────────────────────

    def _bomchk_build_config_tab(self):
        f = self.tab_bomchk_cfg
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        self.txt_bomchk_cfg = scrolledtext.ScrolledText(
            f, font=FONT_TX, wrap='none')
        self.txt_bomchk_cfg.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
        self.txt_bomchk_cfg.insert('end', self.bomchk_cfg.to_text())

        row1 = ttk.Frame(f)
        row1.grid(row=1, column=0, sticky='ew', padx=6, pady=3)
        ttk.Button(row1, text='Save Config',
                   command=self._bomchk_save_config).pack(side='left')
        ttk.Button(row1, text='Reload',
                   command=self._bomchk_reload_config).pack(side='left', padx=6)

    # ── Tab 4: Compare Result ─────────────────────────────────────────────────

    def _bomchk_build_result_tab(self):
        f = self.tab_bomchk_result
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        sub_nb = ttk.Notebook(f)
        sub_nb.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)

        frm_code = ttk.Frame(sub_nb)
        frm_loc  = ttk.Frame(sub_nb)
        frm_tbl  = ttk.Frame(sub_nb)
        frm_note = ttk.Frame(sub_nb)
        sub_nb.add(frm_code, text='依 Code')
        sub_nb.add(frm_loc,  text='依 Location')
        sub_nb.add(frm_tbl,  text='Location 對照表')
        sub_nb.add(frm_note, text='Location 分類')

        for frm, attr in (
                (frm_code, 'txt_bomchk_by_code'),
                (frm_loc,  'txt_bomchk_by_loc')):
            frm.rowconfigure(0, weight=1)
            frm.columnconfigure(0, weight=1)
            txt = scrolledtext.ScrolledText(
                frm, font=FONT_TX, wrap='none', bg='#f8f8f8')
            txt.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)
            setattr(self, attr, txt)
            btn_frm = ttk.Frame(frm)
            btn_frm.grid(row=1, column=0, sticky='ew', padx=4, pady=4)
            ttk.Button(btn_frm, text='Export Excel',
                       command=self._bomchk_save_excel_compare).pack(side='left')

        # Location 對照表 tab
        frm_tbl.rowconfigure(0, weight=1)
        frm_tbl.columnconfigure(0, weight=1)
        self.txt_bomchk_loc_tbl = scrolledtext.ScrolledText(
            frm_tbl, font=FONT_TX, wrap='none', bg='#f8f8f8')
        self.txt_bomchk_loc_tbl.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)
        btn_tbl = ttk.Frame(frm_tbl)
        btn_tbl.grid(row=1, column=0, sticky='ew', padx=4, pady=4)
        ttk.Button(btn_tbl, text='Export Excel',
                   command=self._bomchk_save_excel_compare).pack(side='left')
        self.txt_bomchk_loc_tbl.tag_configure('diff',  background='#FFD700', foreground='#000000')
        self.txt_bomchk_loc_tbl.tag_configure('only_o', background='#FFB3B3', foreground='#000000')
        self.txt_bomchk_loc_tbl.tag_configure('only_s', background='#B3D9FF', foreground='#000000')

        # Location 分類 tab
        frm_note.rowconfigure(0, weight=1)
        frm_note.columnconfigure(0, weight=1)
        self.txt_bomchk_loc_note = scrolledtext.ScrolledText(
            frm_note, font=FONT_TX, wrap='none', bg='#f8f8f8')
        self.txt_bomchk_loc_note.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)
        btn_note = ttk.Frame(frm_note)
        btn_note.grid(row=1, column=0, sticky='ew', padx=4, pady=4)
        ttk.Button(btn_note, text='Export Excel',
                   command=self._bomchk_save_excel_compare).pack(side='left')
        self.txt_bomchk_loc_note.tag_configure('note_swap',     background='#FFD700', foreground='#000000')
        self.txt_bomchk_loc_note.tag_configure('note_new',      background='#FFB3B3', foreground='#000000')
        self.txt_bomchk_loc_note.tag_configure('note_removed',  background='#B3D9FF', foreground='#000000')
        self.txt_bomchk_loc_note.tag_configure('note_existing', background='#C6F0C2', foreground='#000000')

    # ── Tab 5: 使用說明 ───────────────────────────────────────────────────────

    def _bomchk_build_help_tab(self):
        HELP_TEXT = """\
AVABD BOM CHECK Rev0.1  使用說明
══════════════════════════════════════════════════════════════════════════════

【BOM Compare】
  功能：比對 Orcad BOM (.BOM) 與 SAP BOM (.txt 或 .xls) 的差異

  步驟：
    1. 點擊「Open Orcad Bom File」選取 .BOM 檔案
    2. 點擊「Open SAP Bom File」選取 SAP BOM .txt 或 .xls 檔案
    3. 點擊「BOM Compare」執行比對
    4. 結果顯示於下方文字區域（Orcad Only / SAP Only）
    5. 可至「Compare Result」tab 查看依 Code / 依 Location 分類的詳細結果
    6. 點擊「Save Log Result as File ...」可將結果存成 .txt 檔

  注意：預設比對 SAP Layer 2, 3 的料件


【Orcad Bom Check】
  功能：分析 Orcad BOM，標示需確認的料號（黃底警示）

  步驟：
    1. 點擊「Open Orcad Bom File」選取 .BOM 檔案
    2. 點擊「Analyze Orcad BOM」執行分析

  黃底警示條件：
    - 10-Code 為空白
    - 10-Code 為 TBD
    - 10-Code 以 COM 開頭
    - 10-Code 末碼非 M


【Vendor 篩選】
  功能：比對所有料件的 Vendor，標示未在核可清單內的項目（橘色高亮）

  步驟：
    1. 點擊「Open Orcad Bom File」選取 .BOM 檔案
    2. 點擊「Analyze Vendor」執行篩選
    3. 橘色高亮 = Vendor 不在核可清單中
    4. 畫面下方「結論」區段列出所有未核可項目

  核可 Vendor 清單管理（左側面板）：
    - 直接在文字區編輯關鍵字，每行一個，不分大小寫
    - 「儲存」  ：將清單寫入 approved_vendors.txt
    - 「載入」  ：從 approved_vendors.txt 重新讀取（外部編輯後使用）
    - 「還原預設」：恢復內建預設清單

  Vendor 比對規則：
    - 從 VENDOR 欄位括號 () 內取出廠商名稱
      範例：20X030(MURATA MANUFACTURING CO.LTD)
              → 取出 MURATA MANUFACTURING CO.LTD
    - 對核可清單的關鍵字做「包含」比對，不分大小寫
      範例：關鍵字 MURATA 可匹配 MURATA MANUFACTURING CO.LTD
    - VENDOR 欄位空白的料件視為「未核可」（橘色）


【BOM 欄位格式說明】
  Orcad BOM 欄位（Tab 分隔，共 10 欄）：

  Item  10-Code  Value  VENDOR_PN  VENDOR  Description  Footprint  Quantity  Location  NC

  ── OrCAD BOM Template（複製貼入 OrCAD BOM 設定）────────────────────────────

  Header Row（貼入 Header Line 欄位）：
  Item\\t10-Code\\tValue\\tVENDOR_PN\\tVENDOR\\tDescription\\tFootprint\\tQuantity\\tLocation\\tNC

  Data Row（貼入 Data Line 欄位）：
  {Item}\\t{Part_Number}\\t{Value}\\t{VENDOR_PN}\\t{VENDOR}\\t{Description}\\t{PCB Footprint}\\t{Quantity}\\t{Reference}\\t{NC}

  ─────────────────────────────────────────────────────────────────────────────

  欄位序號（bom_check.inf 設定，從 1 開始）：
    Col  1 = Item
    Col  2 = 10-Code      (Orcad Bom Code Column         = 2)
    Col  3 = Value        (Orcad Bom Value Column         = 3)
    Col  4 = VENDOR_PN    (Orcad Bom Vendor PN Column     = 4)
    Col  5 = VENDOR       (Orcad Bom Vendor Column        = 5)
    Col  6 = Description  (Orcad Bom Description Column   = 6)
    Col  7 = Footprint    (Orcad Bom Footprint Column     = 7)
    Col  8 = Quantity     (Orcad Bom Qty Column           = 8)
    Col  9 = Location     (Orcad Bom Location Column      = 9)
    Col 10 = NC           (Orcad Bom NC Column            = 10)
              NC 欄有值時，此料件略過不比對

  忽略規則（bom_check.inf）：
    Orcad Ignore Analyze Value    = NC_;_NC
      → Value 前綴含 NC_ 或後綴含 _NC 的料件略過

    Orcad Ignore Analyze Location = TP;H
      → Location 前綴為 TP 或 H 的料件略過（測試點、機構孔）

    Orcad Ignore 10-Code          = New10Code
      → 10-Code 等於此值的料件略過（尚未申請料號）

"""
        f = self.tab_bomchk_help
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        txt = scrolledtext.ScrolledText(f, font=FONT_TX, wrap='none', bg='#fdfdfd')
        txt.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
        txt.insert('end', HELP_TEXT)
        txt.configure(state='disabled')

    # ── Tab 6: Revision ───────────────────────────────────────────────────────

    def _bomchk_build_revision_tab(self):
        f = self.tab_bomchk_rev
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        txt = scrolledtext.ScrolledText(f, font=FONT_TX, wrap='none')
        txt.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
        txt.insert('end', BOMCHK_REVISION_HISTORY)
        txt.configure(state='disabled')

    # ── Event Handlers ────────────────────────────────────────────────────────

    def _bomchk_open_orcad(self):
        init = self.bomchk_cfg.get('File Path Config', 'Orcad Bom Files Path', '')
        path = filedialog.askopenfilename(
            title='Select Orcad BOM File',
            initialdir=init or None,
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if not path:
            return
        self.bomchk_orcad_file = path
        self.sv_bomchk_orcad.set(path)
        try:
            self.bomchk_orcad_items = self.bomchk_orcad_p.parse(path)
            self._bomchk_log(f'Orcad Bom 解析完成！共 {len(self.bomchk_orcad_items)} 筆十碼\n')
        except Exception as e:
            messagebox.showerror('Orcad Parse Error', str(e))

    def _bomchk_open_sap(self):
        init = self.bomchk_cfg.get('File Path Config', 'SAP Bom Files Path', 'C:\\TEMP')
        path = filedialog.askopenfilename(
            title='Select SAP BOM File',
            initialdir=init or None,
            filetypes=[('SAP BOM', '*.txt *.xls'), ('Text files', '*.txt'),
                       ('Excel', '*.xls'), ('All', '*.*')])
        if not path:
            return
        self.bomchk_sap_file = path
        self.sv_bomchk_sap.set(path)
        try:
            self.bomchk_sap_top, self.bomchk_sap_comps = self.bomchk_sap_p.parse(path)
            self._bomchk_show_sap_summary()
        except Exception as e:
            messagebox.showerror('SAP Parse Error', str(e))

    def _bomchk_show_sap_summary(self):
        self._bomchk_log_clear()
        self._bomchk_log('SAP Bom 分析完成！\n')

        if self.bomchk_sap_top:
            # Collect unique type prefixes found
            found_types = set()
            for t in self.bomchk_sap_top:
                for pfx in self.bomchk_sap_p.top_pfx:
                    if t.code.startswith(pfx):
                        found_types.add(pfx)
                        break
            type_str = ' 或 '.join(sorted(found_types))
            self._bomchk_log(f'共有 {len(self.bomchk_sap_top)}個 {type_str}\n')
            for i, t in enumerate(self.bomchk_sap_top):
                self._bomchk_log(f'[ {i} ]{t.code}    {t.description}\n')

        self._bomchk_log(f'元件總數: {len(self.bomchk_sap_comps)} 筆\n')

    def _bomchk_do_compare(self):
        if not self.bomchk_orcad_items:
            messagebox.showwarning('Warning', '請先開啟 Orcad BOM 檔案')
            return
        if not self.bomchk_sap_comps:
            messagebox.showwarning('Warning', '請先開啟 SAP BOM 檔案')
            return

        lvl_raw = self.sv_bomchk_cmp_level.get()
        levels  = {int(x) for x in re.split(r'[,;]', lvl_raw) if x.strip().isdigit()}
        sap_filtered = [i for i in self.bomchk_sap_comps if i.level in levels] if levels else self.bomchk_sap_comps
        orcad_only, sap_only = self.bomchk_cmp.compare(self.bomchk_orcad_items, sap_filtered)

        result_code = self.bomchk_cmp.format_by_code(
            orcad_only, sap_only, self.bomchk_orcad_file, self.bomchk_sap_file)
        result_loc  = self.bomchk_cmp.format_by_location(
            orcad_only, sap_only, self.bomchk_orcad_file, self.bomchk_sap_file)
        nc_locs  = getattr(self.bomchk_orcad_p, 'nc_locs', set())
        tbl_data = self.bomchk_cmp.loc_table_data(
            self.bomchk_orcad_items, sap_filtered,
            nc_locs=nc_locs, ig_loc_fn=self.bomchk_orcad_p._ig_loc)
        note_data = self.bomchk_cmp.loc_note_data(
            self.bomchk_orcad_items, sap_filtered,
            nc_locs=nc_locs, ig_loc_fn=self.bomchk_orcad_p._ig_loc)

        self._bomchk_last_orcad_only    = orcad_only
        self._bomchk_last_sap_only      = sap_only
        self._bomchk_last_tbl_data      = tbl_data
        self._bomchk_last_note_tbl_data = note_data

        result_note = self.bomchk_cmp.format_loc_note(note_data, self.bomchk_orcad_file, self.bomchk_sap_file)

        self._bomchk_log_clear()
        self._bomchk_log(result_code)
        self._bomchk_log('\n\n' + result_note + '\n')
        self._bomchk_set_txt(self.txt_bomchk_by_code, result_code)
        self._bomchk_set_txt(self.txt_bomchk_by_loc,  result_loc)
        self._bomchk_fill_loc_table(tbl_data)
        self._bomchk_fill_loc_note(note_data)

    # ── Excel export helpers ──────────────────────────────────────────────────

    def _bomchk_xl_check(self) -> bool:
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return False
        return True

    def _bomchk_xl_hdr(self, ws, headers, col_widths):
        fill = PatternFill('solid', fgColor='1F6CBF')
        font = Font(bold=True, color='FFFFFF')
        ws.append(headers)
        for cell, w in zip(ws[1], col_widths):
            cell.fill = fill
            cell.font = font
            ws.column_dimensions[cell.column_letter].width = w
        ws.freeze_panes = 'A2'

    def _bomchk_xl_save(self, wb, prefix):
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_dir = self.bomchk_cfg.get('File Path Config', 'Report Files Path', '')
        path = filedialog.asksaveasfilename(
            initialdir=report_dir or None,
            initialfile=f'{prefix}_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        try:
            wb.save(path)
            messagebox.showinfo('Done', f'已儲存至\n{path}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    # ── Excel export: BOM Compare ─────────────────────────────────────────────

    def _bomchk_save_excel_compare(self):
        if not self._bomchk_xl_check():
            return
        if not self._bomchk_last_orcad_only and not self._bomchk_last_sap_only and not self._bomchk_last_tbl_data \
                and not self._bomchk_last_note_tbl_data:
            messagebox.showwarning('Warning', '請先執行 BOM Compare')
            return

        wb = openpyxl.Workbook()

        # Sheet 1: Orcad Only
        ws1 = wb.active
        ws1.title = 'Orcad Only'
        self._bomchk_xl_hdr(ws1,
            ['Code', 'Value', 'Description', 'Footprint', 'Qty', 'Location'],
            [16, 26, 36, 20, 6, 80])
        for item in self._bomchk_last_orcad_only:
            ws1.append([item.code, item.value, item.description,
                        item.footprint, item.qty, ','.join(item.locations)])

        # Sheet 2: SAP Only
        ws2 = wb.create_sheet('SAP Only')
        self._bomchk_xl_hdr(ws2,
            ['Code', 'Part Name', 'Description', 'Qty', 'Location'],
            [16, 26, 40, 6, 80])
        for item in self._bomchk_last_sap_only:
            ws2.append([item.code, item.part_name, item.description,
                        int(item.qty), ','.join(item.locations)])

        # Sheet 3: Location Table
        ws3 = wb.create_sheet('Location Table')
        self._bomchk_xl_hdr(ws3, ['Location', 'SAP Code', 'BOM Code', 'Status'], [12, 18, 18, 14])
        fill_diff   = PatternFill('solid', fgColor='FFD700')
        fill_only_o = PatternFill('solid', fgColor='FFB3B3')
        fill_only_s = PatternFill('solid', fgColor='B3D9FF')
        for loc, s_code, o_code, match in self._bomchk_last_tbl_data:
            if match:
                continue
            if s_code and not o_code:
                status, fill = 'SAP Only',   fill_only_s
            elif o_code and not s_code:
                status, fill = 'Orcad Only', fill_only_o
            else:
                status, fill = 'Mismatch',   fill_diff
            ws3.append([loc, s_code, o_code, status])
            row = ws3.max_row
            for col in range(1, 5):
                ws3.cell(row, col).fill = fill

        # Sheet 4: Location 分類 (Note)
        ws4 = wb.create_sheet('Location 分類')
        self._bomchk_xl_hdr(ws4, ['Location', 'SAP PN', 'Orcad PN', 'Note'], [12, 18, 18, 24])
        note_fills = {
            '換料':                     PatternFill('solid', fgColor='FFD700'),
            '新增 Part number & 位置':    PatternFill('solid', fgColor='FFB3B3'),
            'SAP位置移除':               PatternFill('solid', fgColor='B3D9FF'),
            'SAP 既有料號':              PatternFill('solid', fgColor='C6F0C2'),
        }
        for loc, s_code, o_code, note in self._bomchk_last_note_tbl_data:
            ws4.append([loc, s_code, o_code, note])
            fill = note_fills.get(note)
            if fill:
                row = ws4.max_row
                for col in range(1, 5):
                    ws4.cell(row, col).fill = fill

        self._bomchk_xl_save(wb, 'BOM_Compare')

    # ── Excel export: Orcad BOM Check ────────────────────────────────────────

    def _bomchk_save_excel_orcad(self):
        if not self._bomchk_xl_check():
            return
        if not self._bomchk_last_orcad_check_items:
            messagebox.showwarning('Warning', '請先執行 Analyze Orcad BOM')
            return

        wb  = openpyxl.Workbook()
        ws  = wb.active
        ws.title = 'Orcad BOM Check'
        self._bomchk_xl_hdr(ws,
            ['No', 'Code', 'Value', 'Vendor PN', 'Vendor', 'Description', 'Qty', 'Location', 'Warning'],
            [5, 16, 26, 20, 36, 36, 6, 80, 8])
        fill_warn = PatternFill('solid', fgColor='FFD700')
        for i, item in enumerate(self._bomchk_last_orcad_check_items):
            warn = self._bomchk_orcad_warn(item.code)
            ws.append([i, item.code, item.value, item.vendor_pn, item.vendor,
                       item.description, item.qty, ','.join(item.locations),
                       '★' if warn else ''])
            if warn:
                row = ws.max_row
                for col in range(1, 10):
                    ws.cell(row, col).fill = fill_warn

        self._bomchk_xl_save(wb, 'Orcad_Check')

    # ── Excel export: Vendor Check ────────────────────────────────────────────

    def _bomchk_save_excel_vendor(self):
        if not self._bomchk_xl_check():
            return
        if not self._bomchk_last_vendor_items:
            messagebox.showwarning('Warning', '請先執行 Analyze Vendor')
            return

        wb  = openpyxl.Workbook()
        ws  = wb.active
        ws.title = 'Vendor Check'
        self._bomchk_xl_hdr(ws,
            ['No', 'Code', 'Value', 'Vendor', 'Qty', 'Location', 'Approved'],
            [5, 16, 26, 40, 6, 80, 8])
        fill_warn = PatternFill('solid', fgColor='FFA500')
        for i, item in enumerate(self._bomchk_last_vendor_items):
            approved    = self._bomchk_is_approved_vendor(item.vendor, self._bomchk_last_vendor_keywords)
            vendor_name = self._bomchk_extract_vendor_name(item.vendor)
            ws.append([i, item.code, item.value, vendor_name, item.qty,
                       ','.join(item.locations), '' if approved else '★'])
            if not approved:
                row = ws.max_row
                for col in range(1, 8):
                    ws.cell(row, col).fill = fill_warn

        self._bomchk_xl_save(wb, 'Vendor_Check')

    def _bomchk_clear_log(self):
        self._bomchk_log_clear()

    def _bomchk_open_orcad_check(self):
        path = filedialog.askopenfilename(
            title='Select Orcad BOM',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if path:
            self.sv_bomchk_orcad2.set(path)

    def _bomchk_do_orcad_check(self):
        path = self.sv_bomchk_orcad2.get()
        if not path:
            messagebox.showwarning('Warning', '請先選擇 Orcad BOM 檔案')
            return
        try:
            items = self.bomchk_orcad_p.parse(path)
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return

        self._bomchk_last_orcad_check_items = items
        self.txt_bomchk_orcad.delete('1.0', 'end')
        self.txt_bomchk_orcad.tag_configure('warn', background='#FFD700', foreground='#000000')

        warn_items = []
        self.txt_bomchk_orcad.insert('end',
            f'Orcad BOM：{os.path.basename(path)}\n'
            f'共 {len(items)} 筆十碼\n\n'
            f'{"No.":<5} {"Code":<15} {"Value":<25} {"Qty":>5}  Location\n'
            f'{"-"*80}\n')
        for i, item in enumerate(items):
            lstr = ','.join(item.locations[:6])
            if len(item.locations) > 6:
                lstr += f',... (+{len(item.locations) - 6})'
            line = (f'{i:<5} {item.code:<15} {item.value[:24]:<25}'
                    f' {item.qty:>5}  {lstr}\n')
            if self._bomchk_orcad_warn(item.code):
                start = self.txt_bomchk_orcad.index('end - 1c linestart')
                self.txt_bomchk_orcad.insert('end', line)
                end = self.txt_bomchk_orcad.index('end - 1c')
                self.txt_bomchk_orcad.tag_add('warn', start, end)
                warn_items.append((i, item))
            else:
                self.txt_bomchk_orcad.insert('end', line)

        if warn_items:
            self.txt_bomchk_orcad.insert('end',
                f'\n{"─"*80}\n'
                f'★ 共 {len(warn_items)} 筆需確認（TBD / COM前綴 / 末碼非M）\n'
                f'{"No.":<5} {"Code":<15} {"Value":<25} {"Qty":>5}  Location\n'
                f'{"-"*80}\n')
            for i, item in warn_items:
                lstr = ','.join(item.locations)
                line = (f'{i:<5} {item.code:<15} {item.value[:24]:<25}'
                        f' {item.qty:>5}  {lstr}\n')
                start = self.txt_bomchk_orcad.index('end - 1c linestart')
                self.txt_bomchk_orcad.insert('end', line)
                end = self.txt_bomchk_orcad.index('end - 1c')
                self.txt_bomchk_orcad.tag_add('warn', start, end)

    @staticmethod
    def _bomchk_orcad_warn(code: str) -> bool:
        if not code or not code.strip():
            return True
        cu = code.strip().upper()
        if cu == 'TBD':
            return True
        if cu.startswith('COM'):
            return True
        if cu[-1] != 'M':
            return True
        return False

    def _bomchk_open_vendor_check(self):
        path = filedialog.askopenfilename(
            title='Select Orcad BOM',
            filetypes=[('BOM / Text', '*.bom *.BOM *.txt'), ('All', '*.*')])
        if path:
            self.sv_bomchk_vendor.set(path)

    def _bomchk_load_vendor_list_to_widget(self):
        vpath = os.path.join(os.path.dirname(self.bomchk_cfg.path), 'approved_vendors.txt')
        if os.path.exists(vpath):
            try:
                with open(vpath, encoding='utf-8') as fh:
                    content = fh.read()
                src = f'載入自：{vpath}'
            except Exception:
                content = '\n'.join(BOMCHK_APPROVED_VENDOR_KEYWORDS)
                src = f'讀取失敗，使用預設\n{vpath}'
        else:
            content = '\n'.join(BOMCHK_APPROVED_VENDOR_KEYWORDS)
            src = f'檔案不存在，使用預設\n期望路徑：{vpath}'
        self.txt_bomchk_vendor_list.delete('1.0', 'end')
        self.txt_bomchk_vendor_list.insert('end', content)
        if hasattr(self, 'lbl_vendor_path'):
            self.lbl_bomchk_vendor_path.configure(text=src)

    def _bomchk_save_vendor_list(self):
        vpath = os.path.join(os.path.dirname(self.bomchk_cfg.path), 'approved_vendors.txt')
        try:
            content = self.txt_bomchk_vendor_list.get('1.0', 'end').rstrip('\n') + '\n'
            with open(vpath, 'w', encoding='utf-8') as fh:
                fh.write(content)
            messagebox.showinfo('Done', f'Vendor 清單已儲存至\n{vpath}')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _bomchk_reset_vendor_list(self):
        self.txt_bomchk_vendor_list.delete('1.0', 'end')
        self.txt_bomchk_vendor_list.insert('end', '\n'.join(BOMCHK_APPROVED_VENDOR_KEYWORDS))

    def _bomchk_do_vendor_check(self):
        path = self.sv_bomchk_vendor.get()
        if not path:
            messagebox.showwarning('Warning', '請先選擇 Orcad BOM 檔案')
            return
        try:
            items = self.bomchk_orcad_p.parse(path)
        except Exception as e:
            messagebox.showerror('Error', str(e))
            return

        keywords = [ln.strip() for ln in self.txt_bomchk_vendor_list.get('1.0', 'end').splitlines()
                    if ln.strip()]

        self._bomchk_last_vendor_items    = items
        self._bomchk_last_vendor_keywords = keywords

        w = self.txt_bomchk_vendor
        w.delete('1.0', 'end')

        unapproved = [item for item in items if not self._bomchk_is_approved_vendor(item.vendor, keywords)]

        w.insert('end',
            f'Orcad BOM：{os.path.basename(path)}\n'
            f'共 {len(items)} 筆零件，其中 {len(unapproved)} 筆 Vendor 不在核可清單\n\n'
            f'{"No.":<5} {"Code":<15} {"Value":<25} {"Vendor":<35} {"Qty":>5}  Location\n'
            f'{"─"*100}\n')

        for i, item in enumerate(items):
            vendor_name = self._bomchk_extract_vendor_name(item.vendor)
            lstr = ','.join(item.locations[:5])
            if len(item.locations) > 5:
                lstr += f',... (+{len(item.locations)-5})'
            line = (f'{i:<5} {item.code:<15} {item.value[:24]:<25} '
                    f'{vendor_name[:34]:<35} {item.qty:>5}  {lstr}\n')
            if not self._bomchk_is_approved_vendor(item.vendor, keywords):
                start = w.index('end - 1c linestart')
                w.insert('end', line)
                end = w.index('end - 1c')
                w.tag_add('vendor_warn', start, end)
            else:
                w.insert('end', line)

        if unapproved:
            w.insert('end', f'\n{"═"*100}\n')
            w.insert('end', f'★ 結論：以下 {len(unapproved)} 筆 Vendor 不在核可清單中\n\n')
            w.insert('end',
                f'{"No.":<5} {"Code":<15} {"Value":<25} {"Vendor":<35} {"Qty":>5}  Location\n'
                f'{"─"*100}\n')
            for item in unapproved:
                vendor_name = self._bomchk_extract_vendor_name(item.vendor)
                lstr = ','.join(item.locations)
                line = (f'{"★":<5} {item.code:<15} {item.value[:24]:<25} '
                        f'{vendor_name[:34]:<35} {item.qty:>5}  {lstr}\n')
                start = w.index('end - 1c linestart')
                w.insert('end', line)
                end = w.index('end - 1c')
                w.tag_add('vendor_warn', start, end)
        else:
            w.insert('end', '\n★ 所有 Vendor 均在核可清單中\n')


    @staticmethod
    def _bomchk_extract_vendor_name(vendor_str: str) -> str:
        m = re.search(r'\((.+)\)', vendor_str)
        return m.group(1).strip() if m else vendor_str.strip()

    @staticmethod
    def _bomchk_is_approved_vendor(vendor_str: str, keywords) -> bool:
        name = BomCheckApp._bomchk_extract_vendor_name(vendor_str).upper()
        if not name:
            return False
        return any(kw.strip().upper() in name for kw in keywords if kw.strip())

    def _bomchk_save_config(self):
        content = self.txt_bomchk_cfg.get('1.0', 'end')
        try:
            with open(self.bomchk_cfg.path, 'w', encoding='utf-8') as fh:
                fh.write(content)
            self.bomchk_cfg._load()
            self.bomchk_orcad_p = OrcadParser(self.bomchk_cfg)
            self.bomchk_sap_p   = SapParser(self.bomchk_cfg)
            messagebox.showinfo('Done', '設定已儲存並重新載入')
        except Exception as e:
            messagebox.showerror('Error', str(e))

    def _bomchk_reload_config(self):
        self.txt_bomchk_cfg.delete('1.0', 'end')
        self.txt_bomchk_cfg.insert('end', self.bomchk_cfg.to_text())

    def _bomchk_fill_loc_table(self, tbl_data):
        w = self.txt_bomchk_loc_tbl
        w.delete('1.0', 'end')
        hdr = f'{"Location":<12} {"SAP Code":<15} {"BOM Code":<15}\n'
        sep = '─' * 44 + '\n'
        w.insert('end', hdr)
        w.insert('end', sep)
        for loc, s_code, o_code, match in tbl_data:
            if match:
                continue
            line = f'{loc:<12} {s_code:<15} {o_code:<15}\n'
            start = w.index('end - 1c linestart')
            w.insert('end', line)
            end = w.index('end - 1c')
            if s_code and not o_code:
                w.tag_add('only_s', start, end)
            elif o_code and not s_code:
                w.tag_add('only_o', start, end)
            else:
                w.tag_add('diff', start, end)

    NOTE_TAGS = {
        '換料':                     'note_swap',
        '新增 Part number & 位置':    'note_new',
        'SAP位置移除':               'note_removed',
        'SAP 既有料號':              'note_existing',
    }

    def _bomchk_fill_loc_note(self, note_data):
        w = self.txt_bomchk_loc_note
        w.delete('1.0', 'end')
        hdr = f'{"Location":<12} {"SAP PN":<16} {"Orcad PN":<16} Note\n'
        sep = '─' * 78 + '\n'
        w.insert('end', hdr)
        w.insert('end', sep)
        for loc, s_code, o_code, note in note_data:
            line = f'{loc:<12} {s_code:<16} {o_code:<16} {note}\n'
            start = w.index('end - 1c linestart')
            w.insert('end', line)
            end = w.index('end - 1c')
            tag = self.NOTE_TAGS.get(note)
            if tag:
                w.tag_add(tag, start, end)

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _bomchk_log(self, text: str):
        self.txt_bomchk_result.insert('end', text)
        self.txt_bomchk_result.see('end')

    def _bomchk_log_clear(self):
        self.txt_bomchk_result.delete('1.0', 'end')

    def _bomchk_set_txt(self, widget, text: str):
        widget.delete('1.0', 'end')
        widget.insert('end', text)



class _PasteRailsDialog(tk.Toplevel):
    """Paste JSON array or TSV (Excel copy) to batch-add rails to a part."""

    # Maps JSON keys (lowercase) → DB rail field for non-name fields.
    # rail_name / power_group are handled explicitly in _json_to_rail.
    _JSON_MAP = {
        'power_rail_crb':           'power_group',     # CRB-format alias
        'input_rail':               'input_rail',
        'pwr_input':                'input_rail',
        'description':              'rail_description',
        'rail_description':         'rail_description',
        'normal_operating_voltage': 'voltage',
        'voltage':                  'voltage',
        'tolerance':                'tolerance_percent',
        'tolerance_percent':        'tolerance_percent',
        'i_max':                    'max_current_a',
        'max_current':              'max_current_a',
        'max_current_a':            'max_current_a',
        'max_current_ma':           'max_current_a',
        'i_typ':                    'typ_current_a',
        'typ_current':              'typ_current_a',
        'typ_current_a':            'typ_current_a',
        'typ_current_ma':           'typ_current_a',
        'condition':                'condition',
        'pin_names':                'pin_names',
    }

    def __init__(self, parent):
        super().__init__(parent)
        self.title('Paste JSON / TSV — Batch Add Rails')
        self.geometry('780x580')
        self.resizable(True, True)
        self.transient(parent)
        self.result: list[dict] | None = None
        self._parsed: list[dict] = []
        self._build()
        self.grab_set()
        self.after(50, self._txt.focus_set)   # ensure focus after window renders

    def _build(self):
        top = ttk.Frame(self, padding=6)
        top.pack(fill='both', expand=True)

        ttk.Label(top, text='Paste JSON array or Excel-copied TSV below, then click Parse:',
                  font=FONT_UI).pack(anchor='w')

        frm_t = ttk.Frame(top)
        frm_t.pack(fill='both', expand=False, pady=(4, 0))
        frm_t.rowconfigure(0, weight=1)
        frm_t.columnconfigure(0, weight=1)

        self._txt = tk.Text(frm_t, height=12, wrap='none', font=('Consolas', 9))
        sb_v = ttk.Scrollbar(frm_t, orient='vertical',   command=self._txt.yview)
        sb_h = ttk.Scrollbar(frm_t, orient='horizontal', command=self._txt.xview)
        self._txt.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        self._txt.grid(row=0, column=0, sticky='nsew')
        sb_v.grid(row=0, column=1, sticky='ns')
        sb_h.grid(row=1, column=0, sticky='ew')

        ttk.Button(top, text='Parse', command=self._parse, width=10).pack(
            anchor='e', pady=4)

        ttk.Label(top, text='Preview (parsed rails):', font=FONT_UI).pack(anchor='w')

        cols = ('rail_name', 'power_group', 'voltage', 'tolerance', 'max_i_ma',
                'typ_i_ma', 'condition', 'rail_desc')
        self._tv = ttk.Treeview(top, columns=cols, show='headings', height=8)
        for cid, hdr, w in [
            ('rail_name',   'Rail Name',    130), ('power_group', 'Power Group', 110),
            ('voltage',     'Voltage(V)',    75),  ('tolerance',   'Tol%',         50),
            ('max_i_ma',    'Max I(mA)',     75),  ('typ_i_ma',    'Typ I(mA)',    75),
            ('condition',   'Condition',    90),   ('rail_desc',   'Description', 130),
        ]:
            self._tv.heading(cid, text=hdr)
            self._tv.column(cid, width=w, minwidth=40, anchor='center')
        vsb = ttk.Scrollbar(top, orient='vertical', command=self._tv.yview)
        hsb = ttk.Scrollbar(top, orient='horizontal', command=self._tv.xview)
        self._tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        frm_tv = ttk.Frame(top)
        frm_tv.pack(fill='both', expand=True, pady=(2, 0))
        self._tv.pack(in_=frm_tv, side='left', fill='both', expand=True)
        vsb.pack(in_=frm_tv, side='right', fill='y')
        hsb.pack(fill='x')

        self._status = ttk.Label(top, text='', foreground='gray', font=FONT_UI)
        self._status.pack(anchor='w', pady=2)

        bf = ttk.Frame(top)
        bf.pack(fill='x', pady=4)
        ttk.Button(bf, text='Import All', command=self._import, width=12).pack(side='right', padx=4)
        ttk.Button(bf, text='Cancel',     command=self.destroy, width=10).pack(side='right')

    # ── parse helpers ─────────────────────────────────────────────────────────

    @staticmethod
    def _parse_voltage(s):
        """'1.8V' → 1.8, '3.3' → 3.3"""
        s = str(s or '').strip().rstrip('Vv')
        try:
            return float(s)
        except ValueError:
            return 0.0

    @staticmethod
    def _parse_current(s):
        """'60mA'→0.060, '0.06A'→0.060, '60'→0.060 (assume mA if no unit)"""
        s = str(s or '').strip()
        if not s:
            return 0.0
        lower = s.lower()
        try:
            if lower.endswith('ma'):
                return float(lower[:-2]) / 1000
            if lower.endswith('a'):
                return float(lower[:-1])
            return float(s) / 1000  # bare number → assume mA
        except ValueError:
            return 0.0

    @staticmethod
    def _parse_tolerance(s):
        """'±8.0%' → 8.0, '5%' → 5.0, '5' → 5.0"""
        s = str(s or '').strip().lstrip('±').rstrip('%')
        try:
            return float(s)
        except ValueError:
            return 0.0

    def _json_to_rail(self, obj: dict) -> dict:
        rail: dict = {
            'rail_name': '', 'power_group': '', 'input_rail': '',
            'rail_description': '',
            'voltage': 0.0, 'tolerance_percent': 0.0,
            'typ_current_a': 0.0, 'max_current_a': 0.0,
            'condition': '', 'pin_names': '',
        }
        lobj = {k.lower(): v for k, v in obj.items()}

        # ── rail_name / power_group: handle two JSON styles ───────────────────
        # Style A (natural):  {"rail_name": "VBAT", "power_group": "VCC_12V", ...}
        # Style B (CRB-export): {"power_group": "AVDD09", "power_rail_crb": "VCC_0V9", ...}
        has_rail_name   = 'rail_name'   in lobj
        has_power_group = 'power_group' in lobj

        if has_rail_name:
            rail['rail_name'] = str(lobj['rail_name'] or '').strip()
            if has_power_group:
                # Style A: both keys present → use power_group as DB power_group
                rail['power_group'] = str(lobj['power_group'] or '').strip()
        elif has_power_group:
            # Style B: only power_group present → treat it as rail_name
            rail['rail_name'] = str(lobj['power_group'] or '').strip()

        # ── All other fields via map ───────────────────────────────────────────
        _skip = {'rail_name', 'power_group'}
        for jk, jv in lobj.items():
            if jk in _skip:
                continue
            field = self._JSON_MAP.get(jk)
            if field is None:
                continue
            if field == 'voltage':
                rail[field] = self._parse_voltage(jv)
            elif field in ('typ_current_a', 'max_current_a'):
                rail[field] = self._parse_current(jv)
            elif field == 'tolerance_percent':
                rail[field] = self._parse_tolerance(jv)
            else:
                rail[field] = str(jv or '').strip()

        if not rail['typ_current_a']:
            rail['typ_current_a'] = rail['max_current_a']
        return rail

    def _tsv_to_rails(self, text: str) -> list[dict]:
        """Parse tab-separated text.  First non-empty row = headers."""
        lines = [ln for ln in text.splitlines() if ln.strip()]
        if len(lines) < 2:
            return []
        headers = [h.strip().lower() for h in lines[0].split('\t')]
        rails = []
        for ln in lines[1:]:
            cells = ln.split('\t')
            obj = {headers[i]: cells[i].strip()
                   for i in range(min(len(headers), len(cells)))}
            r = self._json_to_rail(obj)
            if r['rail_name'] or r['voltage']:
                rails.append(r)
        return rails

    def _parse(self):
        text = self._txt.get('1.0', 'end').strip()
        if not text:
            self._status.config(text='請貼上資料', foreground='red')
            return
        rails: list[dict] = []
        # Try JSON first
        try:
            obj = json.loads(text)
            # Unwrap wrapper objects like {"power_rails": [...]} or {"rails": [...]}
            if isinstance(obj, dict):
                for v in obj.values():
                    if isinstance(v, list):
                        obj = v
                        break
                else:
                    obj = [obj]   # single rail object
            # obj is now a list
            for item in obj:
                if isinstance(item, dict):
                    r = self._json_to_rail(item)
                    if r['rail_name'] or r['voltage']:
                        rails.append(r)
        except json.JSONDecodeError:
            # Fall back to TSV
            rails = self._tsv_to_rails(text)

        if not rails:
            self._status.config(text='無法解析資料，請確認格式', foreground='red')
            return

        self._parsed = rails
        for iid in self._tv.get_children():
            self._tv.delete(iid)
        for r in rails:
            self._tv.insert('', 'end', values=(
                r['rail_name'], r['power_group'],
                f"{r['voltage']:.2f}", f"{r['tolerance_percent']:.1f}",
                f"{r['max_current_a']*1000:.2f}", f"{r['typ_current_a']*1000:.2f}",
                r['condition'], r['rail_description']))
        self._status.config(
            text=f'解析成功：{len(rails)} 條 Rail', foreground='green')

    def _import(self):
        if not self._parsed:
            messagebox.showwarning('Warning', '請先按 Parse 解析資料', parent=self)
            return
        self.result = self._parsed
        self.destroy()


# ── Power Budget Window ───────────────────────────────────────────────────────

class _PwrBudgetWindow(tk.Toplevel):
    def __init__(self, parent, budget_rows, missing, filename, export_fn):
        super().__init__(parent)
        self.title(f'Power Budget — {filename}')
        self.geometry('900x600')
        self.resizable(True, True)
        self._budget_rows = budget_rows
        self._missing     = missing
        self._export_fn   = export_fn
        self._filename    = filename

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=4, pady=4)

        # ── Tab 1: Detail ─────────────────────────────────────────────────────
        frm_d = ttk.Frame(nb)
        nb.add(frm_d, text=f'Per-IC Detail ({len(budget_rows)} rows)')
        frm_d.rowconfigure(0, weight=1); frm_d.columnconfigure(0, weight=1)
        dcols = ('vendor_pn', 'rail_name', 'power_group', 'voltage', 'qty',
                 'typ_i', 'max_i', 'typ_p', 'max_p', 'condition')
        tv_d = ttk.Treeview(frm_d, columns=dcols, show='headings')
        for cid, hdr, w in [
            ('vendor_pn',   'Vendor PN',   120), ('rail_name',   'Rail',       100),
            ('power_group', 'Power Group', 110), ('voltage',     'V(V)',         55),
            ('qty',         'Qty',          45), ('typ_i',       'Typ I(mA)',    75),
            ('max_i',       'Max I(mA)',    75), ('typ_p',       'Typ P(W)',     75),
            ('max_p',       'Max P(W)',     75), ('condition',   'Condition',   100),
        ]:
            tv_d.heading(cid, text=hdr)
            tv_d.column(cid, width=w, anchor='center', minwidth=40)
        vsb_d = ttk.Scrollbar(frm_d, orient='vertical', command=tv_d.yview)
        hsb_d = ttk.Scrollbar(frm_d, orient='horizontal', command=tv_d.xview)
        tv_d.configure(yscrollcommand=vsb_d.set, xscrollcommand=hsb_d.set)
        tv_d.grid(row=0, column=0, sticky='nsew')
        vsb_d.grid(row=0, column=1, sticky='ns')
        hsb_d.grid(row=1, column=0, sticky='ew')
        for r in budget_rows:
            tv_d.insert('', 'end', values=(
                r['vendor_pn'], r['rail_name'], r['power_group'],
                f"{r['voltage']:.2f}", r['qty'],
                f"{r['typ_i']*1000:.2f}", f"{r['max_i']*1000:.2f}",
                f"{r['typ_p']:.4f}", f"{r['max_p']:.4f}",
                r['condition']))

        # ── Tab 2: Summary ────────────────────────────────────────────────────
        from collections import defaultdict
        grp: dict = defaultdict(lambda: {'typ_i': 0, 'max_i': 0, 'typ_p': 0, 'max_p': 0})
        for r in budget_rows:
            key = (r['rail_name'], r['voltage'])
            grp[key]['typ_i'] += r['typ_i']
            grp[key]['max_i'] += r['max_i']
            grp[key]['typ_p'] += r['typ_p']
            grp[key]['max_p'] += r['max_p']

        frm_s = ttk.Frame(nb)
        nb.add(frm_s, text=f'Summary ({len(grp)} groups)')
        frm_s.rowconfigure(0, weight=1); frm_s.columnconfigure(0, weight=1)
        scols = ('rail_name', 'voltage', 'typ_i', 'max_i', 'typ_p', 'max_p')
        tv_s = ttk.Treeview(frm_s, columns=scols, show='headings')
        for cid, hdr, w in [
            ('rail_name',   'Rail Name (Net)',  160), ('voltage', 'Voltage(V)',    80),
            ('typ_i',       'Total Typ I(mA)', 110), ('max_i', 'Total Max I(mA)', 110),
            ('typ_p',       'Total Typ P(W)', 110), ('max_p',  'Total Max P(W)', 110),
        ]:
            tv_s.heading(cid, text=hdr)
            tv_s.column(cid, width=w, anchor='center', minwidth=50)
        tv_s.tag_configure('total', background='#D6EAF8', font=('Microsoft JhengHei UI', 9, 'bold'))
        vsb_s = ttk.Scrollbar(frm_s, orient='vertical', command=tv_s.yview)
        tv_s.configure(yscrollcommand=vsb_s.set)
        tv_s.grid(row=0, column=0, sticky='nsew')
        vsb_s.grid(row=0, column=1, sticky='ns')

        tot_typ_p = tot_max_p = 0
        for (pg, v), d in sorted(grp.items()):
            tv_s.insert('', 'end', values=(
                pg, f'{v:.2f}', f"{d['typ_i']*1000:.2f}", f"{d['max_i']*1000:.2f}",
                f"{d['typ_p']:.4f}", f"{d['max_p']:.4f}"))
            tot_typ_p += d['typ_p']; tot_max_p += d['max_p']
        tv_s.insert('', 'end', tags=('total',), values=(
            'TOTAL', '', '', '', f'{tot_typ_p:.4f}', f'{tot_max_p:.4f}'))

        # ── Tab 3: Input Rail Subtotal ────────────────────────────────────────
        from collections import defaultdict as _dd2
        ir_grp: dict = _dd2(lambda: {'typ_i': 0.0, 'max_i': 0.0,
                                      'typ_p': 0.0, 'max_p': 0.0})
        for r in budget_rows:
            ik = r['input_rail'] if r['input_rail'] else f"{r['voltage']:.3f}V"
            ir_grp[ik]['typ_i'] += r['typ_i']
            ir_grp[ik]['max_i'] += r['max_i']
            ir_grp[ik]['typ_p'] += r['typ_p']
            ir_grp[ik]['max_p'] += r['max_p']

        frm_v = ttk.Frame(nb)
        nb.add(frm_v, text=f'By Input Rail ({len(ir_grp)})')
        frm_v.rowconfigure(0, weight=1); frm_v.columnconfigure(0, weight=1)
        vcols = ('input_rail', 'typ_i', 'max_i', 'typ_p', 'max_p')
        tv_v = ttk.Treeview(frm_v, columns=vcols, show='headings')
        for cid, hdr, w in [
            ('input_rail', 'Input Rail',     160),
            ('typ_i',      'Total Typ I (A)', 140),
            ('max_i',      'Total Max I (A)', 140),
            ('typ_p',      'Total Typ P (W)', 140),
            ('max_p',      'Total Max P (W)', 140),
        ]:
            tv_v.heading(cid, text=hdr)
            tv_v.column(cid, width=w, anchor='center', minwidth=60)
        tv_v.tag_configure('total', background='#D6EAF8',
                            font=('Microsoft JhengHei UI', 9, 'bold'))
        vsb_v = ttk.Scrollbar(frm_v, orient='vertical', command=tv_v.yview)
        tv_v.configure(yscrollcommand=vsb_v.set)
        tv_v.grid(row=0, column=0, sticky='nsew')
        vsb_v.grid(row=0, column=1, sticky='ns')

        gt_i = gt_p = 0.0
        for ik in sorted(ir_grp):
            d = ir_grp[ik]
            tv_v.insert('', 'end', values=(
                ik,
                f"{d['typ_i']:.4f}", f"{d['max_i']:.4f}",
                f"{d['typ_p']:.4f}", f"{d['max_p']:.4f}"))
            gt_i += d['typ_i']; gt_p += d['typ_p']
        tv_v.insert('', 'end', tags=('total',), values=(
            'TOTAL', f'{gt_i:.4f}', '', f'{gt_p:.4f}', ''))

        # ── Bottom bar ────────────────────────────────────────────────────────
        bf = ttk.Frame(self)
        bf.pack(fill='x', padx=8, pady=6)
        miss_txt = f'Missing: {len(missing)} parts not in DB' if missing else 'All parts matched'
        ttk.Label(bf, text=miss_txt,
                  foreground='red' if missing else 'green',
                  font=FONT_UI).pack(side='left')
        ttk.Button(bf, text='Export Excel',
                   command=lambda: export_fn(budget_rows, missing, filename)).pack(side='right')


class _PwrTreeWindow(tk.Toplevel):
    """Hierarchical power-supply tree for the checked parts.

    Layout per root: root supply net -> Power Source(s) drawing from it ->
    the net that source produces -> (its direct rail consumers as leaves,
    plus further Power Source children recursing the same way). Built from
    MouserApp._pwr_build_power_tree_data's root node list.
    """

    def __init__(self, parent, app, checked, part_count):
        super().__init__(parent)
        self._app = app
        self._checked = checked
        self._part_count = part_count
        self._leaf_rail_id = {}   # tv item iid -> power_rails.id (leaf rows only)
        self._root_net_name = {}  # tv item iid -> net name (root rows only)
        self._source_row = {}     # tv item iid -> power_sources row dict (⚡ source rows only)

        self.title('Power Tree')
        self.geometry('880x600')
        self.resizable(True, True)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        cols = ('typ_i', 'max_i', 'typ_p', 'max_p', 'voltage', 'info')
        self.tv = tv = ttk.Treeview(self, columns=cols, show='tree headings')
        tv.heading('#0', text='Net / Source / Consumer')
        tv.column('#0', width=260, anchor='w')
        for cid, hdr, w in [
            ('typ_i',   'Typ I(mA)', 90), ('max_i', 'Max I(mA)', 90),
            ('typ_p',   'Typ P(W)',  80), ('max_p', 'Max P(W)',  80),
            ('voltage', 'V',         55), ('info',  'Info',     220),
        ]:
            tv.heading(cid, text=hdr)
            tv.column(cid, width=w, anchor=('w' if cid == 'info' else 'center'), minwidth=40)
        vsb = ttk.Scrollbar(self, orient='vertical', command=tv.yview)
        hsb = ttk.Scrollbar(self, orient='horizontal', command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tv.bind('<Double-1>', self._on_tree_double_click)

        tv.tag_configure('root',   background='#D6EAF8', font=('Microsoft JhengHei UI', 9, 'bold'))
        tv.tag_configure('source', foreground='#B9770E')
        tv.tag_configure('cycle',  foreground='red')
        tv.tag_configure('leaf',   foreground='#555555')

        bf = ttk.Frame(self)
        bf.grid(row=2, column=0, columnspan=2, sticky='ew', padx=8, pady=6)
        self.lbl_summary = ttk.Label(bf, text='', font=FONT_UI)
        self.lbl_summary.pack(side='left')
        ttk.Label(bf, text='（雙擊消耗列可編輯 Rail Name；雙擊 root 淨或 ⚡ 節點可選用/更換/移除 DC/DC）',
                  foreground='gray', font=FONT_UI).pack(side='left', padx=(12, 0))
        ttk.Button(bf, text='Export Excel', command=self._export_excel).pack(side='right')

        roots = app._pwr_build_power_tree_data(checked)
        self._populate(roots)

    def _populate(self, roots):
        tv = self.tv
        tv.delete(*tv.get_children())
        self._leaf_rail_id.clear()
        self._root_net_name.clear()
        self._source_row.clear()

        def insert_net_children(parent_iid, node):
            for c in node['consumers']:
                iid = tv.insert(parent_iid, 'end', tags=('leaf',),
                                 text=f"{c['vendor_pn']} · {c['power_group']} (x{c['qty']})",
                                 values=(f"{c['typ_i']*1000:.2f}", f"{c['max_i']*1000:.2f}",
                                         f"{c['typ_p']:.4f}", f"{c['max_p']:.4f}",
                                         f"{c['voltage']:.2f}", c['condition']))
                self._leaf_rail_id[iid] = c['rail_id']
            for child in node['children']:
                src = child['source']
                if child['cycle']:
                    tv.insert(parent_iid, 'end', tags=('cycle',),
                              text=f"⚠ circular reference: {src['rail_name']}",
                              values=('0.00', '0.00', '0.0000', '0.0000', '',
                                      'cycle detected — check input_rail / source_rail'))
                    continue
                out = child['out_node']
                eff = src['efficiency'] or 1.0
                info = (f"{src.get('vendor_pn', '')} {src['regulator_type'] or ''} "
                        f"Eff={eff:.0%} Vin={src['source_rail']}")
                sid = tv.insert(parent_iid, 'end', tags=('source',),
                                 text=f"⚡ {src['regulator_type'] or 'Source'} → {out['name'] or '(unnamed rail)'}",
                                 values=(f"{out['typ_i']*1000:.2f}", f"{out['max_i']*1000:.2f}",
                                         f"{out['typ_p']:.4f}", f"{out['max_p']:.4f}",
                                         f"{src['output_voltage']:.2f}", info))
                tv.item(sid, open=True)
                self._source_row[sid] = src
                insert_net_children(sid, out)

        gt_p = 0.0
        for root_node in roots:
            rid = tv.insert('', 'end', tags=('root',),
                             text=root_node['name'] or '(unnamed rail)',
                             values=(f"{root_node['typ_i']*1000:.2f}", f"{root_node['max_i']*1000:.2f}",
                                     f"{root_node['typ_p']:.4f}", f"{root_node['max_p']:.4f}",
                                     '', f"root supply ({len(root_node['consumers'])} direct consumers)"))
            tv.item(rid, open=True)
            self._root_net_name[rid] = root_node['name']
            insert_net_children(rid, root_node)
            gt_p += root_node['typ_p']

        self.lbl_summary.config(
            text=(f'{self._part_count} parts, {len(roots)} root supply net(s), '
                  f'total Typ power ≈ {gt_p:.4f} W (power sums across '
                  f'voltage domains; per-root currents do not — see rows above)'))

    def _on_tree_double_click(self, event):
        iid = self.tv.identify_row(event.y)
        if iid in self._leaf_rail_id:
            self._edit_leaf_rail_name(iid)
        elif iid in self._root_net_name:
            self._pick_dcdc_for_root(iid)
        elif iid in self._source_row:
            self._change_dcdc_for_source(iid)
        # cycle rows: nothing to edit, ignore.

    def _edit_leaf_rail_name(self, iid):
        rail_id = self._leaf_rail_id[iid]
        with self._app.pwr_db._conn() as c:
            r = c.execute('SELECT * FROM power_rails WHERE id=?', (rail_id,)).fetchone()
        if not r:
            return
        choices = self._app.pwr_db.get_all_rail_names()
        dlg = _RailNameDialog(
            self, f"{r['power_group'] or '(no power group)'} 的 Rail Name：",
            r['rail_name'] or '', choices)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        data = dict(r)
        data['rail_name'] = dlg.result.strip()
        self._app.pwr_db.upsert_rail(data)
        self._refresh()

    def _pick_dcdc_for_root(self, iid):
        net_name = self._root_net_name[iid]
        library = self._app.pwr_db.get_library_sources()
        if not library:
            messagebox.showinfo(
                'Info',
                '尚未建立任何 DC/DC 範本。\n'
                '請到 Power DB Builder → Power Sources (DC/DC Library) 分頁新增。',
                parent=self)
            return
        dlg = _DcdcPickerDialog(self, library, net_name)
        self.wait_window(dlg)
        if dlg.remove_id is not None:
            self._app.pwr_db.delete_source(dlg.remove_id)
            self._refresh()
            return
        if dlg.result is None:
            return
        self._app.pwr_db.upsert_source(dlg.result)
        self._refresh()

    def _change_dcdc_for_source(self, iid):
        src = self._source_row[iid]
        library = self._app.pwr_db.get_library_sources()
        dlg = _DcdcPickerDialog(self, library, src['rail_name'], existing_source=src)
        self.wait_window(dlg)
        if dlg.remove_id is not None:
            self._app.pwr_db.delete_source(dlg.remove_id)
            self._refresh()
            return
        if dlg.result is None:
            return
        self._app.pwr_db.upsert_source(dlg.result)
        self._refresh()

    def _refresh(self):
        roots = self._app._pwr_build_power_tree_data(self._checked)
        self._populate(roots)

    def _export_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror('Error', 'openpyxl 未安裝\n請執行：pip install openpyxl')
            return
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = filedialog.asksaveasfilename(
            initialfile=f'PowerTree_{ts}.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Power Tree'
        try:
            ws.sheet_properties.outlinePr.summaryBelow = False
        except Exception:
            pass

        headers = ['Net / Source / Consumer', 'Typ I(mA)', 'Max I(mA)',
                   'Typ P(W)', 'Max P(W)', 'V', 'Info']
        ws.append(headers)
        hdr_fill = PatternFill('solid', fgColor='4472C4')
        hdr_font = Font(bold=True, color='FFFFFF')
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
        for i, w in enumerate([42, 12, 12, 12, 12, 8, 40], start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        tag_fill = {
            'root':  PatternFill('solid', fgColor='D6EAF8'),
            'cycle': PatternFill('solid', fgColor='F5B7B1'),
        }
        tag_font = {
            'root':   Font(bold=True),
            'source': Font(color='B9770E'),
            'cycle':  Font(color='CC0000'),
            'leaf':   Font(color='555555'),
        }

        def walk(iid, depth):
            tags = self.tv.item(iid, 'tags')
            tag = tags[0] if tags else ''
            text = self.tv.item(iid, 'text')
            values = self.tv.item(iid, 'values')
            ws.append([text] + list(values))
            r = ws.max_row
            ws.cell(row=r, column=1).alignment = Alignment(indent=depth)
            if tag in tag_fill or tag in tag_font:
                for c in range(1, len(headers) + 1):
                    cell = ws.cell(row=r, column=c)
                    if tag in tag_fill:
                        cell.fill = tag_fill[tag]
                    if tag in tag_font:
                        cell.font = tag_font[tag]
            ws.row_dimensions[r].outline_level = depth
            for child_iid in self.tv.get_children(iid):
                walk(child_iid, depth + 1)

        for root_iid in self.tv.get_children(''):
            walk(root_iid, 0)

        ws.freeze_panes = 'A2'
        wb.save(path)
        messagebox.showinfo('Done', f'已儲存至\n{path}')


class _DcdcPickerDialog(tk.Toplevel):
    """Pick a pre-designed DC/DC from the independent library to power a given net."""

    def __init__(self, parent, library_rows, net_name, existing_source=None):
        super().__init__(parent)
        verb = 'Change' if existing_source is not None else 'Pick'
        self.title(f'{verb} DC/DC for "{net_name or "(unnamed rail)"}"')
        self.geometry('620x400' if existing_source is not None else '620x360')
        self.resizable(True, True)
        self.grab_set()
        self.result = None
        self.remove_id = None
        self._net_name = net_name
        self._existing = existing_source
        self._rows_by_id = {r['id']: r for r in library_rows}

        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        row0 = 0
        if existing_source is not None:
            eff = existing_source['efficiency'] or 1.0
            ttk.Label(self, foreground='gray', font=FONT_UI, text=(
                f"目前套用：{existing_source['regulator_type'] or '(unknown)'}  "
                f"Eff={eff:.0%}  Vin={existing_source['source_rail']}  "
                f"Vout={existing_source['output_voltage']:.2f}V")).grid(
                row=0, column=0, columnspan=2, sticky='w', padx=8, pady=(8, 0))
            row0 = 1

        cols = ('regulator_type', 'efficiency', 'input_voltage', 'output_voltage',
                 'source_rail', 'remark')
        tv = self.tv = ttk.Treeview(self, columns=cols, show='headings', selectmode='browse')
        for cid, hdr, w in [
            ('regulator_type', 'Type',         90), ('efficiency',    'Eff',        60),
            ('input_voltage',  'Vin(V)',       70), ('output_voltage', 'Vout(V)',   70),
            ('source_rail',    'Template Source Rail', 130), ('remark', 'Remark',  160),
        ]:
            tv.heading(cid, text=hdr)
            tv.column(cid, width=w, anchor='center', minwidth=40)
        vsb = ttk.Scrollbar(self, orient='vertical', command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        tv.grid(row=row0, column=0, sticky='nsew', padx=(8, 0), pady=8)
        vsb.grid(row=row0, column=1, sticky='ns', pady=8)
        for r in library_rows:
            tv.insert('', 'end', iid=str(r['id']), values=(
                r['regulator_type'] or '', f"{r['efficiency']:.2f}",
                f"{r['input_voltage']:.2f}", f"{r['output_voltage']:.2f}",
                r['source_rail'] or '', r['remark'] or ''))
        tv.bind('<<TreeviewSelect>>', self._on_select)

        self.sv_source_rail = tk.StringVar()
        if existing_source is not None:
            self.sv_source_rail.set(existing_source['source_rail'] or '')
            if existing_source['id'] in self._rows_by_id:
                tv.selection_set(str(existing_source['id']))
                tv.see(str(existing_source['id']))

        bottom = ttk.Frame(self)
        bottom.grid(row=row0 + 1, column=0, columnspan=2, sticky='ew', padx=8, pady=(0, 8))
        ttk.Label(bottom, text='Source Rail (Vin)：', font=FONT_UI).pack(side='left')
        ttk.Entry(bottom, textvariable=self.sv_source_rail, width=24).pack(
            side='left', padx=(4, 12))
        if existing_source is not None:
            ttk.Button(bottom, text='Remove Source', command=self._remove).pack(side='left')
        ttk.Button(bottom, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bottom, text='OK', command=self._ok).pack(side='right')

    def _on_select(self, _=None):
        sel = self.tv.selection()
        if not sel:
            return
        r = self._rows_by_id[int(sel[0])]
        self.sv_source_rail.set(r['source_rail'] or '')

    def _remove(self):
        if not messagebox.askyesno(
                'Confirm',
                f'移除這個 Source？\n"{self._net_name}" 之後會變回沒有供電來源的 root。',
                parent=self):
            return
        self.remove_id = self._existing['id']
        self.destroy()

    def _ok(self):
        sel = self.tv.selection()
        if not sel:
            messagebox.showwarning('Warning', '請先選擇一顆 DC/DC', parent=self)
            return
        source_rail = self.sv_source_rail.get().strip()
        if not source_rail:
            messagebox.showwarning('Warning', 'Source Rail (Vin) 不可為空', parent=self)
            return
        tmpl = self._rows_by_id[int(sel[0])]
        self.result = {
            'rail_name':      self._net_name,
            'source_rail':    source_rail,
            'regulator_type': tmpl['regulator_type'] or '',
            'efficiency':     tmpl['efficiency'] or 1.0,
            'input_voltage':  tmpl['input_voltage'] or 0,
            'output_voltage': tmpl['output_voltage'] or 0,
            'remark':         f"applied from library #{tmpl['id']}",
            'part_id':        None,
        }
        if self._existing is not None:
            self.result['id'] = self._existing['id']
        self.destroy()


class _RailNameDialog(tk.Toplevel):
    """Rail Name prompt with an editable combobox of every rail_name already in the DB,
    so renames can be picked consistently instead of retyped (and risk a typo that breaks
    the Power Tree's string-matching)."""

    def __init__(self, parent, prompt, initial, choices):
        super().__init__(parent)
        self.title('Rail Name')
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        ttk.Label(self, text=prompt, font=FONT_UI).grid(
            row=0, column=0, sticky='w', padx=12, pady=(12, 4))
        self.sv = tk.StringVar(value=initial)
        cb = ttk.Combobox(self, textvariable=self.sv, values=choices, width=32)
        cb.grid(row=1, column=0, sticky='ew', padx=12, pady=4)
        cb.focus_set()
        cb.bind('<Return>', lambda _: self._ok())

        bf = ttk.Frame(self)
        bf.grid(row=2, column=0, sticky='e', padx=12, pady=(8, 12))
        ttk.Button(bf, text='Cancel', command=self.destroy).pack(side='right', padx=4)
        ttk.Button(bf, text='OK', command=self._ok).pack(side='right')

    def _ok(self):
        self.result = self.sv.get()
        self.destroy()


def main():
    root = tk.Tk()
    MouserApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
