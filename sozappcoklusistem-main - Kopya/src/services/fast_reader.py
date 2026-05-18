# -*- coding: utf-8 -*-
"""
src/services/fast_reader.py
===========================
Direkt XML parsing ile hızlı Excel okuma modülü.

openpyxl read_only yerine kullan → 4-5x hızlı, 10x daha az bellek.

Yöntem:
  - xlsx = zip arşivi; her platform sayfası = ~1GB XML (sıkıştırılmış ~75MB)
  - SharedStrings tablosu küçük (34KB) → anında yükle
  - İndeks için sadece GENEL satırları bul (191K satır yerine 92 satır parse)
  - SQLite önbellek: ilk yükleme ~15s, sonraki yüklemeler <0.1s

Kullanım:
    reader = FastExcelReader(path)

    # Sözleşme listesi (indeks için)
    index = reader.build_index()          # [{platform, row, no, user, ...}]

    # Tek sözleşmenin tüm satırları (düzenleme için)
    rows = reader.read_contract_block(platform, start_row, end_row)

    # Önbellek varsa oku, yoksa tara
    index = reader.load_or_scan()
"""
from __future__ import annotations

import json
import re
import sqlite3
import time
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Proje içi import (bu modül src/services/ altında çalışır)
try:
    from src.config.app_config import LOG_FOLDER_NAME as _LOG_FOLDER_NAME
except ImportError:
    _LOG_FOLDER_NAME = "sozlesme_takip_sistemi_log"

try:
    from src.config.app_config import EXTRA_SYSTEM_SHEET_NAMES as _EXTRA_SYS
except ImportError:
    _EXTRA_SYS = set()

# ── Sabitler ──────────────────────────────────────────────────────────────────

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DATA_START_ROW = 6   # Platform sayfalarında veri bu satırdan başlar
_GENEL_LABEL  = "GENEL"

# BASE_HEADERS sırasına göre sütun indeksleri (1-tabanlı)
COL_NO             = 1   # Sözleşme No
COL_USER           = 2   # Kullanıcı
COL_YI_YD          = 3   # Yİ/YD
COL_TYPE           = 4   # Sözleşme Tipi
COL_ACTIVITY       = 5   # Faaliyetler (GENEL filtre)
COL_DELIVERY       = 6   # Teslimat / Kabul
COL_CONTENT        = 7   # Sözleşme İçeriği
COL_SIG_DATE       = 8   # İmza Tarihi
COL_T0_DATE        = 9   # T0 Başlangıç
COL_T0_MONTHS      = 10  # T0+Ay
COL_COMPLETION     = 11  # Termin Tarihi
COL_STATUS         = 12  # Durum
COL_ACCEPTANCE     = 13  # Kabul Tarihi
COL_NOTE           = 14  # Not

MAX_INDEX_COL = COL_NOTE   # İndeks için gereken son sütun


# ── Yardımcılar ───────────────────────────────────────────────────────────────

def _col_num(b: bytes) -> int:
    """Sütun harflerini sayıya çevirir: b'A' → 1, b'Z' → 26, b'AA' → 27"""
    n = 0
    for c in b:
        n = n * 26 + (c - 65 + 1)
    return n


# Hücre parser: <c attrs>inner</c> VEYA <c attrs/>  (self-closing)
# Self-closing = boş hücre → atla
_RE_CELL = re.compile(rb'<c\s+([^>]+?)(?:/>|>(.*?)</c>)', re.DOTALL)
_RE_R    = re.compile(rb'\br="([A-Z]+)\d+"')
_RE_T    = re.compile(rb'\bt="([^"]*)"')
_RE_V    = re.compile(rb'<v>([^<]*)</v>')


def _parse_row_seg(seg: bytes, ss: List[str], max_col: Optional[int]) -> dict:
    """
    Bir <row>...</row> segmentinin hücrelerini parse eder.
    Self-closing hücreleri (<c r="M6" s="62"/>) doğru şekilde atlar.
    Döner: {col_num: value, ...}
    """
    rd: dict = {}

    for m in _RE_CELL.finditer(seg):
        attrs = m.group(1)
        inner = m.group(2)       # None → self-closing → değer yok
        if inner is None:
            continue

        rm = _RE_R.search(attrs)
        if not rm:
            continue
        ref = rm.group(1)
        ci = 0
        while ci < len(ref) and ref[ci] >= 65:
            ci += 1
        cn = _col_num(ref)
        if max_col is not None and cn > max_col:
            continue

        vm = _RE_V.search(inner)
        if not vm:
            continue
        raw = vm.group(1).decode('utf-8', 'replace')

        tm = _RE_T.search(attrs)
        if tm and tm.group(1) == b's':
            try:
                idx = int(raw)
                rd[cn] = ss[idx] if 0 <= idx < len(ss) else raw
            except ValueError:
                rd[cn] = raw
        else:
            try:
                fv = float(raw)
                rd[cn] = int(fv) if fv == int(fv) else fv
            except ValueError:
                rd[cn] = raw

    return rd


def _row_to_list(rd: dict, max_col: int) -> List:
    return [rd.get(i + 1) for i in range(max_col)]


# ── Ana sınıf ─────────────────────────────────────────────────────────────────

class FastExcelReader:
    """
    Hızlı Excel okuma — openpyxl kullanmaz, doğrudan XML parse eder.

    - İndeks taraması: 15s (vs openpyxl ~60s)
    - SQLite cache ile: 0.1s
    - Bellek: max 1GB anlık (1 platform), hemen serbest bırakır
    """

    def __init__(self, path: Path):
        self.path = Path(path)
        self._ss: Optional[List[str]] = None            # Shared strings
        self._sheet_map: Optional[Dict[str, str]] = {}  # platform → xml dosyası
        self._genel_pat: Optional[bytes] = None
        self._genel_re: Optional[re.Pattern] = None

    # ── Başlangıç ──────────────────────────────────────────────

    def _load_shared_strings(self) -> List[str]:
        if self._ss is not None:
            return self._ss
        try:
            with zipfile.ZipFile(self.path, 'r') as zf:
                names = zf.namelist()
                # sharedStrings.xml olmayan Excel'ler (tüm değerler inline/sayısal)
                if 'xl/sharedStrings.xml' not in names:
                    self._ss = []
                    return self._ss
                ss_bytes = zf.read('xl/sharedStrings.xml')
            self._ss = [
                s.decode('utf-8', 'replace')
                for s in re.findall(rb'<t[^>]*>([^<]*)</t>', ss_bytes)
            ]
        except Exception:
            self._ss = []
        return self._ss

    def _load_sheet_map(self) -> Dict[str, str]:
        """Platform adı → xl/worksheets/sheetN.xml eşleşmesi"""
        if self._sheet_map:
            return self._sheet_map
        with zipfile.ZipFile(self.path, 'r') as zf:
            wb_xml  = zf.read('xl/workbook.xml').decode('utf-8')
            rel_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')

        sheets = re.findall(r'<sheet\s+name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)
        rels   = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rel_xml))

        mapping: Dict[str, str] = {}
        for name, rid in sheets:
            target = rels.get(rid, '')
            if target:
                xml_path = f"xl/{target}" if not target.startswith('xl/') else target
                mapping[name] = xml_path

        self._sheet_map = mapping
        return mapping

    def _platform_names(self) -> List[str]:
        """Sistem sayfaları dışındaki platform sayfalarını döndürür."""
        sys_sheets = {
            'ANASAYFA', 'Değişiklik Kayıtları', 'Kullanıcılar', 'VeriÇekme',
            'Sistem Bileşenleri', '_Config', '_Data', '_Lists', 'Yedek_SB',
            'Etiketler', 'Config', 'Platform Logolari', 'SistemTipleri',
            '_Meta', '_PlatformConfig',
        }
        return [n for n in self._load_sheet_map() if n not in sys_sheets]

    def _ensure_genel_helpers(self, ss: List[str]):
        if self._genel_pat is not None:
            return
        try:
            idx = ss.index(_GENEL_LABEL)
        except ValueError:
            idx = -1
        if idx >= 0:
            self._genel_pat = f'<v>{idx}</v>'.encode()
            self._genel_re  = re.compile(
                rb'<c\s[^>]*r="E\d+"[^>]*>(?:[^<]*)<v>' +
                str(idx).encode() + rb'</v>'
            )
        else:
            # GENEL inline string ise (nadiren)
            self._genel_pat = b'>GENEL<'
            self._genel_re  = re.compile(rb'r="E\d+"[^>]*>[^<]*GENEL')

    # ── İndeks Taraması ────────────────────────────────────────

    def scan_platform_index(
        self,
        platform: str,
        xml_bytes: bytes,
        ss: List[str],
    ) -> List[dict]:
        """
        Bir platform sayfasının bytes verisinden sözleşme özetlerini çıkarır.
        Sadece col 5 (E) == 'GENEL' olan satırları okur.

        Döner: [{row, platform, no, user, yi_yd, type, status,
                  completion_date, acceptance_date, content, note, ...}]
        """
        self._ensure_genel_helpers(ss)
        genel_pat = self._genel_pat
        genel_re  = self._genel_re
        results: List[dict] = []

        pos = 0
        while True:
            p = xml_bytes.find(genel_pat, pos)
            if p == -1:
                break

            row_start = xml_bytes.rfind(b'<row ', 0, p)
            row_end   = xml_bytes.find(b'</row>', p)
            if row_start == -1 or row_end == -1:
                pos = p + 1
                continue

            seg = xml_bytes[row_start:row_end]

            # Col E kontrolü (sahte GENEL eşleşmelerini filtrele)
            if genel_re and not genel_re.search(seg):
                pos = row_end + 1
                continue

            # Row numarasını al
            rn_s = seg.find(b' r="', 0, 80)
            if rn_s == -1:
                pos = row_end + 6
                continue
            rn_s += 4
            rn_e = seg.index(b'"', rn_s)
            row_num = int(seg[rn_s:rn_e])

            # Tüm sütunları parse et (14'e kadar)
            rd = _parse_row_seg(seg, ss, MAX_INDEX_COL)

            if rd:
                no_raw = rd.get(COL_NO, '')
                results.append({
                    'row':             row_num,
                    'platform':        platform,
                    'no':              str(no_raw) if no_raw is not None else '',
                    'user':            str(rd.get(COL_USER, '') or ''),
                    'yi_yd':           str(rd.get(COL_YI_YD, '') or 'Yİ'),
                    'type':            str(rd.get(COL_TYPE, '') or ''),
                    'status':          str(rd.get(COL_STATUS, '') or ''),
                    'completion_date': str(rd.get(COL_COMPLETION, '') or ''),
                    'acceptance_date': str(rd.get(COL_ACCEPTANCE, '') or ''),
                    'content':         str(rd.get(COL_CONTENT, '') or ''),
                    'note':            str(rd.get(COL_NOTE, '') or ''),
                    'sig_date':        str(rd.get(COL_SIG_DATE, '') or ''),
                    't0_date':         str(rd.get(COL_T0_DATE, '') or ''),
                    't0_months':       rd.get(COL_T0_MONTHS, 0) or 0,
                    'delivery':        str(rd.get(COL_DELIVERY, '') or ''),
                })

            pos = row_end + 6

        return results

    def build_index(
        self,
        progress_cb=None,   # progress_cb(percent: int, message: str)
    ) -> List[dict]:
        """
        Tüm platform sayfalarını tarar, sözleşme indeksini oluşturur.
        Sıralı çalışır (bellek dostu — max 1GB anlık).
        """
        ss = self._load_shared_strings()
        sheet_map = self._load_sheet_map()
        platforms = self._platform_names()

        if not platforms:
            return []

        all_rows: List[dict] = []
        total = len(platforms)

        for i, platform in enumerate(platforms):
            xml_file = sheet_map.get(platform, '')
            if not xml_file:
                continue

            pct = int((i / total) * 80) + 5
            if progress_cb:
                progress_cb(pct, f"Hızlı tarama: {platform} ({i+1}/{total})")

            try:
                with zipfile.ZipFile(self.path, 'r') as zf:
                    xml_bytes = zf.read(xml_file)
            except Exception:
                continue

            platform_rows = self.scan_platform_index(platform, xml_bytes, ss)
            all_rows.extend(platform_rows)
            del xml_bytes

        if progress_cb:
            progress_cb(85, f"İndeks tamamlandı — {len(all_rows)} sözleşme")

        return all_rows

    # ── Tek Sözleşme Bloğu ─────────────────────────────────────

    def read_contract_block(
        self,
        platform: str,
        start_row: int,
        end_row: int,
        max_col: Optional[int] = None,
    ) -> List[Tuple[int, List]]:
        """
        Belirli bir sözleşmenin satırlarını okur (tüm sütunlar).
        start_row..end_row arası satırları döndürür.

        Hızlı: sadece başlangıç pozisyonuna kadar okur (scan-to-row).
        Döner: [(row_num, [col1, col2, ..., col639]), ...]
        """
        ss = self._load_shared_strings()
        sheet_map = self._load_sheet_map()
        xml_file = sheet_map.get(platform, '')
        if not xml_file:
            return []

        # Pattern: <row r="start_row"... ya da daha büyük
        row_pat = re.compile(rb'<row\b[^>]*\br="(\d+)"')
        results: List[Tuple[int, List]] = []

        with zipfile.ZipFile(self.path, 'r') as zf:
            with zf.open(xml_file) as f:
                buf = b''
                in_range = False
                ended = False

                for chunk in iter(lambda: f.read(524288), b''):  # 512KB chunks
                    buf += chunk

                    while b'</row>' in buf:
                        row_end = buf.find(b'</row>')
                        seg = buf[:row_end + 6]
                        buf = buf[row_end + 6:]

                        # Row num bul
                        m = row_pat.search(seg)
                        if not m:
                            continue
                        rn = int(m.group(1))

                        if rn < start_row:
                            continue
                        if rn > end_row:
                            ended = True
                            break

                        rd = _parse_row_seg(seg, ss, max_col)
                        if rd:
                            mc = max(rd.keys()) if max_col is None else max_col
                            results.append((rn, _row_to_list(rd, mc)))

                    if ended:
                        break

        return results

    # ── SQLite Önbellekleme ────────────────────────────────────

    def _cache_db_path(self) -> Path:
        d = self.path.parent / _LOG_FOLDER_NAME
        d.mkdir(parents=True, exist_ok=True)
        return d / 'fast_index_cache.db'

    def _excel_signature(self) -> str:
        try:
            st = self.path.stat()
            return f"{int(st.st_mtime_ns)}_{st.st_size}"
        except Exception:
            return ''

    def load_from_cache(self) -> Optional[List[dict]]:
        """Cache geçerliyse indeksi döndürür, geçersizse None."""
        try:
            db = self._cache_db_path()
            if not db.exists():
                return None
            conn = sqlite3.connect(str(db))
            conn.row_factory = sqlite3.Row
            try:
                sig = conn.execute(
                    "SELECT value FROM meta WHERE key='excel_sig'"
                ).fetchone()
                if not sig or sig[0] != self._excel_signature():
                    return None
                rows = conn.execute("SELECT * FROM contracts ORDER BY rowid").fetchall()
                return [dict(r) for r in rows]
            finally:
                conn.close()
        except Exception:
            return None

    def save_to_cache(self, index: List[dict]) -> None:
        """İndeksi SQLite önbelleğe yazar."""
        try:
            db = self._cache_db_path()
            conn = sqlite3.connect(str(db))
            conn.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS contracts (
                    platform TEXT, row INT, no TEXT, user TEXT, yi_yd TEXT,
                    type TEXT, status TEXT, completion_date TEXT,
                    acceptance_date TEXT, content TEXT, note TEXT,
                    sig_date TEXT, t0_date TEXT, t0_months REAL, delivery TEXT
                )
            """)
            conn.execute("DELETE FROM contracts")
            conn.execute("DELETE FROM meta")

            conn.executemany(
                """INSERT INTO contracts VALUES
                   (:platform,:row,:no,:user,:yi_yd,:type,:status,
                    :completion_date,:acceptance_date,:content,:note,
                    :sig_date,:t0_date,:t0_months,:delivery)""",
                index
            )
            conn.execute(
                "INSERT INTO meta VALUES ('excel_sig', ?)",
                (self._excel_signature(),)
            )
            conn.commit()
            conn.close()
        except Exception:
            pass

    def load_or_scan(self, progress_cb=None) -> List[dict]:
        """
        Önbellek varsa (<0.1s) onu kullan, yoksa tara (~15s) ve kaydet.
        """
        cached = self.load_from_cache()
        if cached is not None:
            if progress_cb:
                progress_cb(90, f"Önbellekten yüklendi — {len(cached)} sözleşme")
            return cached

        index = self.build_index(progress_cb=progress_cb)
        self.save_to_cache(index)
        return index

    # ── Küçük Sayfalar (openpyxl'siz) ─────────────────────────

    def read_small_sheet(self, sheet_name: str) -> List[List]:
        """
        Kullanıcılar, Sistem Bileşenleri gibi küçük sayfaları okur.
        Tüm sütunlar, tüm satırlar.
        """
        ss = self._load_shared_strings()
        sheet_map = self._load_sheet_map()
        xml_file = sheet_map.get(sheet_name, '')
        if not xml_file:
            return []

        with zipfile.ZipFile(self.path, 'r') as zf:
            xml_bytes = zf.read(xml_file)

        rows: List[List] = []
        for seg in xml_bytes.split(b'</row>'):
            rp = seg.rfind(b'<row ')
            if rp == -1:
                continue
            # Row num
            rnp = seg.find(b' r="', rp, rp + 80)
            if rnp == -1:
                continue
            rnp += 4
            rne = seg.index(b'"', rnp)
            row_num = int(seg[rnp:rne])

            content = seg[seg.find(b'>', rp) + 1:]
            rd = _parse_row_seg(content, ss, None)
            if rd:
                mc = max(rd.keys())
                rows.append((row_num, _row_to_list(rd, mc)))

        # row_num bazında sırala, sadece değerleri döndür
        rows.sort(key=lambda x: x[0])
        return [r[1] for r in rows]
