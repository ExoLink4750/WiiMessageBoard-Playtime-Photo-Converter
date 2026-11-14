#!/usr/bin/env python3
"""
playtime_log_to_excel.py
========================
Universal Wii Message Board dump converter. Turns any of these into a single
Excel workbook:

  * a "vff" directory dump (vff/YYYY/MM/DD/HH/MM/<TITLEID>_0N/{txt,log,odh}/*.000)
    - the RECOMMENDED dump: it contains everything the others have, plus game
      title IDs and all embedded photos
  * a raw VFF container file (cdbackup.vff / cdb.vff), bare or zipped/rared
    - same records as a vff dump, without the folder structure parsed for you
  * a raw "cdb" directory dump (cdb/YYYY/MM Mon/DD/...), folder or zip/rar
    - human-readable text export, but lacks title IDs and some photos
  * a zip/rar of "Today's Accomplishments" playtimelog .txt files
  * any TWO of the above combined (duplicated messages are merged; when both
    copies hold different facts - e.g. only the vff copy knows a memo's title
    ID - the merged result keeps the richer information)

Workbook layout:

  * "Combined Totals"  - every game/channel title with play time summed across
    all logs (hh:mm), most-played first, plus a GRAND TOTAL row. Live SUMIF
    formulas over the "All Entries" sheet.
  * "Dump Contents"    - per-dump inventory: log days, entries, achievements,
    letters, personal notes, photos, stationery, title IDs, date span.
  * "All Entries"      - every individual log entry (source dump/file, date,
    title, time, 'Others'-overflow flag).
  * "Achievements"     - game-tied memos (Wii Play medals, Wii Sports records
    and pro status, Wii Sports Resort stamps, Mario Kart trophies, ...) with
    Title ID + attachments when the dump provides them.
  * "Letters & Messages" - Animal Crossing letters, Nintendo mail, personal
    hand-written board notes, and Photo Channel / Check Mii Out messages.
  * one sheet per day  - that day's entries, total, and matched memos.

Times are stored as real Excel time values with the elapsed-time [hh]:mm
format, so totals over 24 hours display correctly. The workbook is set to
recalculate on open, so Excel/WPS computes the formulas automatically.

Usage:
    python3 playtime_log_to_excel.py INPUT [INPUT2] [-o OUTPUT.xlsx]

    INPUT may be a folder, a .zip/.rar archive, or a .vff container file.

Requires: openpyxl (pip install openpyxl)
RAR input additionally needs the 'rarfile' package plus an unrar backend
(WinRAR/UnRAR on Windows, unrar/bsdtar on Linux/macOS).
"""

import argparse
import copy
import os
import re
import struct
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font

TIME_FMT = "[hh]:mm"
SHEET_TOTALS = "Combined Totals"
SHEET_CONTENTS = "Dump Contents"
SHEET_ENTRIES = "All Entries"
SHEET_ACHV = "Achievements"
SHEET_LETTERS = "Letters & Messages"

CATEGORY_LABELS = {"achievement": "Achievement", "letter": "Letter",
                   "personal": "Personal note", "channel": "Channel message"}

WII_EPOCH = datetime(2000, 1, 1)  # Wii timestamps: seconds since 2000-01-01

# Wii title IDs (first 4 chars of the folder name, e.g. "RSPE_01" -> "RSPE")
TITLE_IDS = {
    "HAEA": "Wii Message Board",
    "HAYA": "Photo Channel",
    "HAPE": "Check Mii Out Channel",
    "HABA": "Wii Shop Channel",
    "HACA": "Mii Channel",
    "RSPE": "Wii Sports",
    "RHAE": "Wii Play",
    "RZTE": "Wii Sports Resort",
    "RMCE": "Mario Kart Wii",
    "RUUE": "Animal Crossing: City Folk",
    "RFNE": "Wii Fit",
    "RFP": "Wii Fit Plus",
    "SMNE": "Super Mario Galaxy",
    "SB4E": "Super Mario Galaxy 2",
    "RVOE": "Wii Music",
}
TITLE_ID_PREFIX3 = {  # fallback on the 3-char game code, region letter unknown
    "HAE": "Wii Message Board", "HAY": "Photo Channel", "HAP": "Check Mii Out Channel",
    "HAB": "Wii Shop Channel", "HAC": "Mii Channel",
    "RSP": "Wii Sports", "RHA": "Wii Play", "RZT": "Wii Sports Resort",
    "RMC": "Mario Kart Wii", "RUU": "Animal Crossing: City Folk",
    "RFN": "Wii Fit", "RFP": "Wii Fit Plus",
}


def title_id_name(tid):
    """Human-readable game name for a 4-char Wii title ID (best effort)."""
    if not tid:
        return ""
    if tid in TITLE_IDS:
        return TITLE_IDS[tid]
    if tid[:3] in TITLE_ID_PREFIX3:
        return TITLE_ID_PREFIX3[tid[:3]]
    return f"Unknown title ({tid})"


# ===========================================================================
# Message model
# ===========================================================================

@dataclass
class PlayLog:
    date: datetime            # message (send) date
    entries: list             # [(title, minutes)]
    total: int                # minutes
    source: str               # path within the dump
    dump: str = ""            # dump label
    last_edit: datetime = None
    title_id: str = ""        # always HAEA for board-generated logs, when known

    @property
    def day(self):
        return self.date.strftime("%Y-%m-%d")


@dataclass
class Memo:
    date: datetime
    title: str
    flat: str                 # whitespace-flattened body
    source: str
    dump: str = ""
    title_id: str = ""
    kind: str = ""            # 'txt' or 'odh' (vff record subfolder), or ''
    photos: int = 0           # embedded/attached AJPG images
    arcs: int = 0             # embedded/attached U8 stationery archives


@dataclass
class SourceBundle:
    """A parsed input dump plus its inventory."""
    label: str
    path: str
    kind: str                 # 'cdb-txt' or 'vff-records'
    playlogs: list = field(default_factory=list)
    memos: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    photos: int = 0
    arcs: int = 0

    @property
    def date_span(self):
        ds = [l.date for l in self.playlogs] + [m.date for m in self.memos]
        ds = [d for d in ds if d and d.year > 2000]
        if not ds:
            return ""
        return f"{min(ds):%Y-%m-%d} .. {max(ds):%Y-%m-%d}"

    @property
    def letters(self):
        return [m for m in self.memos if is_letter(m)]

    def category_counts(self):
        c = Counter()
        for m in self.memos:
            c[memo_category(m)] += 1
        return c

    @property
    def title_ids(self):
        return sorted({m.title_id for m in self.memos if m.title_id})


def is_letter(memo):
    """Letters (personal/Animal Crossing/Nintendo mail) vs. achievement memos."""
    t = memo.title.strip()
    if memo.title_id == "RUUE":
        return True
    if t.lower().startswith("from ") or t == "Nintendo":
        return True
    return False


# ===========================================================================
# Plain-text message parsing (cdb dumps / playtimelog zips)
# ===========================================================================

def parse_play_body(body_lines):
    """Parse the 'Today's Play History' body (title + hh:mm lines) into
    (entries[(title, minutes)], total_minutes). Shared by text logs and raw
    CDBFILE records (the record body is the same text in UTF-16BE)."""
    body = [l for l in body_lines if l.strip()]
    if body and body[0].strip() == "Today's Play History":
        body = body[1:]
    entries, total, buf = [], None, []
    for line in body:
        tm = re.match(r"^\s*(\d+):(\d{2})\s*$", line)
        if tm:
            if not buf:
                raise ValueError(f"time line with no title: {line!r}")
            title = " ".join(buf)
            minutes = int(tm.group(1)) * 60 + int(tm.group(2))
            if title == "Total Play Time":
                total = minutes
                break
            entries.append((title, minutes))
            buf = []
        else:
            buf.append(line.strip())
    if total is None:
        raise ValueError("no 'Total Play Time' found")
    return entries, total


def parse_playtime_text(text):
    """Parse a playtimelog .txt file's full text -> (meta, entries, total)."""
    lines = text.splitlines()
    meta = {}
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("======"):
            i += 1
            break
        m = re.match(r"^(Message title|Sender|Send date|Last edit date):\s*(.*)$", line)
        if m:
            meta[m.group(1)] = m.group(2).strip()
        i += 1
    entries, total = parse_play_body(lines[i:])
    return meta, entries, total


def parse_memo_text(text):
    """Parse a ripl_board_record.txt memo's full text -> dict(title/date/flat)."""
    mt = re.search(r"^Message title:\s*(.*)$", text, re.M)
    sd = re.search(r"^Send date:\s*(.*)$", text, re.M)
    if not (mt and sd and "==================" in text):
        raise ValueError("not a message-board memo")
    body = text.split("==================", 1)[1].strip()
    return {"title": mt.group(1).strip(),
            "date": sd.group(1).strip(),
            "flat": " ".join(body.split())}


def parse_board_date(s):
    """'Sun Feb  1 14:48:36 2009' -> datetime (tolerant of padding)."""
    return datetime.strptime(re.sub(r"\s+", " ", s.strip()).replace("  ", " "),
                             "%a %b %d %H:%M:%S %Y") if s else None


# ===========================================================================
# Raw CDBFILE record parsing (vff dumps and VFF containers)
# ===========================================================================

def _read_utf16z(data, off):
    """Read a NUL-terminated big-endian UTF-16 string starting at off."""
    chars = []
    while off + 1 < len(data):
        cp = (data[off] << 8) | data[off + 1]
        if cp == 0:
            break
        if cp in (0x09, 0x0A, 0x0D) or 0x20 <= cp < 0xD800:
            chars.append(chr(cp))
        else:
            break  # hit binary data
        off += 2
    return "".join(chars)


VFF_PATH_RE = re.compile(
    r"(?:^|/)(\d{4})/(\d{1,2})/(\d{1,2})/(\d{1,2})/(\d{1,2})/([^/]+)/"
    r"(log|txt|odh)/[^/]+\.000$", re.I)


def parse_cdbfile(data):
    """Parse one CDBFILE .000 record.

    Returns dict(name, title, body, last_edit, photos, arcs) or None if the
    bytes are not a CDBFILE record."""
    if len(data) < 0x560 or not data.startswith(b"CDBFILE"):
        return None
    nlen = int.from_bytes(data[0x10:0x14], "big")
    name = data[0x14:0x14 + nlen].split(b"\x00")[0].decode("ascii", "replace")
    ts = int.from_bytes(data[0x7C:0x80], "big")
    last_edit = WII_EPOCH + timedelta(seconds=ts) if 0 < ts < 0xF0000000 else None
    title = _read_utf16z(data, 0x548)
    # body follows the NUL-terminated title, padded to an 8-byte boundary
    boff = 0x548 + 2 * (len(title) + 1)
    boff = (boff + 7) & ~7
    body = _read_utf16z(data, boff) if title else ""
    return {"name": name,
            "title": title,
            "body": body,
            "last_edit": last_edit,
            "photos": data.count(b"AJPG"),
            # stationery: presence flag (a record can hold several nested U8
            # magics, so count the archive once per record)
            "arcs": 1 if b"\x55\xAA\x38\x2D" in data else 0}


def vff_record_date(rec, m):
    """Message date for a vff record: folder path date (months are stored
    0-indexed by the Wii), falling back to the record's last-edit stamp."""
    if m:
        y, mo, d, hh, mi = (int(m.group(i)) for i in range(1, 6))
        mo += 1  # Wii stores month folders 0-indexed (0=January)
        if 1 <= mo <= 12 and 1 <= d <= 31 and hh <= 23 and mi <= 59:
            try:
                return datetime(y, mo, d, hh, mi)
            except ValueError:
                pass
    return rec["last_edit"] or datetime(2000, 1, 1)


# ===========================================================================
# VFF container (FAT16) reader
# ===========================================================================

class VFFContainer:
    """Minimal read-only reader for Wii 'VFF ' container files (cdbackup.vff /
    cdb.vff): a 0x20-byte header, FAT16 table(s), a root directory, then
    cluster-chained data. Layout is derived from the header and validated by
    walking the directory tree, so both one-FAT and two-FAT variants work."""

    def __init__(self, data):
        if data[:4] != b"VFF ":
            raise ValueError("not a VFF container (missing 'VFF ' magic)")
        self.data = data
        self.volume_size = struct.unpack_from(">I", data, 8)[0]
        self.cluster_size = struct.unpack_from(">H", data, 12)[0] * 16
        if not self.cluster_size or self.cluster_size & (self.cluster_size - 1):
            raise ValueError("unsupported VFF cluster size")
        n_clusters = self.volume_size // self.cluster_size
        fat_bytes = -(-(n_clusters * 2) // self.cluster_size) * self.cluster_size
        # probe one-FAT and two-FAT layouts; pick whichever yields a valid root
        self.fat_off = 0x20
        self.root_off = None
        for n_fats in (2, 1):
            cand = 0x20 + n_fats * fat_bytes
            if self._valid_dir(cand):
                self.root_off = cand
                break
        if self.root_off is None:
            raise ValueError("could not locate the VFF root directory")
        # The root region is 0x1000 bytes in practice (WiiVFF reads a fixed
        # 0x1000); fall back to the used size rounded up to a cluster, and
        # validate the choice by probing the first subdirectory.
        used = 0
        for off in range(self.root_off, len(data) - 32, 32):
            if data[off] == 0:
                break
            used += 32
        rounded = max(-(-used // self.cluster_size) * self.cluster_size,
                      self.cluster_size)
        first_dir_cluster = next(
            (struct.unpack_from("<H", data, self.root_off + i + 26)[0]
             for i in range(0, used, 32)
             if data[self.root_off + i] not in (0x00, 0xE5)
             and data[self.root_off + i + 11] == 0x10
             and data[self.root_off + i + 26:i + 28 + self.root_off] != b"\x00\x00"),
            None)
        self.data_off = None
        for root_bytes in (0x1000, rounded):
            cand = self.root_off + root_bytes
            if first_dir_cluster is None or self._valid_dir(
                    cand + (first_dir_cluster - 2) * self.cluster_size):
                self.data_off = cand
                break
        if self.data_off is None:
            raise ValueError("could not locate the VFF data region")

    def _cluster_bytes(self, cl):
        off = self.data_off + (cl - 2) * self.cluster_size
        return self.data[off:off + self.cluster_size]

    def _fat_entry(self, cl):
        return struct.unpack_from("<H", self.data, self.fat_off + cl * 2)[0]

    def _chain(self, start, limit=1 << 20):
        cl, seen = start, set()
        while 2 <= cl < 0xFFF8 and cl not in seen and len(seen) < limit:
            seen.add(cl)
            yield cl
            cl = self._fat_entry(cl)

    def _read_chain(self, start, size):
        out = b"".join(self._cluster_bytes(c) for c in self._chain(start))
        return out[:size]

    @staticmethod
    def _entry_ok(e):
        if e[0] in (0x00, 0xE5) or e[11] == 0x0F:
            return False
        if e[11] not in (0x10, 0x20, 0x21, 0x22, 0x30):
            return False
        name = e[0:8]
        return all(c == 0x20 or 0x21 <= c < 0x7F or c >= 0x80 for c in name)

    def _valid_dir(self, off):
        if off < 0 or off + 96 > len(self.data):
            return False
        ents = [self.data[off + i:off + i + 32] for i in range(0, 96, 32)]
        return sum(self._entry_ok(e) for e in ents) >= 2

    def _dir_entries(self, blob):
        i = 0
        while i + 32 <= len(blob):
            e = blob[i:i + 32]
            i += 32
            if e[0] == 0x00:
                break
            if e[0] == 0xE5 or e[11] == 0x0F:  # deleted / LFN fragment
                continue
            name = e[0:8].decode("ascii", "replace").rstrip()
            ext = e[8:11].decode("ascii", "replace").rstrip()
            fname = f"{name}.{ext}" if ext else name
            yield (fname, e[11], struct.unpack_from("<H", e, 26)[0],
                   struct.unpack_from("<I", e, 28)[0])

    def walk(self):
        """Yield (path, file_bytes) for every file in the container."""
        stack = [("", None)]
        while stack:
            prefix, cluster = stack.pop()
            blob = (self.data[self.root_off:self.data_off] if cluster is None
                    else self._read_chain(cluster, 1 << 22))
            for fname, attr, cl, size in self._dir_entries(blob):
                path = f"{prefix}/{fname}" if prefix else fname
                if attr & 0x10:
                    if fname not in (".", ".."):
                        stack.append((path, cl))
                elif size:
                    yield (path, self._read_chain(cl, size))


# ===========================================================================
# Input adapters: folders, zip/rar archives, bare .vff files
# ===========================================================================

class DirAccessor:
    def __init__(self, root):
        self.root = root
        self._names = None

    def namelist(self):
        if self._names is None:
            out = []
            for root, _dirs, files in os.walk(self.root):
                for fn in files:
                    out.append(os.path.relpath(os.path.join(root, fn), self.root))
            self._names = sorted(out)
        return self._names

    def read(self, name):
        with open(os.path.join(self.root, name), "rb") as f:
            return f.read()


class ZipAccessor:
    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)

    def namelist(self):
        return sorted(n for n in self.zf.namelist() if not n.endswith("/"))

    def read(self, name):
        return self.zf.read(name)


class RarAccessor:
    def __init__(self, path):
        try:
            import rarfile
        except ImportError:
            raise RuntimeError(
                "RAR input needs the 'rarfile' package (pip install rarfile) "
                "plus an unrar backend (WinRAR/UnRAR on Windows, "
                "unrar/bsdtar on Linux).")
        try:
            self.rf = rarfile.RarFile(path)
        except Exception as e:
            raise RuntimeError(f"could not open RAR archive: {e}")

    def namelist(self):
        return sorted(n for n in self.rf.namelist() if not n.endswith("/"))

    def read(self, name):
        try:
            return self.rf.read(name)
        except Exception as e:
            raise RuntimeError(
                f"could not extract '{name}' from the RAR archive ({e}). "
                "RAR decompression needs an unrar backend (WinRAR/UnRAR).")


class VFFAccessor:
    """Presents a raw .vff container's contents as a name->bytes tree."""
    def __init__(self, data):
        self.files = dict(VFFContainer(data).walk())

    def namelist(self):
        return sorted(self.files)

    def read(self, name):
        return self.files[name]


def detect_kind(names):
    """'vff-records' if the tree holds raw .000 board records, else 'cdb-txt'."""
    n_vff = sum(1 for n in names if VFF_PATH_RE.search(n.replace("\\", "/")))
    if n_vff:
        return "vff-records"
    if any(n.lower().endswith(".txt") for n in names):
        return "cdb-txt"
    raise ValueError("no playtimelog .txt files or .000 board records found")


def open_source(path):
    """Open any supported input. Returns (label, accessor, kind)."""
    label = os.path.basename(path.rstrip("/\\"))
    if os.path.isdir(path):
        acc = DirAccessor(path)
        return label, acc, detect_kind(acc.namelist())
    low = path.lower()
    with open(path, "rb") as f:
        magic = f.read(4)
    if magic == b"VFF " or low.endswith(".vff"):
        with open(path, "rb") as f:
            acc = VFFAccessor(f.read())
        return label, acc, detect_kind(acc.namelist())
    if magic[:2] == b"PK" or low.endswith(".zip"):
        acc = ZipAccessor(path)
        names = acc.namelist()
        vff_members = [n for n in names if n.lower().endswith(".vff")]
        if vff_members:  # archive wrapping a raw container, e.g. cdbackup.zip
            data = acc.read(vff_members[0])
            acc2 = VFFAccessor(data)
            return label, acc2, detect_kind(acc2.namelist())
        return label, acc, detect_kind(names)
    if magic == b"Rar!" or low.endswith(".rar"):
        acc = RarAccessor(path)
        names = acc.namelist()
        vff_members = [n for n in names if n.lower().endswith(".vff")]
        if vff_members:
            acc2 = VFFAccessor(acc.read(vff_members[0]))
            return label, acc2, detect_kind(acc2.namelist())
        return label, acc, detect_kind(names)
    raise ValueError(f"unsupported input type: {path}")


# ===========================================================================
# Source loading: turn an input into PlayLog / Memo objects
# ===========================================================================

def load_source(path):
    label, acc, kind = open_source(path)
    bundle = SourceBundle(label=label, path=path, kind=kind)
    names = acc.namelist()

    if kind == "cdb-txt":
        # attachment sidecar files: '<msg> ripl_board_record-attachment.<ext>'
        for name in names:
            low = name.lower()
            if "-attachment." in low:
                ext = low.rsplit(".", 1)[-1]
                if ext == "ajpg":
                    bundle.photos += 1
                elif ext == "arc":
                    bundle.arcs += 1
        for name in names:
            if not name.lower().endswith(".txt"):
                continue
            try:
                text = acc.read(name).decode("utf-8", "replace")
            except Exception as e:
                bundle.skipped.append((name, str(e)))
                continue
            try:
                meta, entries, total = parse_playtime_text(text)
                bundle.playlogs.append(PlayLog(
                    date=parse_board_date(meta.get("Send date", "")) or datetime(2000, 1, 1),
                    entries=entries, total=total, source=name, dump=label,
                    last_edit=parse_board_date(meta.get("Last edit date", ""))))
                continue
            except Exception:
                pass
            try:
                memo = parse_memo_text(text)
                bundle.memos.append(Memo(
                    date=parse_board_date(memo["date"]) or datetime(2000, 1, 1),
                    title=memo["title"], flat=memo["flat"], source=name, dump=label))
            except Exception as e:
                bundle.skipped.append((name, str(e)))
    else:  # vff-records
        for name in names:
            norm = name.replace("\\", "/")
            m = VFF_PATH_RE.search(norm)
            if not m:
                continue
            data = acc.read(name)
            rec = parse_cdbfile(data)
            if rec is None:
                bundle.skipped.append((name, "not a CDBFILE record"))
                continue
            title_folder = m.group(6)
            tid = title_folder.split("_")[0].upper()
            sub = m.group(7).lower()
            bundle.photos += rec["photos"]
            bundle.arcs += rec["arcs"]
            if rec["name"] == "playtimelog" or sub == "log":
                try:
                    entries, total = parse_play_body(rec["body"].splitlines())
                except Exception as e:
                    bundle.skipped.append((name, f"unparsable play history: {e}"))
                    continue
                bundle.playlogs.append(PlayLog(
                    date=vff_record_date(rec, m), entries=entries, total=total,
                    source=norm, dump=label, last_edit=rec["last_edit"], title_id=tid))
            else:
                flat = " ".join(rec["body"].split())
                # hand-written board messages store their text in the title
                # field; the body field then only holds the sender's Mii name
                # preceded by binary data (e.g. '�Zac') - drop it
                if (flat and len(flat) <= 12 and not flat[0].isascii()
                        and rec["title"] and not _mostly_garbage(rec["title"])):
                    flat = ""
                bundle.memos.append(Memo(
                    date=vff_record_date(rec, m), title=rec["title"] or "(untitled)",
                    flat=flat, source=norm, dump=label,
                    title_id=tid, kind=sub, photos=rec["photos"], arcs=rec["arcs"]))
    return bundle


# ===========================================================================
# Memo classification (achievements / letters)
# ===========================================================================

def classify_memo(title, flat_body):
    """Map a memo to (game_title, human-readable detail) from its text."""
    t, b = title, flat_body
    if t == "Welcome to Wii Play.":
        m = re.search(r"player has been added to your Wii Play player list: (\S+)", b)
        return "Wii Play", f"Player added to Wii Play: {m.group(1) if m else '?'}"
    if "Medal Earned" in t:
        medal = t.split(" ")[0]
        m = re.search(r"(\S+) has earned an? \w+ medal in (.+?)[.!]", b)
        return "Wii Play", (f"{m.group(1)} earned a {medal.lower()} medal in {m.group(2)}"
                            if m else f"{medal} medal earned")
    if t == "Wii Sports Debut":
        m = re.search(r"player list: (\S+)", b)
        return "Wii Sports", f"Wii Sports debut: {m.group(1).rstrip('.') if m else '?'} added"
    if t == "New Personal Best Set":
        m = re.search(r"personal best in (\w+): (\S+) Score: (.+?) \w+ Info", b)
        return "Wii Sports", (f"{m.group(2)} set a personal best in {m.group(1)}: {m.group(3)}"
                              if m else "New personal best")
    if t in ("Pro Status Earned", "Pro Status Lost"):
        m = re.search(r"become a (\w+) pro: (\S+)", b) or re.search(r"lost the title of (\w+) pro: (\S+)", b)
        verb = "became a" if t.endswith("Earned") else "lost"
        return "Wii Sports", (f"{m.group(2)} {verb} {m.group(1)} pro" if m else t)
    if t == "Wii Fitness Results":
        m = re.search(r"results for (\S+?): (\d+)", b)
        return "Wii Sports", (f"{m.group(1)}'s Wii Fitness age: {m.group(2)}" if m else "Wii Fitness test")
    if "karting" in b or "every trophy" in b:
        m = re.search(r"Congrats, (.+?)!", b)
        return "Mario Kart Wii", f"{m.group(1) if m else '?'} collected every trophy"
    if "From the Photo Channel" in b:
        return "Photo Channel 1.1", "Photo Channel message"
    if "Souvenir Photo" in b:
        return "Check Mii Out Channel", "Souvenir photo received"
    # Wii Sports Resort: stamps (title = stamp name) and Pro-class promotions.
    # Body: 'Wii Sports Resort News <date> (<dow>) <time> <Event> <Player> ...'
    if t == "Advanced to Pro class!":
        m = re.search(r"News .*?[ap]\.m\. (.*?) (\S+) Advanced to Pro class", b)
        return "Wii Sports Resort", (f"{m.group(2)} advanced to Pro class in {m.group(1)}"
                                      if m else "Advanced to Pro class")
    m = re.search(r"Wii Sports Resort News .*?[ap]\.m\. (.*?) Congratulations", b)
    if m:
        # stamp names are the message title; the body tail carries event,
        # player and sometimes score details, so quote it as-is
        return "Wii Sports Resort", f"\"{t}\" stamp: {m.group(1)}"
    return "—", flat_body[:80]


def _mostly_garbage(s):
    """True when a string is mostly non-printable/binary (e.g. the raw Mii-name
    field of a hand-written board message)."""
    t = s.strip()
    if not t:
        return True
    good = sum(1 for c in t if c.isascii() and (c.isalnum() or c in " .,!?'\u2019-"))
    return good / len(t) < 0.5


def memo_category(memo):
    """'letter' | 'personal' | 'channel' | 'achievement'.

    Achievements are memos tied to a game accomplishment (medals, records,
    pro status, stamps, trophies, fitness results). Letters are mail (Animal
    Crossing 'From …', Nintendo); personal messages are hand-written board
    notes; channel messages are Photo Channel / Check Mii Out notifications."""
    if is_letter(memo):
        return "letter"
    if memo.title_id in ("HAYA", "HAPE") or "From the Photo Channel" in memo.flat \
            or "Souvenir Photo" in memo.flat:
        return "channel"
    if memo.title_id in ("", "HAEA") and (
            _mostly_garbage(memo.flat) or _mostly_garbage(memo.title)):
        return "personal"
    game, _detail = classify_memo(memo.title, memo.flat)
    if game != "—" or (memo.title_id and memo.title_id != "HAEA"):
        return "achievement"
    return "personal" if memo.title_id in ("", "HAEA") else "channel"


def memo_game_detail(memo):
    """(game, detail) for a memo, preferring title-ID attribution for letters."""
    cat = memo_category(memo)
    if cat == "letter":
        game = title_id_name(memo.title_id) if memo.title_id else "Letter"
        preview = memo.flat[:90] + ("…" if len(memo.flat) > 90 else "")
        return game, f"Letter: {preview}"
    if cat == "personal":
        text = memo.title if _mostly_garbage(memo.flat) else memo.flat
        return "Personal message", text[:90]
    game, detail = classify_memo(memo.title, memo.flat)
    if game == "—" and memo.title_id and memo.title_id != "HAEA":
        game = title_id_name(memo.title_id)
    return game, detail


def readable_title(title):
    """Some memos (e.g. Mario Kart Wii) have binary/garbled titles."""
    if not title or any(ord(c) > 0x2FFF or c == "�" for c in title):
        return "(unreadable title)"
    return title


def days(minutes):
    """Excel stores times as fractions of a day."""
    return minutes / 1440.0


# ===========================================================================
# Workbook building
# ===========================================================================

def _text_sig(*texts):
    """ASCII-alphanumeric signature of message text; ignores the binary garbage
    that cdb text exports carry in the 'Message title' line of hand-written
    board messages (their real text sits in the body)."""
    return "".join(re.findall(r"[a-z0-9]+", " ".join(texts).lower()))[:120]


def _merge_memo(kept, other):
    """Fold a duplicate memo's metadata into the kept copy: a vff record knows
    its title ID and attachment counts, a cdb export may have the better text
    - merging means input order no longer decides which facts survive."""
    if not kept.title_id and other.title_id:
        kept.title_id = other.title_id
    if not kept.kind and other.kind:
        kept.kind = other.kind
    kept.photos = max(kept.photos, other.photos)
    kept.arcs = max(kept.arcs, other.arcs)
    if _mostly_garbage(kept.title) and not _mostly_garbage(other.title):
        kept.title = other.title
    if not kept.flat and other.flat:
        kept.flat = other.flat
    if not kept.date and other.date:
        kept.date = other.date


def _memo_key(memo):
    """Dedup key: minute-precision date + text signature. The signature counts
    the text once even when title and body hold the same string (hand-written
    board messages, or a memo already merged from two dumps), which keeps the
    key stable under re-runs."""
    sig_t = _text_sig(memo.title)
    sig_f = _text_sig(memo.flat)
    parts = [s for s in (sig_t, sig_f) if s]
    sig = parts[0] if len(parts) == 2 and parts[0] == parts[1] else "".join(parts)
    return (memo.date.replace(second=0, microsecond=0), sig)


def dedup_bundles(bundles):
    """Drop messages that appear in several dumps (same console dumped twice),
    merging metadata into a copy so the richer facts win regardless of input
    order and the input bundles stay unmodified.
    Returns (playlogs, memos, n_log_dupes, n_memo_dupes)."""
    logs, memos, seen_l, seen_m = [], [], set(), {}
    dl = dm = 0
    for b in bundles:
        for lg in b.playlogs:
            key = (lg.day, tuple(lg.entries))
            if key in seen_l:
                dl += 1
                continue
            seen_l.add(key)
            logs.append(lg)
        for mm in b.memos:
            key = _memo_key(mm)
            if key in seen_m:
                dm += 1
                _merge_memo(memos[seen_m[key]], mm)
                continue
            seen_m[key] = len(memos)
            memos.append(copy.copy(mm))
    return logs, memos, dl, dm


def build_workbook(bundles, out_path):
    """bundles: list of SourceBundle (or paths, which get loaded)."""
    bundles = [b if isinstance(b, SourceBundle) else load_source(b) for b in bundles]
    playlogs, memos, dup_logs, dup_memos = dedup_bundles(bundles)
    if not playlogs and not memos:
        raise ValueError("No usable messages found in the input dump(s).")

    playlogs.sort(key=lambda l: (l.date, l.source))
    for b in bundles:
        if b.skipped:
            print(f"NOTE [{b.label}]: skipped {len(b.skipped)} unreadable file(s)")
    if dup_logs or dup_memos:
        print(f"Merged dumps: skipped {dup_logs} duplicate log(s), "
              f"{dup_memos} duplicate memo(s) already covered by another dump")

    # classify memos and group by day, for per-day sheet annotations
    memo_rows = []        # achievements only
    letter_rows = []      # letters, personal notes, channel messages
    memos_by_day = {}
    for mm in memos:
        game, detail = memo_game_detail(mm)
        cat = memo_category(mm)
        att = []
        if mm.photos:
            att.append(f"{mm.photos} photo(s)")
        if mm.arcs:
            att.append(f"{mm.arcs} stationery")
        row = (mm.date, game, mm.title_id, readable_title(mm.title),
               detail, ", ".join(att), mm.dump, mm.source)
        if cat == "achievement":
            memo_rows.append(row)
        else:
            letter_rows.append((mm.date, CATEGORY_LABELS[cat]) + row[1:])
        memos_by_day.setdefault(mm.date.strftime("%Y-%m-%d"), []).append((game, detail))
    memo_rows.sort()
    letter_rows.sort()

    # combined totals per title (only used to sort sheet 1; cells are SUMIFs)
    totals = {}
    for lg in playlogs:
        for title, mins in lg.entries:
            totals[title] = totals.get(title, 0) + mins
    titles_sorted = sorted(totals, key=lambda t: (-totals[t], t))
    n_entries = sum(len(lg.entries) for lg in playlogs)
    last_data_row = n_entries + 1

    wb = Workbook()

    # ---- Sheet 1: Combined Totals ----
    ws = wb.active
    ws.title = SHEET_TOTALS
    ws["A1"] = "Title"
    ws["B1"] = "Total Play Time (hh:mm)"
    for idx, title in enumerate(titles_sorted, start=2):
        ws.cell(row=idx, column=1, value=title)
        c = ws.cell(row=idx, column=2,
                    value=f"=SUMIF('{SHEET_ENTRIES}'!$C$2:$C${last_data_row},A{idx},"
                          f"'{SHEET_ENTRIES}'!$D$2:$D${last_data_row})")
        c.number_format = TIME_FMT
    grand_row = len(titles_sorted) + 2
    ws.cell(row=grand_row, column=1, value="GRAND TOTAL")
    gc = ws.cell(row=grand_row, column=2, value=f"=SUM(B2:B{grand_row - 1})")
    gc.number_format = TIME_FMT
    for cell in (ws["A1"], ws["B1"], ws.cell(row=grand_row, column=1), gc):
        cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 22
    ws.freeze_panes = "A2"

    # ---- Sheet 2: Dump Contents ----
    wsc = wb.create_sheet(SHEET_CONTENTS)
    headers = ["Dump", "Format", "Date Span", "Play-Log Days", "Log Entries",
               "Achievements", "Letters", "Personal Notes", "Photos (AJPG)",
               "Stationery (ARC)", "Title IDs"]
    for col, h in enumerate(headers, start=1):
        wsc.cell(row=1, column=col, value=h).font = Font(bold=True)
    for r, b in enumerate(bundles, start=2):
        cats = b.category_counts()
        wsc.cell(row=r, column=1, value=b.label)
        wsc.cell(row=r, column=2,
                 value="cdb text export" if b.kind == "cdb-txt" else "raw VFF records")
        wsc.cell(row=r, column=3, value=b.date_span)
        wsc.cell(row=r, column=4, value=len({l.day for l in b.playlogs}))
        wsc.cell(row=r, column=5, value=sum(len(l.entries) for l in b.playlogs))
        wsc.cell(row=r, column=6, value=cats.get("achievement", 0))
        wsc.cell(row=r, column=7, value=cats.get("letter", 0) + cats.get("channel", 0))
        wsc.cell(row=r, column=8, value=cats.get("personal", 0))
        wsc.cell(row=r, column=9, value=b.photos)
        wsc.cell(row=r, column=10, value=b.arcs)
        wsc.cell(row=r, column=11, value=", ".join(b.title_ids) or "—")
    note_r = len(bundles) + 3
    if dup_logs or dup_memos:
        wsc.cell(row=note_r, column=1,
                 value=f"Duplicates merged: {dup_logs} log(s), {dup_memos} memo(s) "
                       "appeared in more than one dump and were counted once.").font = Font(italic=True)
        note_r += 1
    wsc.cell(row=note_r, column=1,
             value="Photos use the Wii's proprietary AJPG format (no public decoder); "
                   "stationery archives are U8/LZ77 containers of TPL textures.").font = Font(italic=True)
    for col, w in zip("ABCDEFGHIJK", (26, 17, 24, 14, 12, 13, 9, 14, 15, 17, 34)):
        wsc.column_dimensions[col].width = w
    wsc.freeze_panes = "A2"

    # ---- Sheet 3: All Entries (source data for the SUMIFs) ----
    ws2 = wb.create_sheet(SHEET_ENTRIES)
    headers = ["Source File", "Send Date", "Title", "Play Time (hh:mm)",
               "Day Has 'Others' Overflow", "Source Dump"]
    for col, h in enumerate(headers, start=1):
        ws2.cell(row=1, column=col, value=h).font = Font(bold=True)
    r = 2
    for lg in playlogs:
        has_others = "Yes" if any(t == "Others" for t, _ in lg.entries) else "No"
        for title, mins in lg.entries:
            ws2.cell(row=r, column=1, value=lg.source)
            ws2.cell(row=r, column=2, value=lg.day)
            ws2.cell(row=r, column=3, value=title)
            c = ws2.cell(row=r, column=4, value=days(mins))
            c.number_format = TIME_FMT
            ws2.cell(row=r, column=5, value=has_others)
            ws2.cell(row=r, column=6, value=lg.dump)
            r += 1
    for col, w in zip("ABCDEF", (55, 12, 36, 18, 24, 22)):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"

    # ---- Sheet 4: Achievements (game-tied memos only) ----
    if memo_rows:
        wsa = wb.create_sheet(SHEET_ACHV)
        headers = ["Date", "Matched Game", "Title ID", "Message Title", "Detail",
                   "Attachments", "Source Dump", "Source File"]
        for col, h in enumerate(headers, start=1):
            wsa.cell(row=1, column=col, value=h).font = Font(bold=True)
        for r, (dt, game, tid, title, detail, att, dump, src) in enumerate(memo_rows, start=2):
            wsa.cell(row=r, column=1, value=dt.strftime("%Y-%m-%d %H:%M"))
            wsa.cell(row=r, column=2, value=game)
            wsa.cell(row=r, column=3, value=tid or "—")
            wsa.cell(row=r, column=4, value=title)
            wsa.cell(row=r, column=5, value=detail)
            wsa.cell(row=r, column=6, value=att)
            wsa.cell(row=r, column=7, value=dump)
            wsa.cell(row=r, column=8, value=src)
        for col, w in zip("ABCDEFGH", (17, 26, 9, 22, 70, 18, 20, 50)):
            wsa.column_dimensions[col].width = w
        wsa.freeze_panes = "A2"

    # ---- Sheet 5: Letters & Messages (letters, personal notes, channel mail) ----
    if letter_rows:
        wsl = wb.create_sheet(SHEET_LETTERS)
        headers = ["Date", "Type", "Game / Channel", "Title ID", "Message Title",
                   "Text", "Attachments", "Source Dump", "Source File"]
        for col, h in enumerate(headers, start=1):
            wsl.cell(row=1, column=col, value=h).font = Font(bold=True)
        for r, (dt, cat, game, tid, title, detail, att, dump, src) in enumerate(letter_rows, start=2):
            wsl.cell(row=r, column=1, value=dt.strftime("%Y-%m-%d %H:%M"))
            wsl.cell(row=r, column=2, value=cat)
            wsl.cell(row=r, column=3, value=game)
            wsl.cell(row=r, column=4, value=tid or "—")
            wsl.cell(row=r, column=5, value=title)
            wsl.cell(row=r, column=6, value=detail)
            wsl.cell(row=r, column=7, value=att)
            wsl.cell(row=r, column=8, value=dump)
            wsl.cell(row=r, column=9, value=src)
        for col, w in zip("ABCDEFGHI", (17, 16, 26, 9, 22, 70, 18, 20, 50)):
            wsl.column_dimensions[col].width = w
        wsl.freeze_panes = "A2"

    # ---- One sheet per day ----
    used_names = {SHEET_TOTALS, SHEET_CONTENTS, SHEET_ENTRIES, SHEET_ACHV, SHEET_LETTERS}
    for lg in playlogs:
        name = lg.day
        base, n = name[:31], 2
        while name in used_names:
            suffix = f" ({n})"
            name = base[:31 - len(suffix)] + suffix
            n += 1
        used_names.add(name)

        wsf = wb.create_sheet(name)
        wsf["A1"] = f"Source file: {lg.source}  [{lg.dump}]"
        wsf["A2"] = f"Send date: {lg.date:%a %b %d %H:%M:%S %Y}"
        if lg.last_edit:
            wsf["A3"] = f"Last edit date: {lg.last_edit:%a %b %d %H:%M:%S %Y}"
        wsf["A5"] = "Title"
        wsf["B5"] = "Play Time (hh:mm)"
        wsf["C5"] = "Achievements / Memos"
        for cc in ("A5", "B5", "C5"):
            wsf[cc].font = Font(bold=True)
        day_memos = {}
        for game, detail in memos_by_day.get(lg.day, []):
            day_memos.setdefault(game, []).append(detail)
        row = 6
        for title, mins in lg.entries:
            wsf.cell(row=row, column=1, value=title)
            c = wsf.cell(row=row, column=2, value=days(mins))
            c.number_format = TIME_FMT
            if title in day_memos:
                wsf.cell(row=row, column=3, value="; ".join(day_memos[title]))
            row += 1
        wsf.cell(row=row, column=1, value="Total Play Time").font = Font(bold=True)
        tc = wsf.cell(row=row, column=2, value=f"=SUM(B6:B{row - 1})")
        tc.number_format = TIME_FMT
        tc.font = Font(bold=True)
        wsf.column_dimensions["A"].width = 34
        wsf.column_dimensions["B"].width = 18
        wsf.column_dimensions["C"].width = 60

    # Ask Excel/WPS to recalculate all formulas when the file is opened
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    wb.save(out_path)
    print(f"Saved: {out_path}")
    print(f"  {len(wb.sheetnames)} sheets | {len(playlogs)} day log(s) | "
          f"{len(memos)} memo(s) | {len(titles_sorted)} unique titles | {n_entries} entries")
    return out_path


# ===========================================================================
# CLI
# ===========================================================================

def main():
    ap = argparse.ArgumentParser(
        description="Convert Wii Message Board dumps (playtimelog zip, cdb "
                    "tree, vff tree, or .vff container - folder, zip or rar) "
                    "to an Excel workbook.")
    ap.add_argument("inputs", nargs="+",
                    help="input path(s): folder, .zip/.rar archive, or .vff "
                         "container; two inputs are merged with deduplication")
    ap.add_argument("-o", "--output", default="Playtime Log Summary.xlsx",
                    help="output .xlsx path (default: 'Playtime Log Summary.xlsx')")
    args = ap.parse_args()
    if len(args.inputs) > 2:
        ap.error("at most two input dumps can be combined")
    bundles = [load_source(p) for p in args.inputs]
    build_workbook(bundles, args.output)


if __name__ == "__main__":
    main()
