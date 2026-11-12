#!/usr/bin/env python3
"""eSAP: build a draft Statistical Analysis Plan from a Word template and a metadata workbook.

    python esap.py --config pioneer1.toml            check the metadata, then build the SAP
    python esap.py --config pioneer1.toml --check-only   check only, write nothing but the log

How it works (the "rubber stamp" model described on the workbook's Read me sheet):

* Every content control in the template is named after where its data comes from.
  - "Sheet.Column (Level)" or "Sheet.Column (each)" is a field: it is filled from the
    current row of that sheet and renamed "Sheet.Column [ID]".
  - "Study.Field" is a single field from the Study sheet.
  - A name without a dot is a group (a stamp). It is copied once per matching row
    (rows of that level, belonging to the enclosing row), each copy renamed "Group [ID]".
    A group holding only other groups is a level container: it is kept once.
  - Names listed under [numbering] in the config are computed numbers
    (e.g. Primary Objective 1, 2, ...).
* A group with no matching rows shows "Not applicable".
* Bookmarks are kept; the finished SAP is checked against the metadata before it is saved.

The script never changes its inputs. Exit codes: 0 built (or checked) OK,
1 metadata problems found (nothing built), 2 set-up or template problem.
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import io
import platform
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from importlib import metadata as importlib_metadata
from pathlib import Path

if sys.version_info < (3, 11):
    sys.exit("eSAP needs Python 3.11 or newer.")
import tomllib  # noqa: E402  (standard library from Python 3.11)

try:
    import openpyxl
    from lxml import etree
except ImportError as exc:  # plain message for people double-clicking run_esap.bat
    sys.exit(f"A Python package is missing ({exc.name}). Run run_esap.bat, which installs it.")

ESAP_VERSION = "1.0"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
}
# Characters Word cannot store (XML 1.0 forbids them).
BAD_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
NAME_RE = re.compile(r"^(?P<base>.*?)(?:\s+\((?P<qual>[^()]+)\))?$")
XML_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


class SetupError(Exception):
    """A problem with the config file, the template or the environment (exit code 2)."""


# --------------------------------------------------------------------------------------
# Run log and check report
# --------------------------------------------------------------------------------------
class Report:
    """Collects what was checked, problems (which stop the run) and notes (which do not)."""

    def __init__(self) -> None:
        self.checks: dict[str, list[str]] = {}
        self.notes: list[str] = []

    def check(self, title: str) -> list[str]:
        return self.checks.setdefault(title, [])

    def note(self, text: str) -> None:
        if text not in self.notes:
            self.notes.append(text)

    @property
    def problems(self) -> list[str]:
        return [p for items in self.checks.values() for p in items]


class RunLog:
    def __init__(self) -> None:
        self.lines: list[str] = []

    def __call__(self, text: str = "") -> None:
        print(text)
        self.lines.append(text)

    def write(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n".join(self.lines) + "\n", encoding="utf-8")


def sha256(path_or_bytes) -> str:
    data = path_or_bytes if isinstance(path_or_bytes, bytes) else Path(path_or_bytes).read_bytes()
    return hashlib.sha256(data).hexdigest()


def package_version(name: str) -> str:
    try:
        return importlib_metadata.version(name)
    except importlib_metadata.PackageNotFoundError:
        return "unknown"


# --------------------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------------------
class Config:
    def __init__(self, path: Path) -> None:
        self.path = path.resolve()
        self.base = self.path.parent
        try:
            with open(self.path, "rb") as fh:
                self.raw = tomllib.load(fh)
        except FileNotFoundError:
            raise SetupError(f"Config file not found: {path}")
        except tomllib.TOMLDecodeError as exc:
            raise SetupError(f"Config file {path.name} is not valid TOML: {exc}")
        try:
            files = self.raw["files"]
            self.template = self.base / files["template"]
            self.metadata = self.base / files["metadata"]
            self.output_pattern = files["output"]
            ss = self.raw["study_sheet"]
            self.study_sheet = ss["sheet"]
            self.study_field_col = ss["field_column"]
            self.study_value_col = ss["value_column"]
        except KeyError as exc:
            raise SetupError(f"Config file {path.name} is missing the setting {exc}.")
        self.optional_fields = set(ss.get("optional_fields", []))
        self.label = self.raw.get("study", {}).get("label", self.path.stem)
        doc = self.raw.get("document", {})
        self.doc_title = doc.get("title")
        self.doc_description = doc.get("description")
        self.not_applicable = doc.get("not_applicable", "Not applicable")
        self.date_format = doc.get("date_format", "%-d %B %Y")
        self.remove_paragraphs = list(doc.get("remove_paragraphs_containing", []))
        self.placeholder_style = doc.get("placeholder_style", "PlaceholderText")
        meta = self.raw.get("metadata", {})
        self.trace_col = meta.get("trace_column", "SAP fields")
        self.ignore_cols = set(meta.get("ignore_columns", []))
        self.sheets: dict[str, dict] = self.raw.get("sheets", {})
        for name, s in self.sheets.items():
            if "id_column" not in s:
                raise SetupError(f"Config: [sheets.\"{name}\"] needs an id_column.")
            if ("parent_sheet" in s) != ("parent_column" in s):
                raise SetupError(f"Config: [sheets.\"{name}\"] needs both parent_sheet and parent_column.")
            if s.get("parent_sheet") and s["parent_sheet"] not in self.sheets:
                raise SetupError(f"Config: parent_sheet \"{s['parent_sheet']}\" of \"{name}\" has no [sheets] section.")
        self.numbering: dict[str, str] = {k: v["sheet"] for k, v in self.raw.get("numbering", {}).items()}
        self.splits: dict[str, dict] = {s["group"]: s for s in self.raw.get("split", [])}
        self.checks = self.raw.get("checks", {})
        self.allowed_override: dict[str, list[str]] = self.raw.get("allowed", {})

    def parent_sheet(self, sheet: str) -> str | None:
        return self.sheets.get(sheet, {}).get("parent_sheet")


# --------------------------------------------------------------------------------------
# Metadata workbook
# --------------------------------------------------------------------------------------
@dataclass(eq=False)
class Row:
    sheet: str
    excel_row: int
    id: str
    values: dict[str, str]
    parent: "Row | None" = None
    level: str | None = None
    number: int | None = None

    def where(self) -> str:
        return f"'{self.sheet}' sheet, row {self.excel_row}" + (f" ({self.id})" if self.id else "")


@dataclass
class Sheet:
    name: str
    headers: dict[str, int]
    rows: list[Row] = field(default_factory=list)
    allowed: dict[str, list[str]] = field(default_factory=dict)
    trace: dict[int, object] = field(default_factory=dict)
    ws: object = None


@dataclass
class Book:
    sheets: dict[str, Sheet]
    study: dict[str, Row]          # Study field name -> row (value in values["Value"])
    levels: list[str]


def format_date(value: dt.date, fmt: str) -> str:
    return value.strftime(fmt.replace("%-d", str(value.day)))


def cell_text(value, cfg: Config) -> str:
    if value is None:
        return ""
    if isinstance(value, (dt.datetime, dt.date)):
        return format_date(value, cfg.date_format)
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).replace("\r\n", "\n").replace("\r", "\n").replace("_x000D_", "")
    return text.strip()


def read_book(cfg: Config, rep: Report) -> Book:
    try:
        wb_formulas = openpyxl.load_workbook(cfg.metadata, data_only=False)
        wb_values = openpyxl.load_workbook(cfg.metadata, data_only=True)
    except FileNotFoundError:
        raise SetupError(f"Metadata workbook not found: {cfg.metadata}")
    except Exception as exc:  # noqa: BLE001 - corrupt or locked workbook
        raise SetupError(f"Could not open the metadata workbook {cfg.metadata.name}: {exc}")

    p_sheets = rep.check("Every sheet and column the config and template need is in the workbook")
    p_formula = rep.check("Data cells hold values (not uncalculated formulas)")
    p_chars = rep.check("No hidden control characters in the text")
    needed = [cfg.study_sheet, *cfg.sheets]
    sheets: dict[str, Sheet] = {}
    for name in needed:
        if name not in wb_formulas.sheetnames:
            p_sheets.append(f"The workbook has no '{name}' sheet.")
            continue
        ws_f, ws_v = wb_formulas[name], wb_values[name]
        headers: dict[str, int] = {}
        for idx, cell in enumerate(ws_f[1], start=1):
            head = cell_text(cell.value, cfg)
            if not head:
                continue
            if head in headers:
                p_sheets.append(f"'{name}' sheet: the column heading '{head}' appears twice.")
            headers[head] = idx
        sheet = Sheet(name, headers, ws=ws_f)
        # Drop-down lists (data validation) give the allowed values for a column.
        for dv in ws_f.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1 and dv.formula1.startswith('"'):
                values = [v.strip() for v in dv.formula1.strip('"').split(",") if v.strip()]
                for rng in dv.sqref.ranges:
                    for col in range(rng.min_col, rng.max_col + 1):
                        for head, idx in headers.items():
                            if idx == col:
                                sheet.allowed[head] = values
        for key, values in cfg.allowed_override.items():
            s_name, _, col = key.partition(".")
            if s_name == name:
                sheet.allowed[col] = list(values)
        trace_idx = headers.get(cfg.trace_col)
        for r in range(2, ws_f.max_row + 1):
            values: dict[str, str] = {}
            for head, idx in headers.items():
                if idx == trace_idx:
                    continue
                raw_f = ws_f.cell(row=r, column=idx).value
                raw_v = ws_v.cell(row=r, column=idx).value
                if isinstance(raw_f, str) and raw_f.startswith("="):
                    if raw_v is None:
                        p_formula.append(
                            f"'{name}' sheet, row {r}, column '{head}' is a formula that Excel has not "
                            "calculated yet. Open the workbook in Excel, save it, and run again.")
                    raw = raw_v
                else:
                    raw = raw_f
                text = cell_text(raw, cfg)
                if BAD_CHARS.search(text):
                    p_chars.append(f"'{name}' sheet, row {r}, column '{head}' contains a hidden control "
                                   "character (often from copy-and-paste). Retype the cell.")
                    text = BAD_CHARS.sub("", text)
                values[head] = text
            if not any(values.values()):
                continue  # blank row (possibly with the SAP fields formula already copied down)
            if name == cfg.study_sheet:
                row_id = values.get(cfg.study_field_col, "")
            else:
                row_id = values.get(cfg.sheets[name]["id_column"], "")
            sheet.rows.append(Row(name, r, row_id, values))
            if trace_idx:
                sheet.trace[r] = ws_f.cell(row=r, column=trace_idx).value
        sheets[name] = sheet

    study: dict[str, Row] = {}
    if cfg.study_sheet in sheets:
        st = sheets[cfg.study_sheet]
        for col in (cfg.study_field_col, cfg.study_value_col):
            if col not in st.headers:
                p_sheets.append(f"'{cfg.study_sheet}' sheet has no '{col}' column.")
        p_dupe = rep.check("Study fields are named once each")
        for row in st.rows:
            if not row.id:
                p_dupe.append(f"{row.where()}: the Field cell is empty.")
            elif row.id in study:
                p_dupe.append(f"{row.where()}: the field '{row.id}' is listed twice "
                              f"(also row {study[row.id].excel_row}).")
            else:
                study[row.id] = row

    for name, s_cfg in cfg.sheets.items():
        if name not in sheets:
            continue
        for key in ("id_column", "parent_column", "level_column"):
            col = s_cfg.get(key)
            if col and col not in sheets[name].headers:
                p_sheets.append(f"'{name}' sheet has no '{col}' column (named as {key} in the config).")

    level_col = next((s.get("level_column") for s in cfg.sheets.values() if s.get("level_column")), None)
    level_sheet = next((n for n, s in cfg.sheets.items() if s.get("level_column")), None)
    levels = []
    if level_sheet in sheets and level_col:
        levels = sheets[level_sheet].allowed.get(level_col, [])
    return Book(sheets, study, levels)


def link_rows(book: Book, cfg: Config, rep: Report) -> None:
    """Check IDs, link child rows to parents, inherit levels and number rows within a level."""
    p_ids = rep.check("Every row has an ID, and IDs are unique on each sheet")
    p_links = rep.check("Every link (Objective ID, Estimand ID) points to an existing row")
    index: dict[str, dict[str, Row]] = {}
    for name in cfg.sheets:
        sheet = book.sheets.get(name)
        if not sheet:
            continue
        id_col = cfg.sheets[name]["id_column"]
        index[name] = {}
        for row in sheet.rows:
            if not row.id:
                p_ids.append(f"{row.where()}: the '{id_col}' cell is empty, but the row has other entries.")
            elif row.id in index[name]:
                p_ids.append(f"{row.where()}: the ID '{row.id}' is used twice "
                             f"(also row {index[name][row.id].excel_row}).")
            else:
                index[name][row.id] = row
            if row.id and not re.fullmatch(r"[A-Za-z0-9_.\-]+", row.id):
                p_ids.append(f"{row.where()}: the ID '{row.id}' contains spaces or symbols. "
                             "Use letters, digits, '-' or '_' only.")
    for name, s_cfg in cfg.sheets.items():
        parent_name = s_cfg.get("parent_sheet")
        if not parent_name or name not in book.sheets or parent_name not in index:
            continue
        for row in book.sheets[name].rows:
            ref = row.values.get(s_cfg["parent_column"], "")
            if not ref:
                p_links.append(f"{row.where()}: the '{s_cfg['parent_column']}' cell is empty, so it is not "
                               f"linked to any row on the '{parent_name}' sheet.")
            elif ref not in index[parent_name]:
                p_links.append(f"{row.where()}: '{s_cfg['parent_column']}' is '{ref}', which is not an ID on "
                               f"the '{parent_name}' sheet.")
            else:
                row.parent = index[parent_name][ref]

    def level_of(row: Row, depth: int = 0) -> str | None:
        lc = cfg.sheets[row.sheet].get("level_column")
        if lc:
            return row.values.get(lc) or None
        if row.parent is not None and depth < 20:
            return level_of(row.parent, depth + 1)
        return None

    for name in cfg.sheets:
        if name in book.sheets:
            for row in book.sheets[name].rows:
                row.level = level_of(row)

    for num_name, sheet_name in cfg.numbering.items():
        if sheet_name not in book.sheets:
            raise SetupError(f"Config: [numbering.\"{num_name}\"] uses sheet '{sheet_name}', not in the workbook.")
        counts: Counter = Counter()
        for row in book.sheets[sheet_name].rows:
            counts[row.level] += 1
            row.number = counts[row.level]


# --------------------------------------------------------------------------------------
# Word XML helpers
# --------------------------------------------------------------------------------------
def top_sdts(el) -> list:
    """Content controls inside el that are not inside another content control within el."""
    found = []

    def walk(node):
        for child in node:
            if not isinstance(child.tag, str):
                continue
            if child.tag == w("sdt"):
                found.append(child)
            else:
                walk(child)

    walk(el)
    return found


def sdt_content(sdt):
    return sdt.find(w("sdtContent"))


def sdt_name(sdt) -> str:
    pr = sdt.find(w("sdtPr"))
    if pr is None:
        return ""
    for tag in ("tag", "alias"):
        node = pr.find(w(tag))
        if node is not None and node.get(w("val")):
            return node.get(w("val")).strip()
    return ""


def set_sdt_name(sdt, name: str) -> None:
    pr = sdt.find(w("sdtPr"))
    for tag in ("alias", "tag"):
        node = pr.find(w(tag))
        if node is None:
            node = etree.SubElement(pr, w(tag))
            pr.insert(0 if tag == "alias" else 1, node)
        node.set(w("val"), name)


def split_name(name: str) -> tuple[str, str | None]:
    m = NAME_RE.match(name)
    return m.group("base").strip(), (m.group("qual").strip() if m.group("qual") else None)


def run_text(el) -> str:
    out = []
    for node in el.iter(w("t"), w("br"), w("tab"), w("cr")):
        if node.tag == w("t"):
            out.append(node.text or "")
        elif node.tag == w("tab"):
            out.append("\t")
        else:
            out.append("\n")
    return "".join(out)


def text_of(el) -> str:
    paras = [el] if el.tag == w("p") else list(el.iter(w("p")))
    if not paras:
        return run_text(el)
    return "\n".join(run_text(p) for p in paras)


def clean_rpr(rpr, placeholder_style: str):
    if rpr is None:
        return None
    rpr = copy.deepcopy(rpr)
    for rs in rpr.findall(w("rStyle")):
        if rs.get(w("val")) == placeholder_style:
            rpr.remove(rs)
    return rpr if len(rpr) else None


def make_run(text: str, rpr):
    r = etree.Element(w("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    for i, line in enumerate(text.split("\n")):
        if i:
            etree.SubElement(r, w("br"))
        t = etree.SubElement(r, w("t"))
        t.text = line
        t.set(XML_SPACE, "preserve")
    return r


def make_paragraph(text: str, ppr=None, rpr=None):
    p = etree.Element(w("p"))
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    p.append(make_run(text, rpr))
    return p


def fill_sdt(sdt, value: str, placeholder_style: str) -> None:
    """Replace the content of a field control with value, keeping its formatting."""
    pr = sdt.find(w("sdtPr"))
    for ph in pr.findall(w("showingPlcHdr")):
        pr.remove(ph)
    content = sdt_content(sdt)
    holder = content
    tc = content.find(w("tc"))
    if tc is not None and content.find(w("p")) is None:
        holder = tc  # cell-level control: fill the paragraphs inside the cell
    paras = holder.findall(w("p"))
    if paras:
        model = paras[0]
        ppr = model.find(w("pPr"))
        first_r = model.find(f".//{w('r')}")
        rpr = clean_rpr(first_r.find(w("rPr")) if first_r is not None else None, placeholder_style)
        pos = list(holder).index(paras[0])
        for p in paras:
            holder.remove(p)
        for i, line in enumerate(value.split("\n")):
            holder.insert(pos + i, make_paragraph(line, ppr, rpr))
        return
    runs = holder.findall(w("r"))
    if runs:
        rpr = clean_rpr(runs[0].find(w("rPr")), placeholder_style)
        if rpr is None:
            # The placeholder carries no formatting of its own (only the grey placeholder
            # style), so match the text just before the control in the same paragraph,
            # e.g. the small grey footer text around "Protocol number".
            for sib in sdt.itersiblings(preceding=True):
                if sib.tag == w("r") and sib.find(w("rPr")) is not None:
                    rpr = clean_rpr(sib.find(w("rPr")), placeholder_style)
                    break
        pos = list(holder).index(runs[0])
        for r in runs:
            holder.remove(r)
        holder.insert(pos, make_run(value, rpr))
        return
    raise SetupError(f"Template control '{sdt_name(sdt)}' has no paragraph or text run to fill.")


BOOKMARK_TAGS = (w("bookmarkStart"), w("bookmarkEnd"))


def take_bookmarks(el) -> list:
    """Remove and return the bookmarks inside el (so copies do not duplicate them)."""
    marks = [b for b in el.iter(*BOOKMARK_TAGS)]
    for b in marks:
        b.getparent().remove(b)
    return marks


def put_bookmarks(el, marks: list) -> None:
    """Put bookmarks back around the first paragraph of el."""
    if not marks:
        return
    p = el if el.tag == w("p") else next(el.iter(w("p")), None)
    if p is None:
        return
    starts = [m for m in marks if m.tag == w("bookmarkStart")]
    ends = [m for m in marks if m.tag == w("bookmarkEnd")]
    ppr = p.find(w("pPr"))
    at = 1 if ppr is not None else 0
    for i, m in enumerate(starts):
        p.insert(at + i, m)
    for m in ends:
        p.append(m)


# --------------------------------------------------------------------------------------
# Template analysis
# --------------------------------------------------------------------------------------
@dataclass
class TemplateInfo:
    parts: dict[str, object]                     # part name -> XML root
    raw: dict[str, bytes]                        # every file in the .docx
    infos: list                                  # zip entries, in order
    used_columns: dict[str, set[str]]            # sheet -> columns the template fills
    used_study: set[str]                         # Study fields the template fills
    stamps: dict[str, set[str]]                  # group base name -> qualifiers present
    group_sheets: dict[str, str]                 # group base name -> sheet
    bookmarks: list[str]
    control_count: int


class Classifier:
    def __init__(self, cfg: Config, book: Book) -> None:
        self.cfg = cfg
        self.book = book
        self.sheet_names = sorted(cfg.sheets, key=len, reverse=True)

    def kind(self, base: str) -> tuple[str, str | None, str | None]:
        """Return (kind, sheet, column). kind: study, field, number, group, unknown."""
        if base.startswith(self.cfg.study_sheet + "."):
            return "study", self.cfg.study_sheet, base[len(self.cfg.study_sheet) + 1:]
        for s in self.sheet_names:
            if base.startswith(s + "."):
                return "field", s, base[len(s) + 1:]
        if base in self.cfg.numbering:
            return "number", self.cfg.numbering[base], None
        if "." in base:
            return "unknown", None, None
        return "group", None, None

    def group_sheet(self, sdt) -> tuple[str | None, set[str], bool]:
        """Infer a group's sheet from the fields directly inside it.
        Returns (sheet or None, set of sheets seen, has nested groups)."""
        sheets: set[str] = set()
        nested = False
        for child in top_sdts(sdt_content(sdt)):
            base, _ = split_name(sdt_name(child))
            kind, sheet, _ = self.kind(base)
            if kind in ("field", "number"):
                sheets.add(sheet)
            elif kind == "group":
                nested = True
        return (next(iter(sheets)) if len(sheets) == 1 else None), sheets, nested

    def nested_sheet(self, sdt) -> str | None:
        """For a level container: the sheet of the first group inside it."""
        for child in top_sdts(sdt_content(sdt)):
            base, _ = split_name(sdt_name(child))
            if self.kind(base)[0] == "group":
                sheet, _, nested = self.group_sheet(child)
                return sheet if sheet else self.nested_sheet(child)
        return None


def read_template(cfg: Config) -> tuple[dict, dict, list]:
    try:
        with zipfile.ZipFile(cfg.template) as z:
            infos = z.infolist()
            raw = {i.filename: z.read(i) for i in infos}
    except FileNotFoundError:
        raise SetupError(f"Template not found: {cfg.template}")
    except zipfile.BadZipFile:
        raise SetupError(f"The template {cfg.template.name} is not a valid .docx file.")
    names = [n for n in raw if n == "word/document.xml"
             or re.fullmatch(r"word/(header|footer)\d*\.xml|word/(footnotes|endnotes)\.xml", n)]
    if "word/document.xml" not in names:
        raise SetupError(f"The template {cfg.template.name} has no main document part.")
    parts = {}
    for n in names:
        try:
            parts[n] = etree.fromstring(raw[n], XML_PARSER)
        except etree.XMLSyntaxError as exc:
            raise SetupError(f"The template part {n} could not be read: {exc}")
    return parts, raw, infos


def analyse_template(cfg: Config, book: Book, rep: Report) -> TemplateInfo:
    parts, raw, infos = read_template(cfg)
    cls = Classifier(cfg, book)
    p_tpl = rep.check("Every template control name is understood (Sheet.Column, group or number)")
    p_cols = rep.check("Every Sheet.Column the template uses exists in the workbook")
    used_columns: dict[str, set[str]] = {s: set() for s in cfg.sheets}
    used_study: set[str] = set()
    stamps: dict[str, set[str]] = {}
    group_sheets: dict[str, str] = {}
    bookmarks: list[str] = []
    count = 0
    levels = set(book.levels)

    def visit(el, part: str, ctx_sheets: tuple[str, ...]) -> None:
        nonlocal count
        for sdt in top_sdts(el):
            count += 1
            name = sdt_name(sdt)
            pr = sdt.find(w("sdtPr"))
            alias = pr.find(w("alias")) if pr is not None else None
            tag = pr.find(w("tag")) if pr is not None else None
            if alias is not None and tag is not None and alias.get(w("val")) != tag.get(w("val")):
                rep.note(f"Template control titled '{alias.get(w('val'))}' has a different tag "
                         f"'{tag.get(w('val'))}'; the tag is used.")
            if not name:
                p_tpl.append(f"A content control in {part} has no title or tag, so the script cannot tell "
                             "what goes in it. Give it a name in Word (Developer > Properties).")
                continue
            base, qual = split_name(name)
            kind, sheet, col = cls.kind(base)
            if kind == "unknown":
                p_tpl.append(f"Template control '{name}': '{base.split('.')[0]}' is not a sheet the config "
                             "knows about.")
                continue
            if qual and qual.lower() != "each" and qual not in levels and kind != "study":
                p_tpl.append(f"Template control '{name}': '({qual})' is neither '(each)' nor a level "
                             f"({', '.join(book.levels) or 'none found'}).")
            if kind == "study":
                if col not in book.study:
                    p_cols.append(f"Template control '{name}': the Study sheet has no field called '{col}'.")
                used_study.add(col)
                continue
            if kind in ("field", "number"):
                if kind == "field":
                    sh = book.sheets.get(sheet)
                    if sh is not None and col not in sh.headers:
                        p_cols.append(f"Template control '{name}': the '{sheet}' sheet has no column '{col}'.")
                    used_columns[sheet].add(col)
                if sheet not in ctx_sheets and part == "word/document.xml":
                    p_tpl.append(f"Template control '{name}' sits outside any group for the '{sheet}' sheet, "
                                 "so the script cannot tell which row to use.")
                elif part != "word/document.xml":
                    p_tpl.append(f"Template control '{name}' is in a header or footer; only Study fields can "
                                 "go there.")
                continue
            # group or level container
            g_sheet, seen, nested = cls.group_sheet(sdt)
            if len(seen) > 1:
                p_tpl.append(f"Template group '{name}' holds fields from several sheets "
                             f"({', '.join(sorted(seen))}); a group must hold fields from one sheet.")
                continue
            if g_sheet is None and not nested:
                p_tpl.append(f"Template group '{name}' holds no fields.")
                continue
            if g_sheet is None:  # level container
                inner = cls.nested_sheet(sdt)
                if inner is None:
                    p_tpl.append(f"Template group '{name}' holds no fields.")
                visit(sdt_content(sdt), part, ctx_sheets)
                continue
            stamps.setdefault(base, set()).add(qual or "")
            if base in group_sheets and group_sheets[base] != g_sheet:
                p_tpl.append(f"Template group '{base}' is used for two different sheets.")
            group_sheets[base] = g_sheet
            parent = cfg.parent_sheet(g_sheet)
            outer = [s for s in ctx_sheets if s in cfg.sheets]
            if outer and outer[-1] != g_sheet and parent != outer[-1]:
                p_tpl.append(f"Template group '{name}' ({g_sheet}) sits inside a group for '{outer[-1]}', "
                             f"but the config does not link '{g_sheet}' rows to '{outer[-1]}' rows.")
            visit(sdt_content(sdt), part, ctx_sheets + (g_sheet,))

    for part, root in parts.items():
        visit(root, part, ())
        bookmarks += [b.get(w("name")) for b in root.iter(w("bookmarkStart"))]

    # Splits: the anchor paragraph must exist after each stamp of the group.
    for gname, split in cfg.splits.items():
        if gname not in stamps:
            p_tpl.append(f"Config [[split]] names the group '{gname}', which is not in the template.")
            continue
        for root in parts.values():
            for sdt in root.iter(w("sdt")):
                if split_name(sdt_name(sdt))[0] == gname and find_anchor(sdt, split) is None:
                    p_tpl.append(f"Template group '{sdt_name(sdt)}': no paragraph starting "
                                 f"'{split['after_paragraph_starting']}' follows it, so supplementary rows "
                                 "have nowhere to go.")
    return TemplateInfo(parts, raw, infos, used_columns, used_study, stamps, group_sheets, bookmarks, count)


def find_anchor(sdt, split: dict):
    start = split["after_paragraph_starting"]
    for sib in sdt.itersiblings():
        if sib.tag == w("p") and run_text(sib).strip().startswith(start):
            return sib
    return None


# --------------------------------------------------------------------------------------
# Metadata checks (before building)
# --------------------------------------------------------------------------------------
TRACE_TOKEN = re.compile(r'\s*(?:"((?:[^"]|"")*)"|CHAR\((\d+)\)|\$?([A-Z]{1,3})\$?(\d+)|(&))\s*', re.I)


def eval_trace(value, ws) -> tuple[str | None, set[int]]:
    """Evaluate the simple string formula in the SAP fields column (literals, cells, CHAR, &)."""
    if not isinstance(value, str):
        return None, set()
    if not value.startswith("="):
        return value, set()
    s, pos, out, rows = value[1:], 0, [], set()
    while pos < len(s):
        m = TRACE_TOKEN.match(s, pos)
        if not m or m.end() == pos:
            return None, rows
        pos = m.end()
        lit, ch, col, rw, amp = m.groups()
        if amp:
            continue
        if lit is not None:
            out.append(lit.replace('""', '"'))
        elif ch:
            out.append(chr(int(ch)))
        else:
            rows.add(int(rw))
            v = ws[f"{col.upper()}{rw}"].value
            out.append("" if v is None else str(v).strip())
    return "".join(out), rows


def check_metadata(cfg: Config, book: Book, tpl: TemplateInfo, rep: Report) -> None:
    levels = book.levels or sorted({q for qs in tpl.stamps.values() for q in qs if q and q != "each"})

    # Required values: every column the template fills, plus IDs and links.
    p_blank = rep.check("No blank cells in columns the SAP needs")
    for name, s_cfg in cfg.sheets.items():
        sheet = book.sheets.get(name)
        if not sheet:
            continue
        needed = set(tpl.used_columns.get(name, set()))
        needed |= {c for c in (s_cfg.get("level_column"),) if c}
        for row in sheet.rows:
            for col in sorted(needed):
                if col in sheet.headers and not row.values.get(col):
                    p_blank.append(f"{row.where()}: '{col}' is blank, but the SAP needs it.")
    for fld in sorted(tpl.used_study):
        row = book.study.get(fld)
        if row is None:
            continue
        if not row.values.get(cfg.study_value_col):
            if fld in cfg.optional_fields:
                rep.note(f"Study field '{fld}' is blank. This is allowed by the config, so the SAP keeps the "
                         "template's grey prompt there for you to fill in later.")
            else:
                p_blank.append(f"{row.where()}: the value for '{fld}' is blank, but the SAP needs it. "
                               "(If it can wait, add it to optional_fields in the config.)")

    # Allowed values from the drop-down lists.
    p_allowed = rep.check("Drop-down columns hold only allowed values (Level, Role, Strategy)")
    for name in cfg.sheets:
        sheet = book.sheets.get(name)
        if not sheet:
            continue
        for col, allowed in sheet.allowed.items():
            for row in sheet.rows:
                v = row.values.get(col, "")
                if v and v not in allowed:
                    close = [a for a in allowed if a.lower() == v.lower()]
                    hint = f" Did you mean '{close[0]}'?" if close else ""
                    p_allowed.append(f"{row.where()}: '{col}' is '{v}', which is not one of the allowed "
                                     f"values ({', '.join(allowed)}).{hint}")

    # Every level used in the metadata needs a stamp in the template.
    p_stamp = rep.check("The template has a stamp for every level used in the metadata")
    for base, quals in tpl.stamps.items():
        if "each" in {q.lower() for q in quals} or "" in quals:
            continue
        sheet = book.sheets.get(tpl.group_sheets[base])
        if not sheet:
            continue
        for lvl in sorted({r.level for r in sheet.rows if r.level}):
            if book.levels and lvl not in book.levels:
                continue  # already reported as a value that is not allowed
            if lvl not in quals:
                p_stamp.append(f"The metadata has {lvl} rows on the '{sheet.name}' sheet, but the template "
                               f"has no '{base} ({lvl})' control to print them with.")

    # Role must match the objective's level.
    ck = cfg.checks
    role_sheet, role_col = ck.get("role_sheet"), ck.get("role_column")
    any_level = set(ck.get("roles_any_level", []))
    if role_sheet in book.sheets and role_col:
        p_role = rep.check("Each estimand's Role matches its objective's Level (Supplementary may go anywhere)")
        for row in book.sheets[role_sheet].rows:
            role = row.values.get(role_col, "")
            if role and row.level and role not in any_level and role != row.level:
                parent = row.parent.id if row.parent else "its objective"
                p_role.append(f"{row.where()}: Role is '{role}', but {parent} is a {row.level} objective. "
                              f"Use '{row.level}' or one of: {', '.join(sorted(any_level))}.")

    # SAP version agrees with the Version history.
    v_sheet = ck.get("version_sheet")
    if v_sheet in book.sheets and ck.get("study_version_field") in book.study:
        p_ver = rep.check("The SAP version on the Study sheet is in the Version history")
        version = book.study[ck["study_version_field"]].values.get(cfg.study_value_col, "")
        match = [r for r in book.sheets[v_sheet].rows if r.values.get(ck["version_column"]) == version]
        if not match:
            p_ver.append(f"The Study sheet says SAP version '{version}', but the '{v_sheet}' sheet has no row "
                         "for that version.")
        else:
            if match[-1] is not book.sheets[v_sheet].rows[-1]:
                rep.note(f"Version '{version}' is not the last row of the '{v_sheet}' sheet.")
            date_field = ck.get("study_date_field")
            if date_field in book.study:
                sdate = book.study[date_field].values.get(cfg.study_value_col, "")
                vdate = match[-1].values.get(ck.get("version_date_column", ""), "")
                if sdate and vdate and sdate != vdate:
                    p_ver.append(f"The Study sheet dates version {version} '{sdate}', but the '{v_sheet}' "
                                 f"sheet dates it '{vdate}'.")

    # Coverage notes (do not stop the run).
    obj_sheet = next((n for n, s in cfg.sheets.items() if s.get("level_column")), None)
    est_sheet = next((n for n, s in cfg.sheets.items() if s.get("parent_sheet") == obj_sheet), None)
    ice_sheet = next((n for n, s in cfg.sheets.items() if s.get("parent_sheet") == est_sheet), None)
    if obj_sheet in book.sheets:
        objectives = book.sheets[obj_sheet].rows
        for lvl in levels:
            if not any(r.level == lvl for r in objectives):
                rep.note(f"There are no {lvl} objectives, so that level shows '{cfg.not_applicable}'.")
        if est_sheet in book.sheets:
            ests = book.sheets[est_sheet].rows
            for obj in objectives:
                mine = [e for e in ests if e.parent is obj]
                main = [e for e in mine if e.values.get(role_col, "") not in any_level]
                if not mine:
                    rep.note(f"{obj.id} ({obj.level}) has no estimands; its estimand section shows "
                             f"'{cfg.not_applicable}'.")
                elif not main:
                    rep.note(f"{obj.id} has only supplementary estimands; its main estimand section shows "
                             f"'{cfg.not_applicable}'.")
                end_col = ck.get("endpoint_column")
                for e in mine:
                    if end_col and e.values.get(end_col) and obj.values.get(end_col) \
                            and e.values[end_col] != obj.values[end_col]:
                        rep.note(f"{e.id}: its endpoint differs from {obj.id}'s endpoint "
                                 f"('{e.values[end_col]}' vs '{obj.values[end_col]}').")
            if ice_sheet in book.sheets:
                for e in ests:
                    if not any(i.parent is e for i in book.sheets[ice_sheet].rows):
                        rep.note(f"{e.id} has no intercurrent events; its table shows '{cfg.not_applicable}'.")

    # SAP fields (trace) column.
    p_trace = rep.check("The 'SAP fields' column is filled in on every row and points at its own row")
    for name in [cfg.study_sheet, *cfg.sheets]:
        sheet = book.sheets.get(name)
        if not sheet or cfg.trace_col not in sheet.headers:
            continue
        for row in sheet.rows:
            raw = sheet.trace.get(row.excel_row)
            if raw in (None, ""):
                p_trace.append(f"{row.where()}: the '{cfg.trace_col}' cell is empty. Copy the formula down "
                               "from the row above.")
                continue
            text, refs = eval_trace(raw, sheet.ws)
            if text is None:
                rep.note(f"{row.where()}: could not read the '{cfg.trace_col}' formula, so it was not checked.")
                continue
            other = sorted(refs - {row.excel_row})
            if other:
                p_trace.append(f"{row.where()}: the '{cfg.trace_col}' formula refers to row "
                               f"{', '.join(map(str, other))} instead of its own row, so it names the wrong "
                               "fields. Copy the formula down again.")
                continue
            listed = {n.strip() for n in text.split("\n") if n.strip()}
            if name == cfg.study_sheet:
                expected = {f"{name}.{row.id}"} if row.id in tpl.used_study else set()
            else:
                expected = {f"{name}.{c} [{row.id}]" for c in tpl.used_columns.get(name, set())}
            missing = sorted(expected - listed)
            if missing:
                rep.note(f"{row.where()}: the '{cfg.trace_col}' cell does not list {', '.join(missing)}.")

    # Data the template does not use (a note, so nothing is silently lost).
    for name in cfg.sheets:
        sheet = book.sheets.get(name)
        if not sheet:
            continue
        s_cfg = cfg.sheets[name]
        structural = {s_cfg["id_column"], s_cfg.get("parent_column"), s_cfg.get("level_column"),
                      cfg.trace_col, ck.get("role_column") if name == role_sheet else None}
        unused = [c for c in sheet.headers
                  if c not in tpl.used_columns.get(name, set()) and c not in structural and c not in cfg.ignore_cols]
        if unused:
            rep.note(f"'{name}' sheet: column(s) {', '.join(repr(c) for c in unused)} are not used by this "
                     "template (kept in the metadata for later SAP sections).")
    unused_study = [f for f in book.study if f not in tpl.used_study]
    if unused_study:
        rep.note(f"Study field(s) {', '.join(repr(f) for f in unused_study)} are not used by this template.")


def output_path(cfg: Config, book: Book, rep: Report) -> Path | None:
    p_name = rep.check("The output file name can be made from the Study sheet")

    def value(m):
        fld = m.group(1)
        if fld not in book.study or not book.study[fld].values.get(cfg.study_value_col):
            p_name.append(f"The output name uses the Study field '{fld}', which is missing or blank.")
            return fld
        return book.study[fld].values[cfg.study_value_col]

    rel = re.sub(r"\{([^{}]+)\}", value, cfg.output_pattern)
    rel_path = Path(rel)
    clean_name = re.sub(r"\s+", "", rel_path.name)
    if re.search(r'[<>:"/\\|?*]', clean_name):
        p_name.append(f"The output file name '{clean_name}' contains characters Windows does not allow.")
        return None
    return (cfg.base / rel_path.parent / clean_name).resolve()


def fill_pattern(pattern: str | None, cfg: Config, book: Book) -> str | None:
    if not pattern:
        return None
    return re.sub(r"\{([^{}]+)\}",
                  lambda m: book.study[m.group(1)].values.get(cfg.study_value_col, "") if m.group(1) in book.study
                  else m.group(0), pattern)


# --------------------------------------------------------------------------------------
# Build
# --------------------------------------------------------------------------------------
class Builder:
    def __init__(self, cfg: Config, book: Book, tpl: TemplateInfo) -> None:
        self.cfg = cfg
        self.book = book
        self.tpl = tpl
        self.cls = Classifier(cfg, book)
        self.copies: dict[str, list[str]] = {}
        self.not_applicable: list[str] = []
        self.placeholders_left: list[str] = []
        self.filled = 0

    # -- traversal ---------------------------------------------------------------------
    def process(self, el, ctx: dict[str, Row]) -> None:
        for sdt in top_sdts(el):
            self.handle(sdt, ctx)

    def handle(self, sdt, ctx: dict[str, Row]) -> None:
        name = sdt_name(sdt)
        base, qual = split_name(name)
        kind, sheet, col = self.cls.kind(base)
        if kind == "study":
            value = self.book.study[col].values.get(self.cfg.study_value_col, "")
            if not value:
                self.placeholders_left.append(name)
                return
            self.fill(sdt, name, value)
        elif kind == "field":
            row = ctx.get(sheet)
            if row is None:
                raise SetupError(f"Template control '{name}' has no '{sheet}' row around it.")
            self.fill(sdt, f"{base} [{row.id}]", row.values.get(col, ""))
        elif kind == "number":
            row = ctx.get(sheet)
            if row is None:
                raise SetupError(f"Template control '{name}' has no '{sheet}' row around it.")
            self.fill(sdt, f"{base} [{row.id}]", str(row.number))
        else:
            g_sheet, _, _ = self.cls.group_sheet(sdt)
            if g_sheet is None:
                self.container(sdt, name, qual, ctx)
            else:
                self.group(sdt, base, qual, g_sheet, ctx)

    def fill(self, sdt, new_name: str, value: str) -> None:
        set_sdt_name(sdt, new_name)
        fill_sdt(sdt, value, self.cfg.placeholder_style)
        self.filled += 1

    # -- rows --------------------------------------------------------------------------
    def select(self, sheet: str, qual: str | None, ctx: dict[str, Row]) -> list[Row]:
        rows = self.book.sheets[sheet].rows
        parent = self.cfg.parent_sheet(sheet)
        if parent and parent in ctx:
            rows = [r for r in rows if r.parent is ctx[parent]]
        if qual and qual.lower() != "each":
            rows = [r for r in rows if r.level == qual]
        return rows

    def context_for(self, row: Row, ctx: dict[str, Row]) -> dict[str, Row]:
        new = dict(ctx)
        r = row
        while r is not None:
            new[r.sheet] = r
            r = r.parent
        return new

    # -- groups ------------------------------------------------------------------------
    def make_copy(self, stamp, base: str, row: Row, ctx: dict[str, Row]):
        c = copy.deepcopy(stamp)
        set_sdt_name(c, f"{base} [{row.id}]")
        self.process(sdt_content(c), self.context_for(row, ctx))
        self.copies.setdefault(base, []).append(row.id)
        return c

    def group(self, stamp, base: str, qual: str | None, sheet: str, ctx: dict[str, Row]) -> None:
        rows = self.select(sheet, qual, ctx)
        split = self.cfg.splits.get(base)
        moved: list[Row] = []
        anchor = None
        if split:
            moved = [r for r in rows if r.values.get(split["column"], "") in split["values"]]
            rows = [r for r in rows if r not in moved]
            anchor = find_anchor(stamp, split)
        marks = take_bookmarks(stamp)
        where = f"{sdt_name(stamp)}" + (f" in {', '.join(f'{r.id}' for r in ctx.values())}" if ctx else "")

        main_copies = [self.make_copy(stamp, base, r, ctx) for r in rows]
        moved_copies = [self.make_copy(stamp, base, r, ctx) for r in moved]
        if split and not moved:
            na = copy.deepcopy(stamp)
            self.make_na(na, [])
            moved_copies = [na]
            self.not_applicable.append(f"{where} ({split['after_paragraph_starting']}...)")
        if anchor is not None:
            prev = anchor
            for c in moved_copies:
                prev.addnext(c)
                prev = c
        if main_copies:
            for c in main_copies:
                stamp.addprevious(c)
            put_bookmarks(main_copies[0], marks)
            stamp.getparent().remove(stamp)
        else:
            self.make_na(stamp, marks)
            self.not_applicable.append(where)

    def container(self, sdt, name: str, qual: str | None, ctx: dict[str, Row]) -> None:
        inner = self.cls.nested_sheet(sdt)
        if inner is None or not self.select(inner, qual, ctx):
            self.make_na(sdt, take_bookmarks(sdt))
            self.not_applicable.append(name)
        else:
            self.process(sdt_content(sdt), ctx)

    def make_na(self, sdt, marks: list) -> None:
        """Replace a group's content with a single 'Not applicable' line of the same shape."""
        pr = sdt.find(w("sdtPr"))
        for ph in pr.findall(w("showingPlcHdr")):
            pr.remove(ph)
        content = sdt_content(sdt)
        first = next((c for c in content if isinstance(c.tag, str)), None)
        if first is not None and first.tag == w("tr"):
            cells = first.findall(w("tc"))
            span = sum(int(tc.find(f"{w('tcPr')}/{w('gridSpan')}").get(w("val")))
                       if tc.find(f"{w('tcPr')}/{w('gridSpan')}") is not None else 1 for tc in cells)
            width = sum(int(tc.find(f"{w('tcPr')}/{w('tcW')}").get(w("w"), 0))
                        if tc.find(f"{w('tcPr')}/{w('tcW')}") is not None else 0 for tc in cells)
            p_model = first.find(f".//{w('p')}")
            ppr = copy.deepcopy(p_model.find(w("pPr"))) if p_model is not None else None
            if ppr is not None:
                for kn in ppr.findall(w("keepNext")):
                    ppr.remove(kn)
            tr = etree.Element(w("tr"))
            if first.find(w("trPr")) is not None:
                tr.append(copy.deepcopy(first.find(w("trPr"))))
            tc = etree.SubElement(tr, w("tc"))
            tcpr = etree.SubElement(tc, w("tcPr"))
            tcw = etree.SubElement(tcpr, w("tcW"))
            tcw.set(w("type"), "dxa")
            tcw.set(w("w"), str(width))
            if span > 1:
                etree.SubElement(tcpr, w("gridSpan")).set(w("val"), str(span))
            para = make_paragraph(self.cfg.not_applicable, ppr)
            tc.append(para)
            new = tr
        else:
            para = make_paragraph(self.cfg.not_applicable)
            new = para
        for child in list(content):
            content.remove(child)
        content.append(new)
        put_bookmarks(para, marks)

    # -- whole document ----------------------------------------------------------------
    def build(self) -> dict[str, object]:
        out = {}
        for part, root in self.tpl.parts.items():
            root = copy.deepcopy(root)
            self.process(root, {})
            out[part] = root
        self.removed = remove_paragraphs(out, self.cfg.remove_paragraphs)
        renumber_sdt_ids(out)
        return out


def remove_paragraphs(parts: dict, needles: list[str]) -> int:
    removed = 0
    if not needles:
        return 0
    for root in parts.values():
        for p in list(root.iter(w("p"))):
            text = run_text(p)
            if any(n in text for n in needles) and next(p.iter(*BOOKMARK_TAGS), None) is None:
                p.getparent().remove(p)
                removed += 1
    return removed


def renumber_sdt_ids(parts: dict) -> None:
    """Give every content control a unique id (copies start with their stamp's id)."""
    ids = [i for root in parts.values() for i in root.iter(w("id")) if i.getparent().tag == w("sdtPr")]
    used: set[int] = set()
    next_id = max([int(i.get(w("val"))) for i in ids if i.get(w("val"), "").lstrip("-").isdigit()] + [0]) + 1
    for node in ids:
        val = node.get(w("val"), "")
        num = int(val) if val.lstrip("-").isdigit() else None
        if num is None or num in used:
            num = next_id
            next_id += 1
            node.set(w("val"), str(num))
        used.add(num)


def update_core_properties(raw: bytes, title: str | None, description: str | None) -> bytes:
    root = etree.fromstring(raw, XML_PARSER)
    for tag, value in (("title", title), ("description", description)):
        if value is None:
            continue
        node = root.find(f"dc:{tag}", CORE_NS)
        if node is None:
            node = etree.SubElement(root, f"{{{CORE_NS['dc']}}}{tag}")
        node.text = value
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def package(tpl: TemplateInfo, parts: dict, extra: dict[str, bytes]) -> bytes:
    """Write the .docx: every template file, in the same order, with the changed parts replaced."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for info in tpl.infos:
            if info.filename in parts:
                data = etree.tostring(parts[info.filename], xml_declaration=True, encoding="UTF-8",
                                      standalone=True)
            else:
                data = extra.get(info.filename, tpl.raw[info.filename])
            z.writestr(info, data)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Verification (after building, before saving)
# --------------------------------------------------------------------------------------
def verify(cfg: Config, book: Book, tpl: TemplateInfo, docx: bytes, builder: Builder, rep: Report) -> None:
    # Re-read the finished file from bytes, exactly as it will be saved.
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        p_zip = rep.check("[after] The finished file is a complete, readable Word package")
        bad = z.testzip()
        if bad:
            p_zip.append(f"The finished file is damaged at {bad}.")
        names = z.namelist()
        if names != [i.filename for i in tpl.infos]:
            p_zip.append("The finished file does not contain the same parts as the template.")
        parts = {}
        for n in names:
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    root = etree.fromstring(z.read(n), XML_PARSER)
                except etree.XMLSyntaxError as exc:
                    p_zip.append(f"{n} is not well-formed XML: {exc}")
                    continue
                if n in tpl.parts:
                    parts[n] = root

    controls: dict[str, list[str]] = {}
    placeholders: list[str] = []
    for root in parts.values():
        for sdt in root.iter(w("sdt")):
            name = sdt_name(sdt)
            controls.setdefault(name, []).append(text_of(sdt_content(sdt)))
            if sdt.find(f"{w('sdtPr')}/{w('showingPlcHdr')}") is not None:
                placeholders.append(name)

    # 1. Every metadata value is in the SAP, under its Sheet.Column [ID] name.
    p_val = rep.check("[after] Every metadata value used by the template appears in the SAP, word for word, "
                      "under its Sheet.Column [ID] name")
    checked = 0
    for sheet_name, cols in tpl.used_columns.items():
        sheet = book.sheets.get(sheet_name)
        if not sheet:
            continue
        for row in sheet.rows:
            for col in cols:
                name = f"{sheet_name}.{col} [{row.id}]"
                found = controls.get(name)
                if not found:
                    p_val.append(f"{name} is missing from the SAP.")
                    continue
                bad_texts = [t for t in found if t != row.values.get(col, "")]
                if bad_texts:
                    p_val.append(f"{name} reads '{bad_texts[0]}' instead of '{row.values.get(col, '')}'.")
                checked += 1
    for fld in tpl.used_study:
        expect = book.study[fld].values.get(cfg.study_value_col, "")
        found = controls.get(f"{cfg.study_sheet}.{fld}", [])
        if not found:
            p_val.append(f"{cfg.study_sheet}.{fld} is missing from the SAP.")
        elif expect and any(t != expect for t in found):
            p_val.append(f"{cfg.study_sheet}.{fld} does not read '{expect}' everywhere.")
        checked += len(found) if expect else 0
    rep.checked_values = checked

    # 2. Objective numbers run 1, 2, ... within each level.
    p_num = rep.check("[after] Numbers (e.g. Objective number) run 1, 2, 3 ... within each level")
    for num_name, sheet_name in cfg.numbering.items():
        for row in book.sheets[sheet_name].rows:
            found = controls.get(f"{num_name} [{row.id}]", [])
            if found and any(t != str(row.number) for t in found):
                p_num.append(f"{num_name} [{row.id}] should read {row.number}.")

    # 3. Nothing is left as a stamp, apart from 'Not applicable' and level containers.
    p_left = rep.check("[after] No unfilled stamps or placeholders are left (except optional blanks)")
    cls = Classifier(cfg, book)
    for name, texts in controls.items():
        base, qual = split_name(name)
        kind = cls.kind(base)[0]
        if qual and kind in ("field", "number"):
            p_left.append(f"The field '{name}' was not filled.")
    for name in placeholders:
        base, _ = split_name(name)
        kind, _, col = cls.kind(base)
        if not (kind == "study" and col in cfg.optional_fields):
            p_left.append(f"'{name}' still shows the template's placeholder.")

    # 4. Bookmarks kept: same names as the template, once each, every start has its end.
    p_bm = rep.check("[after] All template bookmarks are kept, once each, and correctly paired")
    starts = [b for root in parts.values() for b in root.iter(w("bookmarkStart"))]
    ends = {b.get(w("id")) for root in parts.values() for b in root.iter(w("bookmarkEnd"))}
    names_out = [b.get(w("name")) for b in starts]
    if sorted(names_out) != sorted(tpl.bookmarks):
        lost = sorted(set(tpl.bookmarks) - set(names_out))
        extra = sorted(n for n, c in Counter(names_out).items() if c > 1)
        p_bm.append(f"Bookmarks differ from the template (lost: {lost or 'none'}; repeated: {extra or 'none'}).")
    for b in starts:
        if b.get(w("id")) not in ends:
            p_bm.append(f"Bookmark '{b.get(w('name'))}' has no end marker.")
    rep.bookmarks_kept = len(names_out)

    # 5. Every content control has its own id.
    p_ids = rep.check("[after] Every content control has a unique id")
    ids = [i.get(w("val")) for root in parts.values() for i in root.iter(w("id"))
           if i.getparent().tag == w("sdtPr")]
    dupes = sorted({i for i in ids if ids.count(i) > 1})
    if dupes:
        p_ids.append(f"Content control ids used more than once: {', '.join(dupes)}.")

    # 6. The names in the 'SAP fields' column can be found in the SAP.
    p_tr = rep.check("[after] Every name in the 'SAP fields' column can be found in the SAP")
    for name in [cfg.study_sheet, *cfg.sheets]:
        sheet = book.sheets.get(name)
        if not sheet:
            continue
        for row in sheet.rows:
            text, _ = eval_trace(sheet.trace.get(row.excel_row), sheet.ws)
            for n in (text or "").split("\n"):
                n = n.strip()
                if not n or n in controls:
                    continue
                base = n.rsplit(" [", 1)[0]
                col = base.split(".", 1)[1] if "." in base else ""
                used = (col in tpl.used_study) if name == cfg.study_sheet else (col in tpl.used_columns.get(name, ()))
                if used:
                    p_tr.append(f"'{n}' (listed on the '{name}' sheet, row {row.excel_row}) is not in the SAP.")


# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
def print_report(log: RunLog, rep: Report, title: str, after: bool) -> None:
    log(title)
    for name, problems in rep.checks.items():
        if name.startswith("[after]") != after:
            continue
        label = name.replace("[after] ", "")
        log(f"  [{'PASS' if not problems else 'FAIL'}] {label}")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Build a draft SAP from a Word template and study metadata.")
    ap.add_argument("--config", default="pioneer1.toml", help="study config file (TOML)")
    ap.add_argument("--check-only", action="store_true", help="check the metadata and template; build nothing")
    args = ap.parse_args(argv)

    log = RunLog()
    rep = Report()
    started = dt.datetime.now()
    log_path: Path | None = None
    try:
        cfg = Config(Path(args.config))
        log_path = cfg.base / "sap" / "esap_run_log.txt"
        log("eSAP run log")
        log("============")
        log(f"Study:   {cfg.label}")
        log(f"Run at:  {started:%Y-%m-%d %H:%M:%S}")
        log(f"Tools:   esap.py {ESAP_VERSION}; Python {platform.python_version()}; "
            f"openpyxl {package_version('openpyxl')}; lxml {package_version('lxml')}")
        log(f"Config:  {cfg.path.name}")
        for label, p in (("Template", cfg.template), ("Metadata", cfg.metadata)):
            if not p.exists():
                raise SetupError(f"{label} file not found: {p}")
            log(f"{label + ':':10}{p.relative_to(cfg.base) if p.is_relative_to(cfg.base) else p}  "
                f"(SHA-256 {sha256(p)})")
        log()

        book = read_book(cfg, rep)
        link_rows(book, cfg, rep)
        tpl = analyse_template(cfg, book, rep)
        check_metadata(cfg, book, tpl, rep)
        out_path = output_path(cfg, book, rep)
        if out_path is not None:
            log_path = out_path.with_name(out_path.stem + "_run_log.txt")

        counts = ", ".join(f"{len(s.rows)} {n}" for n, s in book.sheets.items() if n != cfg.study_sheet)
        log(f"Read {len(book.study)} Study fields, {counts}.")
        log(f"Template: {tpl.control_count} content controls, {len(tpl.bookmarks)} bookmarks; "
            f"levels: {', '.join(book.levels)}.")
        log()
        print_report(log, rep, "Checks before building", after=False)
        if rep.notes:
            log()
            log("Notes (these do not stop the run)")
            for n in rep.notes:
                log(f"  - {n}")
        log()

        problems = rep.problems
        if problems:
            log(f"STOPPED: {len(problems)} problem(s) found in the metadata or template. No SAP was written.")
            log("Please fix these and run again:")
            for i, p in enumerate(problems, 1):
                log(f"  {i}. {p}")
            log.write(log_path)
            print(f"\nRun log: {log_path}")
            return 1
        if args.check_only:
            log("Check only: all checks passed. Nothing was built.")
            log.write(log_path)
            print(f"\nRun log: {log_path}")
            return 0

        builder = Builder(cfg, book, tpl)
        parts = builder.build()
        extra = {}
        if "docProps/core.xml" in tpl.raw:
            extra["docProps/core.xml"] = update_core_properties(
                tpl.raw["docProps/core.xml"], fill_pattern(cfg.doc_title, cfg, book),
                fill_pattern(cfg.doc_description, cfg, book))
        docx = package(tpl, parts, extra)

        log("Build")
        for base, ids in builder.copies.items():
            log(f"  - {base}: {len(ids)} cop{'y' if len(ids) == 1 else 'ies'} ({', '.join(ids)})")
        for where in builder.not_applicable:
            log(f"  - '{cfg.not_applicable}': {where}")
        log(f"  - {builder.filled} fields filled; {len(builder.placeholders_left)} optional blank(s) left as "
            f"placeholder ({', '.join(builder.placeholders_left) or 'none'}); "
            f"{builder.removed} template guidance paragraph(s) removed")
        log()

        verify(cfg, book, tpl, docx, builder, rep)
        print_report(log, rep, "Checks after building", after=True)
        log(f"  ({getattr(rep, 'checked_values', 0)} values compared; "
            f"{getattr(rep, 'bookmarks_kept', 0)} bookmarks kept)")
        log()
        problems = rep.problems
        if problems:
            log(f"STOPPED: the finished SAP failed {len(problems)} check(s), so it was not saved:")
            for i, p in enumerate(problems, 1):
                log(f"  {i}. {p}")
            log.write(log_path)
            print(f"\nRun log: {log_path}")
            return 2

        out_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = out_path.with_suffix(".tmp")
        tmp.write_bytes(docx)
        try:
            tmp.replace(out_path)
        except PermissionError:
            tmp.unlink(missing_ok=True)
            raise SetupError(f"Could not save {out_path.name}. Is it open in Word? Close it and run again.")
        log(f"Output:  {out_path.relative_to(cfg.base) if out_path.is_relative_to(cfg.base) else out_path}  "
            f"(SHA-256 {sha256(docx)})")
        log("RESULT:  SAP built and checked. In Word, update the table of contents when prompted.")
        log.write(log_path)
        print(f"\nRun log: {log_path}")
        return 0
    except SetupError as exc:
        log()
        log(f"STOPPED: {exc}")
        if log_path is not None:
            log.write(log_path)
            print(f"\nRun log: {log_path}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
