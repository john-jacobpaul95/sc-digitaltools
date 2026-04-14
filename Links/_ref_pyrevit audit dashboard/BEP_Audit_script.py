#! python3
# -*- coding: utf-8 -*-
# BEP_Audit_Export.py
# Consolidated BEP Audit Export: Grids, Levels, Coordinates → single XLSX
# PyRevit 6.1 / CPython 3 compatible — no external dependencies
# Uses System.IO.Packaging for XLSX output (same pattern as ModelHealth)
#
# Sheets produced:
#   Grids       — per-grid geometry (origin, angle, length) host vs links, with deltas
#   Levels      — level elevations and parameters (host + links)
#   Coordinates — project locations, base point, survey point (host + links)
#
# OPTIONS (edit below before running):
#   ANGLE_TOL_DEG — tolerance for near-parallel grid detection (used for angle delta)

import os, sys, math, datetime, re

# ---- PyRevit 6.1 CPython stdout compatibility shim ----
# ScriptIO wraps stdout with .Print() instead of .write()
# This shim restores .write() so print() and sys.exit() work correctly
if not hasattr(sys.stdout, 'write'):
    sys.stdout.write = lambda s: sys.stdout.Print(s) if s.strip() else None
if not hasattr(sys.stderr, 'write'):
    sys.stderr.write = lambda s: sys.stderr.Print(s) if s.strip() else None

import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('WindowsBase')   # required for System.IO.Packaging

from System import Array, Environment, Uri, UriKind, Byte
from System.Windows.Forms import (
    SaveFileDialog, DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon,
    Form, Label, TextBox, CheckedListBox, Button,
    AnchorStyles, FormStartPosition, Keys, CheckState, Control
)
from System.Drawing import Size, Point
from System.IO.Packaging import Package, CompressionOption, TargetMode
from System.IO import FileMode, FileAccess, FileShare

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, RevitLinkInstance,
    Grid, Level, XYZ, BasePoint,
    BuiltInCategory, BuiltInParameter
)

# ---- Document access — __revit__ builtin (PyRevit 6.1 CPython safe) ----
# Avoids importing pyrevit.revit which triggers an event handler conflict in 6.1
try:
    uidoc = __revit__.ActiveUIDocument
    doc   = uidoc.Document if uidoc else None
except Exception:
    doc   = None
    uidoc = None

if doc is None:
    print("No active Revit document. Open a model and run again.")
    sys.exit(1)

HOST_MODEL  = doc.Title
EXPORTED_AT = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# ============================================================
#  OPTIONS
# ============================================================
ANGLE_TOL_DEG = 0.3   # degrees — used when reporting angle deltas

# ============================================================
#  UNIT HELPERS
# ============================================================
_HAS_UNITTYPEID = hasattr(DB, 'UnitTypeId')

def mm(v):
    try:
        if _HAS_UNITTYPEID:
            return DB.UnitUtils.ConvertFromInternalUnits(v, DB.UnitTypeId.Millimeters)
        else:
            return DB.UnitUtils.ConvertFromInternalUnits(v, DB.DisplayUnitType.DUT_MILLIMETERS)
    except Exception:
        return v

def r2(x):
    if x in ('', None): return ''
    try:    return round(float(x), 2)
    except: return x

FOOT_TO_MM = 304.8

def rad_to_deg(r):
    return (r * 180.0) / math.pi if r is not None else None

def norm_deg(d):
    if d is None: return None
    d = d % 360.0
    return d + 360.0 if d < 0 else d

def fmt_tz(offset_hours):
    if offset_hours is None: return ""
    sign = "+" if offset_hours >= 0 else "-"
    h = int(abs(offset_hours))
    m = int(round((abs(offset_hours) - h) * 60))
    return "UTC{0}{1:02d}:{2:02d}".format(sign, h, m)

# ============================================================
#  PARAMETER HELPERS
# ============================================================
def _param_as_text(p):
    if p is None: return ''
    try:
        s = p.AsValueString()
        if s not in (None, ''): return s
    except Exception: pass
    try:
        st = p.StorageType
        if   st == DB.StorageType.Integer: return str(p.AsInteger())
        elif st == DB.StorageType.Double:  return str(p.AsDouble())
        else:
            s2 = p.AsString()
            if s2 not in (None, ''): return s2
    except Exception: pass
    return ''

def _find_param_by_names(elem, names):
    if elem is None: return None
    for n in names:
        try:
            p = elem.LookupParameter(n)
            if p: return p
        except Exception: pass
    lname_set = {n.lower() for n in names}
    try:
        for p in elem.Parameters:
            try:
                nm = p.Definition.Name if p.Definition else None
                if nm and nm.lower() in lname_set: return p
            except Exception: pass
    except Exception: pass
    return None

def _map_elev_base(text_or_int):
    if isinstance(text_or_int, int):
        if text_or_int == 0: return "Project Base Point"
        if text_or_int == 1: return "Survey Point"
        return str(text_or_int)
    s = (text_or_int or "").strip().lower()
    if not s: return ""
    if s in {"project", "project base", "project base point", "pbp"}:
        return "Project Base Point"
    if s in {"shared", "shared coordinates", "survey", "survey point", "sp"}:
        return "Survey Point"
    return text_or_int

def level_param_triplet(level, curdoc):
    elev_txt = ""
    try: elev_txt = _param_as_text(level.LookupParameter("Elevation"))
    except Exception: pass

    base_txt = ""
    try:
        lvl_type = curdoc.GetElement(level.GetTypeId())
        if lvl_type:
            p_base = _find_param_by_names(lvl_type, ["Elevation Base"])
            if p_base:
                s = _param_as_text(p_base)
                if s in ("", None):
                    try:    base_txt = _map_elev_base(p_base.AsInteger())
                    except: base_txt = ""
                else:
                    base_txt = _map_elev_base(s)
    except Exception: base_txt = ""

    story_txt = ""
    try:
        p_story = _find_param_by_names(level, ["Building Story", "Building Storey"])
        s = _param_as_text(p_story) if p_story else ""
        if s in ("", None) and p_story is not None:
            try:    story_txt = "Yes" if int(p_story.AsInteger()) == 1 else "No"
            except: story_txt = ""
        else:
            story_txt = s
    except Exception: story_txt = ""

    return elev_txt, base_txt, story_txt

# ============================================================
#  COORDINATE HELPERS (base point / survey point / project location)
# ============================================================
def shared_ENZ_at_point(host_doc, pt):
    try:
        pos = host_doc.ActiveProjectLocation.GetProjectPosition(pt)
        return (mm(pos.EastWest), mm(pos.NorthSouth), mm(pos.Elevation))
    except Exception:
        return ('', '', '')

def _get_param_double(elem, bip):
    try:
        if bip is None: return None
        p = elem.get_Parameter(bip)
        return None if p is None else p.AsDouble()
    except Exception: return None

def _get_param_double_by_names(elem, name_candidates):
    try:
        for p in elem.Parameters:
            nm  = p.Definition.Name if p.Definition else ""
            nml = (nm or "").lower()
            if not nml: continue
            if any(c in nml for c in name_candidates):
                if str(p.StorageType) == "Double": return p.AsDouble()
                try:    return float(p.AsValueString())
                except: pass
    except Exception: pass
    return None

def _bip(name):
    try:    return getattr(DB.BuiltInParameter, name)
    except: return None

# BuiltIn parameter candidates (handles both pre-2022 and post-2022 Revit)
BIP_PBP_E = [_bip('BASEPOINT_EASTWEST_PARAM'),        _bip('BASEPOINT_EASTWEST_SHARED_PARAM')]
BIP_PBP_N = [_bip('BASEPOINT_NORTHSOUTH_PARAM'),      _bip('BASEPOINT_NORTHSOUTH_SHARED_PARAM')]
BIP_PBP_Z = [_bip('BASEPOINT_ELEVATION_PARAM'),       _bip('BASEPOINT_ELEVATION_SHARED_PARAM')]
BIP_PBP_A = [_bip('BASEPOINT_ANGLETON_PARAM')]
BIP_SP_E  = [_bip('BASEPOINT_EASTWEST_SHARED_PARAM'), _bip('BASEPOINT_EASTWEST_PARAM')]
BIP_SP_N  = [_bip('BASEPOINT_NORTHSOUTH_SHARED_PARAM'), _bip('BASEPOINT_NORTHSOUTH_PARAM')]
BIP_SP_Z  = [_bip('BASEPOINT_ELEVATION_SHARED_PARAM'), _bip('BASEPOINT_ELEVATION_PARAM')]

def _first_double(elem, bip_list, name_list):
    for bip in bip_list:
        v = _get_param_double(elem, bip)
        if v is not None: return v
    return _get_param_double_by_names(elem, name_list)

def find_basepoint(d, survey=False):
    # Strategy 1: category filter
    try:
        bic   = DB.BuiltInCategory.OST_SharedBasePoint if survey else DB.BuiltInCategory.OST_ProjectBasePoint
        elems = list(FilteredElementCollector(d).OfCategory(bic).WhereElementIsNotElementType())
        if elems: return elems[0]
    except Exception: pass
    # Strategy 2: API getter (Revit 2022+)
    try:
        bp = BasePoint.GetSurveyPoint(d) if survey else BasePoint.GetProjectBasePoint(d)
        if bp: return bp
    except Exception: pass
    # Strategy 3: scan BasePoint class by IsShared flag
    try:
        cands = list(FilteredElementCollector(d).OfClass(BasePoint))
        if not cands: return None
        for x in cands:
            try:
                if survey     and     x.IsShared: return x
                if not survey and not x.IsShared: return x
            except Exception: continue
        return cands[0]
    except Exception: return None

def pbp_values(d):
    bp = find_basepoint(d, survey=False)
    if not bp: return (None, None, None, None)
    e_ft = _first_double(bp, BIP_PBP_E, ['e/w', 'east/west'])
    n_ft = _first_double(bp, BIP_PBP_N, ['n/s', 'north/south'])
    z_ft = _first_double(bp, BIP_PBP_Z, ['elev', 'elevation'])
    a_rd = _first_double(bp, BIP_PBP_A, ['angle to true north'])
    return (
        None if e_ft is None else e_ft * FOOT_TO_MM,
        None if n_ft is None else n_ft * FOOT_TO_MM,
        None if z_ft is None else z_ft * FOOT_TO_MM,
        None if a_rd is None else norm_deg(rad_to_deg(a_rd))
    )

def sp_values(d):
    bp = find_basepoint(d, survey=True)
    if not bp: return (None, None, None)
    e_ft = _first_double(bp, BIP_SP_E, ['e/w', 'east/west'])
    n_ft = _first_double(bp, BIP_SP_N, ['n/s', 'north/south'])
    z_ft = _first_double(bp, BIP_SP_Z, ['elev', 'elevation'])
    return (
        None if e_ft is None else e_ft * FOOT_TO_MM,
        None if n_ft is None else n_ft * FOOT_TO_MM,
        None if z_ft is None else z_ft * FOOT_TO_MM
    )

def safe_gis_code(site):
    try:
        code = getattr(site, "GeoCoordinateSystemId", None)
        return "" if code is None else str(code)
    except Exception: return ""

def iter_project_locations(plset):
    try:
        it = plset.ForwardIterator(); it.Reset()
        while it.MoveNext(): yield it.Current
    except Exception:
        for pl in plset: yield pl

# ============================================================
#  GRID GEOMETRY HELPERS
# ============================================================
def norm_key(name):
    return (name or '').strip().upper()

def sanitise_grid_label(name):
    """
    Strip ALL whitespace-like and invisible Unicode characters from a grid label,
    including regular spaces and Braille blank U+2800 that Revit sometimes embeds.
    """
    if not name:
        return ''
    return re.sub(r'[\s\u2800\u200b\u200c\u200d\u00a0\ufeff]+', '', name)

def norm_grid_key(name):
    # dot/dash/space/invisible-char-insensitive: A.1 == A-1 == ' A1' == u'\u2800A1' == A1
    s = sanitise_grid_label(name).upper()
    s = s.replace('.', '').replace('-', '')
    return s

def pick_fallback_plan_view(d):
    try:
        for v in FilteredElementCollector(d).OfClass(DB.ViewPlan):
            if not v.IsTemplate and v.ViewType == DB.ViewType.FloorPlan: return v
        for v in FilteredElementCollector(d).OfClass(DB.View):
            if not v.IsTemplate: return v
    except Exception: pass
    return None

def get_grid_curve(d, grid, view=None):
    """Return the first linear curve for a grid (handles multi-segment grids)."""
    try:
        if not view: view = pick_fallback_plan_view(d)
        if getattr(grid, "IsMultiSegment", False):
            for crv in grid.GetCurvesInView(DB.DatumExtentType.Model, view):
                if isinstance(crv, DB.Line): return crv
        else:
            crv = grid.Curve
            if isinstance(crv, DB.Line): return crv
    except Exception:
        crv = getattr(grid, "Curve", None)
        if isinstance(crv, DB.Line): return crv
    return None

def grid_geometry(d, grid, T=None):
    """
    Extract the infinite grid line geometry for distance comparison.

    Returns (p0_XYZ, dir_XYZ, angle_deg, length_mm) where:
      p0  — one point on the line (GetEndPoint(0), transformed to host internal space)
      dir — unit direction vector of the line in host internal space
      angle_deg — bearing 0–180 (direction-agnostic, 0=N 90=E clockwise)
      length_mm — visible segment length in mm

    The perpendicular distance between two infinite grid lines is computed from
    these values directly in internal space using the cross-product formula:
      dist = |( p0_link - p0_host ) x dir_host| / |dir_host|
    This is completely independent of both models' internal origins and of
    segment extent — the only correct way to compare grid positions.

    T is an optional Revit Transform to bring link-model coords into host space.
    Returns None if the grid has no usable linear curve.
    """
    v   = pick_fallback_plan_view(d)
    crv = get_grid_curve(d, grid, v)
    if crv is None:
        return None
    p0 = crv.GetEndPoint(0)
    p1 = crv.GetEndPoint(1)
    if T:
        p0 = T.OfPoint(p0)
        p1 = T.OfPoint(p1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    dz = p1.Z - p0.Z
    length_ft = math.sqrt(dx*dx + dy*dy + dz*dz)
    if length_ft < 1e-12:
        return None
    # Unit direction vector
    ux = dx / length_ft
    uy = dy / length_ft
    uz = dz / length_ft
    # Bearing from true north (0=N, 90=E, clockwise), normalised 0–180
    angle_deg = norm_deg(90.0 - math.degrees(math.atan2(dy, dx)))
    if angle_deg >= 180.0:
        angle_deg -= 180.0
    return p0, XYZ(ux, uy, uz), round(angle_deg, 4), round(length_ft * FOOT_TO_MM, 2)

def collect_grids_by_key(d):
    """
    Return {norm_key: (display_name, grid_element)} for all linear grids in document d.
    """
    out = {}
    for g in FilteredElementCollector(d).OfClass(Grid):
        key  = norm_grid_key(g.Name)
        disp = sanitise_grid_label(g.Name)
        if not key: continue
        if key not in out:
            out[key] = (disp, g)
    return out

# ============================================================
#  WINFORMS INTERACTIVE PICKER (filter + Shift multi-range)
# ============================================================
def pick_many_filterable(title, caption, items_display, prechecked_names=None, precheck_all=True):
    all_items = list(items_display)
    checked   = set(prechecked_names) if prechecked_names is not None else (set(all_items) if precheck_all else set())
    display_items = []

    f = Form()
    f.Text = title; f.Size = Size(560, 640)
    f.StartPosition = FormStartPosition.CenterScreen

    lbl  = Label();  lbl.Text  = caption; lbl.AutoSize = True; lbl.Location = Point(12, 12);  f.Controls.Add(lbl)
    lblf = Label();  lblf.Text = "Filter:"; lblf.AutoSize = True; lblf.Location = Point(12, 40); f.Controls.Add(lblf)
    tb   = TextBox(); tb.Location = Point(60, 36); tb.Size = Size(360, 22)
    tb.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right); f.Controls.Add(tb)

    btn_all    = Button(); btn_all.Text    = "Select All (filtered)"; btn_all.Location    = Point(430, 34); btn_all.Size    = Size(110, 26)
    btn_all.Anchor = (AnchorStyles.Top | AnchorStyles.Right); f.Controls.Add(btn_all)

    clb = CheckedListBox(); clb.Location = Point(12, 68); clb.Size = Size(528, 488)
    clb.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right)
    clb.CheckOnClick = True; f.Controls.Add(clb)

    btn_ok     = Button(); btn_ok.Text     = "OK";               btn_ok.Location     = Point(356, 570)
    btn_clear  = Button(); btn_clear.Text  = "Clear (filtered)"; btn_clear.Location  = Point(240, 570)
    btn_cancel = Button(); btn_cancel.Text = "Cancel";           btn_cancel.Location = Point(460, 570)
    for btn in [btn_ok, btn_clear, btn_cancel]:
        btn.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right); f.Controls.Add(btn)

    last_index = {'idx': None}
    bulk       = {'on': False}

    def rebuild_list():
        filt    = (tb.Text or '').lower()
        visible = [s for s in all_items if filt in s.lower()]
        display_items[:] = visible
        try:
            clb.BeginUpdate(); clb.Items.Clear()
            for s in visible: clb.Items.Add(s, s in checked)
        finally:
            clb.EndUpdate()
        last_index['idx'] = None

    def on_text_changed(sender, args): rebuild_list()
    def on_mouse_down(sender, me):     last_index['idx'] = clb.IndexFromPoint(me.Location)

    def on_item_check(sender, e):
        if bulk['on']: return
        try:
            if (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
                li = last_index['idx']
                if li is None or li < 0:
                    name_i = clb.Items[e.Index]
                    if e.NewValue == CheckState.Checked: checked.add(name_i)
                    else: checked.discard(name_i)
                    return
                start = min(li, e.Index); end = max(li, e.Index)
                new_c = (e.NewValue == CheckState.Checked)
                bulk['on'] = True
                for j in range(start, end + 1):
                    clb.SetItemChecked(j, new_c)
                    if new_c: checked.add(clb.Items[j])
                    else:     checked.discard(clb.Items[j])
            else:
                name_i = clb.Items[e.Index]
                if e.NewValue == CheckState.Checked: checked.add(name_i)
                else: checked.discard(name_i)
        finally:
            bulk['on'] = False

    def on_select_all(sender, args):
        bulk['on'] = True
        try:
            for i in range(clb.Items.Count):
                clb.SetItemChecked(i, True); checked.add(clb.Items[i])
        finally: bulk['on'] = False

    def on_clear_filtered(sender, args):
        bulk['on'] = True
        try:
            for i in range(clb.Items.Count):
                clb.SetItemChecked(i, False); checked.discard(clb.Items[i])
        finally: bulk['on'] = False

    def on_ok(sender, args):
        f.Tag = [s for s in all_items if s in checked]
        f.DialogResult = DialogResult.OK; f.Close()

    def on_cancel(sender, args):
        f.Tag = None; f.DialogResult = DialogResult.Cancel; f.Close()

    tb.TextChanged        += on_text_changed
    clb.MouseDown         += on_mouse_down
    clb.ItemCheck         += on_item_check
    btn_all.Click         += on_select_all
    btn_clear.Click       += on_clear_filtered
    btn_ok.Click          += on_ok
    btn_cancel.Click      += on_cancel

    rebuild_list()
    res = f.ShowDialog()
    return list(f.Tag or []) if res == DialogResult.OK else None

# ============================================================
#  XLSX WRITER — System.IO.Packaging (no openpyxl / no pip)
#  Pattern taken directly from ModelHealth_script.py
# ============================================================
XML_HEADER   = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_illegal_xml = re.compile("[\x00-\x08\x0B\x0C\x0E-\x1F]")

CT_XLSX_MAIN  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_WS         = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_SHARED     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
CT_STYLES     = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
CT_CORE       = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP        = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
REL_OFFICEDOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
REL_SHEET     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
REL_SHARED    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
REL_STYLES    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_CORE      = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_APP       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

def _bytes(b): return Array[Byte](bytearray(b))

def clean_sheet_name(name):
    s = (name or "Sheet")[:31]
    for bad in [":", "\\", "/", "?", "*", "[", "]"]: s = s.replace(bad, "_")
    return s or "Sheet"

def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def xml_escape(v):
    if v is None: return ""
    s = str(v)
    s = _illegal_xml.sub(" ", s)
    s = (s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
          .replace('"',"&quot;").replace("'","&apos;"))
    return s[:32760] if len(s) > 32760 else s

def tokenise_cell(v):
    # Numbers stay as numbers; everything else is a shared string
    return ('n', v) if isinstance(v, (int, float)) and v is not True and v is not False \
           else ('s', "" if v is None else str(v))

def prepare_sheet_tokens(headers, rows):
    return ([tokenise_cell(h) for h in headers],
            [[tokenise_cell(v) for v in row] for row in rows])

def build_shared_strings(prepared):
    sst, idx, total = [], {}, 0
    for sh in prepared:
        for t in sh['headers_tokens']:
            if t[0] == 's':
                total += 1
                if t[1] not in idx: idx[t[1]] = len(sst); sst.append(t[1])
        for r in sh['rows_tokens']:
            for t in r:
                if t[0] == 's':
                    total += 1
                    if t[1] not in idx: idx[t[1]] = len(sst); sst.append(t[1])
    parts = [XML_HEADER,
             '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'count="%d" uniqueCount="%d">' % (total, len(sst))]
    for s in sst:
        parts.append('<si><t xml:space="preserve">%s</t></si>' % xml_escape(s))
    parts.append('</sst>')
    return "".join(parts).encode('utf-8'), idx

def sheet_xml_from_tokens(headers_tokens, rows_tokens, sst_index):
    parts = [XML_HEADER,
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<sheetData>']
    parts.append('<row r="1">')
    for c, tok in enumerate(headers_tokens, 1):
        ref = '%s1' % col_letter(c)
        if tok[0] == 'n': parts.append('<c r="%s"><v>%s</v></c>' % (ref, str(tok[1])))
        else:             parts.append('<c r="%s" t="s"><v>%d</v></c>' % (ref, sst_index.get(tok[1], 0)))
    parts.append('</row>')
    for r, row in enumerate(rows_tokens, 2):
        parts.append('<row r="%d">' % r)
        for c, tok in enumerate(row, 1):
            ref = '%s%d' % (col_letter(c), r)
            if tok[0] == 'n': parts.append('<c r="%s"><v>%s</v></c>' % (ref, str(tok[1])))
            else:             parts.append('<c r="%s" t="s"><v>%d</v></c>' % (ref, sst_index.get(tok[1], 0)))
        parts.append('</row>')
    parts.append('</sheetData></worksheet>')
    return "".join(parts).encode('utf-8')

def styles_xml():
    return (XML_HEADER +
'''<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>''').encode('utf-8')

def core_props_xml():
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (XML_HEADER +
'''<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>BEP Audit Export</dc:title>
  <dc:creator>pyRevit</dc:creator>
  <cp:lastModifiedBy>pyRevit</cp:lastModifiedBy>
  <dcterms:created  xsi:type="dcterms:W3CDTF">''' + now + '''</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">''' + now + '''</dcterms:modified>
</cp:coreProperties>''').encode('utf-8')

def app_props_xml(sheet_names):
    n = str(len(sheet_names))
    parts = [XML_HEADER,
             '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
             'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
             '<Application>pyRevit</Application>'
             '<HeadingPairs><vt:vector size="2" baseType="variant">'
             '<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>'
             '<vt:variant><vt:i4>' + n + '</vt:i4></vt:variant>'
             '</vt:vector></HeadingPairs>'
             '<TitlesOfParts><vt:vector size="' + n + '" baseType="lpstr">']
    for nm in sheet_names:
        parts.append('<vt:lpstr>%s</vt:lpstr>' % xml_escape(nm))
    parts.append('</vt:vector></TitlesOfParts></Properties>')
    return "".join(parts).encode('utf-8')

def build_xlsx_pkg(xlsx_path, sheets_payload):
    prepared = []
    for (name, headers, rows) in sheets_payload:
        htok, rtok = prepare_sheet_tokens(headers, rows)
        prepared.append({'name': clean_sheet_name(name), 'headers_tokens': htok, 'rows_tokens': rtok})
    sst_xml, sst_index = build_shared_strings(prepared)

    folder = os.path.dirname(xlsx_path)
    if folder and not os.path.isdir(folder):
        try: os.makedirs(folder)
        except: pass

    pkg = Package.Open(xlsx_path, FileMode.Create, FileAccess.ReadWrite, getattr(FileShare, 'None'))
    try:
        # Parts
        core = pkg.CreatePart(Uri("/docProps/core.xml",      UriKind.Relative), CT_CORE,      CompressionOption.Normal)
        cb   = core_props_xml(); core.GetStream(FileMode.Create, FileAccess.Write).Write(_bytes(cb), 0, len(cb))

        app  = pkg.CreatePart(Uri("/docProps/app.xml",       UriKind.Relative), CT_APP,       CompressionOption.Normal)
        ab   = app_props_xml([p['name'] for p in prepared]); app.GetStream(FileMode.Create, FileAccess.Write).Write(_bytes(ab), 0, len(ab))

        wb   = pkg.CreatePart(Uri("/xl/workbook.xml",        UriKind.Relative), CT_XLSX_MAIN, CompressionOption.Normal)
        shr  = pkg.CreatePart(Uri("/xl/sharedStrings.xml",   UriKind.Relative), CT_SHARED,    CompressionOption.Normal)
        shr.GetStream(FileMode.Create, FileAccess.Write).Write(_bytes(sst_xml), 0, len(sst_xml))
        sty  = pkg.CreatePart(Uri("/xl/styles.xml",          UriKind.Relative), CT_STYLES,    CompressionOption.Normal)
        sx   = styles_xml(); sty.GetStream(FileMode.Create, FileAccess.Write).Write(_bytes(sx), 0, len(sx))

        sheet_parts = []
        for i, p in enumerate(prepared, 1):
            ws   = pkg.CreatePart(Uri("/xl/worksheets/sheet%d.xml" % i, UriKind.Relative), CT_WS, CompressionOption.Normal)
            data = sheet_xml_from_tokens(p['headers_tokens'], p['rows_tokens'], sst_index)
            s    = ws.GetStream(FileMode.Create, FileAccess.Write)
            s.Write(_bytes(data), 0, len(data)); s.Close()
            sheet_parts.append(ws)

        # Root relationships
        pkg.CreateRelationship(Uri("/xl/workbook.xml",   UriKind.Relative), TargetMode.Internal, REL_OFFICEDOC, "rIdWorkbook")
        pkg.CreateRelationship(Uri("/docProps/core.xml", UriKind.Relative), TargetMode.Internal, REL_CORE,      "rIdCore")
        pkg.CreateRelationship(Uri("/docProps/app.xml",  UriKind.Relative), TargetMode.Internal, REL_APP,       "rIdApp")

        # Sheet relationships with explicit IDs
        sheet_ids = []
        for i, _ in enumerate(sheet_parts, 1):
            rid = "rId%d" % i
            wb.CreateRelationship(Uri("/xl/worksheets/sheet%d.xml" % i, UriKind.Relative), TargetMode.Internal, REL_SHEET, rid)
            sheet_ids.append(rid)
        wb.CreateRelationship(Uri("/xl/sharedStrings.xml", UriKind.Relative), TargetMode.Internal, REL_SHARED, "rIdShared")
        wb.CreateRelationship(Uri("/xl/styles.xml",        UriKind.Relative), TargetMode.Internal, REL_STYLES, "rIdStyles")

        def workbook_xml(sheet_names, rids):
            parts = [XML_HEADER,
                     '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                     '<sheets>']
            for idx, (nm, rid) in enumerate(zip(sheet_names, rids), 1):
                parts.append('<sheet name="%s" sheetId="%d" r:id="%s"/>' % (xml_escape(nm), idx, rid))
            parts.append('</sheets></workbook>')
            return "".join(parts).encode('utf-8')

        wb_bytes = workbook_xml([p['name'] for p in prepared], sheet_ids)
        wbs = wb.GetStream(FileMode.Create, FileAccess.Write)
        wbs.Write(_bytes(wb_bytes), 0, len(wb_bytes)); wbs.Close()

        pkg.Flush()
        return True

    except Exception as ex:
        MessageBox.Show("Failed to create XLSX.\n{}".format(ex), "pyRevit – BEP Audit", MessageBoxButtons.OK)
        return False
    finally:
        pkg.Close()

# ============================================================
#  SAVE AS — single dialog, one XLSX output
# ============================================================
def ask_save_xlsx(default_filename="BEP_Audit_Export.xlsx"):
    dlg = SaveFileDialog()
    dlg.Title            = "Save BEP Audit Export As"
    dlg.Filter           = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    dlg.FileName         = default_filename
    dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    if dlg.ShowDialog() == DialogResult.OK and dlg.FileName:
        return dlg.FileName
    return None

save_path = ask_save_xlsx()
if not save_path:
    print("Export cancelled."); sys.exit(0)

# ============================================================
#  COLLECT LINKED MODELS
# ============================================================
links = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))

# ============================================================
#  PART A — GRIDS  (isolation mode — one row per grid per model)
# ============================================================
print("Collecting grids...")

host_grid_map = collect_grids_by_key(doc)
if not host_grid_map:
    print("No linear grids found in the host model.")
    sys.exit(0)

print("Host grids found: {}".format(len(host_grid_map)))

# --- Compute host grid geometry ---
# p0/dir are in host internal space (feet). E/N are shared coords of p0 for display.
host_geom = {}   # norm_key → {'name', 'p0', 'dir', 'angle', 'length', 'E', 'N'}
for key, (disp, g) in host_grid_map.items():
    geo = grid_geometry(doc, g)
    if geo is None: continue
    p0, dir_v, angle, length = geo
    E, N, _ = shared_ENZ_at_point(doc, p0)
    host_geom[key] = {'name': disp, 'p0': p0, 'dir': dir_v,
                      'angle': angle, 'length': length, 'E': E, 'N': N}

GRIDS_HEADERS = [
    'SourceModel', 'ModelScope', 'GridName',
    'Origin_East_mm', 'Origin_North_mm',
    'Angle_deg', 'Length_mm',
    'Delta_East_mm', 'Delta_North_mm', 'Delta_Pos_mm',
    'Delta_Angle_deg',
    'Status', 'ExportedAt'
]
grids_rows = []

# --- Host rows (no deltas) ---
for key in sorted(host_geom.keys()):
    hg = host_geom[key]
    grids_rows.append([
        HOST_MODEL, 'Host', hg['name'],
        r2(hg['E']), r2(hg['N']),
        hg['angle'], hg['length'],
        '', '', '', '',
        'Host', EXPORTED_AT
    ])

# --- Link rows ---
seen_link_docs = set()
for li in links:
    ld = li.GetLinkDocument()
    if not ld or ld.Title in seen_link_docs: continue
    seen_link_docs.add(ld.Title)

    T             = li.GetTransform()
    link_grid_map = collect_grids_by_key(ld)
    host_keys_set = set(host_geom.keys())
    link_keys_set = set(link_grid_map.keys())

    # Grids present in host — compare geometry
    for key in sorted(host_keys_set):
        hg = host_geom[key]
        if key in link_grid_map:
            _, lg = link_grid_map[key]
            geo   = grid_geometry(ld, lg, T)
            if geo is None:
                grids_rows.append([
                    ld.Title, 'Link', hg['name'],
                    '', '', '', '',
                    '', '', '', '',
                    'MISSING', EXPORTED_AT
                ])
                continue
            p0_l, dir_l, angle_l, length_l = geo
            # Display coords: shared E/N of the link grid's p0 (for human reference only)
            E_l, N_l, _ = shared_ENZ_at_point(doc, p0_l)
            dE = dN = dpos = dang = ''
            # --- 2D perpendicular distance between the two infinite grid lines ---
            # Work entirely in XY (plan view). Z is intentionally ignored — grids
            # in different models often sit at different elevations (a full storey
            # height difference is common) and that is irrelevant for a position audit.
            # Using the full 3D cross product previously caused Z offsets (~1200mm
            # per storey) to dominate Delta_Pos even when the grids were perfectly
            # aligned in plan.
            hp0  = hg['p0']
            hdir = hg['dir']
            # XY vector from any point on host line to any point on link line
            wx = p0_l.X - hp0.X
            wy = p0_l.Y - hp0.Y
            # XY direction of host grid, normalised to unit length in XY plane
            # (hdir is a 3D unit vector; its XY magnitude may be <1 for sloped grids)
            hd_len = math.hypot(hdir.X, hdir.Y)
            if hd_len > 1e-9:
                hd_x = hdir.X / hd_len
                hd_y = hdir.Y / hd_len
            else:
                hd_x, hd_y = 1.0, 0.0   # fallback for vertical grids (rare)
            # Project wx,wy onto the grid direction — this is the along-line component
            # (extent / start-point difference — irrelevant for position check)
            along  = wx * hd_x + wy * hd_y
            # Perpendicular components in internal feet
            perp_x = wx - along * hd_x
            perp_y = wy - along * hd_y
            # Convert to mm
            dE   = r2(perp_x * FOOT_TO_MM)
            dN   = r2(perp_y * FOOT_TO_MM)
            dpos = r2(math.hypot(perp_x, perp_y) * FOOT_TO_MM)
            # dpos == hypot(dE, dN) by construction — no inconsistency possible
            # Angle delta normalised to 0–90 (direction-agnostic)
            raw_da = abs(angle_l - hg['angle'])
            dang   = round(min(raw_da, 180.0 - raw_da) if raw_da > 90.0 else raw_da, 4)
            status = 'FAIL' if (float(dpos) > 1.0 or float(dang) > 0.01) else 'PASS'
            grids_rows.append([
                ld.Title, 'Link', hg['name'],
                r2(E_l), r2(N_l),
                angle_l, length_l,
                dE, dN, dpos, dang,
                status, EXPORTED_AT
            ])
        else:
            # Grid in host but not found in this link
            grids_rows.append([
                ld.Title, 'Link', hg['name'],
                '', '', '', '',
                '', '', '', '',
                'MISSING', EXPORTED_AT
            ])

    # Grids in link but not in host — flag as Added
    for key in sorted(link_keys_set - host_keys_set):
        disp, lg = link_grid_map[key]
        geo = grid_geometry(ld, lg, T)
        E_l = N_l = angle_l = length_l = ''
        if geo is not None:
            p0_l, _, angle_l, length_l = geo
            E_l, N_l, _ = shared_ENZ_at_point(doc, p0_l)
        grids_rows.append([
            ld.Title, 'Link', disp,
            r2(E_l), r2(N_l),
            angle_l, length_l,
            '', '', '', '',
            'ADDED', EXPORTED_AT
        ])

print("Grid rows written: {}".format(len(grids_rows)))

# ============================================================
#  PART B — LEVELS
# ============================================================
print("Collecting levels...")

host_lvls_all = {}
for L in FilteredElementCollector(doc).OfClass(Level):
    disp = L.Name; key = norm_key(disp)
    if not key or key in host_lvls_all: continue
    p_host = XYZ(0.0, 0.0, L.Elevation)
    _, _, Z = shared_ENZ_at_point(doc, p_host)
    elev_txt, base_txt, story_txt = level_param_triplet(L, doc)
    host_lvls_all[key] = {'name': disp, 'Z': Z, 'p_elev': elev_txt, 'p_base': base_txt, 'p_story': story_txt}

level_names_all = [host_lvls_all[k]['name'] for k in sorted(host_lvls_all.keys())]
selLevels = pick_many_filterable(
    "Select Levels to export",
    "Tick the Levels to include in the Levels sheet:",
    level_names_all,
    precheck_all=True
)
if selLevels is None:
    print("Export cancelled (Level selection)."); sys.exit(0)

selLevel_keys = {norm_key(s) for s in selLevels if s is not None}
host_lvls     = {k: v for k, v in host_lvls_all.items() if k in selLevel_keys}

LEVELS_HEADERS = [
    'SourceModel', 'Level Name',
    'Param_Elevation', 'Param_ElevationBase', 'Param_BuildingStorey',
    'ExportedAt'
]
levels_rows = []

# Host block
for key in sorted(host_lvls.keys()):
    levels_rows.append([HOST_MODEL,
                        host_lvls[key]['name'],
                        host_lvls[key]['p_elev'],
                        host_lvls[key]['p_base'],
                        host_lvls[key]['p_story'],
                        EXPORTED_AT])
levels_rows.append([''] * len(LEVELS_HEADERS))  # blank separator

def pick_best_level(host_key, host_Z, link_by_key, link_all, used_ids):
    li = link_by_key.get(host_key)
    if li and li['id'] not in used_ids: return li
    best = None; best_abs = None
    for cand in link_all:
        if cand['id'] in used_ids: continue
        z = cand['Z']
        if z in ('', None) or host_Z in ('', None): continue
        dz = abs(z - host_Z)
        if best_abs is None or dz < best_abs: best = cand; best_abs = dz
    return best

# Link blocks
for li in links:
    ld = li.GetLinkDocument()
    if not ld: continue
    link_by_key = {}; link_all = []
    for L in FilteredElementCollector(ld).OfClass(Level):
        disp = L.Name; key = norm_key(disp)
        p_lnk = XYZ(0.0, 0.0, L.Elevation)
        _, _, Z_own = shared_ENZ_at_point(ld, p_lnk)
        elev_txt, base_txt, story_txt = level_param_triplet(L, ld)
        entry = {'id': L.Id.IntegerValue, 'name': disp, 'key': key, 'Z': Z_own,
                 'p_elev': elev_txt, 'p_base': base_txt, 'p_story': story_txt}
        link_all.append(entry)
        if key and key not in link_by_key: link_by_key[key] = entry

    used_ids = set()
    # Matched levels (host-driven order)
    for key in sorted(host_lvls.keys()):
        best = pick_best_level(key, host_lvls[key]['Z'], link_by_key, link_all, used_ids)
        if best is None:
            levels_rows.append([ld.Title, host_lvls[key]['name'], '', '', '', EXPORTED_AT])
            continue
        used_ids.add(best['id'])
        levels_rows.append([ld.Title, best['name'], best['p_elev'], best['p_base'], best['p_story'], EXPORTED_AT])
    # Remaining link levels not matched to any host level
    for rem in link_all:
        if rem['id'] not in used_ids:
            levels_rows.append([ld.Title, rem['name'], rem['p_elev'], rem['p_base'], rem['p_story'], EXPORTED_AT])
    levels_rows.append([''] * len(LEVELS_HEADERS))  # blank separator between models

# ============================================================
#  PART C — COORDINATES
# ============================================================
print("Collecting coordinates...")

COORDS_HEADERS = [
    "ModelScope", "ModelTitle",
    "ProjectLocationName", "IsActiveProjectLocation",
    "Site_PlaceName", "Site_Latitude_deg", "Site_Longitude_deg",
    "Site_TimeZone", "GIS_CoordinateSystem_Code",
    "PL_TrueNorth_deg", "PL_East_mm", "PL_North_mm", "PL_Elevation_mm",
    "PBP_East_mm", "PBP_North_mm", "PBP_Elevation_mm", "PBP_AngleToTrueNorth_deg",
    "SP_East_mm", "SP_North_mm", "SP_Elevation_mm",
    "ExportedAt"
]
coords_rows = []

def coords_for_doc(d, scope):
    rows  = []
    site  = getattr(d, "SiteLocation", None)
    site_place = site.PlaceName if site else ""
    site_lat   = rad_to_deg(site.Latitude)  if site else None
    site_lon   = rad_to_deg(site.Longitude) if site else None
    site_tz    = fmt_tz(getattr(site, "TimeZone", None)) if site else ""
    gis_code   = safe_gis_code(site) if site else ""

    pbp_e, pbp_n, pbp_z, pbp_ang = pbp_values(d)
    sp_e,  sp_n,  sp_z           = sp_values(d)

    active_pl = getattr(d, "ActiveProjectLocation", None)
    try:    pls = list(iter_project_locations(d.ProjectLocations))
    except: pls = []
    if not pls and active_pl: pls = [active_pl]

    for pl in pls:
        is_active = "Yes" if (active_pl and pl.Id == active_pl.Id) else "No"
        try:
            pos   = pl.GetProjectPosition(XYZ.Zero)
            pl_ang = norm_deg(rad_to_deg(pos.Angle))
            pl_e   = pos.EastWest   * FOOT_TO_MM
            pl_n   = pos.NorthSouth * FOOT_TO_MM
            pl_z   = pos.Elevation  * FOOT_TO_MM
        except Exception:
            pl_ang = pl_e = pl_n = pl_z = None

        rows.append([
            scope, d.Title, pl.Name, is_active,
            site_place, site_lat, site_lon, site_tz, gis_code,
            pl_ang, pl_e, pl_n, pl_z,
            pbp_e, pbp_n, pbp_z, pbp_ang,
            sp_e,  sp_n,  sp_z,
            EXPORTED_AT
        ])
    return rows

coords_rows.extend(coords_for_doc(doc, "Host"))
seen_docs = set()
for li in links:
    ld = li.GetLinkDocument()
    if ld and ld not in seen_docs:
        seen_docs.add(ld)
        coords_rows.extend(coords_for_doc(ld, "Link"))

# ============================================================
#  WRITE XLSX — three sheets, one file
# ============================================================
print("Writing XLSX...")

sheets_payload = [
    ("Grids",       GRIDS_HEADERS,  grids_rows),
    ("Levels",      LEVELS_HEADERS, levels_rows),
    ("Coordinates", COORDS_HEADERS, coords_rows),
]

ok = build_xlsx_pkg(save_path, sheets_payload)
if ok:
    print("BEP Audit Export complete:")
    print("  {}".format(save_path))
    print("  Sheets: Grids | Levels | Coordinates")
else:
    print("Export failed.")
