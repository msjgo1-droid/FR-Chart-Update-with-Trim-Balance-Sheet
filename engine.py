"""
FR TRIM Auto-Fill Engine
------------------------
Takes one or more "Balance Sheet" workbooks (each covering a single style,
sheet named after the style number) and one "FR (Friday Review)" chart
workbook, and fills in the TRIM section of the FR chart for each style
using data pulled from that style's Balance Sheet.

Core rules (confirmed with the user):
- Pull ONLY: Item name, ETD, ETA, Tracking#, Shipped Qty from the Balance
  Sheet's "GENERAL TRIM" + "SPECIAL TRIM..." item table, 1ST ORDER AND SHIP
  DATE columns.
- Blank values ARE brought over for the 1st shipment (nothing is skipped).
- Vendor / Mill-Supplier name is NOT included (kept as condensed as possible).
- Items whose name contains "hangtag" are split into one line PER SIZE
  sub-row (since hangtags/labels differ by size), everything else is one
  line per item block.
- If a style has a 2ND ORDER AND SHIP DATE (or later) for an item, an
  extra line "<item name>-2nd" is added ONLY when ETD/ETA/Tracking#/Shipped
  Qty are ALL filled in for that 2nd shipment (blanks are NOT brought over
  for 2nd+ shipments).
- Destination: the FR chart's TRIM area, which is normally ONE merged cell
  per style, one column wide. This tool unmerges it and turns it into a
  real bordered grid: 5 columns (Item / ETD / ETA / Tracking# / Shipped
  Qty) x N rows (one per trim line). If the style's existing row block is
  shorter than N, extra rows are inserted (with formulas + merges for the
  whole sheet safely adjusted). The 5 columns are created once by
  inserting 4 new columns right after the original TRIM column (whatever
  was there, e.g. "TRIM TEST", is preserved and pushed to the right).
- Header banners are set once: "BULK FB" over the Bulk-Fabric/FB-Test
  columns, "TRIM" over Item/ETD/ETA/Tracking#/Shipped-Qty/Trim-Test, with a
  divider border between them.
"""
import re
import copy
import datetime
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Border, Side, Font, Alignment

CELL_REF_RE = re.compile(r'(\$?)([A-Za-z]{1,3})(\$?)(\d+)')


# --------------------------------------------------------------------------
# generic safe row/col insertion (openpyxl does not fix merges/formulas)
# --------------------------------------------------------------------------

def grow_block_insert_rows(ws, start_row, old_end_row, shortfall):
    """Insert `shortfall` rows right after old_end_row, growing the block
    that starts at start_row and previously ended at old_end_row.
    Any merge that exactly matches (start_row, old_end_row) grows with it;
    merges entirely below the insertion point shift down; formulas are
    patched the same way."""
    insert_at = old_end_row + 1
    merges = [(mc.min_row, mc.min_col, mc.max_row, mc.max_col) for mc in list(ws.merged_cells.ranges)]
    for m in merges:
        ws.unmerge_cells(start_row=m[0], start_column=m[1], end_row=m[2], end_column=m[3])

    formulas = []
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formulas.append((cell.row, cell.column, cell.value))

    ws.insert_rows(insert_at, amount=shortfall)

    def shift_ref_row(rn, belongs_to_block):
        if rn >= insert_at:
            return rn + shortfall
        if belongs_to_block and rn == old_end_row:
            return rn + shortfall
        return rn

    for (r, c, f) in formulas:
        belongs = (r == start_row)
        new_r = r + shortfall if r >= insert_at else r

        def repl(m, belongs=belongs):
            cd, col, rd, row = m.groups()
            return f"{cd}{col}{rd}{shift_ref_row(int(row), belongs)}"

        ws.cell(row=new_r, column=c).value = CELL_REF_RE.sub(repl, f)

    for (min_r, min_c, max_r, max_c) in merges:
        if min_r == start_row and max_r == old_end_row:
            nr = (start_row, min_c, old_end_row + shortfall, max_c)
        elif min_r >= insert_at:
            nr = (min_r + shortfall, min_c, max_r + shortfall, max_c)
        else:
            nr = (min_r, min_c, max_r, max_c)
        ws.merge_cells(start_row=nr[0], start_column=nr[1], end_row=nr[2], end_column=nr[3])


def safe_insert_cols(ws, insert_at_col, amount):
    """insert_cols that also fixes merged-cell column references (openpyxl
    does not do this on its own) and best-effort shifts column widths."""
    merges = [(mc.min_row, mc.min_col, mc.max_row, mc.max_col) for mc in list(ws.merged_cells.ranges)]
    for m in merges:
        ws.unmerge_cells(start_row=m[0], start_column=m[1], end_row=m[2], end_column=m[3])

    ws.insert_cols(insert_at_col, amount=amount)

    for (min_r, min_c, max_r, max_c) in merges:
        if min_c >= insert_at_col:
            nr = (min_r, min_c + amount, max_r, max_c + amount)
        elif max_c >= insert_at_col:
            nr = (min_r, min_c, max_r, max_c + amount)
        else:
            nr = (min_r, min_c, max_r, max_c)
        ws.merge_cells(start_row=nr[0], start_column=nr[1], end_row=nr[2], end_column=nr[3])

    old_dims = {k: v.width for k, v in ws.column_dimensions.items() if v.width}
    for k in list(old_dims.keys()):
        try:
            idx = column_index_from_string(k)
        except Exception:
            continue
        if idx >= insert_at_col + amount:
            ws.column_dimensions[get_column_letter(idx)].width = old_dims[k]


# --------------------------------------------------------------------------
# Balance Sheet parsing
# --------------------------------------------------------------------------

def find_label_cell(ws, text, max_row=90, max_col=50, mode="eq"):
    text_l = text.strip().lower()
    for r in range(1, min(max_row, ws.max_row) + 1):
        for c in range(1, min(max_col, ws.max_column) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            vs = str(v).strip().lower()
            if mode == "eq" and vs == text_l:
                return r, c
            if mode == "startswith" and vs.startswith(text_l):
                return r, c
            if mode == "contains" and text_l in vs:
                return r, c
    return None


def get_style_number(ws):
    title = ws.title.strip()
    if title.lstrip("-").isdigit():
        return int(title) if title.isdigit() else title
    loc = find_label_cell(ws, "STYLE", max_row=5, mode="eq")
    if loc:
        r, c = loc
        v = ws.cell(row=r + 1, column=c).value
        if v is not None:
            return v
    return None


def find_trim_section_layout(ws):
    """Locate the GENERAL TRIM item table and its key columns dynamically."""
    loc = find_label_cell(ws, "GENERAL TRIM", mode="contains")
    if not loc:
        return None
    header_row, _ = loc
    data_start = header_row + 2  # header_row = labels, +1 = sub-labels (ETD/ETA/..), +2 = first data row

    col_item = None
    col_size = None
    col_1st = None
    col_2nd = None
    for c in range(1, min(60, ws.max_column) + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        vs = str(v).strip().lower()
        if col_item is None and vs.startswith("item"):
            col_item = c
        if col_size is None and vs == "size":
            col_size = c
        if col_1st is None and "1st order" in vs:
            col_1st = c
        if col_2nd is None and "2nd order" in vs:
            col_2nd = c

    if col_item is None:
        return None

    layout = dict(
        header_row=header_row,
        data_start=data_start,
        data_end=ws.max_row,
        col_item=col_item,
        col_size=col_size,
        col_etd1=col_1st + 1 if col_1st else None,
        col_eta1=col_1st + 2 if col_1st else None,
        col_track1=col_1st + 3 if col_1st else None,
        col_qty1=col_1st + 4 if col_1st else None,
        col_etd2=col_2nd + 1 if col_2nd else None,
        col_eta2=col_2nd + 2 if col_2nd else None,
        col_track2=col_2nd + 3 if col_2nd else None,
        col_qty2=col_2nd + 4 if col_2nd else None,
    )
    return layout


def _fmt_date(v):
    if isinstance(v, datetime.datetime):
        return f"{v.month}/{v.day}"
    if v is None or str(v).strip() == "":
        return ""
    return str(v)


def _clean(v):
    if v is None:
        return ""
    return str(v).replace("\n", " ").strip()


def _fmt_qty(vals):
    vals = [v for v in vals if v is not None]
    if not vals:
        return ""
    nums = [v for v in vals if isinstance(v, (int, float))]
    others = [v for v in vals if not isinstance(v, (int, float))]
    parts = []
    if nums:
        parts.append(f"{sum(nums):,.0f}")
    for o in others:
        parts.append(str(o))
    return " / ".join(parts) if parts else ""


def _all_filled(*vals):
    return all(v is not None and str(v).strip() != "" for v in vals)


def extract_trim_rows(ws, layout, split_keyword="hangtag"):
    """Returns a list of (item, etd, eta, track, qty) tuples, one per line
    that should end up in the FR trim grid, following the confirmed rules."""

    merges = list(ws.merged_cells.ranges)

    def find_merge(row, col):
        for mc in merges:
            if mc.min_row <= row <= mc.max_row and mc.min_col <= col <= mc.max_col:
                return mc
        return None

    def eff(row, col):
        if col is None:
            return None
        mc = find_merge(row, col)
        if mc:
            return ws.cell(row=mc.min_row, column=mc.min_col).value
        return ws.cell(row=row, column=col).value

    def block_span(row, col):
        mc = find_merge(row, col)
        if mc:
            return mc.min_row, mc.max_row
        return row, row

    col_item = layout["col_item"]
    col_size = layout["col_size"]
    c_etd1, c_eta1, c_track1, c_qty1 = layout["col_etd1"], layout["col_eta1"], layout["col_track1"], layout["col_qty1"]
    c_etd2, c_eta2, c_track2, c_qty2 = layout["col_etd2"], layout["col_eta2"], layout["col_track2"], layout["col_qty2"]

    rows_out = []
    r = layout["data_start"]
    last_row = layout["data_end"]

    while r <= last_row:
        item_val = eff(r, col_item)
        start, end = block_span(r, col_item)
        if item_val is not None and str(item_val).strip() != "":
            item_clean = _clean(item_val)
            split_by_size = split_keyword in item_clean.lower()

            if split_by_size and end > start:
                sub_rows = list(range(start, end + 1))
            else:
                sub_rows = None  # aggregate whole block into one line

            if sub_rows:
                for rr in sub_rows:
                    size = _clean(ws.cell(row=rr, column=col_size).value) if col_size else ""
                    label = f"{item_clean} ({size})" if size else item_clean
                    etd1 = _fmt_date(eff(rr, c_etd1))
                    eta1 = _fmt_date(eff(rr, c_eta1))
                    track1 = _clean(eff(rr, c_track1))
                    qty1 = _fmt_qty([ws.cell(row=rr, column=c_qty1).value]) if c_qty1 else ""
                    rows_out.append((label, etd1, eta1, track1, qty1))

                    if c_etd2:
                        etd2_raw = eff(rr, c_etd2)
                        eta2_raw = eff(rr, c_eta2)
                        track2_raw = eff(rr, c_track2)
                        qty2_raw = ws.cell(row=rr, column=c_qty2).value if c_qty2 else None
                        if _all_filled(etd2_raw, eta2_raw, track2_raw, qty2_raw):
                            rows_out.append((f"{label}-2nd", _fmt_date(etd2_raw), _fmt_date(eta2_raw),
                                              _clean(track2_raw), _fmt_qty([qty2_raw])))
            else:
                etd1 = _fmt_date(eff(start, c_etd1))
                eta1 = _fmt_date(eff(start, c_eta1))
                track1 = _clean(eff(start, c_track1))
                qtys1 = [ws.cell(row=rr, column=c_qty1).value for rr in range(start, end + 1)] if c_qty1 else []
                qty1 = _fmt_qty(qtys1)
                rows_out.append((item_clean, etd1, eta1, track1, qty1))

                if c_etd2:
                    etd2_raw = eff(start, c_etd2)
                    eta2_raw = eff(start, c_eta2)
                    track2_raw = eff(start, c_track2)
                    qtys2 = [ws.cell(row=rr, column=c_qty2).value for rr in range(start, end + 1)] if c_qty2 else []
                    qty2_val = qtys2[0] if qtys2 else None
                    if _all_filled(etd2_raw, eta2_raw, track2_raw, qty2_val):
                        rows_out.append((f"{item_clean}-2nd", _fmt_date(etd2_raw), _fmt_date(eta2_raw),
                                          _clean(track2_raw), _fmt_qty(qtys2)))
        r = end + 1

    return rows_out


# --------------------------------------------------------------------------
# FR chart helpers
# --------------------------------------------------------------------------

def find_fr_layout(ws):
    """Find the STYLE NO column and the (original, single-column) TRIM
    column in the FR chart, and whether the sheet has already been
    expanded into the 5-column grid by a previous run of this tool."""
    style_col = None
    trim_col = None
    for c in range(1, min(60, ws.max_column) + 1):
        v2 = ws.cell(row=2, column=c).value
        if v2 and "style" in str(v2).strip().lower() and style_col is None:
            style_col = c
        v3 = ws.cell(row=3, column=c).value
        if v3 and str(v3).strip().lower() == "trim":
            trim_col = c
    already_expanded = False
    if trim_col:
        v = ws.cell(row=3, column=trim_col).value
        already_expanded = (str(v).strip().lower() == "item")
    return style_col, trim_col, already_expanded


def find_merge_at(ws, row, col):
    for mc in ws.merged_cells.ranges:
        if mc.min_row <= row <= mc.max_row and mc.min_col <= col <= mc.max_col:
            return mc
    return None


def setup_trim_grid_columns(ws, trim_col):
    """One-time transform: insert 4 columns after trim_col for
    ETD/ETA/Tracking#/Shipped-Qty, and (re)build the BULK FB / TRIM
    header banners with a divider border. Returns nothing; mutates ws."""
    fb_fill = copy.copy(ws.cell(row=3, column=trim_col).fill)

    safe_insert_cols(ws, trim_col + 1, 4)

    sub_labels = ["Item", "ETD", "ETA", "Tracking#", "Shipped Qty"]
    thin = Side(style='thin', color='000000')
    for j, label in enumerate(sub_labels):
        cc = ws.cell(row=3, column=trim_col + j)
        cc.value = label
        cc.font = Font(name="Calibri", size=14, bold=True)
        cc.fill = fb_fill
        cc.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cc.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Header banners: BULK FB = the two columns immediately left of the
    # trim grid (Bulk Fabric Schedule + FB Test), TRIM = trim_col through
    # trim_col+5 (Item..Shipped Qty plus whatever was originally right
    # after TRIM, e.g. TRIM TEST, now pushed right by the column insert).
    bulk_start = trim_col - 2
    bulk_end = trim_col - 1
    trim_start = trim_col
    trim_end = trim_col + 5

    for row in (2,):
        for existing in list(ws.merged_cells.ranges):
            if existing.min_row == row and bulk_start <= existing.min_col <= trim_end:
                ws.unmerge_cells(start_row=existing.min_row, start_column=existing.min_col,
                                  end_row=existing.max_row, end_column=existing.max_col)

    if bulk_start >= 1:
        ws.merge_cells(start_row=2, start_column=bulk_start, end_row=2, end_column=bulk_end)
        bulk_cell = ws.cell(row=2, column=bulk_start)
        bulk_cell.value = "BULK FB"
        bulk_cell.font = Font(name="Calibri", size=20, bold=True)
        bulk_cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

    ws.merge_cells(start_row=2, start_column=trim_start, end_row=2, end_column=trim_end)
    trim_cell = ws.cell(row=2, column=trim_start)
    trim_cell.value = "TRIM"
    trim_cell.font = Font(name="Calibri", size=20, bold=True)
    trim_cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

    for cc in range(bulk_start, trim_end + 1):
        cell = ws.cell(row=2, column=cc)
        cell.fill = fb_fill
        left = thin if cc in (bulk_start, trim_start) else None
        right = thin if cc in (bulk_end, trim_end) else None
        cell.border = Border(left=left, right=right, top=thin, bottom=thin)

    widths = {trim_col: 30, trim_col + 1: 10, trim_col + 2: 10, trim_col + 3: 18, trim_col + 4: 12}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def fill_style_trim_grid(ws, style_col, trim_col, style_number, rows_out):
    """Find the row block for `style_number` (freshly, since row positions
    may have shifted from earlier styles being grown) and write rows_out
    into a bordered grid there, growing the block if needed.
    Returns a dict with what happened, or None if the style wasn't found."""
    target_row = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=style_col).value
        if v == style_number or (isinstance(v, (int, float)) and isinstance(style_number, (int, float)) and v == style_number):
            target_row = r
            break
        if v is not None and str(v).strip() == str(style_number).strip():
            target_row = r
            break
    if target_row is None:
        return None

    mc = find_merge_at(ws, target_row, trim_col)
    if mc:
        start_row, end_row = mc.min_row, mc.max_row
    else:
        start_row = end_row = target_row

    available = end_row - start_row + 1
    n_needed = len(rows_out)
    added_rows = 0
    if n_needed > available:
        shortfall = n_needed - available
        grow_block_insert_rows(ws, start_row, end_row, shortfall)
        end_row += shortfall
        added_rows = shortfall

    mc2 = find_merge_at(ws, start_row, trim_col)
    if mc2:
        ws.unmerge_cells(start_row=mc2.min_row, start_column=trim_col, end_row=mc2.max_row, end_column=trim_col)

    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name='Calibri', size=9)
    align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    for i, vals in enumerate(rows_out):
        rr = start_row + i
        for j, v in enumerate(vals):
            c = ws.cell(row=rr, column=trim_col + j)
            c.value = v
            c.border = border
            c.font = font
            c.alignment = align
        ws.row_dimensions[rr].height = 22

    return dict(target_row=start_row, rows_written=n_needed, rows_added=added_rows,
                block_start=start_row, block_end=end_row)


def process(fr_wb, fr_sheet_name, balance_sheets, split_keyword="hangtag", log=None):
    """balance_sheets: list of openpyxl Workbook objects (one per style).
    Mutates fr_wb in place. Returns a list of per-style result dicts."""
    ws = fr_wb[fr_sheet_name]
    style_col, trim_col, already_expanded = find_fr_layout(ws)
    if style_col is None or trim_col is None:
        raise ValueError("FR 차트에서 STYLE NO / TRIM 컬럼을 찾지 못했습니다.")

    if not already_expanded:
        setup_trim_grid_columns(ws, trim_col)

    results = []
    for bs_wb in balance_sheets:
        bs_ws = bs_wb.worksheets[0]
        style_number = get_style_number(bs_ws)
        layout = find_trim_section_layout(bs_ws)
        entry = {"sheet": bs_ws.title, "style": style_number}
        if style_number is None:
            entry["status"] = "style 번호를 찾지 못함"
            results.append(entry)
            continue
        if layout is None:
            entry["status"] = "GENERAL TRIM 섹션을 찾지 못함"
            results.append(entry)
            continue

        rows_out = extract_trim_rows(bs_ws, layout, split_keyword=split_keyword)
        res = fill_style_trim_grid(ws, style_col, trim_col, style_number, rows_out)
        if res is None:
            entry["status"] = f"FR 차트에서 STYLE {style_number} 를 찾지 못함"
            results.append(entry)
            continue
        entry.update(res)
        entry["status"] = "완료"
        entry["lines"] = len(rows_out)
        results.append(entry)

    return results
