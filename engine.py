"""
FR TRIM Auto-Fill Engine
------------------------
Takes one or more "Balance Sheet" workbooks (each covering a single style,
sheet named after the style number) and one "FR (Friday Review)" chart
workbook, and fills in the TRIM section of the FR chart for each style
using data pulled from that style's Balance Sheet.

Core rules (confirmed with the user, most recent version):
- Pulls 7 fields per trim line: Item/Color, Qty Needed, Stock, Shipped Qty,
  ETD, ETA, AWB#(Tracking#) - from the Balance Sheet's "GENERAL TRIM" /
  "SPECIAL TRIM..." item table.
- Every column - on BOTH the Balance Sheet side and the FR chart side - is
  found by MEANING (keyword matching on header text), never by a fixed
  column number or row position, because column order and exact wording
  both vary between files (e.g. "Tracking#" vs "AWB#", "Shipped Qty" vs
  "QTY (PCS)", or a completely different column order inside the "1st
  order and ship date" block).
- The FR chart's TRIM group is found by any header cell whose text
  contains "trim" (so "TRIM", "TRIM SCHEDULE", etc. all match). Whatever
  recognizable sub-columns already exist below it are reused as-is
  (never reordered, so no real data already there ever moves); any of the
  7 fields that are missing get a new column appended.
- A line for a shipment leg (1st order, 2nd order) is only added when at
  least one of its ETD/ETA/Tracking#/Shipped Qty values is filled in -
  skipped entirely if all four are blank. Qty Needed/Stock are per-item
  (not per-shipment), so they ride along on whichever lines do get added.
- Items whose name contains "hangtag" are split into one line PER SIZE
  sub-row; everything else is one line per item block (Qty
  Needed/Stock/Shipped Qty summed across the block's rows when needed).
- Destination: the FR chart's TRIM area, normally one merged cell per
  style, one column wide. This tool unmerges it and turns it into a real
  bordered grid: 7 columns x N rows. If the style's existing row block is
  shorter than N, extra rows are inserted (with formulas + merges for the
  whole sheet safely adjusted).
"""
import re
import json
import copy
import datetime
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill

CELL_REF_RE = re.compile(r'(\$?)([A-Za-z]{1,3})(\$?)(\d+)')

# Canonical order for a freshly-built TRIM grid, and the label written for
# each field. "item" is always first (it anchors the whole section); the
# other 6 follow the order the user asked for.
TRIM_FIELD_ORDER = ["item", "needed_qty", "stock", "shipped_qty", "etd", "eta", "tracking"]
TRIM_FIELD_LABELS = {
    "item": "Item",
    "needed_qty": "Qty Needed",
    "stock": "Stock",
    "shipped_qty": "Shipped Qty",
    "etd": "ETD",
    "eta": "ETA",
    "tracking": "AWB#(Tracking#)",
}


def _classify_trim_field(text):
    """Classify a header cell's text into one of our 7 trim-field
    meanings, by keyword/substring - never by exact wording or position,
    since this has to match differently-worded columns across files
    (e.g. 'Tracking#' vs 'AWB#', 'Shipped Qty' vs 'QTY (PCS)', 'Qty
    Needed' vs 'QTY NEEDED'). Returns one of TRIM_FIELD_ORDER, or None."""
    if text is None:
        return None
    t = str(text).strip().lower()
    if not t:
        return None
    # Most specific keywords first, so e.g. "QTY NEEDED" doesn't fall
    # through to the generic "qty" -> shipped_qty catch-all below.
    if "need" in t:
        return "needed_qty"
    if "stock" in t:
        return "stock"
    if "awb" in t or "track" in t:
        return "tracking"
    if "etd" in t:
        return "etd"
    if "eta" in t:
        return "eta"
    if "ship" in t and ("qty" in t or "quantity" in t or "q'ty" in t):
        return "shipped_qty"
    if "item" in t:
        return "item"
    if "qty" in t or "quantity" in t or "q'ty" in t:
        return "shipped_qty"
    return None


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


def insert_blank_row(ws, at_row):
    """Insert one wholly blank row at `at_row`, shifting every row from
    at_row downward by 1 (merges and formula references patched the same
    way grow_block_insert_rows does). Unlike grow_block_insert_rows, this
    is not tied to growing one specific block - a merge that starts at or
    below at_row shifts down whole; a merge that starts above and ends at
    or below at_row grows by 1 so it keeps covering its original rows
    plus the new blank one inserted in the middle of its span."""
    merges = [(mc.min_row, mc.min_col, mc.max_row, mc.max_col) for mc in list(ws.merged_cells.ranges)]
    for m in merges:
        ws.unmerge_cells(start_row=m[0], start_column=m[1], end_row=m[2], end_column=m[3])

    formulas = []
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formulas.append((cell.row, cell.column, cell.value))

    ws.insert_rows(at_row, amount=1)

    def shift_row(rn):
        return rn + 1 if rn >= at_row else rn

    for (r, c, f) in formulas:
        new_r = shift_row(r)

        def repl(m):
            cd, col, rd, row = m.groups()
            return f"{cd}{col}{rd}{shift_row(int(row))}"

        ws.cell(row=new_r, column=c).value = CELL_REF_RE.sub(repl, f)

    for (min_r, min_c, max_r, max_c) in merges:
        if min_r >= at_row:
            nr = (min_r + 1, min_c, max_r + 1, max_c)
        elif max_r >= at_row:
            nr = (min_r, min_c, max_r + 1, max_c)
        else:
            nr = (min_r, min_c, max_r, max_c)
        ws.merge_cells(start_row=nr[0], start_column=nr[1], end_row=nr[2], end_column=nr[3])


def _extend_header_across_new_row(ws, above_row, new_row, exclude_cols):
    """After a blank row has just been inserted at `new_row` (shifting
    everything from there down by one), every OTHER column's own header
    - everything except the trim field columns themselves, which get
    their own labels written separately - would otherwise show a stray
    blank cell where the new row landed, even though most of this sheet
    already treats its header as spanning two rows (e.g. the STYLE NO
    header's own existing 2-row merge). Re-merge whatever was sitting on
    `above_row` (the row right above the new one) down across the new
    row too, for every such column, so the sheet still reads as one
    continuous header instead of a broken one with a gap row."""
    thin = Side(style='thin', color='000000')
    for c in range(1, ws.max_column + 1):
        if c in exclude_cols:
            continue
        mc = find_merge_at(ws, above_row, c)
        if mc:
            if mc.max_row != above_row:
                continue  # already extends past above_row somehow - leave it
            fill = copy.copy(ws.cell(row=mc.min_row, column=mc.min_col).fill)
            ws.unmerge_cells(start_row=mc.min_row, start_column=mc.min_col,
                              end_row=mc.max_row, end_column=mc.max_col)
            ws.merge_cells(start_row=mc.min_row, start_column=mc.min_col,
                            end_row=new_row, end_column=mc.max_col)
            for cc in range(mc.min_col, mc.max_col + 1):
                ws.cell(row=new_row, column=cc).fill = fill
            continue
        v = ws.cell(row=above_row, column=c).value
        if v is not None and str(v).strip() != "":
            fill = copy.copy(ws.cell(row=above_row, column=c).fill)
            ws.merge_cells(start_row=above_row, start_column=c, end_row=new_row, end_column=c)
            ws.cell(row=new_row, column=c).fill = fill


def write_trim_banner_label(ws, banner_row, field_cols):
    """Writes the 7 field names directly into the trim banner cell itself
    (e.g. "TRIM SCHEDULE"), as a second line below the banner's own text,
    inside the SAME merged cell - rather than a separate row underneath
    for just the labels. A separate row meant every OTHER column's own
    header had to be stretched to cover it too, which is exactly the
    "messy, extra blank-looking row" the chart doesn't need: with no
    extra row ever inserted here, the whole header area - trim columns
    and every other column alike - stays exactly the height it always
    was, and the very first data row (for whichever style starts here,
    including any of its own already-existing hand-typed trim content)
    sits right below the banner, undisturbed.
    By this point ensure_trim_fields has already widened the banner's
    own merge to span the full, final field_cols range, so this only
    changes what text that merge shows."""
    grid_cols = sorted(field_cols.values())
    first_col, last_col = grid_cols[0], grid_cols[-1]
    mc = find_merge_at(ws, banner_row, first_col)
    anchor_row = mc.min_row if mc else banner_row
    anchor_col = mc.min_col if mc else first_col
    cell = ws.cell(row=anchor_row, column=anchor_col)
    existing = str(cell.value).split("\n")[0].strip() if cell.value else "TRIM SCHEDULE"
    label_line = " | ".join(TRIM_FIELD_LABELS[f] for f in TRIM_FIELD_ORDER)
    cell.value = f"{existing}\n{label_line}"
    cell.font = Font(name="Calibri", size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin = Side(style='thin', color='000000')
    for cc in range(first_col, last_col + 1):
        c2 = ws.cell(row=anchor_row, column=cc)
        c2.border = Border(left=thin if cc == first_col else None,
                            right=thin if cc == last_col else None,
                            top=thin, bottom=thin)

    current_h = ws.row_dimensions[anchor_row].height
    if current_h is None or current_h < 55:
        ws.row_dimensions[anchor_row].height = 55


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


def _scan_order_block(ws, sub_header_row, start_col, end_col_exclusive):
    """Classify each column in [start_col, end_col_exclusive) by meaning,
    for one '1st/2nd order and ship date' block. Column order inside the
    block can differ between Balance Sheet files, so each column is read
    independently rather than assumed to sit at a fixed offset from the
    block's own start column."""
    cols = {}
    for c in range(start_col, end_col_exclusive):
        v = ws.cell(row=sub_header_row, column=c).value
        field = _classify_trim_field(v)
        if field in ("etd", "eta", "tracking", "shipped_qty") and field not in cols:
            cols[field] = c
    return cols


def find_trim_section_layout(ws):
    """Locate the GENERAL TRIM item table and its key columns dynamically,
    by meaning rather than fixed position - this has to work across
    Balance Sheets whose "1st/2nd order and ship date" sub-columns are in
    a different order, and which may or may not have "QTY NEEDED" /
    "STOCK" columns at all."""
    loc = find_label_cell(ws, "GENERAL TRIM", mode="contains")
    if not loc:
        return None
    header_row, _ = loc
    data_start = header_row + 2  # header_row = labels, +1 = sub-labels (ETD/ETA/..), +2 = first data row

    col_item = None
    col_size = None
    col_1st = None
    col_2nd = None
    needed_qty_col = None
    stock_col = None
    col_mill = None
    col_code = None
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
        # Qty Needed / Stock are per-item, not per-shipment-leg, so they
        # live alongside the item/size columns rather than inside the
        # 1st/2nd order blocks - found anywhere on this same header row.
        if needed_qty_col is None and "need" in vs:
            needed_qty_col = c
        if stock_col is None and "stock" in vs:
            stock_col = c
        # Mill/Supplier name and the material supplier code, used to
        # annotate each item's name in the FR chart, e.g.
        # "Zipper(YKK,C-29192)" - found by meaning, since "code" alone
        # reliably picks the supplier-code column (e.g. "Material
        # suppier code" [sic]) apart from an unrelated ID column, and
        # "mill" picks the supplier-name column (e.g. "MILL/ SUPPLIER").
        if col_mill is None and "mill" in vs:
            col_mill = c
        if col_code is None and "code" in vs:
            col_code = c

    if col_item is None:
        return None

    sub_header_row = header_row + 1
    order1 = {}
    order2 = {}
    if col_1st:
        end_excl = col_2nd if col_2nd else min(col_1st + 8, ws.max_column + 1)
        order1 = _scan_order_block(ws, sub_header_row, col_1st, end_excl)
    if col_2nd:
        end_excl = min(col_2nd + 8, ws.max_column + 1)
        for c in range(col_2nd + 1, min(60, ws.max_column) + 1):
            v = ws.cell(row=header_row, column=c).value
            if v and ("order" in str(v).strip().lower()):
                end_excl = c
                break
        order2 = _scan_order_block(ws, sub_header_row, col_2nd, end_excl)

    layout = dict(
        header_row=header_row,
        data_start=data_start,
        data_end=ws.max_row,
        col_item=col_item,
        col_size=col_size,
        needed_qty_col=needed_qty_col,
        stock_col=stock_col,
        col_mill=col_mill,
        col_code=col_code,
        order1=order1,
        order2=order2,
    )
    return layout


_EXCEL_ERRORS = {"#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!"}


def _fmt_date(v):
    """Only ever return a clean month/day for a genuine date value - a
    Balance Sheet's ETD/ETA cell can have stray text typed into it by
    mistake (a tracking note, "IN FAC", a broken formula's #REF! error),
    and none of that is a date, so it's dropped rather than passed
    through as-is."""
    if isinstance(v, datetime.datetime):
        return f"{v.month}/{v.day}"
    if isinstance(v, datetime.date):
        return f"{v.month}/{v.day}"
    return ""


def _clean(v):
    if v is None:
        return ""
    s = str(v).replace("\n", " ").strip()
    if s in _EXCEL_ERRORS:
        return ""
    return s


def _fmt_qty(vals):
    """Sum only genuinely numeric values. A broken formula can leave a
    quantity cell holding an error string (#REF!, #VALUE!, ...) instead
    of a number - that and any other non-numeric stray text is dropped
    entirely rather than carried into the output, since a quantity
    field should only ever show a number."""
    nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not nums:
        return ""
    return f"{sum(nums):,.0f}"


def _has_value(v):
    return v is not None and str(v).strip() != ""


def _any_filled(*vals):
    """True if at least one of the given raw values is non-blank. Column
    positions are never assumed here - callers always pass the values found
    via the dynamically-detected ETD/ETA/Tracking#/Shipped-Qty columns for
    that shipment leg, since column numbers differ between Balance Sheet
    files."""
    return any(_has_value(v) for v in vals)


def extract_trim_rows(ws, layout, split_keyword="hangtag"):
    """Returns a list of dicts (one per TRIM_FIELD_ORDER field) for every
    line that should end up in the FR trim grid, following the confirmed
    rules. Qty Needed / Stock are per-item, so the same value rides along
    on both the 1st-order and 2nd-order line for that item when both
    exist; they never gate whether a line is included (only
    ETD/ETA/Tracking#/Shipped Qty do, per-leg)."""

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
    needed_qty_col = layout["needed_qty_col"]
    stock_col = layout["stock_col"]
    col_mill = layout.get("col_mill")
    col_code = layout.get("col_code")
    order1 = layout["order1"]
    order2 = layout["order2"]

    # Every distinct item TYPE gets one sequential number (matching the
    # "1. CARE / 2. ZIPPER / ..." style already used by hand in these
    # charts) - same type, same number, even when it produces several
    # output lines (a Hangtag split into one line per size, or an item
    # with both a 1st- and 2nd-order leg). A different item (even a
    # closely related one, like "Polybag" vs "Polybag UPC Sticker")
    # always gets its own number.
    type_numbers = {}
    next_number = [1]

    def label_for(item_clean, start_row):
        if item_clean not in type_numbers:
            type_numbers[item_clean] = next_number[0]
            next_number[0] += 1
        mill = _clean(eff(start_row, col_mill))
        code = _clean(eff(start_row, col_code))
        if mill and code:
            suffix = f"({mill},{code})"
        elif mill or code:
            suffix = f"({mill or code})"
        else:
            suffix = ""
        return f"{type_numbers[item_clean]}. {item_clean}{suffix}"

    def leg_values_single_row(rr, order):
        etd_raw = eff(rr, order.get("etd"))
        eta_raw = eff(rr, order.get("eta"))
        track_raw = eff(rr, order.get("tracking"))
        qty_raw = ws.cell(row=rr, column=order["shipped_qty"]).value if order.get("shipped_qty") else None
        return etd_raw, eta_raw, track_raw, qty_raw

    def leg_values_block(start, end, order):
        etd_raw = eff(start, order.get("etd"))
        eta_raw = eff(start, order.get("eta"))
        track_raw = eff(start, order.get("tracking"))
        qtys_raw = [ws.cell(row=rr, column=order["shipped_qty"]).value
                    for rr in range(start, end + 1)] if order.get("shipped_qty") else []
        return etd_raw, eta_raw, track_raw, qtys_raw

    rows_out = []
    r = layout["data_start"]
    last_row = layout["data_end"]

    while r <= last_row:
        item_val = eff(r, col_item)
        start, end = block_span(r, col_item)
        if item_val is not None and str(item_val).strip() != "":
            item_clean = _clean(item_val)
            # label_for() is only ever called once a leg is confirmed to
            # actually produce an output line (inside each _any_filled
            # check below) - never up front - so an item type that ends
            # up producing nothing (both legs empty) never consumes a
            # number, and the visible numbering never gets gaps in it.
            split_by_size = split_keyword in item_clean.lower()
            sub_rows = list(range(start, end + 1)) if (split_by_size and end > start) else None

            if sub_rows:
                for rr in sub_rows:
                    size = _clean(ws.cell(row=rr, column=col_size).value) if col_size else ""
                    needed_raw = ws.cell(row=rr, column=needed_qty_col).value if needed_qty_col else None
                    stock_raw = ws.cell(row=rr, column=stock_col).value if stock_col else None
                    needed_fmt = _fmt_qty([needed_raw])
                    stock_fmt = _fmt_qty([stock_raw])

                    etd1, eta1, track1, qty1 = leg_values_single_row(rr, order1)
                    if _any_filled(etd1, eta1, track1, qty1):
                        base = label_for(item_clean, start)
                        label = f"{base} ({size})" if size else base
                        rows_out.append(dict(
                            item=label, needed_qty=needed_fmt, stock=stock_fmt,
                            shipped_qty=_fmt_qty([qty1]), etd=_fmt_date(etd1),
                            eta=_fmt_date(eta1), tracking=_clean(track1),
                        ))

                    if order2:
                        etd2, eta2, track2, qty2 = leg_values_single_row(rr, order2)
                        if _any_filled(etd2, eta2, track2, qty2):
                            base = label_for(item_clean, start)
                            label = f"{base} ({size})" if size else base
                            rows_out.append(dict(
                                item=f"{label}-2nd", needed_qty=needed_fmt, stock=stock_fmt,
                                shipped_qty=_fmt_qty([qty2]), etd=_fmt_date(etd2),
                                eta=_fmt_date(eta2), tracking=_clean(track2),
                            ))
            else:
                needed_raws = [ws.cell(row=rr, column=needed_qty_col).value
                               for rr in range(start, end + 1)] if needed_qty_col else []
                stock_raws = [ws.cell(row=rr, column=stock_col).value
                              for rr in range(start, end + 1)] if stock_col else []
                needed_fmt = _fmt_qty(needed_raws)
                stock_fmt = _fmt_qty(stock_raws)

                etd1, eta1, track1, qtys1 = leg_values_block(start, end, order1)
                if _any_filled(etd1, eta1, track1, *qtys1):
                    rows_out.append(dict(
                        item=label_for(item_clean, start), needed_qty=needed_fmt, stock=stock_fmt,
                        shipped_qty=_fmt_qty(qtys1), etd=_fmt_date(etd1),
                        eta=_fmt_date(eta1), tracking=_clean(track1),
                    ))

                if order2:
                    etd2, eta2, track2, qtys2 = leg_values_block(start, end, order2)
                    if _any_filled(etd2, eta2, track2, *qtys2):
                        rows_out.append(dict(
                            item=f"{label_for(item_clean, start)}-2nd", needed_qty=needed_fmt, stock=stock_fmt,
                            shipped_qty=_fmt_qty(qtys2), etd=_fmt_date(etd2),
                            eta=_fmt_date(eta2), tracking=_clean(track2),
                        ))
        r = end + 1

    return rows_out


# --------------------------------------------------------------------------
# FR chart helpers
# --------------------------------------------------------------------------

def find_merge_at(ws, row, col):
    for mc in ws.merged_cells.ranges:
        if mc.min_row <= row <= mc.max_row and mc.min_col <= col <= mc.max_col:
            return mc
    return None


def find_fr_trim_anchor(ws, max_row=6, max_col=100):
    """Find the (row, col) of the header cell that anchors the real TRIM
    data-grid group - any cell whose text contains "trim"
    (case-insensitive), e.g. "TRIM", "TRIM SCHEDULE". A chart can have more
    than one cell mentioning "trim" (an unrelated "...TRIM DUE" deadline
    column, a wide combined "BULK FB/TRIM SCHEDULE" banner, and the real,
    narrower "TRIM SCHEDULE" group header can all exist on the same
    sheet), so every match in rows 1..max_row is collected and scored:
    the one whose own row's span IMMEDIATELY BELOW already contains
    recognizable sub-columns (item/qty/etd/eta/tracking/...) wins, since
    that is the real grid; ties (e.g. on a brand-new file where no
    sub-columns exist yet) fall back to the widest matching cell/merge,
    since the real group header spans multiple columns while an unrelated
    single-column deadline label does not."""
    candidates = []
    for r in range(1, max_row + 1):
        for c in range(1, min(max_col, ws.max_column) + 1):
            v = ws.cell(row=r, column=c).value
            if v and "trim" in str(v).strip().lower():
                candidates.append((r, c))
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    def score(cand):
        r, c = cand
        mc = find_merge_at(ws, r, c)
        if mc:
            c_start, c_end = mc.min_col, mc.max_col
        else:
            c_start, c_end = c, c
        sub_row = r + 1
        field_count = 0
        for cc in range(c_start, c_end + 1):
            if _classify_trim_field(ws.cell(row=sub_row, column=cc).value):
                field_count += 1
        width = c_end - c_start + 1
        return (field_count, width, r, -c)

    return max(candidates, key=score)


def find_fr_layout(ws, meta_ws=None):
    """Find the STYLE NO column and the TRIM section anchor in the FR
    chart - "trim" appearing anywhere in a header cell's text, not a fixed
    word or row - and map out whichever of our 7 trim fields
    (item/needed_qty/stock/shipped_qty/etd/eta/tracking) already exist as
    real sub-columns there. Works whether the file has never been
    expanded, was expanded by an earlier version of this tool, or already
    has its own pre-built trim columns under different wording (e.g.
    "AWB#" instead of "Tracking#", "QTY (PCS)" instead of "Shipped Qty")
    - every column is recognized by meaning, never a fixed position or
    exact label text.
    Returns (style_col, trim_col, header_row, label_row, already_expanded,
    field_cols). `header_row` is the row the "trim" banner text itself
    sits on; `label_row` is the row right after that banner's own merge
    (its merge may span one row or several, e.g. a two-row-tall header
    banner like the sheet's other column headers) - this is where the 7
    field sub-labels either already sit (already_expanded) or should be
    written (fresh grid), and is deliberately NOT always header_row + 1,
    since that would land back inside a multi-row banner merge instead of
    the row right after it.

    `meta_ws` (optional) is this workbook's hidden `_frtrim_meta` sheet.
    The 7 field labels now live only as a second line inside the banner
    cell itself (write_trim_banner_label), not as their own per-column
    header text, so a plain text-scan can no longer reliably tell an
    already-expanded grid apart from a brand-new one on a file that has
    no other coincidental header-like text sitting in that row. When the
    text-scan below comes back empty, this falls back to whatever
    field_cols layout this tool last saved for this exact sheet name in
    meta_ws (process() saves it after every run) - but only as a
    fallback, so a genuine fresh file (nothing saved yet) still correctly
    falls through to building a brand-new grid.
    """
    style_col = None
    for c in range(1, min(60, ws.max_column) + 1):
        v2 = ws.cell(row=2, column=c).value
        if v2 and "style" in str(v2).strip().lower():
            style_col = c
            break

    anchor = find_fr_trim_anchor(ws)
    if anchor is None:
        return style_col, None, None, None, False, {}

    header_row, header_col = anchor
    mc = find_merge_at(ws, header_row, header_col)
    trim_col = mc.min_col if mc else header_col
    label_row = (mc.max_row + 1) if mc else (header_row + 1)
    max_scan_col = min(trim_col + 30, ws.max_column + 10)

    def scan_fields(row):
        cols = {}
        c = trim_col
        blanks_in_a_row = 0
        while c <= max_scan_col:
            v = ws.cell(row=row, column=c).value
            field = _classify_trim_field(v)
            if field and field not in cols:
                cols[field] = c
                blanks_in_a_row = 0
            elif v is None or str(v).strip() == "":
                blanks_in_a_row += 1
                if blanks_in_a_row >= 3:
                    break
            else:
                blanks_in_a_row = 0
            c += 1
        return cols

    field_cols = scan_fields(label_row)
    # The row right after the banner's own merge is usually where the
    # sub-column labels sit, but a prior run of this tool may have
    # inserted an extra blank row between the banner and the labels
    # (ensure_dedicated_header_row, below) without the banner's own merge
    # growing to match - so if nothing classifies there, check the next
    # couple of rows too and use whichever actually has the labels,
    # instead of trusting the arithmetic blindly.
    if not field_cols:
        for extra in (1, 2):
            candidate = scan_fields(label_row + extra)
            if candidate:
                label_row += extra
                field_cols = candidate
                break

    if not field_cols and meta_ws is not None:
        cached = load_fr_layout_from_meta(meta_ws, ws.title)
        if cached:
            field_cols = cached

    already_expanded = "item" in field_cols
    return style_col, trim_col, header_row, label_row, already_expanded, field_cols


def setup_trim_grid_fresh(ws, trim_col, banner_row, label_row, style_col):
    """Brand-new TRIM section (never expanded by this tool, and no
    pre-existing hand-built sub-columns either): insert 6 columns after
    trim_col for the other 6 fields, building the full 7-column grid in
    canonical order, and (re)build the BULK FB / TRIM header banners with
    a divider border. `banner_row` is the row the original single "TRIM"
    banner text sits on (rebuilt here as two bits, "BULK FB" + "TRIM").
    No separate row is ever used for the field sub-labels - the caller
    writes those into this same banner cell afterward
    (write_trim_banner_label) - so data for whichever style starts right
    after the banner (at `label_row`, including any of its own
    already-existing hand-typed content) is never disturbed here.
    Returns field_cols: {field: column}."""
    fb_fill = copy.copy(ws.cell(row=banner_row, column=trim_col).fill)

    n_other = len(TRIM_FIELD_ORDER) - 1  # 6
    safe_insert_cols(ws, trim_col + 1, n_other)

    field_cols = {}
    for j, field in enumerate(TRIM_FIELD_ORDER):
        field_cols[field] = trim_col + j

    thin = Side(style='thin', color='000000')
    grid_last_col = trim_col + n_other   # last of our 7 columns
    trim_end = grid_last_col + 1         # whatever followed TRIM originally (e.g. "Trim Test")

    bulk_start = trim_col - 2
    bulk_end = trim_col - 1
    trim_start = trim_col

    for existing in list(ws.merged_cells.ranges):
        if existing.min_row == banner_row and bulk_start <= existing.min_col <= trim_end:
            ws.unmerge_cells(start_row=existing.min_row, start_column=existing.min_col,
                              end_row=existing.max_row, end_column=existing.max_col)

    if bulk_start >= 1:
        ws.merge_cells(start_row=banner_row, start_column=bulk_start, end_row=banner_row, end_column=bulk_end)
        bulk_cell = ws.cell(row=banner_row, column=bulk_start)
        bulk_cell.value = "BULK FB"
        bulk_cell.font = Font(name="Calibri", size=20, bold=True)
        bulk_cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

    ws.merge_cells(start_row=banner_row, start_column=trim_start, end_row=banner_row, end_column=trim_end)
    trim_cell = ws.cell(row=banner_row, column=trim_start)
    trim_cell.value = "TRIM"
    trim_cell.font = Font(name="Calibri", size=20, bold=True)
    trim_cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

    for cc in range(bulk_start, trim_end + 1):
        cell = ws.cell(row=banner_row, column=cc)
        cell.fill = fb_fill
        left = thin if cc in (bulk_start, trim_start) else None
        right = thin if cc in (bulk_end, trim_end) else None
        cell.border = Border(left=left, right=right, top=thin, bottom=thin)

    widths = {
        field_cols["item"]: 30,
        field_cols["needed_qty"]: 10,
        field_cols["stock"]: 10,
        field_cols["shipped_qty"]: 12,
        field_cols["etd"]: 10,
        field_cols["eta"]: 10,
        field_cols["tracking"]: 18,
    }
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    return field_cols


def ensure_trim_fields(ws, banner_row, label_row, trim_col, field_cols):
    """Make sure every one of our 7 trim fields has a real column,
    inserting any that are missing (e.g. Qty Needed / Stock on a file
    that already has Item/Tracking/Shipped Qty/ETD/ETA from an earlier
    template or tool run). Existing fields keep their current column at
    this point - reorder_trim_fields (called separately, after this)
    puts everything in the canonical order; a missing field is appended
    right after the current right-most known field purely so there is
    somewhere to put it before that reorder happens. `label_row` is the
    row this style data would otherwise start on (find_fr_layout's
    label_row) - no label text is written there any more (that all
    lives in the banner cell itself now, via write_trim_banner_label),
    this is only used to size the newly-inserted columns' width. The
    banner directly above (e.g. "TRIM SCHEDULE") is widened to cover
    the full, final 7-column grid, since its original width may only
    have covered the fields that already existed before any were
    appended. Returns the completed field_cols (all 7 present)."""
    field_cols = dict(field_cols)
    thin = Side(style='thin', color='000000')

    for field in TRIM_FIELD_ORDER:
        if field in field_cols:
            continue
        insert_at = max(field_cols.values()) + 1
        safe_insert_cols(ws, insert_at, 1)
        for k in list(field_cols.keys()):
            if field_cols[k] >= insert_at:
                field_cols[k] += 1
        field_cols[field] = insert_at
        ws.column_dimensions[get_column_letter(insert_at)].width = 12

    grid_cols = sorted(field_cols.values())
    first_col, last_col = grid_cols[0], grid_cols[-1]
    bmc = find_merge_at(ws, banner_row, trim_col)
    if bmc and bmc.max_col < last_col:
        fill = copy.copy(ws.cell(row=bmc.min_row, column=bmc.min_col).fill)
        value = ws.cell(row=bmc.min_row, column=bmc.min_col).value
        font = copy.copy(ws.cell(row=bmc.min_row, column=bmc.min_col).font)
        alignment = copy.copy(ws.cell(row=bmc.min_row, column=bmc.min_col).alignment)
        ws.unmerge_cells(start_row=bmc.min_row, start_column=bmc.min_col,
                          end_row=bmc.max_row, end_column=bmc.max_col)
        ws.merge_cells(start_row=bmc.min_row, start_column=first_col,
                        end_row=bmc.max_row, end_column=last_col)
        anchor = ws.cell(row=bmc.min_row, column=first_col)
        anchor.value = value
        anchor.font = font
        anchor.alignment = alignment
        for cc in range(first_col, last_col + 1):
            cell = ws.cell(row=bmc.min_row, column=cc)
            cell.fill = fill
            left = thin if cc == first_col else None
            right = thin if cc == last_col else None
            cell.border = Border(left=left, right=right, top=thin, bottom=thin)

    return field_cols


def reorder_trim_fields(ws, field_cols, label_row):
    """Put the 7 trim field columns in the canonical order the user
    asked for (TRIM_FIELD_ORDER), regardless of what order they're
    already sitting in - a column that pre-existed under a different
    order (e.g. this tool's own earlier appended-at-the-end columns, or
    a legacy template) gets physically moved, not just left where it
    was. field_cols is always a contiguous block of 7 real columns by
    this point, so this is a pure left-to-right permutation within that
    same span: every row from `label_row` down to the sheet's last row
    (every style's data, plus the label row itself) is read at the OLD
    per-field column and rewritten at the NEW one. Any merge touching
    this column band first has to go, since a reorder can't coexist
    with cells merged across the columns being swapped - fill_style_
    trim_grid and clean_and_divide_style_blocks both already re-merge
    (or leave unmerged) whatever each row needs on their own next pass,
    so this never leaves anything to clean up afterwards. Returns the
    new field_cols (same columns, remapped to the right fields)."""
    grid_cols = sorted(field_cols.values())
    first_col, last_col = grid_cols[0], grid_cols[-1]
    target = {field: first_col + i for i, field in enumerate(TRIM_FIELD_ORDER)}
    if field_cols == target:
        return field_cols

    last_row = ws.max_row

    for mc in list(ws.merged_cells.ranges):
        if mc.max_row < label_row:
            continue
        if mc.max_col < first_col or mc.min_col > last_col:
            continue
        ws.unmerge_cells(start_row=mc.min_row, start_column=mc.min_col,
                          end_row=mc.max_row, end_column=mc.max_col)

    snapshot = {}
    for row in range(label_row, last_row + 1):
        for field, col in field_cols.items():
            c = ws.cell(row=row, column=col)
            snapshot[(row, field)] = (c.value, copy.copy(c.font), copy.copy(c.fill),
                                       copy.copy(c.border), copy.copy(c.alignment))

    for row in range(label_row, last_row + 1):
        for field, new_col in target.items():
            value, font, fill, border, alignment = snapshot[(row, field)]
            c = ws.cell(row=row, column=new_col)
            c.value = value
            c.font = font
            c.fill = fill
            c.border = border
            c.alignment = alignment

    old_widths = {field: ws.column_dimensions[get_column_letter(col)].width
                  for field, col in field_cols.items()}
    for field, new_col in target.items():
        w = old_widths.get(field)
        if w:
            ws.column_dimensions[get_column_letter(new_col)].width = w

    return target


def enumerate_style_blocks(ws, style_col):
    """Return [(style_value, start_row, end_row), ...] for every style block
    in the FR sheet, using the style_col merges to find each block's row
    span. Data is assumed to start right after the header block that sits
    on top of style_col (found dynamically, never hardcoded)."""
    header_mc = find_merge_at(ws, 2, style_col) or find_merge_at(ws, 1, style_col)
    data_start = (header_mc.max_row + 1) if header_mc else 4

    blocks = []
    r = data_start
    max_row = ws.max_row
    while r <= max_row:
        v = ws.cell(row=r, column=style_col).value
        if v is None or str(v).strip() == "":
            r += 1
            continue
        mc = find_merge_at(ws, r, style_col)
        start, end = (mc.min_row, mc.max_row) if mc else (r, r)
        blocks.append((v, start, end))
        r = end + 1
    return blocks


def clean_and_divide_style_blocks(ws, style_col, field_cols):
    """Runs on every process() call (including re-runs on an already
    once-expanded file, so it also retroactively cleans up a live file the
    user re-uploads), after all Balance Sheets for this run have been
    filled in. For every style block in the sheet:
      - Any BLANK cell to the right of this tool's own TRIM grid (two
        columns past the grid's last field, to leave the "Trim Test"-style
        column alone - derived from the grid's own columns, never a
        hardcoded column letter, since that area's width varies by file)
        has its border cleared and a plain white fill applied, so Excel's
        default gridlines stop showing through it.
      - Every row in the sheet's data range has its row-level default
        border cleared too (Excel renders this across the WHOLE row out
        to its last column when someone selects entire rows and applies a
        border, independent of any actual cell - invisible to, and
        unreachable by, per-cell iteration no matter how far right it
        scans). Any cell with its own explicit style - everything real
        data or this tool ever wrote - keeps that style regardless, since
        a cell's own style always overrides the row default.
      - A style block whose Trim grid is still a single, untouched
        one-column merge (i.e. this tool hasn't filled it from a Balance
        Sheet) has that merge widened across the full grid width and its
        borders cleared too, so it renders as one clean blank cell instead
        of a messy leftover grid.
      - A style block that HAS been filled (fill_style_trim_grid already
        unmerged it and wrote a clean bordered grid) is left as-is.
      - Every block, filled or not, gets a permanent vertical divider line
        right after its own grid, on every row - plus a bottom-border
        divider at its last row spanning from the grid's first column to
        the sheet's last column, so consecutive styles stay easy to tell
        apart.
    """
    no_side = Side(style=None)
    no_border = Border(left=no_side, right=no_side, top=no_side, bottom=no_side)
    thin = Side(style='thin', color='000000')
    white_fill = PatternFill(fill_type="solid", start_color="FFFFFFFF", end_color="FFFFFFFF")

    grid_cols = sorted(field_cols.values())
    trim_col = grid_cols[0]
    grid_last_col = grid_cols[-1]
    outer_start = grid_last_col + 2   # +1 = whatever followed TRIM originally (left alone), +2 = beyond that
    last_col = ws.max_column

    blocks = enumerate_style_blocks(ws, style_col)

    if blocks:
        data_start = blocks[0][1]
        data_end = ws.max_row

        no_row_border = Border(left=no_side, right=no_side, top=no_side, bottom=no_side)
        for rr in range(data_start, data_end + 1):
            ws.row_dimensions[rr].border = no_row_border

        if outer_start <= last_col:
            for rr in range(data_start, data_end + 1):
                for cc in range(outer_start, last_col + 1):
                    cell = ws.cell(row=rr, column=cc)
                    if cell.value is None or str(cell.value).strip() == "":
                        cell.border = no_border
                        cell.fill = white_fill

    for style_value, start, end in blocks:
        mc = find_merge_at(ws, start, trim_col)
        untouched = mc is not None and mc.min_row == start and mc.max_row == end \
            and mc.min_col == trim_col and mc.max_col == trim_col

        if untouched:
            ws.unmerge_cells(start_row=mc.min_row, start_column=mc.min_col,
                              end_row=mc.max_row, end_column=mc.max_col)
            # Any other merge overlapping the target rectangle must be
            # removed first - openpyxl does not check for overlaps, so
            # merging over one would silently create two overlapping
            # merged ranges and corrupt the file.
            for other in list(ws.merged_cells.ranges):
                if other.max_row < start or other.min_row > end \
                        or other.max_col < trim_col or other.min_col > grid_last_col:
                    continue
                ws.unmerge_cells(start_row=other.min_row, start_column=other.min_col,
                                  end_row=other.max_row, end_column=other.max_col)
            ws.merge_cells(start_row=start, start_column=trim_col, end_row=end, end_column=grid_last_col)
            for rr in range(start, end + 1):
                for cc in range(trim_col, grid_last_col + 1):
                    ws.cell(row=rr, column=cc).border = no_border

        # Permanent divider right after our own grid, on every row of the
        # block (independent of whatever the original file's own
        # formatting was there).
        for rr in range(start, end + 1):
            cell = ws.cell(row=rr, column=grid_last_col)
            b = cell.border
            cell.border = Border(left=b.left, top=b.top, bottom=b.bottom, right=thin)

        # style-boundary divider, across the grid + whatever is further right
        for cc in range(trim_col, last_col + 1):
            cell = ws.cell(row=end, column=cc)
            existing = cell.border
            cell.border = Border(left=existing.left, right=existing.right,
                                  top=existing.top, bottom=thin)


def fill_style_trim_grid(ws, style_col, field_cols, style_number, rows_out, preserve_existing=False, prior_structured_count=0, prior_freeform_count=0):
    """Find the row block for `style_number` (freshly, since row positions
    may have shifted from earlier styles being grown) and write rows_out
    (a list of field->value dicts) into a bordered grid there, writing
    each field into its own mapped column (field_cols), growing the block
    if needed. Returns a dict with what happened, or None if the style
    wasn't found.

    `preserve_existing`: when True, whatever is already sitting in the
    field_cols columns for this block (e.g. hand-typed trim notes from
    before this tool ever touched this style) is captured first:
    structured rows to the TOP, a freeform note to the BOTTOM (merged),
    instead of being deleted. When False (the normal case once this
    tool has already claimed the style), only this run's NEW-DATA rows
    are replaced, so repeated re-runs don't stack up forever - but
    `prior_structured_count` / `prior_freeform_count`, if given, still
    carry forward those rows from their own correct ends of the block,
    so neither is ever lost on a later, non-first run."""
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

    # The style's true row span always comes from the STYLE column's own
    # merge, never from the TRIM columns' - their merge state varies (a
    # single untouched column, a widened-but-empty grid left by the
    # cleanup pass, or no merge at all once this tool has already written
    # real data there on an earlier run), so it can't be trusted to
    # reflect the block's actual height, especially when re-processing a
    # style that was already filled before.
    style_mc = find_merge_at(ws, target_row, style_col)
    if style_mc:
        start_row, end_row = style_mc.min_row, style_mc.max_row
    else:
        start_row = end_row = target_row

    # Pre-existing content here isn't always 7 separate per-field values
    # - it can be a single note merged across several rows AND several
    # of our field columns at once (one freeform comment, not
    # structured data), or it can be genuine structured trim data
    # someone already imported another way (several distinct field
    # values per row). The two are kept apart from the moment they're
    # found: structured rows go to the TOP of the grid (real item data,
    # exactly where it originally was), freeform ones to the BOTTOM,
    # merged into one wide cell (out of the way until a real Balance
    # Sheet update for that item comes in).
    structured_preserved = []
    freeform_preserved = []

    def _classify_and_add(vals):
        filled = [f for f, v in vals.items() if v is not None and str(v).strip() != ""]
        (freeform_preserved if len(filled) <= 1 else structured_preserved).append(vals)

    if preserve_existing:
        # Reading each field "merge-aware" would otherwise repeat a wide
        # merged note's text into every field column it happens to
        # cover; only the first field column that hits a given merge
        # keeps its text, the rest stay blank for that row - and once a
        # multi-row merge is captured, skip straight past the rows it
        # already covers instead of re-capturing them as blank duplicates.
        rr = start_row
        while rr <= end_row:
            vals = {}
            any_val = False
            seen_merges = set()
            next_row = rr + 1
            for field, col in field_cols.items():
                mc = find_merge_at(ws, rr, col)
                if mc:
                    key = (mc.min_row, mc.min_col, mc.max_row, mc.max_col)
                    v = None if key in seen_merges else ws.cell(row=mc.min_row, column=mc.min_col).value
                    seen_merges.add(key)
                    next_row = max(next_row, mc.max_row + 1)
                else:
                    v = ws.cell(row=rr, column=col).value
                if v is not None and str(v).strip() != "":
                    any_val = True
                vals[field] = v
            if any_val:
                _classify_and_add(vals)
            rr = next_row
    elif prior_structured_count > 0 or prior_freeform_count > 0:
        # Not this style's first run, but an earlier run already moved
        # some pre-existing content into place: structured rows at the
        # very TOP of the block, a freeform note at the very BOTTOM -
        # exactly where every run (including this one) puts them. That
        # content must keep being carried forward on every later run
        # too, not just captured once and then left for this run's new
        # data and clearing pass to overwrite - so read each group back
        # from its own correct end of the block (as it stands right
        # now), not as one combined guess at "the trailing N rows".
        for rr in range(start_row, start_row + prior_structured_count):
            structured_preserved.append({field: ws.cell(row=rr, column=col).value
                                          for field, col in field_cols.items()})
        for rr in range(end_row - prior_freeform_count + 1, end_row + 1):
            freeform_preserved.append({field: ws.cell(row=rr, column=col).value
                                        for field, col in field_cols.items()})

    preserved_rows = structured_preserved + freeform_preserved
    n_needed = len(rows_out)
    total_needed = n_needed + len(preserved_rows)
    available = end_row - start_row + 1
    added_rows = 0
    if total_needed > available:
        shortfall = total_needed - available
        grow_block_insert_rows(ws, start_row, end_row, shortfall)
        end_row += shortfall
        added_rows = shortfall

    grid_cols = sorted(field_cols.values())
    grid_first_col, grid_last_col = grid_cols[0], grid_cols[-1]

    # Clear out whatever currently occupies the trim grid columns for this
    # block: unmerge anything overlapping that area (an untouched single
    # column, or an already-widened empty grid from the cleanup pass).
    # Values were already captured above (preserved_rows), so this is
    # safe even when preserving - only the merge geometry is touched.
    for other in list(ws.merged_cells.ranges):
        if other.max_row < start_row or other.min_row > end_row \
                or other.max_col < grid_first_col or other.min_col > grid_last_col:
            continue
        ws.unmerge_cells(start_row=other.min_row, start_column=other.min_col,
                          end_row=other.max_row, end_column=other.max_col)

    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name='Calibri', size=20)
    align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # structured_preserved/freeform_preserved were already classified
    # above (either from a fresh first-run capture or carried forward
    # from their own correct ends of the block) - structured goes back
    # to the TOP (real item data, exactly where it originally was,
    # ahead of this run's freshly extracted rows), freeform stays at
    # the BOTTOM, merged into one wide cell.
    tagged_rows = ([("structured", v) for v in structured_preserved]
                   + [("new", v) for v in rows_out]
                   + [("freeform", v) for v in freeform_preserved])
    for i, (kind, vals) in enumerate(tagged_rows):
        rr = start_row + i
        if kind == "freeform":
            for other in list(ws.merged_cells.ranges):
                if other.min_row <= rr <= other.max_row and other.min_col <= grid_last_col and other.max_col >= grid_first_col:
                    ws.unmerge_cells(start_row=other.min_row, start_column=other.min_col,
                                      end_row=other.max_row, end_column=other.max_col)
            ws.merge_cells(start_row=rr, start_column=grid_first_col, end_row=rr, end_column=grid_last_col)
            filled = [f for f, v in vals.items() if v is not None and str(v).strip() != ""]
            c = ws.cell(row=rr, column=grid_first_col)
            c.value = vals.get(filled[0], "") if filled else ""
            c.border = border
            c.font = font
            c.alignment = align
        else:
            for field, col in field_cols.items():
                c = ws.cell(row=rr, column=col)
                c.value = vals.get(field, "")
                c.border = border
                c.font = font
                c.alignment = align
        ws.row_dimensions[rr].height = 50

    # Clear any leftover rows below the newly written data within this
    # block (e.g. an earlier run left more trim lines than this run needs).
    no_side = Side(style=None)
    no_border = Border(left=no_side, right=no_side, top=no_side, bottom=no_side)
    for rr in range(start_row + total_needed, end_row + 1):
        for col in field_cols.values():
            cell = ws.cell(row=rr, column=col)
            cell.value = None
            cell.border = no_border

    return dict(target_row=start_row, rows_written=n_needed, rows_preserved=len(preserved_rows),
                rows_structured_preserved=len(structured_preserved),
                rows_freeform_preserved=len(freeform_preserved),
                rows_added=added_rows, block_start=start_row, block_end=end_row)


META_SHEET_NAME = "_frtrim_meta"


def get_or_create_meta_sheet(wb):
    """A hidden helper sheet used only to remember which style numbers
    this tool has already claimed (i.e. already written its own clean
    bordered grid for) in this FR workbook - never meant to be seen or
    edited by the user. Column A holds one claimed style number per
    row; column B holds how many genuine structured legacy rows (kept
    at the TOP of the grid, one field per column) were captured the
    first time this style was claimed, column C how many freeform
    single-note rows (kept at the BOTTOM, merged into one wide cell)
    were. This is what tells a style's very first run (where any
    pre-existing hand-typed trim content must be captured and
    preserved, not deleted) apart from a routine re-run (where this
    tool's own previous NEW-data rows are simply replaced with the
    latest Balance Sheet numbers - but both of those preserved groups
    must keep being carried forward, from their own correct ends of
    the block, on every run from here on, never just the first, or
    they would be silently overwritten the very next time this tool
    runs)."""
    if META_SHEET_NAME in wb.sheetnames:
        return wb[META_SHEET_NAME]
    meta = wb.create_sheet(META_SHEET_NAME)
    try:
        meta.sheet_state = "veryHidden"
    except Exception:
        meta.sheet_state = "hidden"
    return meta


def is_style_claimed(meta_ws, style_number):
    key = str(style_number).strip()
    for row in meta_ws.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if v is not None and str(v).strip() == key:
            return True
    return False


def get_claimed_preserved_counts(meta_ws, style_number):
    """Returns (structured_count, freeform_count): how many rows at the
    TOP of this style's grid are carried-forward structured legacy data,
    and how many at the BOTTOM are a carried-forward freeform note.
    (0, 0) if the style was never claimed, or had nothing to preserve."""
    key = str(style_number).strip()
    for row in meta_ws.iter_rows(min_col=1, max_col=3):
        v = row[0].value
        if v is not None and str(v).strip() == key:
            b = row[1].value if len(row) > 1 else None
            c = row[2].value if len(row) > 2 else None
            return (int(b) if isinstance(b, (int, float)) else 0,
                    int(c) if isinstance(c, (int, float)) else 0)
    return (0, 0)


def mark_style_claimed(meta_ws, style_number, structured_count=0, freeform_count=0):
    key = str(style_number).strip()
    for row in meta_ws.iter_rows(min_col=1, max_col=3):
        v = row[0].value
        if v is not None and str(v).strip() == key:
            ws = row[0].parent
            ws.cell(row=row[0].row, column=2, value=structured_count)
            ws.cell(row=row[0].row, column=3, value=freeform_count)
            return
    r = 1
    while meta_ws.cell(row=r, column=1).value is not None:
        r += 1
    meta_ws.cell(row=r, column=1, value=key)
    meta_ws.cell(row=r, column=2, value=structured_count)
    meta_ws.cell(row=r, column=3, value=freeform_count)


def _fr_layout_meta_key(sheet_name):
    return f"__layout__:{sheet_name}"


def save_fr_layout_to_meta(meta_ws, sheet_name, field_cols):
    """Persists the FR trim-grid's field_cols mapping (which real column
    holds each of the 7 fields) for this exact sheet name, so a later
    re-scan of the same file (find_fr_layout) can recover it even when
    there is no more per-column label text anywhere on the sheet to scan
    for - the 7 labels now live only inside the banner cell itself. Keyed
    by sheet name in the same column A that style numbers use (via a
    "__layout__:" prefix no real style number can ever match), with the
    mapping itself JSON-encoded into column B so it never collides with
    the per-style structured/freeform counts other rows keep there."""
    key = _fr_layout_meta_key(sheet_name)
    payload = json.dumps(field_cols)
    for row in meta_ws.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if v is not None and str(v).strip() == key:
            meta_ws.cell(row=row[0].row, column=2, value=payload)
            return
    r = 1
    while meta_ws.cell(row=r, column=1).value is not None:
        r += 1
    meta_ws.cell(row=r, column=1, value=key)
    meta_ws.cell(row=r, column=2, value=payload)


def load_fr_layout_from_meta(meta_ws, sheet_name):
    """Returns the field_cols dict saved by save_fr_layout_to_meta for
    this sheet name, or {} if nothing has ever been saved for it (a
    genuinely first-time file)."""
    key = _fr_layout_meta_key(sheet_name)
    for row in meta_ws.iter_rows(min_col=1, max_col=2):
        v = row[0].value
        if v is not None and str(v).strip() == key:
            raw = row[1].value if len(row) > 1 else None
            if not raw:
                return {}
            try:
                data = json.loads(raw)
            except (TypeError, ValueError):
                return {}
            return {k: int(v2) for k, v2 in data.items()}
    return {}


class EngineError(Exception):
    """Raised with a status code (not a human message) so the caller (the
    Streamlit app) can localize it."""
    pass


def process(fr_wb, fr_sheet_name, balance_sheets, split_keyword="hangtag", log=None):
    """balance_sheets: list of openpyxl Workbook objects (one per style).
    Mutates fr_wb in place. Returns a list of per-style result dicts.
    `status` is always a machine-readable code (see STATUS_* constants),
    never a human-language string, so the UI layer can localize it:
      - "ok"
      - "no_style"                 - could not find style number in Balance Sheet
      - "no_trim_section"          - could not find GENERAL TRIM section
      - "style_not_found_in_fr"    - style not present in the FR chart
    """
    ws = fr_wb[fr_sheet_name]
    meta_ws = get_or_create_meta_sheet(fr_wb)
    style_col, trim_col, banner_row, label_row, already_expanded, field_cols = find_fr_layout(ws, meta_ws)
    if style_col is None or trim_col is None:
        raise EngineError("fr_columns_not_found")

    if not already_expanded:
        # Always builds in canonical order, so no reorder is needed after.
        field_cols = setup_trim_grid_fresh(ws, trim_col, banner_row, label_row, style_col)
    else:
        field_cols = ensure_trim_fields(ws, banner_row, label_row, trim_col, field_cols)
        # Whatever order the fields happened to already be in (a legacy
        # template, or this tool's own earlier appended-at-the-end
        # columns) gets physically moved into the canonical order here.
        field_cols = reorder_trim_fields(ws, field_cols, label_row)

    # The 7 field names live as a second line inside the banner cell
    # itself (e.g. "TRIM SCHEDULE") - never a separate row underneath,
    # so no extra row is ever inserted anywhere on the sheet for this,
    # and the very first data row for whichever style starts right
    # after the banner (including any of its own already-existing
    # hand-typed content) is never disturbed.
    write_trim_banner_label(ws, banner_row, field_cols)

    # Save this exact layout so a later re-scan of this same file (a second
    # run) can recover it reliably even if the text-scan in find_fr_layout
    # finds nothing to classify - see save_fr_layout_to_meta / the
    # already_expanded-detection note on find_fr_layout above.
    save_fr_layout_to_meta(meta_ws, ws.title, field_cols)

    results = []
    for bs_wb in balance_sheets:
        bs_ws = bs_wb.worksheets[0]
        style_number = get_style_number(bs_ws)
        layout = find_trim_section_layout(bs_ws)
        entry = {"sheet": bs_ws.title, "style": style_number}
        if style_number is None:
            entry["status"] = "no_style"
            results.append(entry)
            continue
        if layout is None:
            entry["status"] = "no_trim_section"
            results.append(entry)
            continue

        rows_out = extract_trim_rows(bs_ws, layout, split_keyword=split_keyword)
        # The very first time this tool ever claims a given style's trim
        # grid, anything already sitting there (hand-typed notes from
        # before) is captured and preserved below the new rows rather
        # than deleted; once claimed, later re-runs go back to a plain
        # replace for the NEW-DATA rows so repeated processing doesn't
        # stack up old output forever - but whatever got preserved on
        # that first run keeps being carried forward on every run after
        # it too, so it is never silently lost on a later re-run.
        first_time = not is_style_claimed(meta_ws, style_number)
        prior_structured, prior_freeform = (0, 0) if first_time else get_claimed_preserved_counts(meta_ws, style_number)
        res = fill_style_trim_grid(ws, style_col, field_cols, style_number, rows_out,
                                    preserve_existing=first_time,
                                    prior_structured_count=prior_structured,
                                    prior_freeform_count=prior_freeform)
        if res is None:
            entry["status"] = "style_not_found_in_fr"
            results.append(entry)
            continue
        mark_style_claimed(meta_ws, style_number,
                            structured_count=res.get("rows_structured_preserved", 0),
                            freeform_count=res.get("rows_freeform_preserved", 0))
        entry.update(res)
        entry["status"] = "ok"
        entry["lines"] = len(rows_out)
        entry["preserved_existing"] = first_time and res.get("rows_preserved", 0) > 0
        results.append(entry)

    # Runs every time (even on an already-expanded file), so it also
    # retroactively cleans up styles left over from earlier runs.
    clean_and_divide_style_blocks(ws, style_col, field_cols)

    # The "messy from column X onward" look the user sees is not actual
    # cell formatting at all - it is Excel's own default gridline display
    # over the sheet's genuinely empty area beyond its real used range (no
    # cell there to style, so per-cell border/fill clearing can never
    # reach it, however far it's scanned). Turning off gridlines for the
    # whole sheet removes that clutter everywhere in one shot, and leaves
    # every explicitly-formatted cell (which is the entire rest of this
    # heavily-styled chart) looking exactly the same, since explicit
    # borders/fills render independently of this view setting.
    ws.sheet_view.showGridLines = False

    return results
