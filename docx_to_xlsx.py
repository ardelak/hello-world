"""
docx_to_xlsx.py — Parse tables from a DOCX file and write them to an XLSX file.

Each table in the DOCX is written to its own sheet in the XLSX workbook.
Merged cells in the DOCX are preserved as merged cells in the XLSX.

Usage:
    python docx_to_xlsx.py input.docx output.xlsx

    # or from Python:
    from docx_to_xlsx import convert
    convert("input.docx", "output.xlsx")
"""

import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_cell_span(tc):
    """Return (row_span, col_span) for a table-cell XML element."""
    tcPr = tc.find(qn("w:tcPr"))
    col_span = 1
    row_span = 1
    if tcPr is not None:
        grid_span = tcPr.find(qn("w:gridSpan"))
        if grid_span is not None:
            col_span = int(grid_span.get(qn("w:val"), 1))
        v_merge = tcPr.find(qn("w:vMerge"))
        if v_merge is not None:
            val = v_merge.get(qn("w:val"), "")
            if val == "restart":
                row_span = None   # sentinel: will be computed later
            else:
                row_span = 0      # continuation cell — skip
    return row_span, col_span


def _is_bold(run):
    return bool(run.bold)


def _cell_text(cell):
    return "\n".join(p.text for p in cell.paragraphs)


def _cell_bg_color(tc):
    """Return hex fill color or None."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    fill = shd.get(qn("w:fill"))
    if fill and fill.upper() not in ("", "AUTO", "FFFFFF"):
        return fill.upper()
    return None


def _border_side(border_el, tag):
    """Return an openpyxl Side from a w:tcBorders child element, or None."""
    if border_el is None:
        return None
    el = border_el.find(qn(f"w:{tag}"))
    if el is None:
        return None
    val = el.get(qn("w:val"), "none")
    if val in ("none", "nil"):
        return Side(border_style=None)
    color = el.get(qn("w:color"), "000000")
    if color.upper() == "AUTO":
        color = "000000"
    return Side(border_style="thin", color=color.upper())


def _cell_border(tc):
    """Return an openpyxl Border for the cell, or None."""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    tcBorders = tcPr.find(qn("w:tcBorders"))
    top = _border_side(tcBorders, "top")
    bottom = _border_side(tcBorders, "bottom")
    left = _border_side(tcBorders, "left")
    right = _border_side(tcBorders, "right")
    if any(s is not None for s in (top, bottom, left, right)):
        return Border(top=top, bottom=bottom, left=left, right=right)
    return None


# ---------------------------------------------------------------------------
# Core conversion
# ---------------------------------------------------------------------------

def _write_table(ws, table):
    """Write a single docx Table to an openpyxl Worksheet."""
    rows = table.rows
    n_rows = len(rows)

    # Build a 2-D grid so we can track which (ws_row, ws_col) cells are
    # already occupied by a merge from an earlier cell.
    occupied = {}          # (r, c) -> True
    merge_starts = {}      # (r, c) -> (row_span, col_span)

    # ---- first pass: determine row/col spans --------------------------------
    # We need two passes because vMerge "restart" cells don't know their span
    # until we've seen all continuation rows.

    grid = []  # grid[r][c] = (row_span, col_span, cell_obj, tc_el)
    for r, row in enumerate(rows):
        row_data = []
        for tc in row._tr.findall(qn("w:tc")):
            rs, cs = _get_cell_span(tc)
            row_data.append((rs, cs, tc))
        grid.append(row_data)

    # Resolve vMerge "restart" row spans
    for r in range(n_rows):
        col_cursor = 0
        for idx, (rs, cs, tc) in enumerate(grid[r]):
            if rs is None:  # restart — count continuations below
                span = 1
                for r2 in range(r + 1, n_rows):
                    found = False
                    col_offset = 0
                    for rs2, cs2, tc2 in grid[r2]:
                        if col_offset == col_cursor:
                            found = rs2 == 0  # continuation
                            break
                        col_offset += cs2
                    if found:
                        span += 1
                    else:
                        break
                grid[r][idx] = (span, cs, tc)
            col_cursor += cs

    # ---- second pass: write to worksheet ------------------------------------
    # We map docx (row, col) -> xlsx (row, col) accounting for occupied cells.

    xlsx_row = 1
    for r, row_data in enumerate(grid):
        xlsx_col = 1
        for (rs, cs, tc) in row_data:
            # Skip continuation cells
            if rs == 0:
                continue

            # Advance xlsx_col past any already-occupied cells
            while (xlsx_row, xlsx_col) in occupied:
                xlsx_col += 1

            # Extract text and bold flag directly from XML.
            paragraphs = tc.findall(qn("w:p"))
            lines = []
            is_bold = False
            for p in paragraphs:
                runs = p.findall(qn("w:r"))
                parts = []
                for run in runs:
                    t = run.find(qn("w:t"))
                    if t is not None and t.text:
                        parts.append(t.text)
                        rPr = run.find(qn("w:rPr"))
                        if rPr is not None and rPr.find(qn("w:b")) is not None:
                            is_bold = True
                lines.append("".join(parts))
            text = "\n".join(lines)

            cell = ws.cell(row=xlsx_row, column=xlsx_col, value=text or None)

            # Alignment
            cell.alignment = Alignment(wrap_text=True, vertical="top")

            # Bold
            if is_bold:
                cell.font = Font(bold=True)

            # Background fill
            bg = _cell_bg_color(tc)
            if bg:
                cell.fill = PatternFill(
                    fill_type="solid", fgColor=bg, bgColor=bg
                )

            # Border
            border = _cell_border(tc)
            if border:
                cell.border = border

            # Register merge
            row_end = xlsx_row + rs - 1
            col_end = xlsx_col + cs - 1

            if rs > 1 or cs > 1:
                ws.merge_cells(
                    start_row=xlsx_row,
                    start_column=xlsx_col,
                    end_row=row_end,
                    end_column=col_end,
                )
                for mr in range(xlsx_row, row_end + 1):
                    for mc in range(xlsx_col, col_end + 1):
                        if not (mr == xlsx_row and mc == xlsx_col):
                            occupied[(mr, mc)] = True

            xlsx_col = col_end + 1

        xlsx_row += 1

    # Auto-size columns (approximate)
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                max_len = max(max_len, max(len(ln) for ln in str(cell.value).splitlines()))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 60)


def convert(docx_path: str, xlsx_path: str) -> int:
    """
    Convert all tables in *docx_path* to sheets in *xlsx_path*.

    Returns the number of tables converted.
    """
    doc = Document(docx_path)
    wb = Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    tables = doc.tables
    if not tables:
        # Create a blank sheet so the file is valid
        wb.create_sheet("No Tables")
        wb.save(xlsx_path)
        return 0

    for i, table in enumerate(tables, start=1):
        sheet_name = f"Table {i}"
        ws = wb.create_sheet(title=sheet_name)
        _write_table(ws, table)

    wb.save(xlsx_path)
    return len(tables)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 3:
        print("Usage: python docx_to_xlsx.py <input.docx> <output.xlsx>")
        sys.exit(1)

    input_path, output_path = sys.argv[1], sys.argv[2]

    if not Path(input_path).exists():
        print(f"Error: file not found: {input_path}")
        sys.exit(1)

    count = convert(input_path, output_path)
    print(f"Converted {count} table(s) from '{input_path}' to '{output_path}'.")


if __name__ == "__main__":
    main()
