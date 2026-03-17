"""Tests for docx_to_xlsx.py"""

import os
import tempfile

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import load_workbook

from docx_to_xlsx import convert


def _make_simple_docx(path):
    doc = Document()
    table = doc.add_table(rows=3, cols=3)
    data = [
        ["Name", "Age", "City"],
        ["Alice", "30", "New York"],
        ["Bob", "25", "London"],
    ]
    for r, row_data in enumerate(data):
        for c, val in enumerate(row_data):
            table.cell(r, c).text = val
    doc.save(path)


def _make_merged_docx(path):
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).merge(table.cell(0, 1))  # horizontal merge
    table.cell(0, 0).text = "Merged Header"
    table.cell(0, 2).text = "Col C"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "B"
    table.cell(1, 2).text = "C"
    doc.save(path)


def _make_no_table_docx(path):
    doc = Document()
    doc.add_paragraph("No tables here.")
    doc.save(path)


def _make_multi_table_docx(path):
    doc = Document()
    for t in range(3):
        tbl = doc.add_table(rows=2, cols=2)
        for r in range(2):
            for c in range(2):
                tbl.cell(r, c).text = f"T{t+1}R{r+1}C{c+1}"
        doc.add_paragraph("")
    doc.save(path)


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_simple_table():
    with tempfile.TemporaryDirectory() as tmp:
        docx = os.path.join(tmp, "simple.docx")
        xlsx = os.path.join(tmp, "simple.xlsx")
        _make_simple_docx(docx)
        count = convert(docx, xlsx)
        assert count == 1

        wb = load_workbook(xlsx)
        ws = wb["Table 1"]
        assert ws["A1"].value == "Name"
        assert ws["B1"].value == "Age"
        assert ws["C1"].value == "City"
        assert ws["A2"].value == "Alice"
        assert ws["B3"].value == "25"
        assert ws["C3"].value == "London"
    print("test_simple_table PASSED")


def test_merged_cells():
    with tempfile.TemporaryDirectory() as tmp:
        docx = os.path.join(tmp, "merged.docx")
        xlsx = os.path.join(tmp, "merged.xlsx")
        _make_merged_docx(docx)
        count = convert(docx, xlsx)
        assert count == 1

        wb = load_workbook(xlsx)
        ws = wb["Table 1"]
        # A1:B1 should be merged
        merged = {str(r) for r in ws.merged_cells.ranges}
        assert any("A1" in m for m in merged), f"Expected A1 in merged ranges, got {merged}"
        assert ws["A1"].value == "Merged Header"
        assert ws["C1"].value == "Col C"
    print("test_merged_cells PASSED")


def test_no_tables():
    with tempfile.TemporaryDirectory() as tmp:
        docx = os.path.join(tmp, "notables.docx")
        xlsx = os.path.join(tmp, "notables.xlsx")
        _make_no_table_docx(docx)
        count = convert(docx, xlsx)
        assert count == 0
        wb = load_workbook(xlsx)
        assert "No Tables" in wb.sheetnames
    print("test_no_tables PASSED")


def test_multiple_tables():
    with tempfile.TemporaryDirectory() as tmp:
        docx = os.path.join(tmp, "multi.docx")
        xlsx = os.path.join(tmp, "multi.xlsx")
        _make_multi_table_docx(docx)
        count = convert(docx, xlsx)
        assert count == 3

        wb = load_workbook(xlsx)
        assert set(wb.sheetnames) == {"Table 1", "Table 2", "Table 3"}
        assert wb["Table 1"]["A1"].value == "T1R1C1"
        assert wb["Table 3"]["B2"].value == "T3R2C2"
    print("test_multiple_tables PASSED")


if __name__ == "__main__":
    test_simple_table()
    test_merged_cells()
    test_no_tables()
    test_multiple_tables()
    print("\nAll tests passed.")
