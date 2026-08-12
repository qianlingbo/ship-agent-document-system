from __future__ import annotations

from pathlib import Path
from typing import Mapping

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def _append_to_label(table, label: str, value: str) -> bool:
    if not value:
        return False
    for row in table.rows:
        cells = []
        seen = set()
        for cell in row.cells:
            marker = id(cell._tc)
            if marker not in seen:
                seen.add(marker)
                cells.append(cell)
        for index, cell in enumerate(cells):
            text = " ".join(cell.text.split())
            if text == label or text.startswith(label):
                target = cells[index + 1] if index + 1 < len(cells) and not cells[index + 1].text.strip() else cell
                paragraph = target.paragraphs[0]
                run = paragraph.add_run(str(value) if target is not cell else f" {value}")
                run.font.size = Pt(7)
                return True
    return False


def _check(table, row_index: int, keyword: str, checked: bool) -> None:
    if not checked:
        return
    for cell in table.rows[row_index].cells:
        if keyword in cell.text and "□" in cell.text:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.text = run.text.replace("□", "☑")


def build_taskbook(template: str | Path, destination: str | Path, fields: Mapping[str, object], completed: set[str] | None = None) -> Path:
    document = Document(str(template))
    if document.tables:
        table = document.tables[0]
        for label, value in fields.items():
            _append_to_label(table, label, str(value or ""))
        completed = completed or set()
        # These are the checklist rows in the supplied template.
        _check(table, 7, "海事", "海事" in completed)
        _check(table, 11, "海事危险品申报", "危险品" in completed)
        _check(table, 7, "信息和", "单证" in completed)
        _check(table, 7, "PASS", "PASS" in completed)
    if len(document.paragraphs) > 1 and fields.get("调度"):
        document.paragraphs[1].add_run(f" {fields['调度']}")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "PingFang SC")
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "PingFang SC")
    output = Path(destination)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    return output
