from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .certificates import CertificateRecord


RED_FILL = PatternFill(start_color="FFFFCCCC", end_color="FFFFCCCC", fill_type="solid")
DEFAULT_TEMPLATE = Path(__file__).resolve().parents[2] / "assets" / "单证录入标准.xlsx"

CERTIFICATE_TYPES = {
    "01_INTERNATIONAL_REGISTRY_CERT": "1101-船舶国籍证书",
    "02_TONNAGE_CERT": "1201-海上船舶吨位证书",
    "03_LOAD_LINE": "1202-海上船舶载重线证书",
    "04_IOPP_FORM_B": "1205-海上船舶防止油污证书",
    "05_SAFETY_CONSTRUCTION": "2205-货船构造安全证书",
    "06_SAFETY_EQUIPMENT_FORM_E": "2206-货船设备安全证书",
    "07_SAFETY_RADIO_FORM_R": "2207-货船无线电安全证书",
    "08_MINIMUM_MANNING": "1102-船舶最低安全配员证书",
    "09_CLC": "1105-油污损害民事责任保险或其他财务保证证书",
    "10_DOC": "2103-符合证明副本",
    "11_SMC": "1104-安全管理证书",
    "12_ISSC": "2222-国际船舶保安证书",
}


def _clear(sheet, start_row: int, columns: int) -> None:
    for row in range(start_row, max(sheet.max_row, start_row) + 1):
        for column in range(1, columns + 1):
            sheet.cell(row, column).value = None
            sheet.cell(row, column).fill = PatternFill(fill_type=None)


def _write(sheet, row: int, values: list[object], highlight: bool = False) -> None:
    for column, value in enumerate(values, start=1):
        cell = sheet.cell(row, column)
        cell.value = value
        if highlight:
            cell.fill = RED_FILL


def _fill_crew(workbook, crew_rows: list[list[str]]) -> None:
    sheet = workbook["船上非旅客人员清单"]
    _clear(sheet, 3, 16)
    for index, values in enumerate(crew_rows, start=3):
        _write(sheet, index, values, highlight=not values[4])

    items = workbook["船上非旅客人员物品清单"]
    _clear(items, 3, 10)
    for index, crew in enumerate(crew_rows, start=1):
        _write(items, index + 2, [str(index), crew[7], crew[8], "0100", "计算机", 1, "001", "", "", ""])


def _fill_ports(workbook, port_rows: list[list[str]]) -> None:
    activity = workbook["海事船岸活动信息"]
    _clear(activity, 3, 8)
    for index, values in enumerate(port_rows, start=3):
        _write(activity, index, values, highlight=not values[3] or not values[6])

    previous = workbook["前十港信息"]
    _clear(previous, 3, 18)
    for index, values in enumerate(port_rows[:10], start=3):
        row = [values[0], values[1], values[2], "", values[3], values[6], values[4], "", values[7], values[5], "", "", "", "", "", "", "", ""]
        _write(previous, index, row, highlight=not values[3] or not values[6])


def _fill_certificates(workbook, certificates: list[CertificateRecord]) -> None:
    sheet = workbook["船舶证书信息"]
    _clear(sheet, 3, 7)
    ordered = sorted((item for item in certificates if item.category != "13_WRC"), key=lambda item: item.category)
    for index, certificate in enumerate(ordered, start=1):
        cert_type = CERTIFICATE_TYPES.get(certificate.category, certificate.category)
        note_parts = [f"源文件：{Path(certificate.source).name}"]
        if certificate.notes:
            note_parts.append(certificate.notes)
        values = [
            str(index), cert_type, certificate.certificate_number,
            certificate.issue_date, certificate.expiry_date,
            certificate.annual_verification_date, "；".join(note_parts),
        ]
        _write(sheet, index + 2, values)


def build_workbook(
    crew_rows: list[list[str]],
    port_rows: list[list[str]],
    certificates: list[CertificateRecord],
    destination: str | Path,
    template: str | Path = DEFAULT_TEMPLATE,
) -> Path:
    workbook = load_workbook(template)
    _fill_crew(workbook, crew_rows)
    _fill_ports(workbook, port_rows)
    _fill_certificates(workbook, certificates)
    output = Path(destination)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return output
