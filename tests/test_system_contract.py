from __future__ import annotations

import sys
import tempfile
import unittest
import json
from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from ship_agent.certificates import build_certificate_outputs  # noqa: E402
from ship_agent.classify import classify_document  # noqa: E402
from ship_agent.excel_output import build_workbook  # noqa: E402
from ship_agent.agent_mail import AgentMailClient, AgentMailError  # noqa: E402
from ship_agent.run_state import RunStateStore, output_state  # noqa: E402


def blank_pdf(path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    writer.add_metadata({"/Title": path.stem})
    with path.open("wb") as handle:
        writer.write(handle)


class SystemContractTest(unittest.TestCase):
    def test_mail_client_retries_only_transient_exit_codes(self) -> None:
        transient = CompletedProcess([], 4, '{"ok":false,"error":{"message":"network"}}', "")
        success = CompletedProcess([], 0, '{"ok":true,"data":{}}', "")
        with patch("ship_agent.agent_mail.shutil.which", return_value="/tmp/agently-cli"), patch(
            "ship_agent.agent_mail.subprocess.run", side_effect=[transient, success]
        ) as run, patch("ship_agent.agent_mail.time.sleep"):
            self.assertTrue(AgentMailClient(retry_delays=(0,))._run("+me")["ok"])
            self.assertEqual(run.call_count, 2)

        permanent = CompletedProcess([], 3, '{"ok":false,"error":{"message":"auth"}}', "")
        with patch("ship_agent.agent_mail.shutil.which", return_value="/tmp/agently-cli"), patch(
            "ship_agent.agent_mail.subprocess.run", return_value=permanent
        ) as run:
            with self.assertRaises(AgentMailError) as caught:
                AgentMailClient(retry_delays=(0, 0))._run("+me")
            self.assertEqual(caught.exception.code, 3)
            self.assertEqual(run.call_count, 1)

    def test_run_state_records_manual_review(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            store = RunStateStore(root / "运行记录.json")
            run_id = store.start("message-1", "Vessel documents")
            ship = root / "VESSEL"
            ship.mkdir()
            (ship / "处理清单.json").write_text(
                json.dumps({"errors": ["护照 MRZ 需人工核对"]}, ensure_ascii=False), encoding="utf-8"
            )
            state, detail = output_state(ship)
            store.finish(run_id, state, ship, detail)
            record = json.loads((root / "运行记录.json").read_text(encoding="utf-8"))["runs"][0]
            self.assertEqual(record["state"], "needs_manual_review")
            self.assertIn("MRZ", record["error"])

    def test_iopp_form_a_is_classified(self) -> None:
        self.assertEqual(classify_document("04 IOPP AND FORM A.pdf"), "04_IOPP_FORM_B")

    def test_wrc_is_merged_but_not_written_to_excel(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source"
            maritime = root / "海事"
            source.mkdir()
            names = (
                "01 INTERNATIONAL REGISTRY CERT.pdf",
                "04 IOPP AND FORM A.pdf",
                "05 SAFETY CONSTRUCTION.pdf",
                "06 SAFETY EQUIPMENT AND FORM E.pdf",
                "07 SAFETY RADIO AND FORM R.pdf",
                "09 CLC BUNKER OIL.pdf",
                "13 WRC REMOVAL OF WRECKS.pdf",
            )
            files = []
            for name in names:
                path = source / name
                blank_pdf(path)
                files.append(path)

            records = build_certificate_outputs(files, maritime)
            categories = {record.category for record in records}
            self.assertIn("13_WRC", categories)
            self.assertEqual(len(PdfReader(maritime / "船舶证书.pdf").pages), len(names))
            self.assertTrue((maritime / "危险品证书" / "IOPP及附件 I.O.P.P. AND FORM A OR B.pdf").exists())

            workbook_path = root / "单一窗口录入.xlsx"
            build_workbook([], [], records, workbook_path)
            workbook = load_workbook(workbook_path, data_only=False)
            values = [cell.value for row in workbook["船舶证书信息"].iter_rows() for cell in row]
            self.assertFalse(any("WRC" in str(value).upper() for value in values if value))


if __name__ == "__main__":
    unittest.main()
