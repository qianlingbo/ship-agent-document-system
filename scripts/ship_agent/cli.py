from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

from .agent_mail import AgentMailClient, AgentMailError
from .config import Settings
from .pipeline import ShipPipeline
from .run_state import RunStateStore, output_state


SKILL_ROOT = Path(__file__).resolve().parents[2]


def _pipeline(args: argparse.Namespace) -> ShipPipeline:
    settings = Settings.from_env(args.output)
    template = Path(args.template) if args.template else settings.template_path
    crew_info_template = Path(args.crew_info_template) if args.crew_info_template else SKILL_ROOT / "assets" / "船员信息模板.xlsx"
    return ShipPipeline(settings.work_root, template, crew_info_template)


def _common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--output", default="data")
    parser.add_argument("--template", help="船舶计划任务书 DOCX 模板")
    parser.add_argument("--crew-info-template", help="船员信息 XLSX 模板")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="ship-agent", description="船舶代理邮件单证处理系统")
    sub = parser.add_subparsers(dest="command", required=True)
    process = sub.add_parser("process", help="处理 .eml 或本地附件目录")
    process.add_argument("source")
    _common(process)
    process.add_argument("--ship-name")
    watch = sub.add_parser("watch", help="按间隔轮询 Agent Mail 未读附件邮件")
    _common(watch)
    watch.add_argument("--once", action="store_true")
    watch.add_argument("--interval", type=int, help="轮询秒数；默认读取 SHIP_AGENT_POLL_SECONDS")
    watch.add_argument("--mail-cli", help="agently-cli 可执行文件路径")
    mail = sub.add_parser("process-mail", help="处理 Agent Mail 中的一封邮件")
    mail.add_argument("message_id")
    _common(mail)
    mail.add_argument("--mail-cli")
    args = parser.parse_args(argv)
    settings = Settings.from_env(args.output)
    pipeline = _pipeline(args)

    if args.command == "process":
        source = Path(args.source)
        result = pipeline.process_eml(source, args.ship_name) if source.suffix.lower() == ".eml" else pipeline.process_directory(source, args.ship_name)
        print(result)
        return 0

    if args.command == "process-mail":
        runs = RunStateStore(settings.work_root / "运行记录.json")
        run_id = runs.start(args.message_id)
        try:
            result = pipeline.process_agent_mail(AgentMailClient(args.mail_cli), args.message_id)
        except AgentMailError as exc:
            runs.finish(run_id, "failed", error=str(exc))
            parser.error(str(exc))
        except Exception as exc:
            runs.finish(run_id, "failed", error=str(exc))
            raise
        state, detail = output_state(result)
        runs.finish(run_id, state, result, detail)
        print(result)
        return 0

    client = AgentMailClient(args.mail_cli)
    state_path = settings.work_root / "processed_message_ids.json"
    runs = RunStateStore(settings.work_root / "运行记录.json")
    try:
        processed = set(json.loads(state_path.read_text(encoding="utf-8"))) if state_path.exists() else set()
    except (ValueError, OSError):
        processed = set()
    interval = args.interval or settings.poll_seconds
    while True:
        failures = 0
        try:
            messages = client.list_messages(unread_only=True, limit=50)
            for message in messages:
                if message.message_id in processed:
                    continue
                run_id = runs.start(message.message_id, message.subject)
                try:
                    result = pipeline.process_agent_mail(client, message.message_id)
                except AgentMailError as exc:
                    runs.finish(run_id, "failed", error=str(exc))
                    failures += 1
                    if exc.code == 3:
                        raise
                    print(f"FAILED {message.message_id}: {exc}")
                    continue
                except Exception as exc:
                    runs.finish(run_id, "failed", error=str(exc))
                    failures += 1
                    print(f"FAILED {message.message_id}: {exc}")
                    continue
                state, detail = output_state(result)
                runs.finish(run_id, state, result, detail)
                processed.add(message.message_id)
                settings.work_root.mkdir(parents=True, exist_ok=True)
                state_path.write_text(json.dumps(sorted(processed), ensure_ascii=False, indent=2), encoding="utf-8")
                print(result)
        except AgentMailError as exc:
            parser.error(str(exc))
        if args.once:
            return 1 if failures else 0
        time.sleep(interval)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
