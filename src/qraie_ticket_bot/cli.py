from __future__ import annotations

import argparse
from pathlib import Path

from qraie_ticket_bot.bot import QRaieBot
from qraie_ticket_bot.config import load_config
from qraie_ticket_bot.excel_io import load_tickets, write_results


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Create QRaie tickets from Excel.")
    ap.add_argument("--config", required=True, help="Path to config.yaml")
    ap.add_argument("--excel", required=True, help="Path to tickets.xlsx")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    tickets = load_tickets(args.excel, sheet_name=cfg.run.sheet_name)
    bot = QRaieBot(cfg)

    results = bot.run(tickets)
    per_row = {k: (v.status, v.error) for k, v in results.items()}

    write_results(
        template_xlsx=Path(args.excel),
        sheet_name=cfg.run.sheet_name,
        results_xlsx=Path(cfg.output.results_xlsx),
        per_row_results=per_row,
    )

    failed = [k for k, v in results.items() if v.status == "FAILED"]
    return 1 if failed else 0

