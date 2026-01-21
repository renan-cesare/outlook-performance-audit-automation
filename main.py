import argparse
from datetime import datetime
from pathlib import Path

from src.performance_audit.config import Config
from src.performance_audit.logging_utils import Logger
from src.performance_audit.dispatch import dispatch
from src.performance_audit.followup import followup


def build_parser():
    parser = argparse.ArgumentParser(
        prog="outlook-performance-audit-automation",
        description="Automação de envio e acompanhamento de auditorias de desempenho via Outlook + Excel (sanitizado)."
    )
    parser.add_argument(
        "--config",
        default="config.example.json",
        help="Caminho do arquivo de configuração (JSON)."
    )

    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("dispatch", help="Envia auditorias em massa e registra histórico.")
    sub.add_parser("followup", help="Verifica respostas e atualiza status no histórico.")

    return parser


def main():
    args = build_parser().parse_args()
    cfg = Config.load(args.config)

    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    logger = Logger(str(log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.txt"))

    if args.cmd == "dispatch":
        dispatch(cfg, logger)
    elif args.cmd == "followup":
        followup(cfg, logger)


if __name__ == "__main__":
    main()

