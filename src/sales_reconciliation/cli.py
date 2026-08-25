from __future__ import annotations

import argparse
from pathlib import Path

from sales_reconciliation.data_generator import generate_demo_data
from sales_reconciliation.reconcile import (
    export_reconciliation_report,
    read_payments_file,
    read_sales_files,
    reconcile_sales_and_payments,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Genera dati Excel finti e riconcilia vendite e incassi.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    generate_parser = subparsers.add_parser("generate", help="Genera file Excel finti con vendite e incassi.")
    generate_parser.add_argument("--output-dir", type=Path, default=Path("data/raw"), help="Directory di destinazione dei file generati.")

    reconcile_parser = subparsers.add_parser("reconcile", help="Riconcilia file vendite e incassi e salva report Excel/CSV.")
    reconcile_parser.add_argument("--input-dir", type=Path, default=Path("data/raw"), help="Cartella con i file vendite_*.xlsx e incassi.xlsx.")
    reconcile_parser.add_argument("--payments-file", type=Path, default=None, help="Percorso del file incassi.xlsx (opzionale).")
    reconcile_parser.add_argument("--output", type=Path, default=Path("data/output/reconciliation_report.xlsx"), help="File Excel di output.")
    reconcile_parser.add_argument("--csv-output", type=Path, default=None, help="File CSV con i dettagli delle anomalie (opzionale).")

    return parser


def main() -> int:
    args = build_parser().parse_args()

    if args.command == "generate":
        generate_demo_data(args.output_dir)
        print(f"Dati generati in: {args.output_dir}")
        return 0

    input_dir = args.input_dir
    payments_path = args.payments_file or input_dir / "incassi.xlsx"
    sales_df = read_sales_files(input_dir)
    payments_df = read_payments_file(payments_path)
    report = reconcile_sales_and_payments(sales_df, payments_df)
    output_excel, csv_output = export_reconciliation_report(
        report,
        args.output,
        sales_df,
        payments_df,
        csv_path=args.csv_output,
    )

    print(f"Riepilogo: {report}")
    print(f"Excel salvato in: {output_excel}")
    if csv_output is not None:
        print(f"CSV dettagli salvato in: {csv_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
