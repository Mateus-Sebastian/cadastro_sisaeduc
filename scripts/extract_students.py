from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from sisaeduc_pipeline import CSV_COLUMNS, collect_docx_files, parse_student_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extrai fichas de matrícula .docx e gera um CSV consolidado."
    )
    parser.add_argument(
        "--input-dir",
        default="matriculas",
        help="Pasta raiz com os .docx de matrícula.",
    )
    parser.add_argument(
        "--output",
        default="output/csv/students.csv",
        help="Arquivo CSV de saída.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_path = Path(args.output).resolve()

    if not input_dir.exists():
        parser.error(f"Pasta não encontrada: {input_dir}")

    docx_files = collect_docx_files(input_dir)
    if not docx_files:
        parser.error(f"Nenhum .docx elegível encontrado em {input_dir}")

    rows = []
    failures: list[tuple[Path, str]] = []
    for path in docx_files:
        try:
            rows.append(parse_student_docx(path, base_dir=input_dir.parent))
        except Exception as exc:  # noqa: BLE001
            failures.append((path, str(exc)))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)

    print(f"CSV gerado em: {output_path}")
    print(f"Documentos processados: {len(rows)}")
    if failures:
        print(f"Avisos: {len(failures)} arquivo(s) não foram extraídos.", file=sys.stderr)
        for path, message in failures:
            print(f"- {path}: {message}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
