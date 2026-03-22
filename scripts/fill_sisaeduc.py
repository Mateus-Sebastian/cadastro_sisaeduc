from __future__ import annotations

import argparse
import csv
import json
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Preenche o formulário do Sisaeduc via pyautogui a partir de um CSV."
    )
    parser.add_argument("--csv", default="output/csv/students.csv", help="CSV de entrada.")
    parser.add_argument(
        "--config",
        default="config/sisaeduc_tab_order.json",
        help="Configuração da ordem dos campos por Tab.",
    )
    parser.add_argument("--start-row", type=int, default=1, help="Linha inicial do CSV.")
    parser.add_argument("--limit", type=int, default=0, help="Quantidade máxima de linhas.")
    parser.add_argument("--dry-run", action="store_true", help="Mostra o plano sem tocar no navegador.")
    parser.add_argument("--delay", type=float, default=0.15, help="Pausa padrão entre ações.")
    parser.add_argument(
        "--start-delay",
        type=float,
        default=0.0,
        help="Segundos de espera antes de começar a preencher cada linha.",
    )
    parser.add_argument(
        "--resume-delay",
        type=float,
        default=None,
        help="Segundos de espera antes do próximo aluno após sua confirmação manual. Se omitido, reutiliza --start-delay.",
    )
    parser.add_argument(
        "--review-seconds",
        type=float,
        default=0.0,
        help="Segundos para manter a tela preenchida para revisão antes de encerrar a linha sem salvar.",
    )
    parser.add_argument(
        "--auto-status",
        default="",
        help="Se informado, evita prompts interativos e registra esse status no log ao final da linha.",
    )
    parser.add_argument(
        "--log-file",
        default="logs/sisaeduc_fill_log.jsonl",
        help="Arquivo de log JSONL.",
    )
    parser.add_argument(
        "--start-position-file",
        default="config/first_field_position.json",
        help="Arquivo JSON com a posição salva do primeiro campo.",
    )
    parser.add_argument(
        "--save-start-position",
        action="store_true",
        help="Salva a posição atual do mouse como primeiro campo e encerra.",
    )
    parser.add_argument(
        "--no-click-start-position",
        action="store_true",
        help="Não clica automaticamente na posição salva antes de preencher.",
    )
    parser.add_argument(
        "--resume-position-file",
        default="config/first_field_position_resume.json",
        help="Arquivo JSON com a posição salva do primeiro campo quando houver mensagem de sucesso.",
    )
    parser.add_argument(
        "--save-resume-position",
        action="store_true",
        help="Salva a posição atual do mouse como primeiro campo para os alunos seguintes e encerra.",
    )
    return parser


def load_rows(csv_path: Path) -> list[dict[str, str]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def load_config(config_path: Path) -> list[dict[str, Any]]:
    with config_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    fields = payload.get("fields", []) if isinstance(payload, dict) else payload
    if not isinstance(fields, list):
        raise ValueError("Configuração inválida: o campo 'fields' deve ser uma lista.")
    for entry in fields:
        if "field_id" not in entry or "action_type" not in entry:
            raise ValueError(f"Entrada inválida na configuração: {entry}")
    return fields


def load_start_position(position_path: Path) -> dict[str, int] | None:
    if not position_path.exists():
        return None
    with position_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    x = payload.get("x")
    y = payload.get("y")
    if not isinstance(x, int) or not isinstance(y, int):
        raise ValueError(f"Posição inicial inválida: {position_path}")
    return {"x": x, "y": y}


def save_start_position(position_path: Path, x: int, y: int) -> None:
    position_path.parent.mkdir(parents=True, exist_ok=True)
    with position_path.open("w", encoding="utf-8") as handle:
        json.dump({"x": x, "y": y}, handle, ensure_ascii=False, indent=2)


def normalize_key(value: str) -> str:
    return " ".join((value or "").strip().lower().split())


def apply_conversions(value: str, conversions: list[str]) -> str:
    current = value or ""
    for conversion in conversions:
        if conversion == "strip":
            current = current.strip()
        elif conversion == "upper":
            current = current.upper()
        elif conversion == "lower":
            current = current.lower()
        elif conversion == "title":
            current = current.title()
        elif conversion in {"digits", "numbers_only", "phone_digits"}:
            current = "".join(ch for ch in current if ch.isdigit())
        elif conversion == "blank_if_placeholder":
            if current.strip() in {"", "_", "__", "___", "____"}:
                current = ""
        else:
            raise ValueError(f"Conversão não suportada: {conversion}")
    return current


def resolve_value(row: dict[str, str], entry: dict[str, Any]) -> str:
    column = entry.get("csv_column", "")
    value = row.get(column, "") if column else ""
    value = apply_conversions(value, entry.get("conversions", []))
    if not value and "default" in entry:
        value = str(entry["default"])
    return value


def append_log(log_path: Path, payload: dict[str, Any]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False) + "\n")


def copy_to_clipboard(value: str) -> None:
    try:
        import pyperclip  # type: ignore

        pyperclip.copy(value)
        return
    except Exception:  # noqa: BLE001
        pass

    try:
        command = ["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value @'\n" + value + "\n'@"]
        subprocess.run(command, check=True, capture_output=True, text=True)
        return
    except Exception:  # noqa: BLE001
        pass

    raise RuntimeError("Não foi possível copiar para a área de transferência. Instale pyperclip.")


def get_pyautogui():
    try:
        import pyautogui  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pyautogui não está instalado neste Python. Instale com 'pip install pyautogui pyperclip'."
        ) from exc

    pyautogui.FAILSAFE = True
    return pyautogui


def click_start_position(pyautogui: Any, position: dict[str, int], delay: float) -> None:
    pyautogui.press("home")
    time.sleep(delay)
    pyautogui.scroll(4000)
    time.sleep(delay)
    pyautogui.click(position["x"], position["y"])
    time.sleep(delay)


def clear_current_field(pyautogui: Any, delay: float) -> None:
    pyautogui.hotkey("ctrl", "a")
    time.sleep(delay)
    pyautogui.press("backspace")
    time.sleep(delay)


def paste_text(pyautogui: Any, value: str, delay: float) -> None:
    copy_to_clipboard(value)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(delay)


def press_keys(pyautogui: Any, keys: list[str], delay: float) -> None:
    for key in keys:
        if key.startswith("sleep:"):
            time.sleep(float(key.split(":", 1)[1]))
            continue
        if key.startswith("hotkey:"):
            pyautogui.hotkey(*key.split(":", 1)[1].split("+"))
        else:
            pyautogui.press(key)
        time.sleep(delay)


def execute_entry(pyautogui: Any, entry: dict[str, Any], value: str, delay: float) -> None:
    action_type = entry["action_type"]
    clear_first = bool(entry.get("clear_first", action_type == "text"))

    if action_type == "text":
        if clear_first:
            clear_current_field(pyautogui, delay)
        if value:
            paste_text(pyautogui, value, delay)
    elif action_type == "choice":
        if value:
            choices = entry.get("choices", {})
            mapped = choices.get(value) or choices.get(normalize_key(value))
            if mapped:
                press_keys(pyautogui, list(mapped), delay)
            elif entry.get("choice_fallback", "type_value") == "type_value":
                if clear_first:
                    clear_current_field(pyautogui, delay)
                paste_text(pyautogui, value, delay)
        else:
            should_apply_extra_tabs = False
    elif action_type != "skip":
        raise ValueError(f"action_type não suportado: {action_type}")

    if entry.get("tab_after", True):
        pyautogui.press("tab")
        time.sleep(delay)
        extra_tabs = int(entry.get("extra_tabs", 0)) if should_apply_extra_tabs else 0
        for _ in range(extra_tabs):
            pyautogui.press("tab")
            time.sleep(delay)

    if entry.get("pause_after"):
        time.sleep(float(entry["pause_after"]))


def preview_entry(entry: dict[str, Any], value: str) -> str:
    column = entry.get("csv_column", "")
    return f"{entry['field_id']}: action={entry['action_type']} column={column or '-'} value={value!r}"


def execute_entry_safe(pyautogui: Any, entry: dict[str, Any], value: str, delay: float) -> None:
    action_type = entry["action_type"]
    clear_first = bool(entry.get("clear_first", action_type == "text"))
    should_apply_extra_tabs = True

    if action_type == "text":
        if clear_first:
            clear_current_field(pyautogui, delay)
        if value:
            paste_text(pyautogui, value, delay)
    elif action_type == "choice":
        if value:
            choices = entry.get("choices", {})
            mapped = choices.get(value) or choices.get(normalize_key(value))
            if mapped:
                press_keys(pyautogui, list(mapped), delay)
            elif entry.get("choice_fallback", "type_value") == "type_value":
                if clear_first:
                    clear_current_field(pyautogui, delay)
                paste_text(pyautogui, value, delay)
    elif action_type == "skip":
        if clear_first:
            clear_current_field(pyautogui, delay)
    else:
        raise ValueError(f"action_type nÃ£o suportado: {action_type}")

    if entry.get("tab_after", True):
        pyautogui.press("tab")
        time.sleep(delay)
        for _ in range(int(entry.get("extra_tabs", 0))):
            pyautogui.press("tab")
            time.sleep(delay)

    if entry.get("pause_after"):
        time.sleep(float(entry["pause_after"]))


def row_subset(rows: list[dict[str, str]], start_row: int, limit: int) -> list[tuple[int, dict[str, str]]]:
    start_index = max(start_row - 1, 0)
    sliced = rows[start_index:] if limit <= 0 else rows[start_index : start_index + limit]
    return [(start_index + index + 1, row) for index, row in enumerate(sliced)]


def run_dry_run(rows: list[tuple[int, dict[str, str]]], config: list[dict[str, Any]]) -> int:
    for row_number, row in rows:
        print(f"\n[Linha {row_number}] {row.get('aluno_nome', '').strip()}")
        for entry in config:
            value = resolve_value(row, entry)
            print(" - " + preview_entry(entry, value))
    return 0


def run_live(
    rows: list[tuple[int, dict[str, str]]],
    config: list[dict[str, Any]],
    delay: float,
    log_file: Path,
    start_delay: float,
    review_seconds: float,
    auto_status: str,
    start_position: dict[str, int] | None,
    resume_position: dict[str, int] | None,
) -> int:
    pyautogui = get_pyautogui()
    for row_number, row in rows:
        aluno_nome = row.get("aluno_nome", "").strip()
        print(f"\n[Linha {row_number}] {aluno_nome}")
        print("Posicione o navegador no formulário em branco e foque o primeiro campo.")
        if start_delay > 0:
            print(f"Iniciando em {start_delay:.1f}s...")
            time.sleep(start_delay)
        else:
            input("Pressione Enter para iniciar o preenchimento desta linha...")

        try:
            for entry in config:
                value = resolve_value(row, entry)
                execute_entry_safe(pyautogui, entry, value, delay)

            if auto_status:
                if review_seconds > 0:
                    print(f"Preenchimento concluído. Aguardando {review_seconds:.1f}s para revisão antes de encerrar.")
                    time.sleep(review_seconds)
                status = auto_status
                message = "Preenchido automaticamente para revisão; nenhum salvamento foi acionado."
            else:
                decision = input(
                    "Preenchimento concluído. Revise no navegador, salve manualmente e pressione Enter. "
                    "Digite 'skip' para pular ou 'error:motivo' para registrar erro: "
                ).strip()
                if decision.lower().startswith("error:"):
                    status = "error"
                    message = decision.split(":", 1)[1].strip()
                elif decision.lower() == "skip":
                    status = "skipped"
                    message = "Pulado manualmente pelo operador."
                else:
                    status = "success"
                    message = "Preenchido; confirmação final realizada manualmente."

            append_log(
                log_file,
                {
                    "row_number": row_number,
                    "student_name": aluno_nome,
                    "status": status,
                    "message": message,
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                },
            )
        except Exception as exc:  # noqa: BLE001
            append_log(
                log_file,
                {
                    "row_number": row_number,
                    "student_name": aluno_nome,
                    "status": "error",
                    "message": str(exc),
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                },
            )
            raise
    return 0


def run_live_with_start_position(
    rows: list[tuple[int, dict[str, str]]],
    config: list[dict[str, Any]],
    delay: float,
    log_file: Path,
    start_delay: float,
    resume_delay: float | None,
    review_seconds: float,
    auto_status: str,
    start_position: dict[str, int] | None,
    resume_position: dict[str, int] | None,
) -> int:
    pyautogui = get_pyautogui()
    for row_index, (row_number, row) in enumerate(rows):
        aluno_nome = row.get("aluno_nome", "").strip()
        current_start_delay = start_delay if row_index == 0 else (resume_delay if resume_delay is not None else start_delay)
        current_position = start_position if row_index == 0 else (resume_position or start_position)
        print(f"\n[Linha {row_number}] {aluno_nome}")
        if current_position:
            print("Posicione o navegador no formulário em branco; o script clicará no primeiro campo salvo.")
        else:
            print("Posicione o navegador no formulário em branco e foque o primeiro campo.")
        if current_start_delay > 0:
            print(f"Iniciando em {current_start_delay:.1f}s...")
            time.sleep(current_start_delay)
        else:
            input("Pressione Enter para iniciar o preenchimento desta linha...")
        if current_position:
            click_start_position(pyautogui, current_position, delay)

        try:
            for entry in config:
                value = resolve_value(row, entry)
                execute_entry_safe(pyautogui, entry, value, delay)

            if auto_status:
                if review_seconds > 0:
                    print(f"Preenchimento concluído. Aguardando {review_seconds:.1f}s para revisão antes de encerrar.")
                    time.sleep(review_seconds)
                status = auto_status
                message = "Preenchido automaticamente para revisão; nenhum salvamento foi acionado."
            else:
                decision = input(
                    "Preenchimento concluído. Revise no navegador, salve manualmente e pressione Enter. "
                    "Digite 'skip' para pular ou 'error:motivo' para registrar erro: "
                ).strip()
                if decision.lower().startswith("error:"):
                    status = "error"
                    message = decision.split(":", 1)[1].strip()
                elif decision.lower() == "skip":
                    status = "skipped"
                    message = "Pulado manualmente pelo operador."
                else:
                    status = "success"
                    message = "Preenchido; confirmação final realizada manualmente."

            append_log(
                log_file,
                {
                    "row_number": row_number,
                    "student_name": aluno_nome,
                    "status": status,
                    "message": message,
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                },
            )
        except Exception as exc:  # noqa: BLE001
            append_log(
                log_file,
                {
                    "row_number": row_number,
                    "student_name": aluno_nome,
                    "status": "error",
                    "message": str(exc),
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                },
            )
            raise
    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    csv_path = Path(args.csv).resolve()
    config_path = Path(args.config).resolve()
    log_file = Path(args.log_file).resolve()
    start_position_path = Path(args.start_position_file).resolve()
    resume_position_path = Path(args.resume_position_file).resolve()

    if not csv_path.exists():
        parser.error(f"CSV não encontrado: {csv_path}")
    if not config_path.exists():
        parser.error(f"Configuração não encontrada: {config_path}")

    if args.save_start_position:
        if args.start_delay > 0:
            print(f"Salvando a posição inicial em {args.start_delay:.1f}s...")
            time.sleep(args.start_delay)
        pyautogui = get_pyautogui()
        x, y = pyautogui.position()
        save_start_position(start_position_path, int(x), int(y))
        print(f"Posição inicial salva em: {start_position_path}")
        print(f"x={int(x)} y={int(y)}")
        return 0
    if args.save_resume_position:
        if args.start_delay > 0:
            print(f"Salvando a posição de retomada em {args.start_delay:.1f}s...")
            time.sleep(args.start_delay)
        pyautogui = get_pyautogui()
        x, y = pyautogui.position()
        save_start_position(resume_position_path, int(x), int(y))
        print(f"PosiÃ§Ã£o de retomada salva em: {resume_position_path}")
        print(f"x={int(x)} y={int(y)}")
        return 0

    rows = load_rows(csv_path)
    config = load_config(config_path)
    start_position = None if args.no_click_start_position else load_start_position(start_position_path)
    resume_position = None if args.no_click_start_position else load_start_position(resume_position_path)
    subset = row_subset(rows, args.start_row, args.limit)
    if not subset:
        parser.error("Nenhuma linha selecionada para processar.")

    unknown_columns = sorted(
        {
            entry["csv_column"]
            for entry in config
            if entry.get("csv_column") and entry["csv_column"] not in rows[0]
        }
    )
    if unknown_columns:
        parser.error("Colunas ausentes no CSV: " + ", ".join(unknown_columns))

    if args.dry_run:
        return run_dry_run(subset, config)
    return run_live_with_start_position(
        subset,
        config,
        args.delay,
        log_file,
        args.start_delay,
        args.resume_delay,
        args.review_seconds,
        args.auto_status.strip(),
        start_position,
        resume_position,
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nExecução interrompida pelo usuário.", file=sys.stderr)
        raise SystemExit(130)
