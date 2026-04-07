from __future__ import annotations

from datetime import date, datetime
from difflib import SequenceMatcher
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Iterable

CSV_COLUMNS = [
    "source_file",
    "aluno_nome_original",
    "aluno_nome_corrigido_por_arquivo",
    "aluno_nome",
    "aluno_data_nascimento",
    "aluno_nome_social",
    "aluno_id",
    "aluno_nis",
    "aluno_cartao_sus",
    "aluno_naturalidade",
    "aluno_uf_naturalidade",
    "aluno_nacionalidade",
    "aluno_sexo",
    "aluno_cor_raca",
    "certidao_tipo",
    "certidao_numero",
    "certidao_emissao",
    "certidao_cartorio",
    "certidao_municipio",
    "certidao_uf",
    "aluno_cpf",
    "aluno_zona",
    "aluno_endereco",
    "aluno_municipio",
    "aluno_uf",
    "aluno_cep",
    "ponto_referencia",
    "telefone_contato",
    "tem_alergias",
    "alergias_descricao",
    "mae_nome",
    "mae_endereco",
    "mae_municipio",
    "mae_uf",
    "mae_telefone",
    "mae_rg",
    "mae_orgao_emissor",
    "mae_data_emissao",
    "mae_cpf",
    "mae_profissao",
    "pai_nome",
    "pai_endereco",
    "pai_municipio",
    "pai_uf",
    "pai_telefone",
    "pai_rg",
    "pai_orgao_emissor",
    "pai_data_emissao",
    "pai_cpf",
    "pai_profissao",
    "responsavel_nome",
    "responsavel_endereco",
    "responsavel_municipio",
    "responsavel_uf",
    "responsavel_whatsapp",
    "responsavel_parentesco",
    "responsavel_email",
    "irmaos_rede_municipal",
    "irmaos_quantidade",
    "irmaos_escolas",
    "matricula_etapa_serie",
    "matricula_turno",
    "vulnerabilidade_social",
    "bolsa_familia",
    "utiliza_transporte",
    "programa_escolar",
    "programa_escolar_qual",
    "escola_procedencia",
    "itinerancia",
    "pessoa_com_deficiencia",
    "deficiencia_tipos",
    "atendimento_especializado",
    "cuidador",
    "observacoes_raw",
]

SECTION_HEADERS = {
    "2 - DADOS DE IDENTIFICACAO DO ALUNO",
    "3 - DADOS DO RESPONSAVEL PELO ALUNO",
    "4 - DADOS DA MATRICULA DO ALUNO",
    "5 - DADOS COMPLEMENTARES DO ALUNO",
    "6 - OBSERVACOES",
}

BASE_LABELS = {
    "NOME COMPLETO DO ALUNO:",
    "DATA DE NASCIMENTO:",
    "NOME SOCIAL:",
    "N DO ID DO ALUNO :",
    "NO DO ID DO ALUNO :",
    "NUMERO DE IDENTIFICACAO SOCIAL (NIS)",
    "N DO CARTAO SUS:",
    "NO DO CARTAO SUS:",
    "NATURAL DE:",
    "UF:",
    "NACIONALIDADE:",
    "SEXO:",
    "COR/RACA:",
    "TIPO DE CERTIDAO:",
    "TERMO:",
    "FOLHA:",
    "LIVRO:",
    "DATA DA EMISSAO:",
    "CERTIDAO DE",
    "NOME DO CARTORIO:",
    "MUNICIPIO:",
    "RG DO ALUNO:",
    "ORGAO EMISSOR:",
    "CPF DO ALUNO:",
    "ZONA RESIDENCIAL:",
    "ENDERECO DO ALUNO COM O N:",
    "ENDERECO DO ALUNO COM O NO:",
    "CEP:",
    "FILIACAO",
    "NOME DA MAE:",
    "ENDERECO:",
    "TELEFONE:",
    "TELEFONE",
    "RG:",
    "CPF:",
    "PROFISSAO:",
    "NOME DO PAI:",
    "NOME DO RESPONSAVEL:",
    "WHATSAPP:",
    "GRAU DE PARENTESCO:",
    "E-MAIL:",
    "POSSUI IRMAOS MATRICULADOS NA REDE MUNICIPAL DE ALAGOA NOVA:",
    "QUANTIDADE DE IRMAOS MATRICULADOS NO MUNICIPIO:",
    "EM QUE ESCOLA(S) O(S) IRMAO(S) ESTAO MATRICULADOS:",
    "MATRICULA - ETAPA / SERIE",
    "SITUACAO DO ALUNO NO ANO ANTERIOR:",
    "ALUNO EM VULNERABILIDADE SOCIAL:",
    "EJA:",
    "TURNO:",
    "PARTICIPANTE DE ALGUM PROGRAMA ESCOLAR:",
    "QUAL PROGRAMA QUE A CRIANCA RECEBE APOIO / SUPORTE:",
    "ESCOLA DE PROCEDENCIA: (ULTIMA ESCOLA QUE ESTUDOU)",
    "EM QUE ANO:",
    "SITUACAO DE INTINERANCIA:",
    "PESSOA COM DEFICIENCIA:",
    "TIPO:",
    "RECURSOS PARA AVALIACOES DO INEP:",
    "CID N:",
    "RECEBE ATENDIMENTO ESPECIALIZADO:",
    "INSTITUICAO QUE RECEBE ATENDIMENTO ESPECIALIZADO:",
    "PROFESSOR DO AEE:",
    "TURNO DO ATENDIENTO EDUCACIONAL ESPECIALIZADO: (CONTRA TURNO)",
    "TIPO DE ATENDIMENTO EDUCACIONAL ESPECIALIZADO:",
    "NECESSITA DE CUIDADOR:",
    "FAZ USO DE ALGUM TIPO DE MEDICAMENTO:",
    "QUAL MEDICAMENTO:",
    "UTILIZA TRANSPORTE ESCOLAR",
    "TRANSPORTE ESCOLAR MANTIDO:",
    "TRANSPORTE:",
    "PUBLICO:",
    "TIPO DE TRANSPORTE:",
    "RECEBE BENEFICIO SOCIAL:",
    "TIPO DE BENEFICIO:",
}

SERIE_OPTIONS = [
    "Berçário",
    "Berçário A",
    "Berçário B",
    "Maternal I",
    "Maternal II",
    "Pré I",
    "Pré II",
    "1º ano",
    "2º ano",
    "3º ano",
    "4º ano",
    "5º ano",
    "6º ano",
    "7º ano",
    "8º ano",
    "9º ano",
]

DEFICIENCIA_OPTIONS = [
    "Baixa Visão",
    "Cegueira",
    "Surdez",
    "Surdo-cegueira",
    "Autista",
    "Cadeirante",
    "Def. Intelectual",
    "Def. Auditiva",
    "Def. Física",
    "Def. Múltiplas",
    "Altas Habilidades",
    "Diabetes",
    "Intolerância a lactose",
]

LOWERCASE_WORDS = {
    "a",
    "ao",
    "aos",
    "as",
    "com",
    "da",
    "das",
    "de",
    "do",
    "dos",
    "e",
    "em",
    "na",
    "nas",
    "no",
    "nos",
    "para",
    "por",
}

UPPERCASE_TOKENS = {
    "AEE",
    "CID",
    "CPF",
    "DDD",
    "EJA",
    "INEP",
    "LOA",
    "PB",
    "RG",
    "SUS",
    "UF",
}


def fold_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = normalized.replace("º", "O").replace("ª", "A")
    normalized = normalized.replace("–", "-").replace("—", "-")
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip().upper()


def normalize_label_text(value: str) -> str:
    return (
        (value or "")
        .replace("Âº", "")
        .replace("Â°", "")
        .replace("Âª", "")
        .replace("º", "")
        .replace("°", "")
        .replace("ª", "")
    )


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    normalized = (
        normalized.replace("\u00a0", " ")
        .replace("\u2007", " ")
        .replace("\u202f", " ")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2212", "-")
    )
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def cleanup_value(value: str) -> str:
    cleaned = normalize_text(value)
    if fold_text(cleaned) in {"", "_", "__", "___", "____", "//", "////", "N/A", "NAO TEM", "NAO POSSUI"}:
        return ""
    if re.fullmatch(r"[-\s]+", cleaned):
        return ""
    return cleaned


def cleanup_point_of_reference(value: str) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    folded = re.sub(r"[^A-Z0-9 ]+", "", fold_text(cleaned)).strip()
    if folded == "PROXIMO":
        return ""
    return cleaned


def is_plausible_name_text(value: str) -> bool:
    cleaned = cleanup_value(value)
    if not cleaned:
        return False
    digits = re.sub(r"\D", "", cleaned)
    if digits:
        return False
    letters = [ch for ch in cleaned if ch.isalpha()]
    if len(letters) < 3:
        return False
    folded = fold_text(cleaned)
    forbidden_tokens = {
        "NO DO ID DO ALUNO",
        "NUMERO DE IDENTIFICACAO SOCIAL",
        "N DO CARTAO SUS",
        "NO DO CARTAO SUS",
        "NATURAL DE",
        "UF",
        "NACIONALIDADE",
    }
    return not any(token in folded for token in forbidden_tokens)


def strip_parenthetical_text(value: str) -> str:
    return normalize_text(re.sub(r"\s*\([^)]*\)", "", value or ""))


def smart_title(value: str) -> str:
    cleaned = normalize_text(value)
    if not cleaned:
        return ""

    words = cleaned.split(" ")
    formatted: list[str] = []
    for index, word in enumerate(words):
        if not word:
            continue
        parts = re.split(r"([-/])", word)
        rebuilt: list[str] = []
        for part in parts:
            if part in {"-", "/"}:
                rebuilt.append(part)
                continue
            upper = part.upper()
            if upper in UPPERCASE_TOKENS:
                rebuilt.append(upper)
            else:
                lower = part.lower()
                if index > 0 and lower in LOWERCASE_WORDS:
                    rebuilt.append(lower)
                else:
                    rebuilt.append(lower[:1].upper() + lower[1:])
        formatted.append("".join(rebuilt))
    return " ".join(formatted)


def extract_name_from_filename(docx_path: Path) -> str:
    stem = normalize_text(docx_path.stem)
    stem = re.sub(r"(?i)^matr[íi]cula(?:\s+de)?\s+", "", stem).strip()
    return smart_title(stem)


def collapse_repeated_letters(value: str) -> str:
    return re.sub(r"([A-Z])\1+", r"\1", fold_text(value))


def should_replace_student_name_with_filename(extracted_name: str, filename_name: str) -> bool:
    extracted = smart_title(strip_parenthetical_text(extracted_name))
    filename = smart_title(strip_parenthetical_text(filename_name))
    if not filename:
        return False
    if not extracted:
        return True
    if fold_text(extracted) == fold_text(filename):
        return False

    extracted_tokens = extracted.split()
    filename_tokens = filename.split()
    if len(extracted_tokens) != len(filename_tokens):
        return False

    differing_tokens = 0
    repeated_letter_fix_found = False
    for extracted_token, filename_token in zip(extracted_tokens, filename_tokens):
        extracted_folded = fold_text(extracted_token)
        filename_folded = fold_text(filename_token)
        if extracted_folded == filename_folded:
            continue
        if extracted_folded in LOWERCASE_WORDS or filename_folded in LOWERCASE_WORDS:
            return False
        collapsed_match = collapse_repeated_letters(extracted_token) == collapse_repeated_letters(filename_token)
        if not collapsed_match:
            return False
        if extracted_folded[:1] != filename_folded[:1]:
            return False
        if len(extracted_folded) <= len(filename_folded):
            return False
        differing_tokens += 1
        repeated_letter_fix_found = True

    overall_ratio = SequenceMatcher(None, fold_text(extracted), fold_text(filename)).ratio()
    return repeated_letter_fix_found and differing_tokens <= 2 and overall_ratio >= 0.94


def normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if not digits:
        return ""
    if digits.startswith("55") and len(digits) > 11:
        digits = digits[2:]
    if len(digits) < 8:
        return ""
    if len(digits) in {8, 9}:
        digits = "83" + digits
    if len(digits) != 11:
        return ""
    return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"


def normalize_address(value: str) -> str:
    cleaned = normalize_text(value)
    if not cleaned:
        return ""
    cleaned = re.sub(r"(?i)\bS/\s*(?:N(?:[º°]|\b))?", "S/N", cleaned)
    folded = fold_text(cleaned)
    prefix = ""
    body = cleaned
    if folded.startswith("AV") or folded.startswith("AVENIDA"):
        prefix = "Avenida"
        body = re.sub(r"^(AVENIDA|AV)[\s:;.-]*", "", cleaned, flags=re.IGNORECASE)
    elif folded.startswith("RUA") or re.match(r"^R[\s:;.-]", cleaned, flags=re.IGNORECASE):
        prefix = "Rua"
        body = re.sub(r"^(RUA|R)[\s:;.-]*", "", cleaned, flags=re.IGNORECASE)

    match = re.match(r"^(.*?)(?:,\s*|\s+)(\d+)\s*[-–]?\s*(.*)$", body)
    if not match:
        match = re.match(r"^(.*?)\s+-\s*(\d+)(?:\s*[-â€“]\s*(.*))?$", body)
    if match:
        street, number, district = match.groups()
        street = lowercase_leading_preposition(smart_title(street.strip(" ,;-")))
        district = smart_title((district or "").strip(" ,;-"))
        prefix_text = f"{prefix} " if prefix else ""
        if district:
            result = f"{prefix_text}{street}, {number} - {district}".strip()
        else:
            result = f"{prefix_text}{street}, {number}".strip()
        result = re.sub(r"\bS/N\w*\b", "S/N", result, flags=re.IGNORECASE)
        result = re.sub(r"S/N(?=-)", "S/N ", result, flags=re.IGNORECASE)
        result = re.sub(r"^(Rua)\s+Rua:?\s+", r"\1 ", result, flags=re.IGNORECASE)
        result = re.sub(r"^(Avenida)\s+(Avenida|Av):?\s+", r"\1 ", result, flags=re.IGNORECASE)
        return result

    prefix_text = f"{prefix} " if prefix else ""
    result = f"{prefix_text}{lowercase_leading_preposition(smart_title(body))}".strip()
    result = re.sub(r"\bS/N\w*\b", "S/N", result, flags=re.IGNORECASE)
    result = re.sub(r"S/N(?=-)", "S/N ", result, flags=re.IGNORECASE)
    result = re.sub(r"^(Rua)\s+Rua:?\s+", r"\1 ", result, flags=re.IGNORECASE)
    result = re.sub(r"^(Avenida)\s+(Avenida|Av):?\s+", r"\1 ", result, flags=re.IGNORECASE)
    return result


def normalize_student_address(value: str) -> str:
    normalized = normalize_address(value)
    if not normalized:
        return ""
    normalized = re.sub(r"\s*-\s*Zona Rural\b", "", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\s*,\s*Alagoa Nova\b", "", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\s*-\s*Alagoa Nova\b", "", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\s{2,}", " ", normalized)
    return normalized.strip(" ,-")


def lowercase_leading_preposition(value: str) -> str:
    cleaned = normalize_text(value)
    if not cleaned:
        return ""
    parts = cleaned.split(" ", 1)
    first = parts[0].lower()
    if first in LOWERCASE_WORDS:
        return f"{first} {parts[1]}" if len(parts) > 1 else first
    return cleaned


def sentence_case(value: str) -> str:
    cleaned = normalize_text(value)
    if not cleaned:
        return ""
    lowered = cleaned.lower()
    return lowered[:1].upper() + lowered[1:]


def normalize_alergia_item(value: str) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    cleaned = re.sub(r"(?i)^alergias?\s*:\s*", "", cleaned)
    cleaned = re.sub(r"(?i)^alergia\s+a\s+", "", cleaned)
    cleaned = re.sub(r"(?i)^al[ée]rgic[ao]\s+a\s+", "", cleaned)
    cleaned = re.sub(r"(?i)^al[ée]rgic[ao]\s*:\s*", "", cleaned)
    cleaned = cleaned.strip(" .;,:-")
    return sentence_case(cleaned)


def extract_alergias(deficiencia_tipos: str, observacoes: str = "") -> str:
    alergias: list[str] = []
    for item in [part.strip() for part in (deficiencia_tipos or "").split(";") if part.strip()]:
        folded = fold_text(item)
        if "INTOLERANCIA" in folded or "ALERG" in folded:
            alergias.append(item)
    for item in [cleanup_value(part) for part in (observacoes or "").split("|") if cleanup_value(part)]:
        folded = fold_text(item)
        if "INTOLERANCIA" in folded or "ALERG" in folded:
            alergias.append(item)
    if not alergias and "ALERG" in fold_text(observacoes):
        alergias.append("Alergia")
    normalized = [normalize_alergia_item(item) for item in alergias if normalize_alergia_item(item)]
    return "; ".join(dict.fromkeys(normalized))


def normalize_grade(value: str) -> str:
    folded = fold_text(value)
    if folded.startswith("BERCARIO"):
        return "Creche 1"
    if folded.startswith("MATERNAL II"):
        return "Creche 3"
    if folded.startswith("MATERNAL I"):
        return "Creche 2"
    if folded.startswith("MATERNAL"):
        return "Creche 2"
    ano_match = re.fullmatch(r"([1-9])\s*(?:O)?\s*ANO", folded)
    if ano_match:
        return f"{ano_match.group(1)}º ano do ensino fundamental"
    return smart_title(value)


def infer_grade_from_birthdate(value: str, reference_date: date | None = None) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    try:
        birthdate = datetime.strptime(cleaned, "%d/%m/%Y").date()
    except ValueError:
        return ""
    ref = reference_date or date.today()
    age = ref.year - birthdate.year - ((ref.month, ref.day) < (birthdate.month, birthdate.day))
    if age <= 1:
        return "Creche 1"
    if age == 2:
        return "Creche 2"
    if age == 3:
        return "Creche 3"
    return ""


def preferred_document(cpf: str, rg: str) -> str:
    return cleanup_value(cpf) or cleanup_value(rg)


def preferred_phone(*phones: str) -> str:
    for phone in phones:
        normalized = normalize_phone(phone)
        if normalized:
            return normalized
    return ""


def is_valid_cns(value: str) -> bool:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) != 15:
        return False

    if digits[0] in {"1", "2"}:
        pis = digits[:11]
        total = sum(int(digit) * weight for digit, weight in zip(pis, range(15, 4, -1)))
        remainder = total % 11
        dv = 11 - remainder

        if dv == 11:
            dv = 0
        if dv == 10:
            total = sum(int(digit) * weight for digit, weight in zip(pis, range(15, 4, -1))) + 2
            remainder = total % 11
            dv = 11 - remainder
            result = f"{pis}001{int(dv)}"
        else:
            result = f"{pis}000{int(dv)}"
        return digits == result

    if digits[0] in {"5", "7", "8", "9"}:
        total = sum(int(digit) * weight for digit, weight in zip(digits, range(15, 0, -1)))
        return total % 11 == 0

    return False


def normalize_sus_card(value: str) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    digits = re.sub(r"\D", "", cleaned)
    if len(digits) != 15:
        return ""
    if not is_valid_cns(digits):
        return ""
    return f"{digits[:3]} {digits[3:7]} {digits[7:11]} {digits[11:]}"


def normalize_cpf(value: str) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    digits = re.sub(r"\D", "", cleaned)
    if len(digits) != 11:
        return ""
    return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"


def normalize_student_id(value: str) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    digits = re.sub(r"\D", "", cleaned)
    if len(digits) < 11 or len(digits) > 13:
        return ""
    return digits


def normalize_nis(value: str) -> str:
    cleaned = cleanup_value(value)
    if not cleaned:
        return ""
    digits = re.sub(r"\D", "", cleaned)
    if len(digits) < 10 or len(digits) > 16:
        return ""
    return digits


def format_record(record: dict[str, str]) -> dict[str, str]:
    name_fields = {
        "aluno_nome_original",
        "aluno_nome",
        "aluno_nome_social",
        "mae_nome",
        "pai_nome",
        "responsavel_nome",
    }
    address_fields = {
        "aluno_endereco",
        "mae_endereco",
        "pai_endereco",
        "responsavel_endereco",
    }
    title_fields = {
        "aluno_naturalidade",
        "aluno_nacionalidade",
        "aluno_sexo",
        "aluno_cor_raca",
        "aluno_zona",
        "aluno_municipio",
        "aluno_nome_corrigido_por_arquivo",
        "certidao_tipo",
        "certidao_cartorio",
        "certidao_municipio",
        "ponto_referencia",
        "mae_municipio",
        "mae_profissao",
        "pai_municipio",
        "pai_profissao",
        "responsavel_municipio",
        "responsavel_parentesco",
        "irmaos_escolas",
        "matricula_etapa_serie",
        "programa_escolar_qual",
        "escola_procedencia",
        "deficiencia_tipos",
        "alergias_descricao",
    }
    phone_fields = {"telefone_contato", "mae_telefone", "pai_telefone", "responsavel_whatsapp"}
    upper_fields = {"aluno_uf_naturalidade", "certidao_uf", "aluno_uf", "mae_uf", "pai_uf", "responsavel_uf"}

    for field in name_fields:
        record[field] = smart_title(strip_parenthetical_text(record.get(field, "")))
    for field in address_fields:
        record[field] = normalize_address(record.get(field, ""))
    if not record.get("aluno_endereco"):
        record["aluno_endereco"] = (
            record.get("mae_endereco", "")
            or record.get("pai_endereco", "")
            or record.get("responsavel_endereco", "")
        )
    record["aluno_endereco"] = normalize_student_address(record.get("aluno_endereco", ""))
    for field in title_fields:
        if field == "matricula_etapa_serie":
            record[field] = normalize_grade(record.get(field, ""))
        else:
            record[field] = smart_title(record.get(field, ""))
    for field in phone_fields:
        record[field] = normalize_phone(record.get(field, ""))
    for field in upper_fields:
        record[field] = cleanup_value(record.get(field, "")).upper()
    record["aluno_cpf"] = normalize_cpf(record.get("aluno_cpf", ""))
    record["aluno_id"] = normalize_student_id(record.get("aluno_id", ""))
    record["aluno_nis"] = normalize_nis(record.get("aluno_nis", ""))
    record["aluno_cartao_sus"] = normalize_sus_card(record.get("aluno_cartao_sus", ""))

    cpf_digits = re.sub(r"\D", "", record.get("aluno_cpf", ""))
    sus_digits = re.sub(r"\D", "", record.get("aluno_cartao_sus", ""))
    if record["aluno_nis"] and (
        record["aluno_nis"] == cpf_digits
        or record["aluno_nis"] == record["aluno_id"]
        or record["aluno_nis"] == sus_digits
    ):
        record["aluno_nis"] = ""

    record["alergias_descricao"] = extract_alergias(
        record.get("deficiencia_tipos", ""),
        record.get("observacoes_raw", ""),
    )
    record["tem_alergias"] = "Sim" if record["alergias_descricao"] else "Não"

    if not record.get("matricula_etapa_serie"):
        record["matricula_etapa_serie"] = infer_grade_from_birthdate(record.get("aluno_data_nascimento", ""))

    return record


def strip_trailing_uf(value: str) -> str:
    return re.sub(r"\s*[-]\s*[A-Z]{2}$", "", value).strip(" ,")


def iter_docx_lines(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")

    stack: list[str] = []
    current_parts: list[str] = []
    lines: list[str] = []
    paragraph_capture = False

    for event, elem in ET.iterparse(BytesIO(xml_bytes), events=("start", "end")):
        tag = elem.tag.rsplit("}", 1)[-1]
        if event == "start":
            stack.append(tag)
            if tag == "p":
                paragraph_capture = "txbxContent" in stack or "tc" in stack
                current_parts = []
        else:
            if tag == "t" and paragraph_capture:
                current_parts.append(elem.text or "")
            elif tag == "p" and paragraph_capture:
                line = normalize_text("".join(current_parts))
                if line:
                    lines.append(line)
                current_parts = []
                paragraph_capture = False
                elem.clear()
            if stack:
                stack.pop()
    return lines


def iter_docx_paragraph_lines(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")

    current_parts: list[str] = []
    lines: list[str] = []

    for event, elem in ET.iterparse(BytesIO(xml_bytes), events=("start", "end")):
        tag = elem.tag.rsplit("}", 1)[-1]
        if event == "start" and tag == "p":
            current_parts = []
        elif event == "end":
            if tag == "t":
                current_parts.append(elem.text or "")
            elif tag == "p":
                line = normalize_text("".join(current_parts))
                if line:
                    lines.append(line)
                current_parts = []
                elem.clear()
    return lines


def find_index(lines: list[str], target: str, start: int = 0) -> int:
    folded_target = fold_text(normalize_label_text(target))
    for index in range(start, len(lines)):
        if fold_text(normalize_label_text(lines[index])) == folded_target:
            return index
    return -1


def find_index_prefix(lines: list[str], target: str, start: int = 0) -> int:
    folded_target = fold_text(normalize_label_text(target))
    for index in range(start, len(lines)):
        if fold_text(normalize_label_text(lines[index])).startswith(folded_target):
            return index
    return -1


def find_any_index(lines: list[str], targets: Iterable[str], start: int = 0) -> int:
    indexes = [find_index(lines, target, start=start) for target in targets]
    indexes = [index for index in indexes if index >= 0]
    return min(indexes) if indexes else -1


def find_any_index_prefix(lines: list[str], targets: Iterable[str], start: int = 0) -> int:
    indexes = [find_index_prefix(lines, target, start=start) for target in targets]
    indexes = [index for index in indexes if index >= 0]
    return min(indexes) if indexes else -1


def find_all_indices(lines: list[str], target: str, start: int = 0) -> list[int]:
    folded_target = fold_text(target)
    return [index for index in range(start, len(lines)) if fold_text(lines[index]) == folded_target]


def is_label_line(line: str) -> bool:
    folded = fold_text(line)
    if not folded:
        return False
    if folded in BASE_LABELS or folded in SECTION_HEADERS:
        return True
    if re.match(r"^\d+ - ", folded):
        return True
    if folded.endswith(":"):
        return True
    return False


def next_value(lines: list[str], start_index: int, stop_index: int | None = None, max_lines: int = 1) -> str:
    if start_index < 0:
        return ""
    values: list[str] = []
    limit = len(lines) if stop_index is None else min(stop_index, len(lines))
    for index in range(start_index + 1, limit):
        candidate = cleanup_value(lines[index])
        if not candidate:
            continue
        if is_label_line(candidate):
            break
        values.append(candidate)
        if len(values) >= max_lines:
            break
    return " ".join(values).strip()


def value_after_label(lines: list[str], start_index: int, stop_index: int | None = None, max_lines: int = 1) -> str:
    if start_index < 0:
        return ""
    line = cleanup_value(lines[start_index])
    if ":" in line:
        inline_value = cleanup_value(line.split(":", 1)[1])
        if inline_value and not is_label_line(inline_value):
            return inline_value
    return next_value(lines, start_index, stop_index, max_lines=max_lines)


def extract_header_numeric_fields(ident_lines: list[str]) -> dict[str, str]:
    social_index = find_index_prefix(ident_lines, "NOME SOCIAL:")
    naturalidade_index = find_index_prefix(ident_lines, "NATURAL DE:")
    if social_index < 0 or naturalidade_index < 0 or naturalidade_index <= social_index:
        return {"aluno_id": "", "aluno_nis": "", "aluno_cartao_sus": ""}

    id_index = find_any_index_prefix(
        ident_lines,
        [
            "N DO ID DO ALUNO :",
            "NO DO ID DO ALUNO :",
            "Nº DO ID DO ALUNO :",
            "Nº DO ID DO ALUNO:",
            "NO DO ID DO ALUNO:",
            "N DO ID DO ALUNO:",
        ],
    )
    nis_index = find_any_index_prefix(
        ident_lines,
        [
            "NUMERO DE IDENTIFICACAO SOCIAL (NIS):",
            "NUMERO DE IDENTIFICACAO SOCIAL (NIS)",
            "N?MERO DE IDENTIFICA??O SOCIAL (NIS):",
            "N?MERO DE IDENTIFICA??O SOCIAL (NIS)",
            "NÚMERO DE IDENTIFICAÇÃO SOCIAL (NIS):",
            "NÚMERO DE IDENTIFICAÇÃO SOCIAL (NIS)",
        ],
    )
    sus_index = find_any_index_prefix(
        ident_lines,
        [
            "NO DO CARTAO SUS:",
            "NO DO CARTAO SUS",
            "N DO CARTAO SUS:",
            "N DO CARTAO SUS",
            "N? DO CART?O SUS:",
            "No DO CART?O SUS:",
            "Nº DO CARTÃO SUS:",
            "Nº DO CARTÃO SUS",
        ],
    )

    extracted = {"aluno_id": "", "aluno_nis": "", "aluno_cartao_sus": ""}

    explicit_labels_complete = all(index >= 0 for index in (id_index, nis_index, sus_index))
    if explicit_labels_complete:
        id_stop = nis_index if nis_index > id_index >= 0 else naturalidade_index
        nis_stop = sus_index if sus_index > nis_index >= 0 else naturalidade_index
        sus_stop = naturalidade_index

        extracted["aluno_id"] = cleanup_value(value_after_label(ident_lines, id_index, id_stop))
        extracted["aluno_nis"] = cleanup_value(value_after_label(ident_lines, nis_index, nis_stop))
        extracted["aluno_cartao_sus"] = cleanup_value(value_after_label(ident_lines, sus_index, sus_stop))
        return extracted

    raw_candidates: list[str] = []
    previous_folded = ""
    for line in ident_lines[social_index + 1 : naturalidade_index]:
        cleaned = cleanup_value(line)
        if not cleaned:
            continue
        folded = fold_text(cleaned)
        if folded == previous_folded:
            continue
        previous_folded = folded
        if is_label_line(cleaned):
            continue
        raw_candidates.append(cleaned)

    digit_candidates: list[tuple[str, str]] = []
    seen_digits: set[str] = set()
    for candidate in raw_candidates:
        digits = re.sub(r"\D", "", candidate)
        if not digits or digits in seen_digits:
            continue
        seen_digits.add(digits)
        digit_candidates.append((candidate, digits))
        if len(digits) == 15 and not extracted["aluno_cartao_sus"]:
            extracted["aluno_cartao_sus"] = candidate
        elif len(digits) in {12, 13} and not extracted["aluno_id"]:
            extracted["aluno_id"] = candidate
        elif len(digits) == 11 and not extracted["aluno_nis"]:
            extracted["aluno_nis"] = candidate

    # Some scans lose the ID/NIS/SUS labels but keep the values in order after
    # "NOME SOCIAL": first ID, then NIS, then SUS.
    if digit_candidates:
        if not extracted["aluno_id"] and len(digit_candidates) >= 1:
            first_candidate, first_digits = digit_candidates[0]
            if len(first_digits) in {11, 12, 13}:
                extracted["aluno_id"] = first_candidate
        if not extracted["aluno_nis"] and len(digit_candidates) >= 2:
            second_candidate, second_digits = digit_candidates[1]
            if 10 <= len(second_digits) <= 16:
                extracted["aluno_nis"] = second_candidate
        if not extracted["aluno_cartao_sus"] and len(digit_candidates) >= 3:
            third_candidate, third_digits = digit_candidates[2]
            if len(third_digits) == 15:
                extracted["aluno_cartao_sus"] = third_candidate
    return extracted


def extract_student_cpf(ident_lines: list[str], zona_index: int) -> str:
    cpf_index = find_index_prefix(ident_lines, "CPF DO ALUNO:")
    cpf_value = cleanup_value(value_after_label(ident_lines, cpf_index, zona_index))
    if cpf_value:
        return cpf_value

    search_start = find_any_index_prefix(
        ident_lines,
        ["CPF DO ALUNO:", "ÓRGÃO EMISSOR:", "ORGAO EMISSOR:", "RG DO ALUNO:"],
    )
    if search_start < 0:
        return ""
    limit = len(ident_lines) if zona_index < 0 else zona_index
    for line in ident_lines[search_start + 1 : limit]:
        cleaned = cleanup_value(line)
        digits = re.sub(r"\D", "", cleaned)
        if len(digits) == 11:
            return cleaned
    return ""


def lines_between(lines: list[str], start_label: str, end_label: str | None = None, start: int = 0) -> list[str]:
    start_index = find_index(lines, start_label, start)
    if start_index < 0:
        return []
    end_index = len(lines)
    if end_label is not None:
        found = find_index(lines, end_label, start_index + 1)
        if found >= 0:
            end_index = found
    return lines[start_index + 1 : end_index]


def parse_marked_options(block_lines: Iterable[str], options: Iterable[str]) -> list[str]:
    block_text = " ".join(block_lines)
    folded_block = fold_text(block_text)
    selected: list[str] = []
    for option in options:
        pattern = re.compile(rf"{re.escape(fold_text(option))}\s*\(([^)]*)\)")
        for match in pattern.finditer(folded_block):
            if "X" in match.group(1):
                selected.append(option)
                break
    return selected


def parse_marked_option(block_lines: Iterable[str], options: Iterable[str], prefer_last: bool = False) -> str:
    selected = parse_marked_options(block_lines, options)
    if not selected:
        return ""
    return selected[-1] if prefer_last else selected[0]


def section_slice(lines: list[str], start_label: str, end_label: str | None) -> list[str]:
    start_index = find_index(lines, start_label)
    if start_index < 0:
        return []
    end_index = len(lines)
    if end_label is not None:
        found = find_index(lines, end_label, start_index + 1)
        if found >= 0:
            end_index = found
    return lines[start_index:end_index]


def parse_person_block(lines: list[str], name_label: str, whatsapp_label: str | None = None) -> dict[str, str]:
    name_index = find_index_prefix(lines, name_label)
    if name_index < 0:
        return {
            "nome": "",
            "municipio": "",
            "endereco": "",
            "uf": "",
            "telefone": "",
            "rg": "",
            "orgao_emissor": "",
            "data_emissao": "",
            "cpf": "",
            "profissao": "",
            "whatsapp": "",
            "parentesco": "",
            "email": "",
        }

    positions = {
        "nome": name_index,
        "municipio": find_index_prefix(lines, "MUNICÍPIO:", start=name_index),
        "endereco": find_index_prefix(lines, "ENDEREÇO:", start=name_index),
        "uf": find_index_prefix(lines, "UF:", start=name_index),
        "telefone": find_any_index_prefix(lines, ["TELEFONE:", "TELEFONE", "WHATSAPP:"], start=name_index),
        "rg": find_index_prefix(lines, "RG:", start=name_index),
        "orgao_emissor": find_any_index_prefix(lines, ["ÓRGÃO EMISSOR:", "ORGAO EMISSOR:"], start=name_index),
        "data_emissao": find_any_index_prefix(lines, ["DATA DA EMISSÃO:", "DATA DA EMISSÃO", "DATA DE EMISSÃO:", "DATA DE EMISSÃO"], start=name_index),
        "cpf": find_any_index_prefix(lines, ["CPF:", "CPF DO ALUNO:"], start=name_index),
        "profissao": find_index_prefix(lines, "PROFISSÃO:", start=name_index),
        "whatsapp": "",
        "parentesco": "",
        "email": "",
    }

    if whatsapp_label:
        positions["whatsapp"] = find_index_prefix(lines, whatsapp_label, start=name_index)
        positions["parentesco"] = find_index_prefix(lines, "GRAU DE PARENTESCO:", start=name_index)
        positions["email"] = find_index_prefix(lines, "E-MAIL:", start=name_index)

    ordered = sorted((index, key) for key, index in positions.items() if isinstance(index, int) and index >= 0)
    data: dict[str, str] = {}
    for order_index, (line_index, key) in enumerate(ordered):
        next_stop = ordered[order_index + 1][0] if order_index + 1 < len(ordered) else None
        data[key] = value_after_label(lines, line_index, next_stop)

    return {key: cleanup_value(data.get(key, "")) for key in positions}


def extract_certidao_numero(ident_lines: list[str]) -> str:
    model_index = find_index(ident_lines, "(Matrícula-Modelo Novo)")
    if model_index < 0:
        return ""
    emissao_indexes = find_all_indices(ident_lines, "DATA DA EMISSÃO:", start=model_index + 1)
    stop_index = emissao_indexes[0] if emissao_indexes else len(ident_lines)
    candidates: list[str] = []
    for line in ident_lines[model_index + 1 : stop_index]:
        cleaned = cleanup_value(line)
        if len(re.sub(r"\D", "", cleaned)) >= 20:
            candidates.append(cleaned)
    unique_candidates = list(dict.fromkeys(candidates))
    if not unique_candidates:
        return ""
    return min(unique_candidates, key=len)


def extract_certidao_numero_modelo(ident_lines: list[str]) -> str:
    model_index = find_index(ident_lines, "(MatrÃ­cula-Modelo Novo)")
    if model_index < 0:
        return ""
    stop_index = find_index(ident_lines, "NOME DO CARTÃ“RIO:", start=model_index + 1)
    if stop_index < 0:
        stop_index = len(ident_lines)
    candidates: list[str] = []
    for line in ident_lines[model_index + 1 : stop_index]:
        cleaned = cleanup_value(line)
        if len(re.sub(r"\D", "", cleaned)) >= 20:
            candidates.append(cleaned)
    unique_candidates = list(dict.fromkeys(candidates))
    if not unique_candidates:
        return ""
    return min(unique_candidates, key=len)


def extract_certidao_emissao_modelo(ident_lines: list[str]) -> str:
    model_index = find_index(ident_lines, "(MatrÃ­cula-Modelo Novo)")
    if model_index < 0:
        return ""
    stop_index = find_index(ident_lines, "NOME DO CARTÃ“RIO:", start=model_index + 1)
    if stop_index < 0:
        stop_index = len(ident_lines)
    for line in ident_lines[model_index + 1 : stop_index]:
        cleaned = cleanup_value(line)
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", cleaned):
            return cleaned
    return ""


def extract_certidao_numero_bloco(ident_lines: list[str]) -> str:
    start_index = find_any_index(ident_lines, ["CERTIDAO DE", "(MATRICULA-MODELO NOVO)"])
    if start_index < 0:
        return ""
    stop_index = find_any_index(ident_lines, ["NOME DO CARTORIO:", "NOME DO CARTORIO"], start=start_index + 1)
    if stop_index < 0:
        stop_index = len(ident_lines)
    candidates: list[str] = []
    for line in ident_lines[start_index + 1 : stop_index]:
        cleaned = cleanup_value(line)
        if len(re.sub(r"\D", "", cleaned)) >= 20:
            candidates.append(cleaned)
    unique_candidates = list(dict.fromkeys(candidates))
    if unique_candidates:
        return min(unique_candidates, key=len)

    parts: list[str] = []
    for line in ident_lines[start_index + 1 : stop_index]:
        cleaned = cleanup_value(line)
        if not cleaned:
            continue
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", cleaned):
            break
        folded = fold_text(cleaned)
        if folded in {"(MATRICULA-MODELO NOVO)", "NASCIMENTO:", "/"}:
            continue
        if is_label_line(cleaned):
            continue
        parts.append(cleaned)
        if len(parts) >= 3:
            break

    if len(parts) >= 3:
        return f"Termo: {parts[0]} Folha: {parts[1]} Livro: {parts[2]}"
    return ""


def extract_certidao_emissao_bloco(ident_lines: list[str]) -> str:
    start_index = find_index(ident_lines, "CERTIDÃO DE")
    if start_index < 0:
        start_index = find_index(ident_lines, "(MatrÃ­cula-Modelo Novo)")
    if start_index < 0:
        return ""
    stop_index = find_index(ident_lines, "NOME DO CARTÃ“RIO:", start=start_index + 1)
    if stop_index < 0:
        stop_index = len(ident_lines)
    for line in ident_lines[start_index + 1 : stop_index]:
        cleaned = cleanup_value(line)
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", cleaned):
            return cleaned
    return ""


def extract_observacoes(comp_lines: list[str]) -> str:
    start_index = find_any_index(comp_lines, ["6 - OBSERVA??ES", "6 - OBSERVACOES"], start=0)
    if start_index >= 0:
        relevant = [cleanup_value(line) for line in comp_lines[start_index + 1 :] if cleanup_value(line)]
        relevant = [line for line in relevant if not is_label_line(line)]
        return " | ".join(dict.fromkeys(relevant))

    start_index = find_any_index(comp_lines, ["FAZ USO DE ALGUM TIPO DE MEDICAMENTO:", "6 - OBSERVA??ES"], start=0)
    if start_index < 0:
        return ""
    relevant = [cleanup_value(line) for line in comp_lines[start_index:] if cleanup_value(line)]
    return " | ".join(relevant)


def extract_point_of_reference(paragraph_lines: list[str]) -> str:
    for line in paragraph_lines:
        if "PONTO DE REFERENCIA:" not in fold_text(line):
            continue
        value = line.split(":", 1)[1] if ":" in line else ""
        return cleanup_point_of_reference(value)
    stop_index = len(paragraph_lines)
    for index, line in enumerate(paragraph_lines):
        if "4 - DADOS DA MATRICULA DO ALUNO" in fold_text(line):
            stop_index = index
            break
    for line in paragraph_lines[:stop_index]:
        cleaned = cleanup_value(line)
        if not cleaned or is_label_line(cleaned):
            continue
        folded = fold_text(cleaned)
        if "PROXIM" in folded or folded.startswith("CASA ") or folded.startswith("RESIDENCIA "):
            return cleanup_point_of_reference(cleaned)
    return ""


def collect_docx_files(input_dir: Path) -> list[Path]:
    files = []
    for path in sorted(input_dir.glob("*.docx")):
        if "FICHA EM BRANCO" in fold_text(path.name):
            continue
        files.append(path)
    return files


def parse_student_docx(docx_path: Path, base_dir: Path | None = None) -> dict[str, str]:
    lines = iter_docx_lines(docx_path)
    paragraph_lines = iter_docx_paragraph_lines(docx_path)
    if not any(
        (
            "EDUCAÇÃO INFANTIL - 2026" in line
            or "EDUCAÇÃO INFANTIL - 2025" in line
            or "FUNDAMENTAL II - 2026" in line
            or "FUNDAMENTAL II - 2025" in line
        )
        for line in lines[:40]
    ):
        raise ValueError("Layout de ficha não suportado pelo extrator atual.")

    ident_lines = section_slice(lines, "2 - DADOS DE IDENTIFICAÇÃO DO ALUNO", "3 - DADOS DO RESPONSÁVEL PELO ALUNO")
    if not ident_lines:
        ident_start = find_index_prefix(lines, "NOME COMPLETO DO ALUNO:")
        ident_end = find_any_index(
            lines,
            [
                "3 - DADOS DO RESPONSÁVEL PELO ALUNO",
                "3 - DADOS DO RESPONSAVEL PELO ALUNO",
                "NOME DO RESPONSÁVEL:",
                "NOME DO RESPONSAVEL:",
            ],
            start=ident_start + 1 if ident_start >= 0 else 0,
        )
        if ident_start >= 0:
            ident_lines = lines[ident_start : ident_end if ident_end >= 0 else len(lines)]
    resp_lines = section_slice(lines, "3 - DADOS DO RESPONSÁVEL PELO ALUNO", "4 - DADOS DA MATRÍCULA DO ALUNO")
    matricula_lines = section_slice(lines, "4 - DADOS DA MATRÍCULA DO ALUNO", "5 - DADOS COMPLEMENTARES DO ALUNO")
    comp_lines = section_slice(lines, "5 - DADOS COMPLEMENTARES DO ALUNO", None)

    mae_start = find_index_prefix(ident_lines, "NOME DA MÃE:")
    pai_start = find_index_prefix(ident_lines, "NOME DO PAI:")
    mae_lines = ident_lines[mae_start:pai_start] if mae_start >= 0 and pai_start >= 0 else []
    pai_lines = ident_lines[pai_start:] if pai_start >= 0 else []

    mae = parse_person_block(mae_lines, "NOME DA MÃE:")
    pai = parse_person_block(pai_lines, "NOME DO PAI:")
    responsavel = parse_person_block(resp_lines, "NOME DO RESPONSÁVEL:", whatsapp_label="WHATSAPP:")

    naturalidade_index = find_index_prefix(ident_lines, "NATURAL DE:")
    uf_naturalidade_index = find_index_prefix(ident_lines, "UF:", start=naturalidade_index if naturalidade_index >= 0 else 0)
    nacionalidade_index = find_index_prefix(ident_lines, "NACIONALIDADE:")
    cartorio_index = find_index(ident_lines, "NOME DO CARTÓRIO:")
    municipio_certidao_index = find_any_index(ident_lines, ["MUNICÍPIO:,", "MUNICÍPIO:"], start=cartorio_index if cartorio_index >= 0 else 0)
    uf_certidao_index = find_index(ident_lines, "UF:", start=municipio_certidao_index if municipio_certidao_index >= 0 else 0)
    zona_index = find_index(ident_lines, "ZONA RESIDENCIAL:")
    municipio_index = find_index(ident_lines, "MUNICÍPIO:", start=zona_index if zona_index >= 0 else 0)
    aluno_uf_index = find_index(ident_lines, "UF:", start=municipio_index if municipio_index >= 0 else 0)
    endereco_index = find_index(ident_lines, "ENDEREÇO DO ALUNO COM O Nº:", start=aluno_uf_index if aluno_uf_index >= 0 else 0)
    if endereco_index < 0:
        endereco_index = find_index(ident_lines, "ENDEREÇO DO ALUNO COM O N°:", start=aluno_uf_index if aluno_uf_index >= 0 else 0)
    if endereco_index < 0:
        endereco_index = find_index(ident_lines, "ENDEREÃ‡O DO ALUNO COM O No:", start=aluno_uf_index if aluno_uf_index >= 0 else 0)
    if endereco_index < 0:
        endereco_index = find_index(ident_lines, "ENDEREÇO DO ALUNO COM O No:", start=aluno_uf_index if aluno_uf_index >= 0 else 0)
    cep_index = find_index(ident_lines, "CEP:", start=endereco_index if endereco_index >= 0 else 0)

    programa_index = find_index(matricula_lines, "PARTICIPANTE DE ALGUM PROGRAMA ESCOLAR:")
    programa_qual_index = find_index(matricula_lines, "QUAL PROGRAMA QUE A CRIANÇA RECEBE APOIO / SUPORTE:", start=programa_index if programa_index >= 0 else 0)
    escola_procedencia_index = find_index(matricula_lines, "ESCOLA DE PROCEDÊNCIA: (Última escola que estudou)", start=programa_qual_index if programa_qual_index >= 0 else 0)
    em_que_ano_index = find_index(matricula_lines, "EM QUE ANO:", start=escola_procedencia_index if escola_procedencia_index >= 0 else 0)
    itinerancia_index = find_index(matricula_lines, "SITUAÇÃO DE INTINERANCIA:", start=em_que_ano_index if em_que_ano_index >= 0 else 0)

    cid_index = find_index(comp_lines, "CID Nº:")
    if cid_index < 0:
        cid_index = find_index(comp_lines, "CID No:")
    medicamento_index = find_index(comp_lines, "FAZ USO DE ALGUM TIPO DE MEDICAMENTO:")
    cuidador_index = find_index(comp_lines, "NECESSITA DE CUIDADOR:")

    marcacoes_identidade_block = lines_between(ident_lines, "TIPO DE CERTIDÃO:", "TERMO:")
    certidao_block = marcacoes_identidade_block
    sexo_block = marcacoes_identidade_block
    cor_block = marcacoes_identidade_block
    zona_block = lines_between(ident_lines, "ZONA RESIDENCIAL:", "MUNICÍPIO:")
    irmaos_block = lines_between(resp_lines, "POSSUI IRMÃOS MATRICULADOS NA REDE MUNICIPAL DE ALAGOA NOVA:", "QUANTIDADE DE IRMÃOS MATRICULADOS NO MUNICÍPIO:")
    serie_block = lines_between(matricula_lines, "SITUAÇÃO DO ALUNO NO ANO ANTERIOR:", "ALUNO EM VULNERABILIDADE SOCIAL:")
    vulnerabilidade_block = lines_between(matricula_lines, "ALUNO EM VULNERABILIDADE SOCIAL:", "EJA:")
    turno_block = lines_between(matricula_lines, "TURNO:", "PARTICIPANTE DE ALGUM PROGRAMA ESCOLAR:")
    programa_block = lines_between(matricula_lines, "PARTICIPANTE DE ALGUM PROGRAMA ESCOLAR:", "QUAL PROGRAMA QUE A CRIANÇA RECEBE APOIO / SUPORTE:")
    itinerancia_block = lines_between(matricula_lines, "SITUAÇÃO DE INTINERANCIA:", "ENDEREÇO:")
    deficiencia_block = lines_between(comp_lines, "TIPO:", "RECURSOS PARA AVALIAÇÕES DO INEP:")
    pessoa_deficiencia_block = comp_lines[:cid_index] if cid_index >= 0 else comp_lines
    atendimento_block = lines_between(comp_lines, "RECEBE ATENDIMENTO ESPECIALIZADO:", "INSTITUIÇÃO QUE RECEBE ATENDIMENTO ESPECIALIZADO:")
    cuidador_block = comp_lines[cuidador_index + 1 : medicamento_index] if cuidador_index >= 0 and medicamento_index > cuidador_index else []
    transporte_publico_block = lines_between(comp_lines, "PÚBLICO:", "ESTADO (")
    if not transporte_publico_block:
        transporte_publico_block = lines_between(comp_lines, "PÚBLICO:", "Estado (")

    beneficio_tipo_index = find_index(comp_lines, "TIPO DE BENEFÍCIO:")
    beneficio_tipo = cleanup_value(next_value(comp_lines, beneficio_tipo_index))

    quantidade_index = find_index(resp_lines, "QUANTIDADE DE IRMÃOS MATRICULADOS NO MUNICÍPIO:")
    escolas_index = find_index(resp_lines, "EM QUE ESCOLA(S) O(S) IRMÃO(S) ESTÃO MATRICULADOS:")

    emissao_indices = find_all_indices(ident_lines, "DATA DA EMISSÃO:")
    certidao_municipio = cleanup_value(next_value(ident_lines, municipio_certidao_index, uf_certidao_index))

    record = {column: "" for column in CSV_COLUMNS}
    extracted_student_name = cleanup_value(value_after_label(ident_lines, find_index_prefix(ident_lines, "NOME COMPLETO DO ALUNO:")))
    filename_student_name = extract_name_from_filename(docx_path)
    use_filename_student_name = should_replace_student_name_with_filename(extracted_student_name, filename_student_name)
    record["source_file"] = str(docx_path.relative_to(base_dir) if base_dir else docx_path)
    record["aluno_nome_original"] = extracted_student_name
    record["aluno_nome_corrigido_por_arquivo"] = "Sim" if use_filename_student_name else "N?o"
    record["aluno_nome"] = filename_student_name if use_filename_student_name else extracted_student_name
    record["aluno_data_nascimento"] = cleanup_value(value_after_label(ident_lines, find_index_prefix(ident_lines, "DATA DE NASCIMENTO:")))
    aluno_nome_social = cleanup_value(value_after_label(ident_lines, find_index_prefix(ident_lines, "NOME SOCIAL:")))
    record["aluno_nome_social"] = aluno_nome_social if is_plausible_name_text(aluno_nome_social) else ""
    header_numeric_fields = extract_header_numeric_fields(ident_lines)
    record["aluno_id"] = cleanup_value(
        value_after_label(
            ident_lines,
            find_any_index_prefix(ident_lines, ["N DO ID DO ALUNO :", "NO DO ID DO ALUNO :"]),
        )
    ) or header_numeric_fields["aluno_id"]
    record["aluno_nis"] = cleanup_value(
        value_after_label(
            ident_lines,
            find_any_index_prefix(
                ident_lines,
                [
                    "NUMERO DE IDENTIFICACAO SOCIAL (NIS):",
                    "NUMERO DE IDENTIFICACAO SOCIAL (NIS)",
                    "N?MERO DE IDENTIFICA??O SOCIAL (NIS):",
                    "N?MERO DE IDENTIFICA??O SOCIAL (NIS)",
                ],
            ),
        )
    ) or header_numeric_fields["aluno_nis"]
    record["aluno_cartao_sus"] = cleanup_value(
        value_after_label(
            ident_lines,
            find_any_index_prefix(
                ident_lines,
                [
                    "NO DO CARTAO SUS:",
                    "NO DO CARTAO SUS",
                    "N DO CARTAO SUS:",
                    "N DO CARTAO SUS",
                    "N? DO CART?O SUS:",
                    "No DO CART?O SUS:",
                ],
            ),
        )
    ) or header_numeric_fields["aluno_cartao_sus"]
    if not record["aluno_cartao_sus"]:
        record["aluno_cartao_sus"] = cleanup_value(
            value_after_label(
                ident_lines,
                find_any_index_prefix(
                    ident_lines,
                    [
                        "NO DO CARTAO SUS:",
                        "NO DO CARTAO SUS",
                        "N DO CARTAO SUS:",
                        "N DO CARTAO SUS",
                        "N? DO CART?O SUS:",
                        "No DO CART?O SUS:",
                    ],
                ),
            )
        ) or header_numeric_fields["aluno_cartao_sus"]
    record["aluno_naturalidade"] = cleanup_value(value_after_label(ident_lines, naturalidade_index, uf_naturalidade_index))
    record["aluno_uf_naturalidade"] = cleanup_value(value_after_label(ident_lines, uf_naturalidade_index, nacionalidade_index))
    record["aluno_nacionalidade"] = cleanup_value(value_after_label(ident_lines, nacionalidade_index))
    record["aluno_sexo"] = parse_marked_option(sexo_block, ["Masculino", "Feminino"])
    record["aluno_cor_raca"] = parse_marked_option(cor_block, ["Branca", "Preta", "Parda", "Amarela", "Indígena"])
    record["certidao_tipo"] = parse_marked_option(certidao_block, ["Nascimento", "Casamento"])
    record["certidao_numero"] = cleanup_value(extract_certidao_numero_bloco(ident_lines))
    record["certidao_emissao"] = cleanup_value(extract_certidao_emissao_bloco(ident_lines))
    record["certidao_cartorio"] = cleanup_value(next_value(ident_lines, cartorio_index, municipio_certidao_index))
    record["certidao_municipio"] = strip_trailing_uf(certidao_municipio)
    record["certidao_uf"] = cleanup_value(next_value(ident_lines, uf_certidao_index))
    record["aluno_cpf"] = cleanup_value(extract_student_cpf(ident_lines, zona_index))
    record["aluno_zona"] = parse_marked_option(zona_block, ["Urbana", "Rural"])
    record["aluno_endereco"] = cleanup_value(next_value(ident_lines, endereco_index, cep_index))
    record["aluno_municipio"] = cleanup_value(next_value(ident_lines, municipio_index, aluno_uf_index))
    record["aluno_uf"] = cleanup_value(next_value(ident_lines, aluno_uf_index, endereco_index))
    record["aluno_cep"] = cleanup_value(next_value(ident_lines, cep_index, mae_start if mae_start >= 0 else None))
    record["ponto_referencia"] = extract_point_of_reference(paragraph_lines)
    record["telefone_contato"] = preferred_phone(mae.get("telefone", ""), pai.get("telefone", ""), responsavel.get("whatsapp", ""))

    record["mae_nome"] = mae["nome"]
    record["mae_endereco"] = mae["endereco"]
    record["mae_municipio"] = mae["municipio"]
    record["mae_uf"] = mae["uf"]
    record["mae_telefone"] = mae["telefone"]
    record["mae_rg"] = preferred_document(mae["cpf"], mae["rg"])
    record["mae_orgao_emissor"] = mae["orgao_emissor"]
    record["mae_data_emissao"] = mae["data_emissao"]
    record["mae_cpf"] = mae["cpf"]
    record["mae_profissao"] = mae["profissao"]

    record["pai_nome"] = pai["nome"]
    record["pai_endereco"] = pai["endereco"]
    record["pai_municipio"] = pai["municipio"]
    record["pai_uf"] = pai["uf"]
    record["pai_telefone"] = pai["telefone"]
    record["pai_rg"] = preferred_document(pai["cpf"], pai["rg"])
    record["pai_orgao_emissor"] = pai["orgao_emissor"]
    record["pai_data_emissao"] = pai["data_emissao"]
    record["pai_cpf"] = pai["cpf"]
    record["pai_profissao"] = pai["profissao"]

    record["responsavel_nome"] = responsavel["nome"]
    record["responsavel_endereco"] = responsavel["endereco"]
    record["responsavel_municipio"] = responsavel["municipio"]
    record["responsavel_uf"] = responsavel["uf"]
    record["responsavel_whatsapp"] = responsavel["whatsapp"]
    record["responsavel_parentesco"] = responsavel["parentesco"]
    record["responsavel_email"] = responsavel["email"]

    record["irmaos_rede_municipal"] = parse_marked_option(irmaos_block, ["Sim", "Não"])
    record["irmaos_quantidade"] = cleanup_value(next_value(resp_lines, quantidade_index, escolas_index))
    record["irmaos_escolas"] = cleanup_value(next_value(resp_lines, escolas_index))

    record["matricula_etapa_serie"] = parse_marked_option(serie_block, SERIE_OPTIONS)
    record["matricula_turno"] = parse_marked_option(turno_block, ["M", "T", "N"])
    record["vulnerabilidade_social"] = parse_marked_option(vulnerabilidade_block, ["Sim", "Não"])
    record["bolsa_familia"] = "Sim" if "BOLSA FAM" in fold_text(beneficio_tipo) else "Não"
    record["utiliza_transporte"] = parse_marked_option(transporte_publico_block, ["Sim", "Não"])
    record["programa_escolar"] = parse_marked_option(programa_block, ["Sim", "Não"])
    record["programa_escolar_qual"] = cleanup_value(next_value(matricula_lines, programa_qual_index, escola_procedencia_index))
    record["escola_procedencia"] = cleanup_value(next_value(matricula_lines, escola_procedencia_index, em_que_ano_index))
    record["itinerancia"] = parse_marked_option(itinerancia_block, ["Sim", "Não"])

    record["pessoa_com_deficiencia"] = parse_marked_option(pessoa_deficiencia_block, ["Sim", "Não"], prefer_last=True)
    record["deficiencia_tipos"] = "; ".join(parse_marked_options(deficiencia_block, DEFICIENCIA_OPTIONS))
    record["atendimento_especializado"] = parse_marked_option(atendimento_block, ["Sim", "Não"], prefer_last=True)
    record["cuidador"] = parse_marked_option(cuidador_block, ["Sim", "Não"], prefer_last=True)
    record["observacoes_raw"] = extract_observacoes(comp_lines)

    cleaned_record = {key: cleanup_value(value) for key, value in record.items()}
    return format_record(cleaned_record)
