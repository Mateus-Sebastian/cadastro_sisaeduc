from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PLACEHOLDER_IMAGE_SIZES = {10028}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extrai a foto do aluno dos .docx e salva em uma pasta dedicada."
    )
    parser.add_argument(
        "--input-dir",
        default="matriculas",
        help="Pasta com os .docx de matrícula.",
    )
    parser.add_argument(
        "--output-dir",
        default="matriculas/fotos_dos_alunos",
        help="Pasta onde as fotos serão salvas.",
    )
    return parser


def collect_docx_files(input_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in input_dir.glob("*.docx")
        if "FICHA EM BRANCO" not in path.name.upper()
    )


def choose_photo_target(doc_xml: bytes, rels_xml: bytes, archive: ZipFile) -> str | None:
    doc = ET.fromstring(doc_xml)
    rels = ET.fromstring(rels_xml)
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    candidates: list[tuple[int, int, int, str]] = []
    containers = doc.findall(".//wp:anchor", NS) + doc.findall(".//wp:inline", NS)
    for container in containers:
        doc_pr = container.find("wp:docPr", NS)
        c_nv_pr = container.find(".//pic:cNvPr", NS)
        extent = container.find("wp:extent", NS)
        blip = container.find(".//a:blip", NS)
        if blip is None:
            continue
        rel_id = blip.attrib.get(DOC_REL_NS + "embed")
        if not rel_id or rel_id not in rel_map:
            continue
        target = "word/" + rel_map[rel_id].lstrip("/")
        if not target.startswith("word/media/"):
            continue
        size = archive.getinfo(target).file_size
        if size in PLACEHOLDER_IMAGE_SIZES:
            continue
        name_parts = []
        if doc_pr is not None:
            name_parts.append(doc_pr.attrib.get("name", ""))
        if c_nv_pr is not None:
            name_parts.append(c_nv_pr.attrib.get("name", ""))
        name = " ".join(part.strip() for part in name_parts if part).upper()
        cx = int(extent.attrib.get("cx", "0")) if extent is not None else 0
        image_priority = 0 if "IMAGEM" in name else 1
        candidates.append((image_priority, -size, -cx, target))

    if not candidates:
        return None

    candidates.sort()
    return candidates[0][3]


def extract_photo(docx_path: Path) -> tuple[str | None, bytes | None]:
    with ZipFile(docx_path) as archive:
        doc_xml = archive.read("word/document.xml")
        rels_xml = archive.read("word/_rels/document.xml.rels")
        target = choose_photo_target(doc_xml, rels_xml, archive)
        if not target:
            return None, None
        suffix = Path(target).suffix.lower() or ".jpg"
        return suffix, archive.read(target)


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    if not input_dir.exists():
        parser.error(f"Pasta não encontrada: {input_dir}")

    docx_files = collect_docx_files(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    extracted = 0
    skipped: list[str] = []

    for docx_path in docx_files:
        suffix, payload = extract_photo(docx_path)
        if not suffix or payload is None:
            skipped.append(docx_path.name)
            continue
        output_path = output_dir / f"{docx_path.stem}{suffix}"
        output_path.write_bytes(payload)
        extracted += 1

    print(f"Fotos extraídas: {extracted}")
    print(f"Pasta de saída: {output_dir}")
    if skipped:
        print(f"Sem foto extraída: {len(skipped)}")
        for name in skipped:
            print(f"- {name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
