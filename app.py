from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple



STANDARD_TITLES = [
    "Resumo Executivo",
    "Objetivo Geral",
    "Metodologia",
    "Análise dos Resultados",
    "Conclusão e Recomendações",
]

TITLE_ALIASES = {
    "Resumo Executivo": [
        "resumo executivo",
        "resumo",
        "sumario executivo",
    ],
    "Objetivo Geral": [
        "objetivo geral",
        "objetivo",
        "objetivos",
    ],
    "Metodologia": [
        "metodologia",
        "metodo",
        "método",
    ],
    "Análise dos Resultados": [
        "analise dos resultados",
        "análise dos resultados",
        "resultados",
        "analise",
    ],
    "Conclusão e Recomendações": [
        "conclusao e recomendacoes",
        "conclusão e recomendações",
        "conclusao",
        "recomendacoes",
        "recomendações",
    ],
}


@dataclass
class ReportSection:
    title: str
    content: str



def _load_reportlab():
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "Dependência ausente: instale 'reportlab' para gerar PDF (pip install -r requirements.txt)."
        ) from exc

    return {
        "colors": colors,
        "A4": A4,
        "ParagraphStyle": ParagraphStyle,
        "getSampleStyleSheet": getSampleStyleSheet,
        "cm": cm,
        "Paragraph": Paragraph,
        "SimpleDocTemplate": SimpleDocTemplate,
        "Spacer": Spacer,
        "Table": Table,
        "TableStyle": TableStyle,
    }


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.lower()
    value = re.sub(r"\s+", " ", value).strip()
    value = re.sub(r"[^a-z0-9 ]", "", value)
    return value


def resolve_standard_title(raw_title: str) -> str:
    normalized_raw = normalize_text(raw_title)

    for standard, aliases in TITLE_ALIASES.items():
        alias_set = {normalize_text(standard), *(normalize_text(a) for a in aliases)}
        if normalized_raw in alias_set:
            return standard

    for standard in STANDARD_TITLES:
        if normalized_raw == normalize_text(standard):
            return standard

    return raw_title.strip().title()


def parse_sections(raw_text: str) -> List[ReportSection]:
    pattern = re.compile(r"^\s*([^:\n]{2,80})\s*:\s*(.*)$")
    sections: Dict[str, List[str]] = {}
    current_title = "Informações Gerais"

    for line in raw_text.splitlines():
        match = pattern.match(line)
        if match:
            current_title = resolve_standard_title(match.group(1))
            content = match.group(2).strip()
            sections.setdefault(current_title, [])
            if content:
                sections[current_title].append(content)
        else:
            clean_line = line.strip()
            if clean_line:
                sections.setdefault(current_title, []).append(clean_line)

    ordered: List[Tuple[str, List[str]]] = []
    for title in STANDARD_TITLES:
        if title in sections:
            ordered.append((title, sections.pop(title)))

    for title, content in sections.items():
        ordered.append((title, content))

    return [ReportSection(title=title, content="\n".join(content)) for title, content in ordered]


def build_styles(rl):
    styles = rl["getSampleStyleSheet"]()
    return {
        "cover_title": rl["ParagraphStyle"](
            "CoverTitle",
            parent=styles["Heading1"],
            fontSize=26,
            leading=30,
            textColor=rl["colors"].HexColor("#0F2A43"),
            alignment=1,
            spaceAfter=12,
        ),
        "cover_subtitle": rl["ParagraphStyle"](
            "CoverSubtitle",
            parent=styles["Normal"],
            fontSize=12,
            leading=16,
            textColor=rl["colors"].HexColor("#38566F"),
            alignment=1,
        ),
        "heading": rl["ParagraphStyle"](
            "SectionHeading",
            parent=styles["Heading2"],
            fontSize=14,
            textColor=rl["colors"].HexColor("#133C5A"),
            spaceBefore=14,
            spaceAfter=8,
        ),
        "body": rl["ParagraphStyle"](
            "SectionBody",
            parent=styles["BodyText"],
            fontSize=11,
            leading=16,
            textColor=rl["colors"].HexColor("#1C2A39"),
            alignment=4,
        ),
        "toc": rl["ParagraphStyle"](
            "TOC",
            parent=styles["Normal"],
            fontSize=10,
            leading=14,
            textColor=rl["colors"].HexColor("#2F4B63"),
        ),
    }


def draw_header_footer(canvas, doc, rl):
    canvas.saveState()
    canvas.setStrokeColor(rl["colors"].HexColor("#D2D9E1"))
    canvas.setLineWidth(0.5)
    canvas.line(doc.leftMargin, rl["A4"][1] - 1.8 * rl["cm"], rl["A4"][0] - doc.rightMargin, rl["A4"][1] - 1.8 * rl["cm"])

    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(rl["colors"].HexColor("#4A6178"))
    canvas.drawString(doc.leftMargin, rl["A4"][1] - 1.4 * rl["cm"], "Relatório Técnico")
    canvas.drawRightString(rl["A4"][0] - doc.rightMargin, rl["A4"][1] - 1.4 * rl["cm"], datetime.now().strftime("%d/%m/%Y"))

    canvas.line(doc.leftMargin, 1.7 * rl["cm"], rl["A4"][0] - doc.rightMargin, 1.7 * rl["cm"])
    canvas.setFillColor(rl["colors"].HexColor("#4A6178"))
    canvas.drawString(doc.leftMargin, 1.2 * rl["cm"], "Documento gerado automaticamente")
    canvas.drawRightString(rl["A4"][0] - doc.rightMargin, 1.2 * rl["cm"], f"Página {doc.page}")
    canvas.restoreState()


def generate_pdf(input_text: str, output_path: Path):
    rl = _load_reportlab()
    sections = parse_sections(input_text)
    styles = build_styles(rl)

    doc = rl["SimpleDocTemplate"](
        str(output_path),
        pagesize=rl["A4"],
        leftMargin=2.2 * rl["cm"],
        rightMargin=2.2 * rl["cm"],
        topMargin=2.6 * rl["cm"],
        bottomMargin=2.2 * rl["cm"],
        title="Relatório",
    )

    story = []
    story.append(rl["Spacer"](1, 5 * rl["cm"]))
    story.append(rl["Paragraph"]("Relatório", styles["cover_title"]))
    story.append(rl["Paragraph"]("Modelo profissional com estrutura padronizada", styles["cover_subtitle"]))
    story.append(rl["Spacer"](1, 1.2 * rl["cm"]))

    info_table = rl["Table"](
        [
            ["Data de geração", datetime.now().strftime("%d/%m/%Y %H:%M")],
            ["Total de seções", str(len(sections))],
        ],
        colWidths=[5 * rl["cm"], 8 * rl["cm"]],
    )
    info_table.setStyle(
        rl["TableStyle"](
            [
                ("BACKGROUND", (0, 0), (-1, -1), rl["colors"].HexColor("#F4F7FA")),
                ("BOX", (0, 0), (-1, -1), 0.8, rl["colors"].HexColor("#C4CFDA")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, rl["colors"].HexColor("#D5DEE7")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 0), (-1, -1), rl["colors"].HexColor("#25394D")),
            ]
        )
    )
    story.append(info_table)
    story.append(rl["Spacer"](1, 2 * rl["cm"]))

    story.append(rl["Paragraph"]("Sumário de Seções", styles["heading"]))
    for idx, section in enumerate(sections, start=1):
        story.append(rl["Paragraph"](f"{idx}. {section.title}", styles["toc"]))

    story.append(rl["Spacer"](1, 0.8 * rl["cm"]))

    for section in sections:
        story.append(rl["Paragraph"](section.title, styles["heading"]))
        body = section.content.replace("\n", "<br/>")
        story.append(rl["Paragraph"](body, styles["body"]))

    doc.build(
        story,
        onFirstPage=lambda canvas, doc: draw_header_footer(canvas, doc, rl),
        onLaterPages=lambda canvas, doc: draw_header_footer(canvas, doc, rl),
    )


def main():
    parser = argparse.ArgumentParser(description="Gerador de relatório PDF profissional")
    parser.add_argument("--input", required=True, help="Arquivo texto de entrada")
    parser.add_argument("--output", required=True, help="Arquivo PDF de saída")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    text = input_path.read_text(encoding="utf-8")
    generate_pdf(text, output_path)
    print(f"PDF gerado com sucesso em: {output_path}")


if __name__ == "__main__":
    main()
