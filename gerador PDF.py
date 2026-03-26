import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
import json, os, re, unicodedata, threading, tempfile, subprocess, sys

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import *
from reportlab.lib.styles import *
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.pdfbase import pdfmetrics

from pypdf import PdfReader, PdfWriter
import fitz  # pymupdf
from PIL import Image as PILImage, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

CONFIG_FILE = "config.json"
SECOES_EDITAVEIS = ["descricao", "detalhamento", "diagnostico", "acoes", "resultado", "estado"]
SECOES_PAINEL = ["cabecalho", *SECOES_EDITAVEIS]

SECTION_HEADERS = {
    "descricao": "1 – ESCOPO DO ATENDIMENTO",
    "detalhamento": "2 – DETALHAMENTO DO PROBLEMA",
    "diagnostico": "3 – DIAGNÓSTICO",
    "acoes": "4 – AÇÕES CORRETIVAS",
    "resultado": "5 – RESULTADO",
    "estado": "6 – ESTADO FINAL",
}

PHOTO_LAYOUT_MODES = ["Dividir página", "Página inteira"]
WATERMARK_MODES = ["Central", "Timbrado aleatório", "Central + Timbrado aleatório"]
WATERMARK_RANDOM_COUNT_OPTIONS = ["4", "6", "8", "10", "12", "14", "16", "20"]
HORARIOS_DESCRICAO_PRESETS = ["", "Hora técnica", "Deslocamento"]
DOCUMENT_TYPES = ["Relatório técnico", "Certificado de treinamento"]

INFO_ALIASES = {
    "equipamento": "equipamento",
    "fonte": "fonte",
    "cliente": "cliente",
    "empresa": "cliente",
    "empresa cliente": "cliente",
    "empresa / cliente": "cliente",
    "empresa/cliente": "cliente",
    "cnc": "cnc",
    "thc": "thc",
    "fabricante": "fabricante",
    "contato cliente": "contato_cliente",
    "data": "data",
    "tecnico": "tecnico",
    "técnico": "tecnico",
    "atendente": "tecnico",
    "atendente / tecnico": "tecnico",
    "atendente/tecnico": "tecnico",
    "acompanhamento remoto": "acompanhamento",
    "acompanhamento": "acompanhamento",
    "horario de inicio": "inicio",
    "horário de início": "inicio",
    "horario de início": "inicio",
    "inicio": "inicio",
    "início": "inicio",
    "horario de termino": "fim",
    "horário de término": "fim",
    "fim": "fim",
    "tempo do atendimento": "tempo_atendimento",
    "tempo atendimento": "tempo_atendimento",
    "tempo em espera": "tempo_espera",
    "tempo espera": "tempo_espera",
    "contato": "contato_cliente",
    "nome do cliente": "cliente",
    "contato no cliente": "contato_cliente",
    "modelo da fonte": "modelo_fonte",
    "controle de altura": "controle_altura",
    "marca do cnc": "marca_cnc",
    "data inicio": "data_inicio",
    "data início": "data_inicio",
    "data final": "data_final",
    "tecnico responsavel": "tecnico",
    "técnico responsável": "tecnico",
    "motivo do chamado": "motivo_chamado",
}

INFO_FIELDS = [
    ("Equipamento", "equipamento"),
    ("Fonte", "fonte"),
    ("Empresa / Cliente", "cliente"),
    ("CNC", "cnc"),
    ("THC", "thc"),
    ("Fabricante", "fabricante"),
    ("Contato Cliente", "contato_cliente"),
    ("Data", "data"),
    ("Atendente / Técnico", "tecnico"),
    ("Acompanhamento remoto", "acompanhamento"),
    ("Início", "inicio"),
    ("Fim", "fim"),
    ("Tempo Atendimento", "tempo_atendimento"),
    ("Tempo Espera", "tempo_espera"),
]
MANDATORY_INFO_FIELDS = ("tecnico", "cliente")

HEADER_DEFAULT_ROWS = [
    ("Nome do Cliente", "cliente"),
    ("Contato no cliente", "contato_cliente"),
    ("Equipamento", "equipamento"),
    ("Modelo da Fonte", "modelo_fonte"),
    ("Controle de Altura", "controle_altura"),
    ("Marca do CNC", "marca_cnc"),
    ("Fabricante", "fabricante"),
    ("Data início", "data_inicio"),
    ("Data final", "data_final"),
    ("Técnico Responsável", "tecnico"),
    ("Motivo do chamado", "motivo_chamado"),
]

MAX_RECENT = 7


# =========================
# CONFIG
# =========================
def load_config():
    defaults = {
        "template_path": "",
        "last_dir": "",
        "pdf_save_dir": "",
        "image_dir": "",
        "preview_visible": True,
        "preview_auto": True,
        "zoom_factor": 1.0,
        "window_geometry": "1280x750",
        "default_tecnico": "",
        "foto_cols": "2",
        "foto_max_height_cm": "8.1",
        "section_offsets_cm": {},
        "watermark_path": "",
        "watermark_opacity": "0.12",
        "watermark_scale": "0.75",
        "watermark_mode": "Central",
        "watermark_random_count": "8",
        "cover_header_scale": "1.8",
        "signature_page": True,
        "recent_pdfs": [],
    }
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            defaults.update(data)
    return defaults


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def add_recent_pdf(config, path):
    recents = config.get("recent_pdfs", [])
    if path in recents:
        recents.remove(path)
    recents.insert(0, path)
    config["recent_pdfs"] = recents[:MAX_RECENT]
    save_config(config)


# =========================
# PARSER  (melhorado)
# =========================
def normalize(text):
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c)).lower().strip()


def _section_number(text):
    """Extrai o número inicial de um título de seção normalizado, ex: '1' de '1 – escopo...'"""
    m = re.match(r"^(\d+)", normalize(text))
    return m.group(1) if m else None


# Mapeamento: número da seção → chave interna
SECTION_NUMBER_MAP = {
    "1": "descricao",
    "2": "detalhamento",
    "3": "diagnostico",
    "4": "acoes",
    "5": "resultado",
    "6": "estado",
}


def parse_text(text):
    sections = {"info": {}, "info_extra": []}

    for line in text.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            key = normalize(k)
            value = v.strip()
            if not value:
                continue
            if key in INFO_ALIASES:
                field_key = INFO_ALIASES[key]
                sections["info"][field_key] = value
                # Auto-populate header fields from parsed info
                _try_autofill_header(sections, field_key, value)
            else:
                sections["info_extra"].append((k.strip(), value))

    # Split on numbered section headings (tolerant: "1 –", "1 -", "1." etc.)
    parts = re.split(r"\n\s*(\d+\s*[–\-\.]\s*.+)", text)

    for i in range(1, len(parts), 2):
        title_raw = parts[i]
        content = parts[i + 1].strip()
        num = _section_number(title_raw)
        if num and num in SECTION_NUMBER_MAP:
            key = SECTION_NUMBER_MAP[num]
            sections[key] = content

    return sections


def _try_autofill_header(sections, field_key, value):
    """Tenta preencher campos do cabeçalho a partir de campos de info parseados."""
    mapping = {
        "inicio": "data_inicio",
        "fim": "data_final",
        "data": "data_inicio",
        "motivo_chamado": "motivo_chamado",
    }
    if field_key in mapping:
        target = mapping[field_key]
        if not sections["info"].get(target):
            sections["info"][target] = value


def limpar_texto(texto):
    texto = re.sub(r"\*\*(.*?)\*\*", r"\1", texto)
    texto = re.sub(r"#+\s*", "", texto)
    texto = re.sub(r"---+", "", texto)
    texto = re.sub(r"━+", "", texto)          # ← remove separadores ━━━
    texto = re.sub(r"📄\s*", "", texto)       # ← remove emoji de cabeçalho
    texto = re.sub(r"\*(\s*)", "", texto)

    substituicoes = {
        "\u201c": '"', "\u201d": '"',
        "\u2018": "'",
        "\u2013": "-", "\u2014": "-",
        "\u2022": "-", "\xa0": " ",
    }
    for k, v in substituicoes.items():
        texto = texto.replace(k, v)

    return texto


# =========================
# LISTAS AUTOMÁTICAS  (melhoradas — suporte a listas aninhadas)
# =========================
def _linha_tem_topico(linha):
    return bool(re.match(r"^(\d+[\.\)]|-|•|\u2022|°)\s*", linha.strip()))


def _limpar_marcador_topico(linha):
    return re.sub(r"^(\d+[\.\)]|-|•|\u2022|°)\s*", "", linha.strip())


def _valor_info_preenchido(valor):
    if valor is None:
        return False
    texto = str(valor).strip()
    return bool(texto and texto != ".")


def processar_lista(texto, styles):
    """
    Processa texto em elementos ReportLab com suporte a:
    - Listas com marcadores explícitos (-, •, 1., 2.)
    - Listas numeradas com indentação
    - Parágrafos separados por linha em branco
    - Sub-grupos separados por ponto final
    """
    elementos = []
    if not texto or not texto.strip():
        return [Spacer(1, 10)]

    # Separar por linhas em branco em blocos
    blocos = re.split(r"\n\s*\n", texto.strip())

    for bloco in blocos:
        linhas = [l.strip() for l in bloco.splitlines() if l.strip()]
        if not linhas:
            continue

        tem_topicos = any(_linha_tem_topico(l) for l in linhas)

        if tem_topicos:
            itens = []
            item_atual = ""

            for linha in linhas:
                if _linha_tem_topico(linha):
                    if item_atual:
                        itens.append(item_atual.strip())
                    item_atual = _limpar_marcador_topico(linha)
                else:
                    item_atual += " " + linha if item_atual else linha

            if item_atual:
                itens.append(item_atual.strip())

            if itens:
                lista = [
                    ListItem(
                        Paragraph(item, styles["Body"]),
                        leftIndent=12,
                        bulletText="°"
                    )
                    for item in itens if item
                ]
                elementos.append(
                    ListFlowable(
                        lista,
                        bulletType='bullet',
                        leftIndent=8,
                        bulletFontName="Helvetica",
                        bulletFontSize=10
                    )
                )
        else:
            # Texto corrido: verificar se tem múltiplas frases para separar
            texto_bloco = " ".join(linhas)
            # Dividir por ponto + maiúscula como heurística de frases
            frases = re.split(r'(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÀÃÕÇ])', texto_bloco)
            for frase in frases:
                frase = frase.strip()
                if frase:
                    elementos.append(Paragraph(frase, styles["Body"]))
                    elementos.append(Spacer(1, 3))

        elementos.append(Spacer(1, 6))

    return elementos if elementos else [Spacer(1, 10)]


# =========================
# PDF
# =========================
def _coerce_offset_cm(value):
    try:
        return max(0.0, float(value))
    except Exception:
        return 0.0


def _extract_section_body(raw_text, section_key):
    text = (raw_text or "").strip()
    if not text:
        return ""
    header = SECTION_HEADERS.get(section_key, "")
    normalized_header = normalize(header) if header else ""
    lines = text.splitlines()
    if lines:
        first = lines[0].strip().lstrip("#").strip()
        first = re.sub(r"^\*+|\*+$", "", first).strip()
        if normalized_header and normalize(first) == normalized_header:
            return "\n".join(lines[1:]).strip()
    return text


def _compose_full_text_with_sections(base_text, sections):
    header_pattern = re.compile(r"^\s*\d+\s*[–-]\s*.+$")
    if sections.get("info") or sections.get("info_extra"):
        info_lines = _compose_info_text(
            sections.get("info", {}),
            sections.get("info_extra", []),
        ).splitlines()
    else:
        info_lines = []
        for line in (base_text or "").splitlines():
            if ":" in line and not header_pattern.match(line.strip()):
                info_lines.append(line.rstrip())

    blocks = []
    if info_lines:
        blocks.append("\n".join(info_lines).strip())

    for key in SECOES_EDITAVEIS:
        body = (sections.get(key) or "").strip()
        if not body:
            continue
        header = SECTION_HEADERS.get(key, key.title())
        blocks.append(f"{header}\n{body}".strip())

    return "\n\n".join(block for block in blocks if block).strip()


def _compose_info_text(info, info_extra=None, include_required_empty=False):
    info = info or {}
    info_extra = info_extra or []
    lines = []
    for label, key in INFO_FIELDS:
        value = info.get(key)
        if _valor_info_preenchido(value):
            lines.append(f"{label}: {str(value).strip()}")
        elif include_required_empty and key in MANDATORY_INFO_FIELDS:
            lines.append(f"{label}: ")
    for label, value in info_extra:
        if _valor_info_preenchido(value):
            lines.append(f"{str(label).strip()}: {str(value).strip()}")
    return "\n".join(lines).strip()


def _parse_info_from_editor(text):
    info = {}
    info_extra = []
    for line in (text or "").splitlines():
        if ":" not in line:
            continue
        k, v = line.split(":", 1)
        normalized_key = normalize(k)
        value = v.strip()
        if not value:
            continue
        info_key = INFO_ALIASES.get(normalized_key)
        if info_key:
            info[info_key] = value
        else:
            info_extra.append((k.strip(), value))
    return info, info_extra


def _parse_header_info_text(text):
    info, _ = _parse_info_from_editor(text)
    return info


def _compose_header_rows(info, info_extra=None):
    info = info or {}
    rows = []
    for label, key in HEADER_DEFAULT_ROWS:
        value = str(info.get(key, "") or "").strip()
        rows.append([label, value])
    return rows


def _parse_header_rows(rows):
    info = {}
    info_extra = []
    header_rows = []
    for row in rows or []:
        label, value = (list(row) + ["", ""])[:2]
        label = str(label or "").strip()
        value = str(value or "").strip()
        if not label:
            continue
        header_rows.append([label, value])
        normalized_key = normalize(label)
        info_key = INFO_ALIASES.get(normalized_key)
        if info_key:
            if value:
                info[info_key] = value
        elif value:
            info_extra.append((label, value))
    return info, info_extra, header_rows


def _parse_horarios_table(raw_text):
    def _normalize_date(text):
        t = str(text or "").strip()
        if not t:
            return ""
        digits = re.sub(r"\D", "", t)
        if len(digits) == 8:
            return f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"
        if len(digits) == 6:
            return f"{digits[:2]}/{digits[2:4]}/20{digits[4:]}"
        return t

    def _normalize_time(text):
        t = str(text or "").strip()
        if not t:
            return ""
        digits = re.sub(r"\D", "", t)
        if len(digits) >= 4:
            hh = digits[:2]
            mm = digits[2:4]
            return f"{hh}:{mm}"
        if len(digits) <= 2:
            return f"{digits.zfill(2)}:00"
        return t

    linhas = []
    for raw_line in (raw_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.lower().startswith("data"):
            continue
        parts = [p.strip() for p in re.split(r"[|;]", line) if p.strip()]
        if len(parts) < 3:
            parts = [p.strip() for p in re.split(r"\s+", line) if p.strip()]
        if len(parts) < 3:
            continue
        if len(parts) == 3:
            parts.append("")
        if len(parts) == 4:
            parts.append("")
        date, inicio, fim, intervalos, descricao = parts[:5]
        # ← MELHORIA: só inclui linha se tiver pelo menos data ou horário válido
        if not date and not inicio and not fim:
            continue
        linhas.append([
            _normalize_date(date),
            _normalize_time(inicio),
            _normalize_time(fim),
            intervalos,
            descricao,
        ])
    return linhas


def _compose_horarios_table_text(rows=None):
    rows = rows or []
    lines = ["Data | Início | Fim | Intervalos | Descrição"]
    for row in rows:
        date, inicio, fim, intervalos, descricao = (list(row) + ["", "", "", "", ""])[:5]
        lines.append(f"{date} | {inicio} | {fim} | {intervalos} | {descricao}")
    return "\n".join(lines)


# =========================
# PÁGINA DE ASSINATURA
# =========================
def _build_signature_page(styles, info=None):
    """Constrói a página de assinatura do técnico e cliente."""
    info = info or {}
    tecnico = str(info.get("tecnico", "") or "").strip()
    cliente = str(info.get("cliente", "") or "").strip()

    story = []
    story.append(Spacer(1, 2.5 * cm))
    story.append(Paragraph("<b>ASSINATURA E CONFIRMAÇÃO DO ATENDIMENTO</b>", styles["Secao"]))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(
        "Declaro que os serviços descritos neste relatório foram realizados conforme acordado "
        "e que as informações aqui contidas são verdadeiras.",
        styles["Body"]
    ))
    story.append(Spacer(1, 1.8 * cm))

    largura_util = A4[0] - (5.0 * cm)
    col_w = (largura_util - 1.5 * cm) / 2

    linha_style = ParagraphStyle(
        "AssinaturaLabel",
        parent=styles["Body"],
        fontSize=9.5,
        textColor=colors.HexColor("#4A5568"),
        alignment=1,
    )
    nome_style = ParagraphStyle(
        "AssinaturaNome",
        parent=styles["Body"],
        fontSize=10,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#0E2A44"),
        alignment=1,
        spaceAfter=2,
    )

    dados_assinatura = [
        [
            [
                Spacer(1, 1.8 * cm),
                HRFlowable(width=col_w * 0.85, thickness=1, color=colors.HexColor("#9FB3C7"), hAlign="CENTER"),
                Spacer(1, 5),
                Paragraph(f"<b>{tecnico or 'Técnico Responsável'}</b>", nome_style),
                Paragraph("Técnico / Assistência Técnica", linha_style),
            ],
            [
                Spacer(1, 1.8 * cm),
                HRFlowable(width=col_w * 0.85, thickness=1, color=colors.HexColor("#9FB3C7"), hAlign="CENTER"),
                Spacer(1, 5),
                Paragraph(f"<b>{cliente or 'Cliente'}</b>", nome_style),
                Paragraph("Responsável pelo Cliente", linha_style),
            ],
        ]
    ]

    tbl = Table(dados_assinatura, colWidths=[col_w, col_w])
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 2.0 * cm))

    # Campo de data/local
    story.append(Paragraph("Local e Data: _______________________________________", styles["Body"]))
    story.append(Spacer(1, 1.5 * cm))

    # Observações
    story.append(Paragraph("<b>Observações:</b>", ParagraphStyle(
        "ObsLabel", parent=styles["Body"], fontName="Helvetica-Bold", fontSize=10
    )))
    story.append(Spacer(1, 0.3 * cm))
    for _ in range(4):
        story.append(HRFlowable(width=largura_util, thickness=0.5, color=colors.HexColor("#C5D0DC")))
        story.append(Spacer(1, 0.55 * cm))

    return story


# =========================
# GERADOR PDF PRINCIPAL
# =========================
def gerar_pdf(
    sections,
    template_path,
    output_path,
    fotos=None,
    foto_cols=2,
    foto_max_height_cm=8.1,
    section_offsets_cm=None,
    horarios=None,
    watermark_path="",
    watermark_opacity=0.12,
    watermark_scale=0.75,
    watermark_mode="Central",
    watermark_random_count=8,
    cover_header_scale=1.8,
    include_signature_page=False,
):
    import uuid
    temp_pdf = output_path + f".tmp_{uuid.uuid4().hex[:8]}.pdf"
    fotos = fotos or []
    foto_cols = 1 if int(foto_cols) == 1 else 2

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(name="Titulo",
                              fontSize=18, alignment=1, spaceAfter=6,
                              textColor=colors.HexColor("#0E2A44"), leading=22))
    styles.add(ParagraphStyle(name="CoverTitle",
                              fontSize=28, alignment=1, spaceAfter=12,
                              textColor=colors.HexColor("#0A2238"), leading=32, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="Secao",
                              fontSize=12.5, spaceBefore=14, spaceAfter=7,
                              textColor=colors.HexColor("#123A5A"), leading=15, keepWithNext=True))
    styles.add(ParagraphStyle(name="Body",
                              fontSize=10.5, leading=15, textColor=colors.HexColor("#1F2B37")))

    story = []

    info = sections.get("info", {})
    header_rows = sections.get("header_rows")
    if not header_rows:
        header_rows = _compose_header_rows(info, sections.get("info_extra", []))
    if header_rows:
        story.append(Spacer(1, 0.9 * cm))
        story.append(Paragraph("Relatório de Atendimento Técnico", styles["CoverTitle"]))
        story.append(Spacer(1, 0.6 * cm))
        cover_scale = max(1.0, min(3.0, float(cover_header_scale or 1.8)))
        cover_font_size = min(13.5, 10.2 * cover_scale)
        cover_padding = max(4, int(6 * cover_scale))
        largura_total = A4[0] - (5.0 * cm)
        largura_label_min = largura_total * 0.34
        largura_label_max = largura_total * 0.56
        maior_rotulo = max((str(label).strip() for label, _ in header_rows), key=len, default="")
        largura_rotulo_pt = pdfmetrics.stringWidth(maior_rotulo, "Helvetica-Bold", cover_font_size)
        largura_label_necessaria = largura_rotulo_pt + (cover_padding * 2) + 6
        largura_label = min(max(largura_label_necessaria, largura_label_min), largura_label_max)
        largura_valor = largura_total - largura_label
        dados_tabela = [
            [
                Paragraph(f"<b>{label}</b>", ParagraphStyle(
                    "HeaderLabelCell",
                    parent=styles["Body"],
                    fontName="Helvetica-Bold",
                    fontSize=cover_font_size,
                    leading=cover_font_size * 1.25,
                    wordWrap="CJK",
                )),
                Paragraph(str(valor).strip(), ParagraphStyle(
                    "HeaderValueCell",
                    parent=styles["Body"],
                    fontName="Helvetica",
                    fontSize=cover_font_size,
                    leading=cover_font_size * 1.25,
                    wordWrap="CJK",
                )),
            ]
            for label, valor in header_rows
        ]
        tabela = Table(dados_tabela, colWidths=[largura_label, largura_valor], repeatRows=0)
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FCFDFE")),
            ("BOX", (0, 0), (-1, -1), 1.6, colors.HexColor("#9FB3C7")),
            ("INNERGRID", (0, 0), (-1, -1), 0.9, colors.HexColor("#C2D1DE")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E9F0F6")),
            ("LEFTPADDING", (0, 0), (-1, -1), cover_padding),
            ("RIGHTPADDING", (0, 0), (-1, -1), cover_padding),
            ("TOPPADDING", (0, 0), (-1, -1), max(4, int(7 * cover_scale))),
            ("BOTTOMPADDING", (0, 0), (-1, -1), max(4, int(7 * cover_scale))),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(Spacer(1, 0.5 * cm))
        story.append(tabela)
    else:
        story.append(Spacer(1, 4.5 * cm))
        story.append(Paragraph("<b>CABEÇALHO DO ATENDIMENTO</b>", styles["Titulo"]))
        story.append(Spacer(1, 0.7 * cm))
        story.append(Paragraph("Sem dados preenchidos no cabeçalho.", styles["Body"]))
    story.append(PageBreak())

    section_offsets_cm = section_offsets_cm or {}

    def add_secao(titulo, conteudo, section_key):
        offset_cm = _coerce_offset_cm(section_offsets_cm.get(section_key, 0))
        if offset_cm > 0:
            story.append(Spacer(1, offset_cm * cm))
        story.append(Paragraph(f"<b>{titulo}</b>", styles["Secao"]))
        story.extend(processar_lista(conteudo, styles))

    if sections.get("descricao"):
        add_secao("1 – ESCOPO DO ATENDIMENTO", sections["descricao"], "descricao")
    if sections.get("detalhamento"):
        add_secao("2 – DETALHAMENTO DO PROBLEMA", sections["detalhamento"], "detalhamento")
    if sections.get("diagnostico"):
        add_secao("3 – DIAGNÓSTICO", sections["diagnostico"], "diagnostico")
    if sections.get("acoes"):
        add_secao("4 – AÇÕES CORRETIVAS", sections["acoes"], "acoes")
    if sections.get("resultado"):
        add_secao("5 – RESULTADO", sections["resultado"], "resultado")
    if sections.get("estado"):
        add_secao("6 – ESTADO FINAL", sections["estado"], "estado")

    # Tabela de horários — só inclui se tiver linhas válidas
    horarios = [r for r in (horarios or []) if any(str(c).strip() for c in r)]
    if horarios:
        story.append(Paragraph("<b>TABELA DE HORÁRIOS DO ATENDIMENTO</b>", styles["Secao"]))
        story.append(Spacer(1, 8))
        dados_horarios = [["Data", "Início", "Fim", "Intervalos", "Descrição"], *horarios]
        tabela_horarios = Table(dados_horarios, colWidths=[3.2 * cm, 2.4 * cm, 2.4 * cm, 4.2 * cm, 3.8 * cm])
        tabela_horarios.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE8F2")),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FBFD")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#C7D4DF")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DCE5EC")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (1, 1), (2, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(KeepTogether([tabela_horarios]))

    if fotos:
        story.append(Paragraph("<b>ANEXOS FOTOGRÁFICOS</b>", styles["Secao"]))
        story.append(Spacer(1, 6))

        largura_util = A4[0] - (5 * cm)
        espacamento = 0.8 * cm
        largura_metade = (largura_util - espacamento) / 2
        pending_half_blocks = []

        def _build_photo_block(foto_item, foto_idx, largura_max):
            titulo = foto_item.get("title") or f"Foto {foto_idx}"
            comentario = (foto_item.get("comment") or "").strip()
            caminho = foto_item.get("path")
            ajuste_altura_cm = _coerce_offset_cm(foto_item.get("max_height_cm", foto_max_height_cm))
            altura_max = (ajuste_altura_cm if ajuste_altura_cm > 0 else foto_max_height_cm) * cm
            ajuste_largura_pct = max(30.0, min(130.0, float(foto_item.get("width_percent", 100.0)))) / 100.0
            largura_limite = largura_max * ajuste_largura_pct

            bloco = [Paragraph(f"<b>{foto_idx}. {titulo}</b>", styles["Body"]), Spacer(1, 4)]
            try:
                img_reader = ImageReader(caminho)
                largura_original, altura_original = img_reader.getSize()
                escala = min(largura_limite / largura_original, altura_max / altura_original)
                largura = largura_original * escala
                altura = altura_original * escala
                imagem = Image(caminho, width=largura, height=altura)
                imagem.hAlign = "CENTER"
                bloco.append(imagem)
            except Exception:
                bloco.append(Paragraph("Não foi possível carregar esta imagem.", styles["Body"]))

            if comentario:
                bloco.append(Spacer(1, 4))
                bloco.append(Paragraph(f"<i>Comentário:</i> {comentario}", styles["Body"]))
            return bloco

        def _flush_half_blocks():
            nonlocal pending_half_blocks
            if not pending_half_blocks:
                return
            row = pending_half_blocks
            if len(row) == 1:
                row = [row[0], ""]
            tbl = Table([row], colWidths=[largura_metade, largura_metade])
            tbl.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 10))
            pending_half_blocks = []

        for idx, foto in enumerate(fotos, start=1):
            modo = str(foto.get("layout", "Dividir página"))
            if modo == "Página inteira":
                _flush_half_blocks()
                bloco = _build_photo_block(foto, idx, largura_util)
                story.extend(bloco)
                story.append(Spacer(1, 10))
            else:
                pending_half_blocks.append(_build_photo_block(foto, idx, largura_metade))
                if len(pending_half_blocks) == 2:
                    _flush_half_blocks()
        _flush_half_blocks()

    # ← NOVA: Página de assinatura
    if include_signature_page:
        story.append(PageBreak())
        story.extend(_build_signature_page(styles, info=info))

    def _draw_page_chrome(canvas_obj, page_number):
        canvas_obj.saveState()
        if page_number > 1:
            canvas_obj.setFillColor(colors.HexColor("#F5F8FB"))
            canvas_obj.rect(2.3 * cm, A4[1] - 2.65 * cm, A4[0] - (4.6 * cm), 1.2 * cm, stroke=0, fill=1)
            canvas_obj.setStrokeColor(colors.HexColor("#D1DCE6"))
            canvas_obj.rect(2.3 * cm, A4[1] - 2.65 * cm, A4[0] - (4.6 * cm), 1.2 * cm, stroke=1, fill=0)
            canvas_obj.setFont("Helvetica-Bold", 12)
            canvas_obj.setFillColor(colors.HexColor("#0E2A44"))
            canvas_obj.drawCentredString(A4[0] / 2, A4[1] - 1.95 * cm, "RELATÓRIO TÉCNICO DE ATENDIMENTO")

        canvas_obj.setFont("Helvetica", 8.8)
        canvas_obj.setFillColor(colors.HexColor("#5B6E7D"))
        canvas_obj.drawString(2.5 * cm, 1.4 * cm, "Relatório técnico")
        canvas_obj.drawRightString(18.5 * cm, 1.4 * cm, f"Página {page_number}")
        canvas_obj.restoreState()

    class FinalPageCanvas(pdf_canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved_page_states)
            for idx, state in enumerate(self._saved_page_states, start=1):
                self.__dict__.update(state)
                _draw_page_chrome(self, idx)
                if idx == total:
                    self.setFont("Helvetica-Bold", 11)
                    self.setFillColor(colors.HexColor("#0E2A44"))
                    self.drawCentredString(A4[0] / 2, 2.05 * cm, "Assistência Técnica Grupo BAW")
                super().showPage()
            super().save()

    doc = SimpleDocTemplate(temp_pdf, pagesize=A4,
                            leftMargin=2.5 * cm, rightMargin=2.5 * cm,
                            topMargin=3.4 * cm, bottomMargin=2.5 * cm)

    doc.build(story, canvasmaker=FinalPageCanvas)

    def _make_watermark_background_pdf(content_pdf_path):
        wm_path = str(watermark_path or "").strip()
        if not wm_path or not os.path.exists(wm_path):
            return None
        try:
            content_reader = PdfReader(content_pdf_path)
            page_count = len(content_reader.pages)
            if page_count <= 0:
                return None
            op = max(0.02, min(0.95, float(watermark_opacity)))
            scale = max(0.2, min(2.5, float(watermark_scale)))
            temp_wm = output_path + f".wm_{uuid.uuid4().hex[:8]}.pdf"
            c = pdf_canvas.Canvas(temp_wm, pagesize=A4)
            image = ImageReader(wm_path)
            iw, ih = image.getSize()
            wm_mode = str(watermark_mode or "Central")
            apply_random = wm_mode in ("Timbrado aleatório", "Central + Timbrado aleatório")
            apply_central = wm_mode in ("Central", "Central + Timbrado aleatório")
            for page_number in range(1, page_count + 1):
                if apply_random:
                    import random

                    rng = random.Random(page_number * 1777)
                    logos = max(1, min(60, int(watermark_random_count)))
                    min_factor = max(0.07, 0.11 * scale)
                    max_factor = max(min_factor + 0.02, min(0.28, 0.19 * scale + 0.08))
                    for _ in range(logos):
                        factor = rng.uniform(min_factor, max_factor)
                        draw_w = iw * factor
                        draw_h = ih * factor
                        max_x = max(1.0, A4[0] - draw_w)
                        max_y = max(1.0, A4[1] - draw_h)
                        x = rng.uniform(0, max_x)
                        y = rng.uniform(0, max_y)
                        rot = rng.uniform(-25, 25)
                        c.saveState()
                        c.setFillAlpha(max(0.02, min(0.30, op * 0.85)))
                        c.translate(x + draw_w / 2, y + draw_h / 2)
                        c.rotate(rot)
                        c.drawImage(image, -draw_w / 2, -draw_h / 2, draw_w, draw_h, preserveAspectRatio=True, mask="auto")
                        c.restoreState()
                if apply_central:
                    max_w = A4[0] * scale
                    max_h = A4[1] * scale
                    factor = min(max_w / iw, max_h / ih)
                    draw_w = iw * factor
                    draw_h = ih * factor
                    x = (A4[0] - draw_w) / 2
                    y = (A4[1] - draw_h) / 2
                    c.saveState()
                    c.setFillAlpha(op)
                    c.drawImage(image, x, y, draw_w, draw_h, preserveAspectRatio=True, mask="auto")
                    c.restoreState()
                c.showPage()
            c.save()
            return temp_wm
        except Exception:
            return None

    wm_background_pdf = _make_watermark_background_pdf(temp_pdf)

    template_path_real = template_path if template_path else ""
    content_reader = PdfReader(temp_pdf)
    template_reader = PdfReader(template_path_real) if (template_path_real and os.path.exists(template_path_real)) else None
    wm_reader = PdfReader(wm_background_pdf) if (wm_background_pdf and os.path.exists(wm_background_pdf)) else None
    writer = PdfWriter()

    template_base = template_reader.pages[0] if template_reader else None
    for idx, page in enumerate(content_reader.pages):
        if template_base:
            writer.add_page(template_base)
        elif wm_reader:
            writer.add_page(wm_reader.pages[min(idx, len(wm_reader.pages) - 1)])
        else:
            writer.add_page(page)
            continue
        out_page = writer.pages[-1]
        if wm_reader and template_base:
            out_page.merge_page(wm_reader.pages[min(idx, len(wm_reader.pages) - 1)])
        out_page.merge_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)
    if wm_background_pdf and os.path.exists(wm_background_pdf):
        os.remove(wm_background_pdf)


# =========================
# VALIDAÇÃO
# =========================
def validar_sections(sections, fotos):
    """Retorna lista de avisos de validação (strings). Lista vazia = tudo OK."""
    avisos = []
    info = sections.get("info", {})

    if not _valor_info_preenchido(info.get("tecnico")):
        avisos.append("Técnico responsável não preenchido.")
    if not _valor_info_preenchido(info.get("cliente")):
        avisos.append("Nome do cliente não preenchido.")
    if not sections.get("descricao"):
        avisos.append("Seção '1 – Escopo do Atendimento' está vazia.")

    for idx, foto in enumerate(fotos or [], start=1):
        path = foto.get("path", "")
        if path and not os.path.exists(path):
            avisos.append(f"Foto {idx} não encontrada: {os.path.basename(path)}")

    return avisos


# =========================
# ABRIR PDF (multiplataforma)
# =========================
def abrir_pdf(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


# =========================
# PREVIEW ENGINE
# =========================
class PreviewEngine:
    def __init__(self, on_update, debounce_ms=1000):
        self.on_update = on_update
        self.debounce_ms = debounce_ms
        self._timer = None
        self._lock = threading.Lock()
        self._current_text = ""
        self._current_sections = {}
        self._current_fotos = []
        self._current_foto_cols = 2
        self._current_foto_h = 8.1
        self._current_section_offsets_cm = {}
        self._current_template_path = ""
        self._current_horarios = []
        self._current_watermark_path = ""
        self._current_watermark_opacity = 0.12
        self._current_watermark_scale = 0.75
        self._current_watermark_mode = "Central"
        self._current_watermark_random_count = 8
        self._current_cover_header_scale = 1.8
        self._current_signature_page = False
        self._temp_dir = tempfile.mkdtemp()
        self._running = True

    def schedule_update(
        self,
        text,
        sections=None,
        fotos=None,
        foto_cols=2,
        foto_max_height_cm=8.1,
        section_offsets_cm=None,
        template_path="",
        horarios=None,
        watermark_path="",
        watermark_opacity=0.12,
        watermark_scale=0.75,
        watermark_mode="Central",
        watermark_random_count=8,
        cover_header_scale=1.8,
        include_signature_page=False,
    ):
        with self._lock:
            self._current_text = text
            self._current_sections = sections or {}
            self._current_fotos = [dict(f) for f in (fotos or [])]
            self._current_foto_cols = foto_cols
            self._current_foto_h = foto_max_height_cm
            self._current_section_offsets_cm = section_offsets_cm or {}
            self._current_template_path = template_path or ""
            self._current_horarios = horarios or []
            self._current_watermark_path = watermark_path or ""
            self._current_watermark_opacity = watermark_opacity
            self._current_watermark_scale = watermark_scale
            self._current_watermark_mode = watermark_mode or "Central"
            self._current_watermark_random_count = max(1, min(60, int(watermark_random_count)))
            self._current_cover_header_scale = cover_header_scale
            self._current_signature_page = include_signature_page
            if self._timer is not None:
                self._timer.cancel()
            self._timer = threading.Timer(self.debounce_ms / 1000.0, self._generate)
            self._timer.daemon = True
            self._timer.start()

    def _generate(self):
        if not self._running:
            return
        with self._lock:
            text = self._current_text
            sections = self._current_sections
            fotos = self._current_fotos
            foto_cols = self._current_foto_cols
            foto_max_height_cm = self._current_foto_h
            section_offsets_cm = self._current_section_offsets_cm
            template_path = self._current_template_path
            horarios = self._current_horarios
            watermark_path = self._current_watermark_path
            watermark_opacity = self._current_watermark_opacity
            watermark_scale = self._current_watermark_scale
            watermark_mode = self._current_watermark_mode
            watermark_random_count = self._current_watermark_random_count
            cover_header_scale = self._current_cover_header_scale
            include_signature_page = self._current_signature_page

        import uuid
        temp_pdf = os.path.join(self._temp_dir, f"preview_{uuid.uuid4().hex[:8]}.pdf")
        try:
            if not sections:
                texto = limpar_texto(text)
                sections = parse_text(texto)
            gerar_pdf(
                sections,
                template_path,
                temp_pdf,
                fotos=fotos,
                foto_cols=foto_cols,
                foto_max_height_cm=foto_max_height_cm,
                section_offsets_cm=section_offsets_cm,
                horarios=horarios,
                watermark_path=watermark_path,
                watermark_opacity=watermark_opacity,
                watermark_scale=watermark_scale,
                watermark_mode=watermark_mode,
                watermark_random_count=watermark_random_count,
                cover_header_scale=cover_header_scale,
                include_signature_page=include_signature_page,
            )

            doc = fitz.open(temp_pdf)
            images = []
            mat = fitz.Matrix(2.0, 2.0)
            for page in doc:
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                images.append(img)
            doc.close()
            self.on_update(images, None)
        except Exception as e:
            self.on_update(None, str(e))
        finally:
            try:
                if os.path.exists(temp_pdf):
                    os.remove(temp_pdf)
            except Exception:
                pass

    def stop(self):
        self._running = False
        if self._timer:
            self._timer.cancel()


class PreviewPanel(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self._label_status = ctk.CTkLabel(self, text="A pré-visualização aparecerá aqui...", text_color="#8A9BAD")

        self._nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._btn_prev = ctk.CTkButton(self._nav_frame, text="◀", width=36, command=self._prev_page)
        self._lbl_page = ctk.CTkLabel(self._nav_frame, text="", width=60)
        self._btn_next = ctk.CTkButton(self._nav_frame, text="▶", width=36, command=self._next_page)
        self._lbl_sep = ctk.CTkLabel(self._nav_frame, text=" | ", text_color="#A0B0C0")
        self._btn_zoom_out = ctk.CTkButton(self._nav_frame, text="−", width=30, command=self._zoom_out)
        self._lbl_zoom = ctk.CTkLabel(self._nav_frame, text="100%", width=45)
        self._btn_zoom_in = ctk.CTkButton(self._nav_frame, text="+", width=30, command=self._zoom_in)

        self._btn_prev.pack(side="left", padx=2)
        self._lbl_page.pack(side="left", padx=2)
        self._btn_next.pack(side="left", padx=2)
        self._lbl_sep.pack(side="left", padx=4)
        self._btn_zoom_out.pack(side="left", padx=2)
        self._lbl_zoom.pack(side="left", padx=2)
        self._btn_zoom_in.pack(side="left", padx=2)

        self._canvas_container = ctk.CTkFrame(self, fg_color="transparent")
        self._canvas = tk.Canvas(self._canvas_container, bg="#E8EEF3", highlightthickness=0)
        self._vsb = ctk.CTkScrollbar(self._canvas_container, orientation="vertical", command=self._canvas.yview)
        self._hsb = ctk.CTkScrollbar(self._canvas_container, orientation="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=self._vsb.set, xscrollcommand=self._hsb.set)

        self._hsb.pack(side="bottom", fill="x", padx=(0, 16))
        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._canvas.bind("<MouseWheel>", self._on_mousewheel)
        self._canvas.bind("<Button-4>", self._on_mousewheel)
        self._canvas.bind("<Button-5>", self._on_mousewheel)
        self._canvas.bind("<Button-1>", self._on_pan_start)
        self._canvas.bind("<B1-Motion>", self._on_pan_move)

        self._nav_frame.pack(pady=(8, 2))
        self._canvas_container.pack(expand=True, fill="both", padx=8, pady=4)
        self._label_status.place(relx=0.5, rely=0.5, anchor="center")

        self._pages = []
        self._tk_images = []
        self._current_page = 0
        self._panel_width = 380
        self._zoom_factor = 1.0

    def get_zoom_factor(self):
        return self._zoom_factor

    def set_zoom_factor(self, factor):
        self._zoom_factor = max(0.4, min(4.0, float(factor)))
        if self._pages:
            self._show_page()

    def _on_mousewheel(self, event):
        if event.state & 0x0004:
            if event.delta > 0 or event.num == 4:
                self._zoom_in()
            else:
                self._zoom_out()
            return "break"

        if event.state & 0x0001:
            if event.delta:
                self._canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self._canvas.xview_scroll(-1, "units")
            elif event.num == 5:
                self._canvas.xview_scroll(1, "units")
        else:
            if event.delta:
                self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self._canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self._canvas.yview_scroll(1, "units")

    def _on_pan_start(self, event):
        self._canvas.scan_mark(event.x, event.y)

    def _on_pan_move(self, event):
        self._canvas.scan_dragto(event.x, event.y, gain=1)

    def update_pages(self, images):
        self._label_status.place_forget()
        self._pages = images
        if self._current_page >= len(self._pages):
            self._current_page = max(0, len(self._pages) - 1)
        self._show_page()

    def show_status(self, msg):
        self._label_status.configure(text=msg)
        if self._pages:
            self._label_status.place(relx=0.5, rely=0.08, anchor="n")
            return
        self._canvas.delete("all")
        self._lbl_page.configure(text="")
        self._label_status.place(relx=0.5, rely=0.5, anchor="center")

    def show_generating(self):
        if self._pages:
            self._label_status.configure(text="⏳ A atualizar pré-visualização...")
            self._label_status.place(relx=0.5, rely=0.08, anchor="n")
        else:
            self.show_status("⏳ A atualizar pré-visualização...")

    def _show_page(self):
        if not self._pages:
            return
        idx = self._current_page
        img = self._pages[idx]

        base_panel_w = max(self._panel_width - 35, 200)
        fit_ratio = base_panel_w / img.width
        final_ratio = fit_ratio * self._zoom_factor

        target_w = int(img.width * final_ratio)
        target_h = int(img.height * final_ratio)

        resized = img.resize((target_w, target_h), PILImage.LANCZOS)
        tk_img = ImageTk.PhotoImage(resized)
        self._tk_images = [tk_img]

        self._canvas.delete("all")
        self._canvas.create_image(0, 0, anchor="nw", image=tk_img)
        self._canvas.config(scrollregion=(0, 0, target_w, target_h))

        total = len(self._pages)
        self._lbl_page.configure(text=f"{idx + 1} / {total}")
        self._lbl_zoom.configure(text=f"{int(self._zoom_factor * 100)}%")

        self._btn_prev.configure(state="normal" if idx > 0 else "disabled")
        self._btn_next.configure(state="normal" if idx < total - 1 else "disabled")

    def _prev_page(self):
        if self._current_page > 0:
            self._current_page -= 1
            self._show_page()

    def _next_page(self):
        if self._current_page < len(self._pages) - 1:
            self._current_page += 1
            self._show_page()

    def _zoom_in(self):
        self.set_zoom_factor(self._zoom_factor + 0.25)

    def _zoom_out(self):
        self.set_zoom_factor(self._zoom_factor - 0.25)

    def set_width(self, w):
        if abs(self._panel_width - w) > 15:
            self._panel_width = w
            if self._pages:
                self._show_page()


class HorariosTableEditor(ctk.CTkFrame):
    COLUMNS = ("data", "inicio", "fim", "intervalos", "descricao")

    def __init__(self, master, on_change=None, **kwargs):
        super().__init__(master, **kwargs)
        self._on_change = on_change
        self._inline_editor = None

        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.pack(fill="x", padx=6, pady=(6, 2))
        ctk.CTkLabel(
            controls,
            text="Edite direto na tabela (duplo clique na célula).",
            text_color="#8a8a8a",
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(controls, text="Adicionar linha", width=120, command=self._add_empty_row).pack(side="left", padx=(0, 6))
        ctk.CTkButton(controls, text="Remover", width=90, command=self._remove_selected).pack(side="left")

        self.tree = ttk.Treeview(self, columns=self.COLUMNS, show="headings", height=4)
        headings = ("Data", "Início", "Fim", "Intervalos", "Descrição")
        widths = (100, 70, 70, 180, 150)
        for col, heading, width in zip(self.COLUMNS, headings, widths):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center" if col != "intervalos" else "w", stretch=True)

        self.tree.pack(fill="both", expand=True, padx=6, pady=(2, 6))
        self.tree.bind("<Delete>", lambda _e: self._remove_selected())
        self.tree.bind("<Double-1>", self._begin_inline_edit)
        self.tree.bind("<Double-Button-1>", self._begin_inline_edit)
        self.tree.bind("<Return>", self._begin_inline_edit_from_focus)
        self.tree.bind("<F2>", self._begin_inline_edit_from_focus)

    def _add_empty_row(self):
        row_id = self.tree.insert("", "end", values=["", "", "", "", ""])
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        if self._on_change:
            self._on_change()

    def _begin_inline_edit_from_focus(self, _event=None):
        row_id = self.tree.focus() or (self.tree.selection()[0] if self.tree.selection() else "")
        if not row_id:
            return
        bbox = self.tree.bbox(row_id, "#1")
        if not bbox:
            return
        x, y, w, h = bbox

        class _Ev:
            pass

        fake = _Ev()
        fake.x = x + 4
        fake.y = y + (h // 2)
        self._begin_inline_edit(fake)

    def _remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            self.tree.delete(item)
        if self._on_change:
            self._on_change()

    def _begin_inline_edit(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        col_index = int(col_id.replace("#", "")) - 1
        values = list(self.tree.item(row_id, "values"))
        while len(values) < len(self.COLUMNS):
            values.append("")
        if self._inline_editor:
            self._inline_editor.destroy()
            self._inline_editor = None
        if self.COLUMNS[col_index] == "descricao":
            editor_var = tk.StringVar(value=values[col_index])
            editor = ctk.CTkComboBox(
                self.tree,
                values=HORARIOS_DESCRICAO_PRESETS,
                variable=editor_var,
                width=max(110, w),
                height=h,
            )
            editor.place(x=x, y=y)
            editor.focus_set()

            def _save(_event=None):
                values[col_index] = editor_var.get().strip()
                self.tree.item(row_id, values=values)
                editor.destroy()
                self._inline_editor = None
                if self._on_change:
                    self._on_change()

            def _cancel(_event=None):
                editor.destroy()
                self._inline_editor = None

            editor.bind("<Return>", _save)
            editor.bind("<FocusOut>", _save)
            editor.bind("<Escape>", _cancel)
            self._inline_editor = editor
            return

        editor = ctk.CTkEntry(self.tree, width=w, height=h)
        editor.insert(0, values[col_index])
        editor.place(x=x, y=y)
        editor.focus_set()
        editor.select_range(0, "end")

        def _save(_event=None):
            values[col_index] = editor.get().strip()
            self.tree.item(row_id, values=values)
            editor.destroy()
            self._inline_editor = None
            if self._on_change:
                self._on_change()

        def _cancel(_event=None):
            editor.destroy()
            self._inline_editor = None

        editor.bind("<Return>", _save)
        editor.bind("<FocusOut>", _save)
        editor.bind("<Escape>", _cancel)
        self._inline_editor = editor

    def get_rows(self):
        return [list(self.tree.item(item, "values")) for item in self.tree.get_children("")]

    def set_rows(self, rows):
        self.clear()
        for row in rows or []:
            self.tree.insert("", "end", values=(list(row) + ["", "", "", "", ""])[:5])

    def clear(self):
        for item in self.tree.get_children(""):
            self.tree.delete(item)


class HeaderTableEditor(ctk.CTkFrame):
    COLUMNS = ("topico", "resposta")

    def __init__(self, master, on_change=None, **kwargs):
        super().__init__(master, **kwargs)
        self._on_change = on_change
        self._inline_editor = None

        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.pack(fill="x", padx=6, pady=(6, 2))
        ctk.CTkLabel(
            controls,
            text="Edite direto na tabela (duplo clique na célula).",
            text_color="#8a8a8a",
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(controls, text="Adicionar linha", width=120, command=self._add_empty_row).pack(side="left", padx=(0, 6))
        ctk.CTkButton(controls, text="Remover", width=90, command=self._remove_selected).pack(side="left")

        self.tree = ttk.Treeview(self, columns=self.COLUMNS, show="headings", height=9)
        self.tree.heading("topico", text="Tópico")
        self.tree.heading("resposta", text="Resposta")
        self.tree.column("topico", width=260, anchor="w", stretch=True)
        self.tree.column("resposta", width=420, anchor="w", stretch=True)
        self.tree.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        self.tree.bind("<Delete>", lambda _e: self._remove_selected())
        self.tree.bind("<Double-1>", self._begin_inline_edit)
        self.tree.bind("<Double-Button-1>", self._begin_inline_edit)
        self.tree.bind("<Return>", self._begin_inline_edit_from_focus)
        self.tree.bind("<F2>", self._begin_inline_edit_from_focus)

    def _add_empty_row(self):
        row_id = self.tree.insert("", "end", values=["", ""])
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        if self._on_change:
            self._on_change()

    def _begin_inline_edit(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        col_index = int(col_id.replace("#", "")) - 1
        values = list(self.tree.item(row_id, "values"))
        while len(values) < len(self.COLUMNS):
            values.append("")
        if self._inline_editor:
            self._inline_editor.destroy()
            self._inline_editor = None
        editor = ctk.CTkEntry(self.tree, width=w, height=h)
        editor.insert(0, values[col_index])
        editor.place(x=x, y=y)
        editor.focus_set()
        editor.select_range(0, "end")

        def _save(_event=None):
            values[col_index] = editor.get().strip()
            if values[0]:
                self.tree.item(row_id, values=values[:2])
                self.tree.selection_set(row_id)
                if self._on_change:
                    self._on_change()
            editor.destroy()
            self._inline_editor = None

        def _cancel(_event=None):
            editor.destroy()
            self._inline_editor = None

        editor.bind("<Return>", _save)
        editor.bind("<FocusOut>", _save)
        editor.bind("<Escape>", _cancel)
        self._inline_editor = editor

    def _remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            self.tree.delete(item)
        if self._on_change:
            self._on_change()

    def _begin_inline_edit_from_focus(self, _event=None):
        row_id = self.tree.focus() or (self.tree.selection()[0] if self.tree.selection() else "")
        if not row_id:
            return
        bbox = self.tree.bbox(row_id, "#1")
        if not bbox:
            return
        x, y, w, h = bbox

        class _Ev:
            pass

        fake = _Ev()
        fake.x = x + 4
        fake.y = y + (h // 2)
        self._begin_inline_edit(fake)

    def get_rows(self):
        return [list(self.tree.item(item, "values")) for item in self.tree.get_children("")]

    def set_rows(self, rows):
        self.clear()
        for row in rows or []:
            label, value = (list(row) + ["", ""])[:2]
            self.tree.insert("", "end", values=[label, value])

    def clear(self):
        for item in self.tree.get_children(""):
            self.tree.delete(item)


# =========================
# APP PRINCIPAL
# =========================
class App(TkinterDnD.Tk if TkinterDnD else ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Relatórios")
        self.config_data = load_config()
        self.geometry(self.config_data.get("window_geometry", "1280x750"))
        self.minsize(980, 600)
        self.config_data.setdefault("preview_visible", True)
        self.config_data.setdefault("preview_auto", True)
        self.config_data.setdefault("zoom_factor", 1.0)
        self.config_data.setdefault("last_dir", "")
        self.config_data.setdefault("pdf_save_dir", "")
        self.config_data.setdefault("image_dir", "")
        self.config_data.setdefault("default_tecnico", "")
        self.config_data.setdefault("signature_page", True)
        self.config_data.setdefault("watermark_mode", "Central")
        self.config_data.setdefault("watermark_random_count", "8")
        self.config_data.setdefault("recent_pdfs", [])

        self.fotos = []
        self._thumb_cache = []
        self._suspend_section_events = False
        self._last_generated_pdf = ""
        self._ensure_tecnico_login()

        # ── Barra superior ──────────────────────────────────────────────
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(10, 4))

        self.label = ctk.CTkLabel(top, text="Template: não selecionado", font=ctk.CTkFont(size=12))
        self.label.pack(side="left")

        ctk.CTkButton(top, text="Selecionar Template", width=160, command=self.select_template).pack(side="left", padx=8)
        self._tecnico_btn = ctk.CTkButton(top, text="Trocar técnico", width=120, command=self._change_tecnico_login)
        self._tecnico_btn.pack(side="left", padx=4)
        self._tecnico_label = ctk.CTkLabel(top, text="", text_color="#7D93A8")
        self._tecnico_label.pack(side="left", padx=(4, 8))
        ctk.CTkButton(top, text="Atualizar prévia", width=130, command=self.force_preview_update).pack(side="left", padx=4)

        self.chk_auto_var = tk.BooleanVar(value=bool(self.config_data.get("preview_auto", True)))
        ctk.CTkCheckBox(top, text="Prévia automática", variable=self.chk_auto_var, command=self._on_auto_preview_toggle).pack(side="left", padx=10)
        current_doc_type = str(self.config_data.get("document_type", DOCUMENT_TYPES[0]))
        if current_doc_type not in DOCUMENT_TYPES:
            current_doc_type = DOCUMENT_TYPES[0]
        self.document_type_var = tk.StringVar(value=current_doc_type)
        ctk.CTkLabel(top, text="Modelo:").pack(side="left", padx=(10, 4))
        ctk.CTkOptionMenu(
            top,
            variable=self.document_type_var,
            values=DOCUMENT_TYPES,
            width=240,
            command=self._on_document_type_change,
        ).pack(side="left")

        # Botão "Abrir PDF" — aparece só depois de gerar
        self._btn_abrir_pdf = ctk.CTkButton(top, text="📂 Abrir PDF", width=110, command=self._abrir_ultimo_pdf,
                                             fg_color="#2A6040", hover_color="#1E4A30")
        # não pack ainda — aparece após gerar

        ctk.CTkButton(top, text="Gerar PDF", width=120, command=self.generate,
                      fg_color="#1A4A7A", hover_color="#133A62").pack(side="right", padx=4)
        ctk.CTkButton(top, text="Limpar", width=90, command=self._limpar).pack(side="right", padx=4)

        self._preview_visible = bool(self.config_data.get("preview_visible", True))
        self._toggle_btn = ctk.CTkButton(top, text="◀ Ocultar prévia", width=130, command=self._toggle_preview)
        self._toggle_btn.pack(side="right", padx=8)

        # ── Barra de opções ─────────────────────────────────────────────
        options = ctk.CTkFrame(self, fg_color="transparent")
        options.pack(fill="x", padx=14, pady=(0, 6))
        self.report_options_frame = options

        ctk.CTkLabel(options, text="Fotos:").pack(side="left")
        self.foto_cols_var = tk.StringVar(value=str(self.config_data.get("foto_cols", "2")))
        ctk.CTkOptionMenu(options, variable=self.foto_cols_var, values=["1", "2"], width=65,
                          command=lambda _v: self._on_photo_layout_change()).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(options, text="col | alt máx:").pack(side="left", padx=(4, 4))
        self.foto_h_var = tk.StringVar(value=str(self.config_data.get("foto_max_height_cm", "8.1")))
        ctk.CTkOptionMenu(options, variable=self.foto_h_var,
                          values=["6.0", "7.0", "8.1", "9.5", "11.0"], width=76,
                          command=lambda _v: self._on_photo_layout_change()).pack(side="left")

        ctk.CTkLabel(options, text="  Marca d'água:").pack(side="left", padx=(6, 4))
        ctk.CTkButton(options, text="Selecionar", width=88, command=self.select_watermark).pack(side="left", padx=(0, 3))
        ctk.CTkButton(options, text="Limpar", width=60, command=self.clear_watermark).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(options, text="op:").pack(side="left", padx=(2, 2))
        self.watermark_opacity_var = tk.StringVar(value=str(self.config_data.get("watermark_opacity", "0.12")))
        ctk.CTkOptionMenu(options, variable=self.watermark_opacity_var,
                          values=["0.05", "0.08", "0.12", "0.16", "0.20", "0.25", "0.30", "0.35", "0.40"],
                          width=80, command=lambda _v: self._on_watermark_change()).pack(side="left")
        ctk.CTkLabel(options, text="tam:").pack(side="left", padx=(6, 2))
        self.watermark_scale_var = tk.StringVar(value=str(self.config_data.get("watermark_scale", "0.75")))
        ctk.CTkOptionMenu(options, variable=self.watermark_scale_var,
                          values=["0.40", "0.55", "0.75", "0.90", "1.05", "1.25", "1.45"],
                          width=76, command=lambda _v: self._on_watermark_change()).pack(side="left")
        ctk.CTkLabel(options, text="modo:").pack(side="left", padx=(6, 2))
        self.watermark_mode_var = tk.StringVar(value=str(self.config_data.get("watermark_mode", "Central")))
        ctk.CTkOptionMenu(options, variable=self.watermark_mode_var,
                          values=WATERMARK_MODES,
                          width=170, command=lambda _v: self._on_watermark_change()).pack(side="left")
        ctk.CTkLabel(options, text="qtd:").pack(side="left", padx=(6, 2))
        self.watermark_random_count_var = tk.StringVar(value=str(self.config_data.get("watermark_random_count", "8")))
        ctk.CTkOptionMenu(
            options,
            variable=self.watermark_random_count_var,
            values=WATERMARK_RANDOM_COUNT_OPTIONS,
            width=70,
            command=lambda _v: self._on_watermark_change(),
        ).pack(side="left")

        ctk.CTkLabel(options, text="  Capa:").pack(side="left", padx=(6, 2))
        self.cover_header_scale_var = tk.StringVar(value=str(self.config_data.get("cover_header_scale", "1.8")))
        ctk.CTkOptionMenu(options, variable=self.cover_header_scale_var,
                          values=["1.0", "1.2", "1.5", "1.8", "2.1", "2.4", "2.8"],
                          width=76, command=lambda _v: self._on_cover_scale_change()).pack(side="left")

        # ← NOVA: checkbox folha de assinatura
        self.signature_var = tk.BooleanVar(value=bool(self.config_data.get("signature_page", False)))
        ctk.CTkCheckBox(
            options,
            text="Folha de assinatura",
            variable=self.signature_var,
            command=self._on_signature_toggle,
        ).pack(side="left", padx=(12, 4))

        # ── Recentes ──────────────────────────────────────────────────
        self._build_recents_bar()

        # ── Área principal ─────────────────────────────────────────────
        self._main = ctk.CTkFrame(self, fg_color="transparent")
        self._main.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self.report_main_frame = self._main

        left = ctk.CTkFrame(self._main, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(left, text="Cole ou digite o texto do relatório:", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(0, 4))
        self.text = ctk.CTkTextbox(left, wrap="word", font=ctk.CTkFont(family="Courier New", size=12), height=170)
        self.text.pack(fill="x", expand=False)
        self.text.bind("<<Modified>>", self._on_text_change)

        ctk.CTkLabel(left, text="Fotos anexadas:").pack(anchor="w", pady=(8, 4))
        self.fotos_list = ctk.CTkScrollableFrame(left, height=130)
        self.fotos_list.pack(fill="x")
        ctk.CTkButton(left, text="Adicionar fotos", command=self._add_fotos).pack(anchor="w", pady=(4, 2))

        ctk.CTkLabel(left, text="Seções extraídas (editáveis):").pack(anchor="w", pady=(8, 4))
        self.sections_tabs = ttk.Notebook(left)
        self.sections_tabs.pack(fill="both", expand=True)
        self.section_widgets = {}
        self.section_offset_vars = {}
        nomes = {
            "cabecalho": "Cabeçalho",
            "descricao": "Escopo",
            "detalhamento": "Detalhamento",
            "diagnostico": "Diagnóstico",
            "acoes": "Ações",
            "resultado": "Resultado",
            "estado": "Estado final",
            "horarios": "Horários",
        }
        panel_keys = [*SECOES_PAINEL, "horarios"]
        saved_offsets = self.config_data.get("section_offsets_cm", {}) or {}
        for key in panel_keys:
            frame = ctk.CTkFrame(self.sections_tabs)
            if key != "cabecalho":
                controls_frame = ctk.CTkFrame(frame, fg_color="transparent")
                controls_frame.pack(fill="x", padx=6, pady=(6, 0))
                if key == "horarios":
                    ctk.CTkLabel(controls_frame, text="Preencha a tabela abaixo.", text_color="#8EA1B2").pack(side="left")
                else:
                    ctk.CTkLabel(controls_frame, text="Deslocamento antes do tópico (cm):").pack(side="left", padx=(0, 4))
                    offset_var = tk.StringVar(value=str(saved_offsets.get(key, "0.0")))
                    ctk.CTkOptionMenu(
                        controls_frame, variable=offset_var,
                        values=["0.0", "0.5", "1.0", "1.5", "2.0", "2.5", "3.0"],
                        width=90,
                        command=lambda _v, k=key: self._on_offset_change(k),
                    ).pack(side="left")
                    self.section_offset_vars[key] = offset_var
            if key == "horarios":
                box = HorariosTableEditor(frame, on_change=self._on_horarios_table_change)
                box.pack(fill="both", expand=True, padx=6, pady=6)
            elif key == "cabecalho":
                ctk.CTkLabel(frame, text="Edite, exclua ou adicione tópicos do cabeçalho.", text_color="#8EA1B2").pack(anchor="w", padx=8, pady=(6, 0))
                box = HeaderTableEditor(frame, on_change=self._on_header_table_change)
                box.pack(fill="both", expand=True, padx=6, pady=6)
            else:
                box = ctk.CTkTextbox(frame, wrap="word", height=110)
                box.pack(fill="both", expand=True, padx=6, pady=6)
                box.bind("<<Modified>>", self._on_section_text_edit)
            self.sections_tabs.add(frame, text=nomes[key])
            self.section_widgets[key] = box

        # Painel de prévia
        right = ctk.CTkFrame(self._main, fg_color="transparent")
        right.pack(side="right", fill="both", padx=(10, 0))

        self._preview_panel = PreviewPanel(right, width=520, fg_color=("#E8EEF3", "#1E2C38"), corner_radius=8)
        self._preview_panel.pack(fill="both", expand=True)

        self._engine = PreviewEngine(on_update=self._on_preview_ready, debounce_ms=900)
        self._preview_panel.set_zoom_factor(self.config_data.get("zoom_factor", 1.0))

        self.status_label = ctk.CTkLabel(self, text="Pronto", anchor="w")
        self.status_label.pack(fill="x", padx=14, pady=(0, 8))
        self._build_drop_support()
        self._build_certificate_screen()

        self.update_label()
        if not self._preview_visible:
            self._toggle_preview(initial=True)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_resize)
        self._update_tecnico_label()
        self._on_document_type_change(self.document_type_var.get())

        # Atalhos de teclado
        self.bind("<Control-g>", lambda e: self.generate())
        self.bind("<Control-G>", lambda e: self.generate())
        self.bind("<Control-l>", lambda e: self._limpar())
        self.bind("<Control-L>", lambda e: self._limpar())
        self.bind("<Control-t>", lambda e: self.select_template())
        self.bind("<Control-T>", lambda e: self.select_template())
        self.bind("<Control-s>", lambda e: self._save_text_to_file())
        self.bind("<Control-S>", lambda e: self._save_text_to_file())
        self.bind("<Control-o>", lambda e: self._abrir_ultimo_pdf())
        self.bind("<Control-O>", lambda e: self._abrir_ultimo_pdf())

    # ── Recentes ──────────────────────────────────────────────────────
    def _build_recents_bar(self):
        recents = [p for p in self.config_data.get("recent_pdfs", []) if os.path.exists(p)]
        if not recents:
            return
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=14, pady=(0, 4))
        ctk.CTkLabel(bar, text="Recentes:", text_color="#7D93A8", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        for path in recents[:5]:
            nome = os.path.basename(path)
            btn = ctk.CTkButton(
                bar, text=nome[:28] + ("…" if len(nome) > 28 else ""),
                width=10, height=24, font=ctk.CTkFont(size=10),
                fg_color="transparent", border_width=1,
                text_color=("#1A4A7A", "#7EB3E8"),
                command=lambda p=path: abrir_pdf(p),
            )
            btn.pack(side="left", padx=3)

    def _build_certificate_screen(self):
        self.certificate_frame = ctk.CTkFrame(self, fg_color="transparent")

        header = ctk.CTkFrame(self.certificate_frame, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(6, 6))
        ctk.CTkLabel(
            header,
            text="Modelo: Certificado de treinamento (exemplo de tela)",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Preencha os campos-base. A geração deste modelo será implementada no próximo passo.",
            text_color="#8EA1B2",
        ).pack(anchor="w", pady=(2, 0))

        form = ctk.CTkFrame(self.certificate_frame)
        form.pack(fill="both", expand=True, padx=14, pady=(0, 8))

        self.certificate_template_path_var = tk.StringVar(value=str(self.config_data.get("certificate_template_path", "")))
        row1 = ctk.CTkFrame(form, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(row1, text="Template padrão do certificado:", width=240, anchor="w").pack(side="left")
        ctk.CTkEntry(row1, textvariable=self.certificate_template_path_var).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkButton(row1, text="Carregar template", width=150, command=self.select_certificate_template).pack(side="left")

        row2 = ctk.CTkFrame(form, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(row2, text="Nome do treinado:", width=240, anchor="w").pack(side="left")
        self.certificate_trainee_var = tk.StringVar(value=str(self.config_data.get("certificate_trainee_name", "")))
        ctk.CTkEntry(row2, textvariable=self.certificate_trainee_var).pack(side="left", fill="x", expand=True)

        row3 = ctk.CTkFrame(form, fg_color="transparent")
        row3.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(row3, text="Nome do instrutor:", width=240, anchor="w").pack(side="left")
        self.certificate_instructor_var = tk.StringVar(value=str(self.config_data.get("certificate_instructor_name", "")))
        ctk.CTkEntry(row3, textvariable=self.certificate_instructor_var).pack(side="left", fill="x", expand=True)

        row4 = ctk.CTkFrame(form, fg_color="transparent")
        row4.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        ctk.CTkLabel(row4, text="Texto padrão do certificado:", anchor="w").pack(anchor="w")
        self.certificate_text = ctk.CTkTextbox(row4, wrap="word", height=120)
        self.certificate_text.pack(fill="x", pady=(4, 8))
        self.certificate_text.insert("1.0", str(self.config_data.get("certificate_default_text", "")))

        ctk.CTkLabel(row4, text="Assuntos abordados (contra capa):", anchor="w").pack(anchor="w")
        self.certificate_topics = ctk.CTkTextbox(row4, wrap="word")
        self.certificate_topics.pack(fill="both", expand=True, pady=(4, 0))
        self.certificate_topics.insert("1.0", str(self.config_data.get("certificate_topics", "")))

    def _on_document_type_change(self, selected_type):
        selected = selected_type or DOCUMENT_TYPES[0]
        if hasattr(self, "certificate_frame"):
            self._persist_certificate_fields()
        self.config_data["document_type"] = selected
        if selected == "Certificado de treinamento":
            self.report_options_frame.pack_forget()
            self.report_main_frame.pack_forget()
            self.certificate_frame.pack(fill="both", expand=True, padx=0, pady=(0, 0))
            self._set_status("Modo certificado selecionado.")
            self._preview_panel.show_status("Prévia indisponível neste exemplo de certificado.")
        else:
            self.certificate_frame.pack_forget()
            self.report_options_frame.pack(fill="x", padx=14, pady=(0, 6))
            self.report_main_frame.pack(fill="both", expand=True, padx=14, pady=(0, 6))
            self._set_status("Modo relatório técnico selecionado.")
            if self.chk_auto_var.get():
                self.force_preview_update()

    def select_certificate_template(self):
        file = filedialog.askopenfilename(initialdir=self._pick_initial_dir(), filetypes=[("PDF", "*.pdf")])
        if not file:
            return
        self.certificate_template_path_var.set(file)
        self.config_data["certificate_template_path"] = file
        self._remember_dir(file)
        self._set_status("Template do certificado carregado.")

    # ── Helpers ───────────────────────────────────────────────────────
    def _set_status(self, msg):
        self.status_label.configure(text=msg)

    def _persist_certificate_fields(self):
        self.config_data["certificate_template_path"] = self.certificate_template_path_var.get().strip()
        self.config_data["certificate_trainee_name"] = self.certificate_trainee_var.get().strip()
        self.config_data["certificate_instructor_name"] = self.certificate_instructor_var.get().strip()
        self.config_data["certificate_default_text"] = self.certificate_text.get("1.0", "end").strip()
        self.config_data["certificate_topics"] = self.certificate_topics.get("1.0", "end").strip()

    def _ensure_tecnico_login(self):
        tecnico_salvo = str(self.config_data.get("default_tecnico", "")).strip()
        if tecnico_salvo:
            return
        nome = simpledialog.askstring("Login do técnico", "Digite o nome do técnico:", parent=self)
        if nome and nome.strip():
            self.config_data["default_tecnico"] = nome.strip()
            save_config(self.config_data)

    def _change_tecnico_login(self):
        atual = str(self.config_data.get("default_tecnico", "")).strip()
        nome = simpledialog.askstring("Login do técnico", "Nome do técnico:", initialvalue=atual, parent=self)
        if nome is None:
            return
        self.config_data["default_tecnico"] = nome.strip()
        save_config(self.config_data)
        self._update_tecnico_label()
        self._refresh_sections_panel()
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _update_tecnico_label(self):
        tecnico = str(self.config_data.get("default_tecnico", "")).strip() or "(não definido)"
        self._tecnico_label.configure(text=f"Técnico: {tecnico}")

    def _build_drop_support(self):
        if not (TkinterDnD and DND_FILES):
            return
        try:
            self.text.drop_target_register(DND_FILES)
            self.text.dnd_bind("<<Drop>>", self._on_drop_text_file)
        except Exception:
            pass

    def _on_drop_text_file(self, event):
        raw = event.data.strip()
        paths = self.tk.splitlist(raw)
        if not paths:
            return
        path = paths[0].strip("{}")
        ext = os.path.splitext(path)[1].lower()
        if ext not in {".txt", ".md"}:
            self._set_status("Apenas arquivos .txt e .md são suportados no arrastar/soltar.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = f.read()
            self.text.delete("1.0", "end")
            self.text.insert("1.0", data)
            self._set_status(f"Arquivo carregado: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar arquivo: {e}")

    def _on_section_text_edit(self, event):
        event.widget.edit_modified(False)
        if self._suspend_section_events:
            return
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_offset_change(self, _section_key):
        self.config_data["section_offsets_cm"] = self._get_section_offsets_cm()
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_horarios_table_change(self):
        if self._suspend_section_events:
            return
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_header_table_change(self):
        if self._suspend_section_events:
            return
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_signature_toggle(self):
        self.config_data["signature_page"] = bool(self.signature_var.get())
        estado = "ativada" if self.signature_var.get() else "desativada"
        self._set_status(f"Folha de assinatura {estado}.")
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _get_section_offsets_cm(self):
        return {key: _coerce_offset_cm(var.get()) for key, var in self.section_offset_vars.items()}

    def _get_sections_from_ui(self):
        texto = limpar_texto(self.text.get("1.0", "end").strip())
        sections = parse_text(texto)
        header_rows = self.section_widgets["cabecalho"].get_rows()
        info, info_extra, parsed_rows = _parse_header_rows(header_rows)
        sections["info"] = info
        sections["info_extra"] = info_extra
        sections["header_rows"] = parsed_rows
        tecnico_salvo = str(self.config_data.get("default_tecnico", "")).strip()
        if tecnico_salvo and not _valor_info_preenchido(sections["info"].get("tecnico")):
            sections["info"]["tecnico"] = tecnico_salvo
            for row in sections["header_rows"]:
                if normalize(row[0]) in {"tecnico responsavel", "técnico responsável"} and not str(row[1]).strip():
                    row[1] = tecnico_salvo
                    break
        for key, box in self.section_widgets.items():
            if key in {"cabecalho", "horarios"}:
                continue
            val = _extract_section_body(box.get("1.0", "end"), key)
            if val:
                sections[key] = val
            elif key in sections:
                sections.pop(key)
        return sections

    def _get_horarios_from_ui(self):
        return self.section_widgets["horarios"].get_rows()

    def _refresh_sections_panel(self):
        texto = self.text.get("1.0", "end").strip()
        sections = parse_text(limpar_texto(texto)) if texto else {}
        info = dict(sections.get("info", {}))
        tecnico_salvo = str(self.config_data.get("default_tecnico", "")).strip()
        if tecnico_salvo and not _valor_info_preenchido(info.get("tecnico")):
            info["tecnico"] = tecnico_salvo
        sections["info"] = info
        self._suspend_section_events = True
        try:
            for key, box in self.section_widgets.items():
                if key == "cabecalho":
                    rows = _compose_header_rows(sections.get("info", {}), sections.get("info_extra", []))
                    box.set_rows(rows)
                elif key == "horarios":
                    box.set_rows([])
                else:
                    box.delete("1.0", "end")
                    header = SECTION_HEADERS.get(key, key.title())
                    body = sections.get(key, "").strip()
                    composed = f"{header}\n{body}".strip()
                    box.insert("1.0", composed)
                    box.edit_modified(False)
        finally:
            self._suspend_section_events = False

    def _limpar(self):
        self.text.delete("1.0", "end")
        self._refresh_sections_panel()
        for var in self.section_offset_vars.values():
            var.set("0.0")
        self.fotos = []
        self._render_fotos_list()
        self._preview_panel.show_status("A pré-visualização aparecerá aqui...")
        self._set_status("Campos limpos.")

    def _on_text_change(self, event=None):
        self.text.edit_modified(False)
        self._refresh_sections_panel()
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_auto_preview_toggle(self):
        self.config_data["preview_auto"] = bool(self.chk_auto_var.get())

    def _on_photo_layout_change(self):
        self.config_data["foto_cols"] = str(self.foto_cols_var.get())
        self.config_data["foto_max_height_cm"] = str(self.foto_h_var.get())
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_watermark_change(self):
        self.config_data["watermark_opacity"] = str(self.watermark_opacity_var.get())
        self.config_data["watermark_scale"] = str(self.watermark_scale_var.get())
        self.config_data["watermark_mode"] = str(self.watermark_mode_var.get())
        self.config_data["watermark_random_count"] = str(self.watermark_random_count_var.get())
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_cover_scale_change(self):
        self.config_data["cover_header_scale"] = str(self.cover_header_scale_var.get())
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _get_watermark_random_count(self):
        try:
            return max(1, min(60, int(self.watermark_random_count_var.get())))
        except Exception:
            return 8

    def select_watermark(self):
        file = filedialog.askopenfilename(
            initialdir=self._pick_image_dir(),
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")],
        )
        if not file:
            return
        self.config_data["watermark_path"] = file
        self._remember_image_dir(file)
        self._set_status("Marca d'água atualizada.")
        if self.chk_auto_var.get():
            self.force_preview_update()

    def clear_watermark(self):
        self.config_data["watermark_path"] = ""
        self._set_status("Marca d'água removida.")
        if self.chk_auto_var.get():
            self.force_preview_update()

    def force_preview_update(self):
        if self.document_type_var.get() == "Certificado de treinamento":
            self._preview_panel.show_status("Prévia indisponível neste exemplo de certificado.")
            return
        texto = self.text.get("1.0", "end").strip()
        sections = self._get_sections_from_ui() if texto else {}
        if not texto and not self.fotos:
            self._preview_panel.show_status("A pré-visualização aparecerá aqui...")
            return
        self._preview_panel.show_generating()
        self._set_status("Gerando pré-visualização...")
        self._engine.schedule_update(
            limpar_texto(texto),
            sections=sections,
            fotos=self.fotos,
            foto_cols=int(self.foto_cols_var.get()),
            foto_max_height_cm=float(self.foto_h_var.get()),
            section_offsets_cm=self._get_section_offsets_cm(),
            template_path=self.config_data.get("template_path", ""),
            horarios=self._get_horarios_from_ui(),
            watermark_path=self.config_data.get("watermark_path", ""),
            watermark_opacity=float(self.watermark_opacity_var.get()),
            watermark_scale=float(self.watermark_scale_var.get()),
            watermark_mode=str(self.watermark_mode_var.get()),
            watermark_random_count=self._get_watermark_random_count(),
            cover_header_scale=float(self.cover_header_scale_var.get()),
            include_signature_page=bool(self.signature_var.get()),
        )

    def _on_preview_ready(self, images, error):
        def _update():
            if error:
                self._set_status("Falha ao gerar prévia.")
                self._preview_panel.show_status(f"⚠ {error[:120]}")
            elif images:
                self._set_status("Prévia atualizada.")
                self._preview_panel.update_pages(images)
        self.after(0, _update)

    def _toggle_preview(self, initial=False):
        if self._preview_visible:
            self._preview_panel.pack_forget()
            self._toggle_btn.configure(text="▶ Mostrar prévia")
            self._preview_visible = False
        else:
            self._preview_panel.pack(fill="both", expand=True)
            self._toggle_btn.configure(text="◀ Ocultar prévia")
            self._preview_visible = True
        if not initial:
            self.config_data["preview_visible"] = self._preview_visible

    def _on_resize(self, event=None):
        w = self._preview_panel.winfo_width()
        if w > 50:
            self._preview_panel.set_width(w)

    def _on_close(self):
        self._persist_certificate_fields()
        self.config_data["window_geometry"] = self.geometry()
        self.config_data["preview_visible"] = self._preview_visible
        self.config_data["zoom_factor"] = self._preview_panel.get_zoom_factor()
        self.config_data["preview_auto"] = bool(self.chk_auto_var.get())
        self.config_data["foto_cols"] = str(self.foto_cols_var.get())
        self.config_data["foto_max_height_cm"] = str(self.foto_h_var.get())
        self.config_data["section_offsets_cm"] = self._get_section_offsets_cm()
        self.config_data["watermark_opacity"] = str(self.watermark_opacity_var.get())
        self.config_data["watermark_scale"] = str(self.watermark_scale_var.get())
        self.config_data["watermark_mode"] = str(self.watermark_mode_var.get())
        self.config_data["watermark_random_count"] = str(self.watermark_random_count_var.get())
        self.config_data["cover_header_scale"] = str(self.cover_header_scale_var.get())
        self.config_data["signature_page"] = bool(self.signature_var.get())
        self.config_data["document_type"] = self.document_type_var.get()
        save_config(self.config_data)
        self._engine.stop()
        self.destroy()

    def update_label(self):
        path = self.config_data.get("template_path", "")
        self.label.configure(text=f"Template: {os.path.basename(path)}" if path else "Template: não selecionado")

    def _pick_initial_dir(self):
        return self.config_data.get("last_dir") or os.path.expanduser("~")

    def _pick_pdf_save_dir(self):
        return self.config_data.get("pdf_save_dir") or self._pick_initial_dir()

    def _pick_image_dir(self):
        return self.config_data.get("image_dir") or self._pick_initial_dir()

    def _remember_dir(self, filepath):
        if filepath:
            self.config_data["last_dir"] = os.path.dirname(filepath)
            save_config(self.config_data)

    def _remember_pdf_save_dir(self, filepath):
        if filepath:
            self.config_data["pdf_save_dir"] = os.path.dirname(filepath)
            self._remember_dir(filepath)

    def _remember_image_dir(self, filepath):
        if filepath:
            self.config_data["image_dir"] = os.path.dirname(filepath)
            self._remember_dir(filepath)

    def select_template(self):
        file = filedialog.askopenfilename(initialdir=self._pick_initial_dir(), filetypes=[("PDF", "*.pdf")])
        if file:
            self.config_data["template_path"] = file
            self._remember_dir(file)
            self.update_label()
            self._set_status("Template atualizado.")
            if self.chk_auto_var.get():
                self.force_preview_update()

    def _save_text_to_file(self):
        content = self.text.get("1.0", "end").strip()
        sections_ui = self._get_sections_from_ui()
        content = _compose_full_text_with_sections(content, sections_ui)
        if not content:
            messagebox.showwarning("Salvar", "Não há texto para salvar.")
            return
        path = filedialog.asksaveasfilename(
            initialdir=self._pick_initial_dir(),
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Markdown", "*.md")],
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        self._remember_dir(path)
        self._set_status(f"Texto salvo em {os.path.basename(path)}.")

    def _abrir_ultimo_pdf(self):
        if self._last_generated_pdf and os.path.exists(self._last_generated_pdf):
            abrir_pdf(self._last_generated_pdf)
        else:
            self._set_status("Nenhum PDF gerado nesta sessão.")

    def _add_fotos(self):
        arquivos = filedialog.askopenfilenames(
            initialdir=self._pick_image_dir(),
            title="Selecione as fotos",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")],
        )
        if not arquivos:
            return
        self._remember_image_dir(arquivos[0])
        for arquivo in arquivos:
            self.fotos.append({
                "path": arquivo,
                "title": os.path.basename(arquivo),
                "comment": "",
                "layout": "Dividir página",
                "max_height_cm": float(self.foto_h_var.get()),
                "width_percent": 100.0,
            })
        self._render_fotos_list()
        self._set_status(f"{len(arquivos)} foto(s) adicionada(s).")
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _move_foto(self, idx, delta):
        novo = idx + delta
        if novo < 0 or novo >= len(self.fotos):
            return
        self.fotos[idx], self.fotos[novo] = self.fotos[novo], self.fotos[idx]
        self._render_fotos_list()
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _remove_foto(self, idx):
        self.fotos.pop(idx)
        self._render_fotos_list()
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _render_fotos_list(self):
        for w in self.fotos_list.winfo_children():
            w.destroy()
        self._thumb_cache = []
        if not self.fotos:
            ctk.CTkLabel(self.fotos_list, text="Nenhuma foto anexada.", text_color="#7E8D9A").pack(anchor="w", padx=4, pady=4)
            return

        for idx, foto in enumerate(self.fotos):
            row = ctk.CTkFrame(self.fotos_list)
            row.pack(fill="x", padx=2, pady=2)
            thumb_lbl = ctk.CTkLabel(row, text="", width=48)
            thumb_lbl.pack(side="left", padx=4)
            try:
                img = PILImage.open(foto["path"])
                img.thumbnail((48, 48))
                tk_img = ImageTk.PhotoImage(img)
                self._thumb_cache.append(tk_img)
                thumb_lbl.configure(image=tk_img)
            except Exception:
                thumb_lbl.configure(text="N/A")

            entry = ctk.CTkEntry(row)
            entry.insert(0, foto.get("title", ""))
            entry.pack(side="left", fill="x", expand=True, padx=4)
            entry.bind("<KeyRelease>", lambda e, i=idx, ent=entry: self._set_foto_title(i, ent.get()))

            comentario_entry = ctk.CTkEntry(row, placeholder_text="Comentário")
            comentario_entry.insert(0, foto.get("comment", ""))
            comentario_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
            comentario_entry.bind("<KeyRelease>", lambda e, i=idx, ent=comentario_entry: self._set_foto_comment(i, ent.get()))

            layout_var = tk.StringVar(value=foto.get("layout", PHOTO_LAYOUT_MODES[0]))
            ctk.CTkOptionMenu(row, variable=layout_var, values=PHOTO_LAYOUT_MODES, width=140,
                              command=lambda v, i=idx: self._set_foto_layout(i, v)).pack(side="left", padx=3)

            height_var = tk.StringVar(value=f"{float(foto.get('max_height_cm', self.foto_h_var.get())):.1f}")
            ctk.CTkOptionMenu(row, variable=height_var,
                              values=["5.0", "6.0", "7.0", "8.1", "9.5", "11.0", "13.0", "16.0", "20.0"],
                              width=85, command=lambda v, i=idx: self._set_foto_height(i, v)).pack(side="left", padx=3)

            width_var = tk.StringVar(value=f"{int(float(foto.get('width_percent', 100)))}%")
            ctk.CTkOptionMenu(row, variable=width_var,
                              values=["70%", "80%", "90%", "100%", "110%", "120%", "130%", "150%", "170%", "200%"],
                              width=78, command=lambda v, i=idx: self._set_foto_width(i, v)).pack(side="left", padx=3)

            ctk.CTkButton(row, text="↑", width=28, command=lambda i=idx: self._move_foto(i, -1)).pack(side="left", padx=1)
            ctk.CTkButton(row, text="↓", width=28, command=lambda i=idx: self._move_foto(i, 1)).pack(side="left", padx=1)
            ctk.CTkButton(row, text="✕", width=28, command=lambda i=idx: self._remove_foto(i)).pack(side="left", padx=1)

    def _set_foto_title(self, idx, title):
        if 0 <= idx < len(self.fotos):
            self.fotos[idx]["title"] = title.strip()
            if self.chk_auto_var.get():
                self.force_preview_update()

    def _set_foto_layout(self, idx, layout):
        if 0 <= idx < len(self.fotos):
            self.fotos[idx]["layout"] = layout
            if self.chk_auto_var.get():
                self.force_preview_update()

    def _set_foto_comment(self, idx, comment):
        if 0 <= idx < len(self.fotos):
            self.fotos[idx]["comment"] = comment.strip()
            if self.chk_auto_var.get():
                self.force_preview_update()

    def _set_foto_height(self, idx, value):
        if 0 <= idx < len(self.fotos):
            try:
                self.fotos[idx]["max_height_cm"] = float(value)
            except Exception:
                self.fotos[idx]["max_height_cm"] = float(self.foto_h_var.get())
            if self.chk_auto_var.get():
                self.force_preview_update()

    def _set_foto_width(self, idx, value):
        if 0 <= idx < len(self.fotos):
            try:
                self.fotos[idx]["width_percent"] = float(str(value).replace("%", ""))
            except Exception:
                self.fotos[idx]["width_percent"] = 100.0
            if self.chk_auto_var.get():
                self.force_preview_update()

    # ── GERAR PDF (em thread separada com validação) ───────────────────
    def generate(self):
        if self.document_type_var.get() == "Certificado de treinamento":
            self._persist_certificate_fields()
            save_config(self.config_data)
            messagebox.showinfo(
                "Modelo em preparação",
                "A tela de Certificado de treinamento foi criada.\n"
                "A geração de PDF deste modelo será implementada no próximo passo.",
            )
            return
        texto = self.text.get("1.0", "end").strip()
        if not texto:
            messagebox.showerror("Erro", "Cole o texto primeiro.")
            return

        sections = self._get_sections_from_ui()

        # ← Validação antes de gerar
        avisos = validar_sections(sections, self.fotos)
        if avisos:
            msg = "Atenção antes de gerar:\n\n" + "\n".join(f"• {a}" for a in avisos)
            msg += "\n\nDeseja continuar mesmo assim?"
            if not messagebox.askyesno("Validação", msg):
                return

        save = filedialog.asksaveasfilename(
            initialdir=self._pick_pdf_save_dir(),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return
        self._remember_pdf_save_dir(save)

        # ← Geração em thread separada para não travar a UI
        self._set_status("Gerando PDF...")
        self._btn_abrir_pdf.pack_forget()

        def _worker():
            try:
                gerar_pdf(
                    sections,
                    self.config_data.get("template_path", ""),
                    save,
                    fotos=self.fotos,
                    foto_cols=int(self.foto_cols_var.get()),
                    foto_max_height_cm=float(self.foto_h_var.get()),
                    section_offsets_cm=self._get_section_offsets_cm(),
                    horarios=self._get_horarios_from_ui(),
                    watermark_path=self.config_data.get("watermark_path", ""),
                    watermark_opacity=float(self.watermark_opacity_var.get()),
                    watermark_scale=float(self.watermark_scale_var.get()),
                    watermark_mode=str(self.watermark_mode_var.get()),
                    watermark_random_count=self._get_watermark_random_count(),
                    cover_header_scale=float(self.cover_header_scale_var.get()),
                    include_signature_page=bool(self.signature_var.get()),
                )
                self.after(0, lambda: self._on_generate_success(save))
            except Exception as e:
                self.after(0, lambda err=str(e): self._on_generate_error(err))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_generate_success(self, path):
        self._last_generated_pdf = path
        add_recent_pdf(self.config_data, path)
        self._set_status(f"PDF gerado: {os.path.basename(path)}")
        # Mostra botão "Abrir PDF" na barra superior
        self._btn_abrir_pdf.pack(side="left", padx=6)
        if messagebox.askyesno("Sucesso", f"PDF gerado com sucesso!\n{os.path.basename(path)}\n\nAbrir agora?"):
            abrir_pdf(path)

    def _on_generate_error(self, err):
        self._set_status("Erro ao gerar PDF.")
        messagebox.showerror("Erro ao gerar PDF", err)


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
