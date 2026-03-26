import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
import json, os, re, unicodedata, threading, tempfile

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import *
from reportlab.lib.styles import *
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdf_canvas

from pypdf import PdfReader, PdfWriter
import fitz  # pymupdf — instale com: pip install pymupdf
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
    "acompanhamento remoto": "acompanhamento",
    "acompanhamento": "acompanhamento",
    "horario de inicio": "inicio",
    "inicio": "inicio",
    "horario de termino": "fim",
    "fim": "fim",
    "tempo do atendimento": "tempo_atendimento",
    "tempo atendimento": "tempo_atendimento",
    "tempo em espera": "tempo_espera",
    "tempo espera": "tempo_espera",
    "contato": "contato_cliente",
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
    ("Técnico", "tecnico"),
    ("Acompanhamento remoto", "acompanhamento"),
    ("Início", "inicio"),
    ("Fim", "fim"),
    ("Tempo Atendimento", "tempo_atendimento"),
    ("Tempo Espera", "tempo_espera"),
]
MANDATORY_INFO_FIELDS = ("tecnico", "cliente")


# =========================
# CONFIG
# =========================
def load_config():
    defaults = {
        "template_path": "",
        "last_dir": "",
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
        "cover_header_scale": "1.8",
    }
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            defaults.update(data)
    return defaults


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# =========================
# PARSER
# =========================
def normalize(text):
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c)).lower().strip()


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
                sections["info"][INFO_ALIASES[key]] = value
            else:
                sections["info_extra"].append((k.strip(), value))

    parts = re.split(r"\n\s*(\d+\s*[–-]\s*.+)", text)

    for i in range(1, len(parts), 2):
        title = normalize(parts[i])
        content = parts[i + 1].strip()

        if "descricao" in title:
            sections["descricao"] = content
        elif "detalhamento" in title:
            sections["detalhamento"] = content
        elif "diagnostico" in title:
            sections["diagnostico"] = content
        elif "acoes" in title:
            sections["acoes"] = content
        elif "resultado" in title:
            sections["resultado"] = content
        elif "estado" in title:
            sections["estado"] = content

    return sections


def limpar_texto(texto):
    texto = re.sub(r"\*\*(.*?)\*\*", r"\1", texto)
    texto = re.sub(r"#+\s*", "", texto)
    texto = re.sub(r"---+", "", texto)
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
# LISTAS AUTOMÁTICAS
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
    elementos = []
    linhas = [linha.strip() for linha in texto.split("\n") if linha.strip()]
    if not linhas:
        return [Spacer(1, 10)]

    tem_topicos_explicitos = any(_linha_tem_topico(linha) for linha in linhas)
    itens = []
    item_atual = ""

    def adicionar_item(item):
        texto_item = item.strip()
        if texto_item:
            itens.append(texto_item)

    for linha in linhas:
        if tem_topicos_explicitos and _linha_tem_topico(linha):
            if item_atual:
                adicionar_item(item_atual)
            item_atual = _limpar_marcador_topico(linha)
            continue

        if item_atual:
            if tem_topicos_explicitos:
                item_atual += " " + linha
            else:
                if item_atual.rstrip().endswith("."):
                    adicionar_item(item_atual)
                    item_atual = linha
                else:
                    item_atual += " " + linha
        else:
            item_atual = _limpar_marcador_topico(linha) if tem_topicos_explicitos else linha

    if item_atual:
        adicionar_item(item_atual)

    if itens:
        lista = [
            ListItem(
                Paragraph(item, styles["Body"]),
                leftIndent=12,
                bulletText="°"
            )
            for item in itens
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

    elementos.append(Spacer(1, 10))
    return elementos


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
    return _parse_info_from_editor(text)


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
        date, inicio, fim, intervalos = parts[:4]
        linhas.append([
            _normalize_date(date),
            _normalize_time(inicio),
            _normalize_time(fim),
            intervalos,
        ])
    return linhas


def _compose_horarios_table_text(rows=None):
    rows = rows or []
    lines = ["Data | Início | Fim | Intervalos"]
    for row in rows:
        date, inicio, fim, intervalos = (list(row) + ["", "", "", ""])[:4]
        lines.append(f"{date} | {inicio} | {fim} | {intervalos}")
    return "\n".join(lines)


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
    cover_header_scale=1.8,
):
    import uuid
    temp_pdf = output_path + f".tmp_{uuid.uuid4().hex[:8]}.pdf"
    fotos = fotos or []
    foto_cols = 1 if int(foto_cols) == 1 else 2

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(name="Titulo",
                              fontSize=18, alignment=1, spaceAfter=6,
                              textColor=colors.HexColor("#0E2A44"), leading=22))

    styles.add(ParagraphStyle(name="Secao",
                              fontSize=12.5, spaceBefore=14, spaceAfter=7,
                              textColor=colors.HexColor("#123A5A"), leading=15, keepWithNext=True))

    styles.add(ParagraphStyle(name="Body",
                              fontSize=10.5, leading=15, textColor=colors.HexColor("#1F2B37")))

    story = []

    info = sections.get("info", {})
    dados = []

    for label, key in INFO_FIELDS:
        valor = info.get(key)
        if _valor_info_preenchido(valor):
            dados.append([label, str(valor).strip()])

    if dados:
        cover_scale = max(1.0, min(3.0, float(cover_header_scale or 1.8)))
        cover_font_size = min(13.5, 10.2 * cover_scale)
        cover_padding = max(4, int(6 * cover_scale))
        largura_total = A4[0] - (5.0 * cm)
        largura_label = largura_total * 0.34
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
            for label, valor in dados
        ]
        tabela = Table(dados_tabela, colWidths=[largura_label, largura_valor], repeatRows=0)
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FCFDFE")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#C7D4DF")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DCE5EC")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E9F0F6")),
            ("LEFTPADDING", (0, 0), (-1, -1), cover_padding),
            ("RIGHTPADDING", (0, 0), (-1, -1), cover_padding),
            ("TOPPADDING", (0, 0), (-1, -1), max(4, int(7 * cover_scale))),
            ("BOTTOMPADDING", (0, 0), (-1, -1), max(4, int(7 * cover_scale))),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(Spacer(1, max(1.4, (3.4 - cover_scale)) * cm))
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

    horarios = horarios or []
    if horarios:
        story.append(Paragraph("<b>TABELA DE HORÁRIOS DO ATENDIMENTO</b>", styles["Secao"]))
        story.append(Spacer(1, 8))
        dados_horarios = [["Data", "Início", "Fim", "Intervalos"], *horarios]
        tabela_horarios = Table(dados_horarios, colWidths=[4 * cm, 3.2 * cm, 3.2 * cm, 5.6 * cm])
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

    def _draw_page_chrome(canvas_obj, page_number):
        canvas_obj.saveState()
        wm_path = str(watermark_path or "").strip()
        if wm_path and os.path.exists(wm_path):
            try:
                op = max(0.02, min(0.95, float(watermark_opacity)))
            except Exception:
                op = 0.12
            try:
                scale = max(0.2, min(2.5, float(watermark_scale)))
            except Exception:
                scale = 0.75
            try:
                image = ImageReader(wm_path)
                iw, ih = image.getSize()
                max_w = A4[0] * scale
                max_h = A4[1] * scale
                factor = min(max_w / iw, max_h / ih)
                draw_w = iw * factor
                draw_h = ih * factor
                x = (A4[0] - draw_w) / 2
                y = (A4[1] - draw_h) / 2
                canvas_obj.setFillAlpha(op)
                canvas_obj.drawImage(image, x, y, draw_w, draw_h, preserveAspectRatio=True, mask="auto")
                canvas_obj.setFillAlpha(1)
            except Exception:
                pass
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

    template_path_real = template_path if template_path else ""
    if template_path_real and os.path.exists(template_path_real):
        template_reader = PdfReader(template_path_real)
        content_reader = PdfReader(temp_pdf)
        writer = PdfWriter()

        template_base = template_reader.pages[0]
        for page in content_reader.pages:
            writer.add_page(template_base)
            out_page = writer.pages[-1]
            out_page.merge_page(page)

        with open(output_path, "wb") as f:
            writer.write(f)
        os.remove(temp_pdf)
    else:
        if os.path.exists(output_path):
            os.remove(output_path)
        os.rename(temp_pdf, output_path)


# =========================
# PREVIEW ENGINE
# =========================
class PreviewEngine:
    """Gera imagens de pré-visualização do PDF num thread em segundo plano com debounce."""

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
        self._current_cover_header_scale = 1.8
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
        cover_header_scale=1.8,
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
            self._current_cover_header_scale = cover_header_scale
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
            cover_header_scale = self._current_cover_header_scale

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
                cover_header_scale=cover_header_scale,
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
        self._label_status.pack(expand=True)

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
        self._label_status.pack_forget()
        self._nav_frame.pack(pady=(8, 2))
        self._canvas_container.pack(expand=True, fill="both", padx=8, pady=4)
        self._pages = images
        self._show_page()

    def show_status(self, msg):
        self._nav_frame.pack_forget()
        self._canvas_container.pack_forget()
        self._label_status.configure(text=msg)
        self._label_status.pack(expand=True)

    def show_generating(self):
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
    COLUMNS = ("data", "inicio", "fim", "intervalos")

    def __init__(self, master, on_change=None, **kwargs):
        super().__init__(master, **kwargs)
        self._on_change = on_change

        self.tree = ttk.Treeview(self, columns=self.COLUMNS, show="headings", height=4)
        headings = ("Data", "Início", "Fim", "Intervalos")
        widths = (100, 70, 70, 220)
        for col, heading, width in zip(self.COLUMNS, headings, widths):
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor="center" if col != "intervalos" else "w", stretch=True)
        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.pack(fill="x", padx=6, pady=(6, 2))

        self.inputs = {}
        for label, key, width in (("Data", "data", 90), ("Início", "inicio", 75), ("Fim", "fim", 75), ("Intervalos", "intervalos", 210)):
            ctk.CTkLabel(controls, text=label).pack(side="left", padx=(0, 4))
            entry = ctk.CTkEntry(controls, width=width, placeholder_text="dd/mm/aaaa" if key == "data" else "")
            entry.pack(side="left", padx=(0, 8))
            entry.bind("<Return>", self._add_row_from_inputs)
            self.inputs[key] = entry

        ctk.CTkButton(controls, text="Adicionar", width=90, command=self._add_row_from_inputs).pack(side="left", padx=(0, 6))
        ctk.CTkButton(controls, text="Remover", width=90, command=self._remove_selected).pack(side="left")

        self.tree.pack(fill="both", expand=True, padx=6, pady=(2, 6))
        self.tree.bind("<Delete>", lambda _e: self._remove_selected())

    def _add_row_from_inputs(self, _event=None):
        raw = " | ".join(self.inputs[key].get().strip() for key in self.COLUMNS)
        parsed = _parse_horarios_table(raw)
        if not parsed:
            return
        self.tree.insert("", "end", values=parsed[0])
        for entry in self.inputs.values():
            entry.delete(0, "end")
        if self._on_change:
            self._on_change()

    def _remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            self.tree.delete(item)
        if self._on_change:
            self._on_change()

    def get_rows(self):
        return [list(self.tree.item(item, "values")) for item in self.tree.get_children("")]

    def set_rows(self, rows):
        self.clear()
        for row in rows or []:
            self.tree.insert("", "end", values=(list(row) + ["", "", "", ""])[:4])

    def clear(self):
        for item in self.tree.get_children(""):
            self.tree.delete(item)


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
        self.config_data.setdefault("default_tecnico", "")

        self.fotos = []
        self._thumb_cache = []
        self._suspend_section_events = False
        self._ensure_tecnico_login()

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
        ctk.CTkCheckBox(top, text="Atualizar prévia automaticamente", variable=self.chk_auto_var, command=self._on_auto_preview_toggle).pack(side="left", padx=10)

        ctk.CTkButton(top, text="Gerar PDF", width=120, command=self.generate).pack(side="right", padx=4)
        ctk.CTkButton(top, text="Limpar", width=90, command=self._limpar).pack(side="right", padx=4)

        self._preview_visible = bool(self.config_data.get("preview_visible", True))
        self._toggle_btn = ctk.CTkButton(top, text="◀ Ocultar prévia", width=130, command=self._toggle_preview)
        self._toggle_btn.pack(side="right", padx=8)

        options = ctk.CTkFrame(self, fg_color="transparent")
        options.pack(fill="x", padx=14, pady=(0, 6))
        ctk.CTkLabel(options, text="Layout das fotos:").pack(side="left")
        self.foto_cols_var = tk.StringVar(value=str(self.config_data.get("foto_cols", "2")))
        ctk.CTkOptionMenu(
            options,
            variable=self.foto_cols_var,
            values=["1", "2"],
            width=80,
            command=lambda _v: self._on_photo_layout_change(),
        ).pack(side="left", padx=6)
        ctk.CTkLabel(options, text="colunas | altura máx (cm):").pack(side="left", padx=(8, 4))
        self.foto_h_var = tk.StringVar(value=str(self.config_data.get("foto_max_height_cm", "8.1")))
        ctk.CTkOptionMenu(
            options,
            variable=self.foto_h_var,
            values=["6.0", "7.0", "8.1", "9.5", "11.0"],
            width=90,
            command=lambda _v: self._on_photo_layout_change(),
        ).pack(side="left")
        ctk.CTkLabel(options, text="Marca d'água:").pack(side="left", padx=(10, 4))
        ctk.CTkButton(options, text="Selecionar", width=92, command=self.select_watermark).pack(side="left", padx=(0, 4))
        ctk.CTkButton(options, text="Limpar", width=68, command=self.clear_watermark).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(options, text="opacidade:").pack(side="left", padx=(4, 4))
        self.watermark_opacity_var = tk.StringVar(value=str(self.config_data.get("watermark_opacity", "0.12")))
        ctk.CTkOptionMenu(
            options,
            variable=self.watermark_opacity_var,
            values=["0.05", "0.08", "0.12", "0.16", "0.20", "0.25", "0.30", "0.35", "0.40"],
            width=86,
            command=lambda _v: self._on_watermark_change(),
        ).pack(side="left")
        ctk.CTkLabel(options, text="tamanho:").pack(side="left", padx=(8, 4))
        self.watermark_scale_var = tk.StringVar(value=str(self.config_data.get("watermark_scale", "0.75")))
        ctk.CTkOptionMenu(
            options,
            variable=self.watermark_scale_var,
            values=["0.40", "0.55", "0.75", "0.90", "1.05", "1.25", "1.45"],
            width=84,
            command=lambda _v: self._on_watermark_change(),
        ).pack(side="left")

        self._main = ctk.CTkFrame(self, fg_color="transparent")
        self._main.pack(fill="both", expand=True, padx=14, pady=(0, 6))

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
                controls = ctk.CTkFrame(frame, fg_color="transparent")
                controls.pack(fill="x", padx=6, pady=(6, 0))
                if key == "horarios":
                    ctk.CTkLabel(
                        controls,
                        text="Preencha a tabela abaixo (mesmo layout do PDF).",
                        text_color="#8EA1B2",
                    ).pack(side="left", padx=(0, 4))
                else:
                    ctk.CTkLabel(controls, text="Deslocamento antes do tópico (cm):").pack(side="left", padx=(0, 4))
                    offset_var = tk.StringVar(value=str(saved_offsets.get(key, "0.0")))
                    offset = ctk.CTkOptionMenu(
                        controls,
                        variable=offset_var,
                        values=["0.0", "0.5", "1.0", "1.5", "2.0", "2.5", "3.0"],
                        width=90,
                        command=lambda _v, k=key: self._on_offset_change(k),
                    )
                    offset.pack(side="left")
                    self.section_offset_vars[key] = offset_var
            if key == "horarios":
                box = HorariosTableEditor(frame, on_change=self._on_horarios_table_change)
                box.pack(fill="both", expand=True, padx=6, pady=6)
            else:
                box = ctk.CTkTextbox(frame, wrap="word", height=110)
                box.pack(fill="both", expand=True, padx=6, pady=6)
                box.bind("<<Modified>>", self._on_section_text_edit)
            self.sections_tabs.add(frame, text=nomes[key])
            self.section_widgets[key] = box
        ctk.CTkLabel(options, text="Escala da capa (cabeçalho):").pack(side="left", padx=(10, 4))
        self.cover_header_scale_var = tk.StringVar(value=str(self.config_data.get("cover_header_scale", "1.8")))
        ctk.CTkOptionMenu(
            options,
            variable=self.cover_header_scale_var,
            values=["1.0", "1.2", "1.5", "1.8", "2.1", "2.4", "2.8"],
            width=86,
            command=lambda _v: self._on_cover_scale_change(),
        ).pack(side="left")

        right = ctk.CTkFrame(self._main, fg_color="transparent")
        right.pack(side="right", fill="both", padx=(10, 0))

        self._preview_panel = PreviewPanel(right, width=520, fg_color=("#E8EEF3", "#1E2C38"), corner_radius=8)
        self._preview_panel.pack(fill="both", expand=True)

        self._engine = PreviewEngine(on_update=self._on_preview_ready, debounce_ms=900)
        self._preview_panel.set_zoom_factor(self.config_data.get("zoom_factor", 1.0))

        self.status_label = ctk.CTkLabel(self, text="Pronto", anchor="w")
        self.status_label.pack(fill="x", padx=14, pady=(0, 8))
        self._build_drop_support()

        self.update_label()
        self._update_template_preview_image()
        if not self._preview_visible:
            self._toggle_preview(initial=True)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_resize)
        self._update_tecnico_label()

        self.bind("<Control-g>", lambda e: self.generate())
        self.bind("<Control-G>", lambda e: self.generate())
        self.bind("<Control-l>", lambda e: self._limpar())
        self.bind("<Control-L>", lambda e: self._limpar())
        self.bind("<Control-t>", lambda e: self.select_template())
        self.bind("<Control-T>", lambda e: self.select_template())
        self.bind("<Control-s>", lambda e: self._save_text_to_file())
        self.bind("<Control-S>", lambda e: self._save_text_to_file())

    def _set_status(self, msg):
        self.status_label.configure(text=msg)

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
        self._tecnico_label.configure(text=f"Técnico atual: {tecnico}")

    def _build_drop_support(self):
        if not (TkinterDnD and DND_FILES):
            self._set_status("Arrastar e soltar desativado (tkinterdnd2 não instalado).")
            return
        try:
            self.text.drop_target_register(DND_FILES)
            self.text.dnd_bind("<<Drop>>", self._on_drop_text_file)
        except Exception:
            self._set_status("Não foi possível ativar arrastar e soltar.")

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
            self._set_status(f"Arquivo carregado via arrastar/soltar: {os.path.basename(path)}")
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

    def _get_section_offsets_cm(self):
        return {key: _coerce_offset_cm(var.get()) for key, var in self.section_offset_vars.items()}

    def _get_sections_from_ui(self):
        texto = limpar_texto(self.text.get("1.0", "end").strip())
        sections = parse_text(texto)
        cabecalho = self.section_widgets["cabecalho"].get("1.0", "end")
        info, info_extra = _parse_info_from_editor(cabecalho)
        sections["info"] = info
        sections["info_extra"] = info_extra
        tecnico_salvo = str(self.config_data.get("default_tecnico", "")).strip()
        if tecnico_salvo and not _valor_info_preenchido(sections["info"].get("tecnico")):
            sections["info"]["tecnico"] = tecnico_salvo
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
                    box.delete("1.0", "end")
                    composed = _compose_info_text(
                        sections.get("info", {}),
                        sections.get("info_extra", []),
                        include_required_empty=True,
                    )
                    box.insert("1.0", composed)
                    box.edit_modified(False)
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
        estado = "ativada" if self.chk_auto_var.get() else "desativada"
        self._set_status(f"Prévia automática {estado}.")

    def _on_photo_layout_change(self):
        self.config_data["foto_cols"] = str(self.foto_cols_var.get())
        self.config_data["foto_max_height_cm"] = str(self.foto_h_var.get())
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_watermark_change(self):
        self.config_data["watermark_opacity"] = str(self.watermark_opacity_var.get())
        self.config_data["watermark_scale"] = str(self.watermark_scale_var.get())
        if self.chk_auto_var.get():
            self.force_preview_update()

    def _on_cover_scale_change(self):
        self.config_data["cover_header_scale"] = str(self.cover_header_scale_var.get())
        if self.chk_auto_var.get():
            self.force_preview_update()

    def select_watermark(self):
        file = filedialog.askopenfilename(
            initialdir=self._pick_initial_dir(),
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")],
        )
        if not file:
            return
        self.config_data["watermark_path"] = file
        self._remember_dir(file)
        self._set_status("Marca d'água atualizada.")
        if self.chk_auto_var.get():
            self.force_preview_update()

    def clear_watermark(self):
        self.config_data["watermark_path"] = ""
        self._set_status("Marca d'água removida.")
        if self.chk_auto_var.get():
            self.force_preview_update()

    def force_preview_update(self):
        texto = self.text.get("1.0", "end").strip()
        sections = self._get_sections_from_ui() if texto else {}
        if not texto and not self.fotos:
            self._preview_panel.show_status("A pré-visualização aparecerá aqui...")
            return
        self._preview_panel.show_generating()
        self._set_status("Gerando pré-visualização...")
        preview_text = limpar_texto(texto)
        self._engine.schedule_update(
            preview_text,
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
            cover_header_scale=float(self.cover_header_scale_var.get()),
        )

    def _on_preview_ready(self, images, error):
        def _update():
            if error:
                self._set_status("Falha ao gerar prévia.")
                self._preview_panel.show_status(f"⚠ {error[:80]}")
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
        self.config_data["window_geometry"] = self.geometry()
        self.config_data["preview_visible"] = self._preview_visible
        self.config_data["zoom_factor"] = self._preview_panel.get_zoom_factor()
        self.config_data["preview_auto"] = bool(self.chk_auto_var.get())
        self.config_data["foto_cols"] = str(self.foto_cols_var.get())
        self.config_data["foto_max_height_cm"] = str(self.foto_h_var.get())
        self.config_data["section_offsets_cm"] = self._get_section_offsets_cm()
        self.config_data["watermark_opacity"] = str(self.watermark_opacity_var.get())
        self.config_data["watermark_scale"] = str(self.watermark_scale_var.get())
        self.config_data["cover_header_scale"] = str(self.cover_header_scale_var.get())
        save_config(self.config_data)
        self._engine.stop()
        self.destroy()

    def update_label(self):
        path = self.config_data.get("template_path", "")
        self.label.configure(text=f"Template: {path}" if path else "Template: não selecionado")

    def _pick_initial_dir(self):
        return self.config_data.get("last_dir") or os.path.expanduser("~")

    def _remember_dir(self, filepath):
        if filepath:
            self.config_data["last_dir"] = os.path.dirname(filepath)
            save_config(self.config_data)

    def _update_template_preview_image(self):
        # Prévia dedicada do template removida: agora mostramos apenas a pré-visualização final do PDF.
        return

    def select_template(self):
        file = filedialog.askopenfilename(initialdir=self._pick_initial_dir(), filetypes=[("PDF", "*.pdf")])
        if file:
            self.config_data["template_path"] = file
            self._remember_dir(file)
            self.update_label()
            self._update_template_preview_image()
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

    def _add_fotos(self):
        arquivos = filedialog.askopenfilenames(
            initialdir=self._pick_initial_dir(),
            title="Selecione as fotos",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")],
        )
        if not arquivos:
            return
        self._remember_dir(arquivos[0])
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
            comentario_entry = ctk.CTkEntry(row, placeholder_text="Comentário da foto")
            comentario_entry.insert(0, foto.get("comment", ""))
            comentario_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
            comentario_entry.bind("<KeyRelease>", lambda e, i=idx, ent=comentario_entry: self._set_foto_comment(i, ent.get()))

            layout_var = tk.StringVar(value=foto.get("layout", PHOTO_LAYOUT_MODES[0]))
            layout_menu = ctk.CTkOptionMenu(
                row,
                variable=layout_var,
                values=PHOTO_LAYOUT_MODES,
                width=140,
                command=lambda v, i=idx: self._set_foto_layout(i, v),
            )
            layout_menu.pack(side="left", padx=3)

            height_var = tk.StringVar(value=f"{float(foto.get('max_height_cm', self.foto_h_var.get())):.1f}")
            height_menu = ctk.CTkOptionMenu(
                row,
                variable=height_var,
                values=["5.0", "6.0", "7.0", "8.1", "9.5", "11.0", "13.0", "16.0", "20.0"],
                width=85,
                command=lambda v, i=idx: self._set_foto_height(i, v),
            )
            height_menu.pack(side="left", padx=3)

            width_var = tk.StringVar(value=f"{int(float(foto.get('width_percent', 100)))}%")
            width_menu = ctk.CTkOptionMenu(
                row,
                variable=width_var,
                values=["70%", "80%", "90%", "100%", "110%", "120%", "130%", "150%", "170%", "200%"],
                width=78,
                command=lambda v, i=idx: self._set_foto_width(i, v),
            )
            width_menu.pack(side="left", padx=3)

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

    def generate(self):
        texto = self.text.get("1.0", "end").strip()
        if not texto:
            messagebox.showerror("Erro", "Cole o texto primeiro")
            return

        save = filedialog.asksaveasfilename(
            initialdir=self._pick_initial_dir(),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return
        self._remember_dir(save)

        sections = self._get_sections_from_ui()
        self._set_status("Gerando PDF...")
        self.update_idletasks()
        try:
            self._set_status("Processando imagens...")
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
                cover_header_scale=float(self.cover_header_scale_var.get()),
            )
            self._set_status("PDF gerado com sucesso.")
            messagebox.showinfo("Sucesso", "PDF gerado!")
        except Exception as e:
            self._set_status("Erro ao gerar PDF.")
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
