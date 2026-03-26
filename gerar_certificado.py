"""
Módulo: Gerador de Certificados
Compatível com o sistema modular do Gerador de Relatórios (Grupo BAW).
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import json, os, re, threading, tempfile, subprocess, sys, uuid
from copy import deepcopy

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak, ListFlowable, ListItem, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.pdfbase import pdfmetrics

from pypdf import PdfReader, PdfWriter
import fitz
from PIL import Image as PILImage, ImageTk

try:
    from module_registry import discover_modules
    HAS_REGISTRY = True
except ImportError:
    HAS_REGISTRY = False

# ─── Arquivo de estado ──────────────────────────────────────────────────────
LAUNCHER_STATE_FILE = os.path.join(os.path.dirname(__file__), "launcher_state.json")
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config_certificado.json")

PAGE_SIZE = landscape(A4)   # Certificado em paisagem
W, H = PAGE_SIZE


def save_last_module(module_filename):
    try:
        with open(LAUNCHER_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_module": module_filename}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ─── Config ─────────────────────────────────────────────────────────────────
def load_config() -> dict:
    defaults = {
        "template_frente": "",
        "template_verso": "",
        "lista_nomes_path": "",
        "texto_certificado": "",
        "topicos": "",
        "instrutor": "",
        "carga_horaria": "",
        "data_realizacao": "",
        "local": "",
        "last_dir": "",
        "window_geometry": "1500x860",
        "preview_auto": True,
        "zoom_factor": 1.0,
    }
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            defaults.update(data)
    return defaults


def save_config(cfg: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _sanitize_filename(value: str) -> str:
    nome = re.sub(r'[\\/:*?"<>|]+', "_", (value or "").strip())
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome or "Participante"


# ─── PDF do Certificado ──────────────────────────────────────────────────────
def _build_frente_story(styles, aprendiz, texto, instrutor, carga, data, local):
    """Constrói o story ReportLab para a frente do certificado."""
    story = []

    # Título
    story.append(Spacer(1, 1.1 * cm))
    story.append(Paragraph(
        "<b>CERTIFICADO DE PARTICIPAÇÃO</b>",
        ParagraphStyle("CertTitulo",
                       parent=styles["Normal"],
                       fontSize=30, alignment=1,
                       textColor=colors.HexColor("#0A2238"),
                       leading=36, spaceAfter=6,
                       fontName="Helvetica-Bold")
    ))
    story.append(HRFlowable(width="70%", thickness=2,
                             color=colors.HexColor("#9FB3C7"), hAlign="CENTER"))
    story.append(Spacer(1, 0.7 * cm))

    # "Certificamos que"
    story.append(Paragraph(
        "Certificamos que",
        ParagraphStyle("CertSub", parent=styles["Normal"],
                       fontSize=14, alignment=1,
                       textColor=colors.HexColor("#4A5568"), leading=20)
    ))
    story.append(Spacer(1, 0.35 * cm))

    # Nome do aprendiz — destaque
    story.append(Paragraph(
        f"<b>{aprendiz or '[Nome do Participante]'}</b>",
        ParagraphStyle("CertNome", parent=styles["Normal"],
                       fontSize=26, alignment=1,
                       textColor=colors.HexColor("#0E2A44"),
                       leading=32, spaceAfter=4,
                       fontName="Helvetica-Bold")
    ))
    story.append(HRFlowable(width="50%", thickness=1,
                             color=colors.HexColor("#C2D1DE"), hAlign="CENTER"))
    story.append(Spacer(1, 0.5 * cm))

    # Texto livre
    corpo_style = ParagraphStyle("CertCorpo", parent=styles["Normal"],
                                 fontSize=12, alignment=1,
                                 textColor=colors.HexColor("#2D3748"),
                                 leading=18, spaceAfter=6)
    if texto and texto.strip():
        for linha in texto.strip().splitlines():
            linha = linha.strip()
            if linha:
                story.append(Paragraph(linha, corpo_style))
        story.append(Spacer(1, 0.4 * cm))

    # Informações adicionais
    infos = []
    if carga:
        infos.append(f"Carga horária: <b>{carga}</b>")
    if data:
        infos.append(f"Realizado em: <b>{data}</b>")
    if local:
        infos.append(f"Local: <b>{local}</b>")
    if infos:
        story.append(Paragraph(
            "  |  ".join(infos),
            ParagraphStyle("CertInfo", parent=styles["Normal"],
                           fontSize=10, alignment=1,
                           textColor=colors.HexColor("#5B6E7D"), leading=14)
        ))
        story.append(Spacer(1, 0.6 * cm))

    # Assinaturas
    story.append(Spacer(1, 0.3 * cm))
    largura_util = W - 5 * cm
    col = (largura_util - 2 * cm) / 2

    sig_label = ParagraphStyle("SigLabel", parent=styles["Normal"],
                                fontSize=10, alignment=1,
                                textColor=colors.HexColor("#4A5568"))
    sig_nome = ParagraphStyle("SigNome", parent=styles["Normal"],
                               fontSize=11, alignment=1, fontName="Helvetica-Bold",
                               textColor=colors.HexColor("#0E2A44"))

    dados_sig = [[
        [
            Spacer(1, 1.4 * cm),
            HRFlowable(width=col * 0.80, thickness=1,
                       color=colors.HexColor("#9FB3C7"), hAlign="CENTER"),
            Spacer(1, 4),
            Paragraph(f"<b>{instrutor or 'Instrutor'}</b>", sig_nome),
            Paragraph("Instrutor Responsável", sig_label),
        ],
        [
            Spacer(1, 1.4 * cm),
            HRFlowable(width=col * 0.80, thickness=1,
                       color=colors.HexColor("#9FB3C7"), hAlign="CENTER"),
            Spacer(1, 4),
            Paragraph(f"<b>{aprendiz or 'Participante'}</b>", sig_nome),
            Paragraph("Participante", sig_label),
        ],
    ]]
    tbl = Table(dados_sig, colWidths=[col, col])
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.5 * cm))

    return story


def _build_verso_story(styles, aprendiz, topicos, instrutor, carga, data):
    """Constrói o story ReportLab para o verso do certificado."""
    story = []
    story.append(Spacer(1, 1.2 * cm))

    story.append(Paragraph(
        "<b>CONTEÚDO PROGRAMÁTICO</b>",
        ParagraphStyle("VersoTitulo", parent=styles["Normal"],
                       fontSize=20, alignment=1,
                       textColor=colors.HexColor("#0A2238"),
                       leading=26, spaceAfter=4,
                       fontName="Helvetica-Bold")
    ))
    story.append(HRFlowable(width="55%", thickness=1.5,
                             color=colors.HexColor("#9FB3C7"), hAlign="CENTER"))
    story.append(Spacer(1, 0.8 * cm))

    topico_style = ParagraphStyle("TopStyle", parent=styles["Normal"],
                                  fontSize=11, leading=17,
                                  textColor=colors.HexColor("#1F2B37"))

    if topicos and topicos.strip():
        linhas = [l.strip() for l in topicos.strip().splitlines() if l.strip()]
        itens = [
            ListItem(Paragraph(l, topico_style), leftIndent=14, bulletText="▸")
            for l in linhas
        ]
        story.append(ListFlowable(
            itens,
            bulletType="bullet",
            leftIndent=10,
            bulletFontName="Helvetica",
            bulletFontSize=11,
        ))
    else:
        story.append(Paragraph(
            "Nenhum tópico informado.",
            ParagraphStyle("SemTop", parent=styles["Normal"],
                           fontSize=11, textColor=colors.HexColor("#8A9BAD"))
        ))

    story.append(Spacer(1, 0.9 * cm))
    story.append(HRFlowable(width="80%", thickness=0.6,
                             color=colors.HexColor("#C2D1DE"), hAlign="CENTER"))
    story.append(Spacer(1, 0.5 * cm))

    # Rodapé verso
    rodape_partes = []
    if aprendiz:
        rodape_partes.append(f"Participante: <b>{aprendiz}</b>")
    if instrutor:
        rodape_partes.append(f"Instrutor: <b>{instrutor}</b>")
    if carga:
        rodape_partes.append(f"Carga: <b>{carga}</b>")
    if data:
        rodape_partes.append(f"Data: <b>{data}</b>")
    if rodape_partes:
        story.append(Paragraph(
            "  ·  ".join(rodape_partes),
            ParagraphStyle("VersoRodape", parent=styles["Normal"],
                           fontSize=9.5, alignment=1,
                           textColor=colors.HexColor("#5B6E7D"), leading=14)
        ))

    return story


def _pick_template_page(template_path: str, fallback_path: str = ""):
    """Retorna a página de template e o reader associado (ou None)."""
    for path in [template_path, fallback_path]:
        if not path or not os.path.exists(path):
            continue
        try:
            reader = PdfReader(path)
            if reader.pages:
                return reader.pages[0], reader
        except Exception:
            continue
    return None, None


def gerar_certificado_pdf(
    output_path: str,
    aprendiz: str,
    texto: str,
    topicos: str,
    instrutor: str,
    carga: str,
    data: str,
    local: str,
    template_frente: str = "",
    template_verso: str = "",
):
    """Gera o PDF do certificado (frente + verso) com templates opcionais.

    O conteúdo textual é sempre desenhado acima do template para evitar
    que edições prévias no template escondam o texto do certificado.
    """
    tmp_id = uuid.uuid4().hex[:8]
    tmp_frente = output_path + f".frente_{tmp_id}.pdf"
    tmp_verso = output_path + f".verso_{tmp_id}.pdf"

    styles = getSampleStyleSheet()

    doc_frente = SimpleDocTemplate(
        tmp_frente, pagesize=PAGE_SIZE,
        leftMargin=2.5 * cm, rightMargin=2.5 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
    )
    doc_frente.build(
        _build_frente_story(styles, aprendiz, texto, instrutor, carga, data, local),
    )

    doc_verso = SimpleDocTemplate(
        tmp_verso, pagesize=PAGE_SIZE,
        leftMargin=2.5 * cm, rightMargin=2.5 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
    )
    doc_verso.build(
        _build_verso_story(styles, aprendiz, topicos, instrutor, carga, data),
    )

    # ── Junta frente + verso aplicando template como camada inferior ───────
    writer = PdfWriter()
    page_configs = [
        (tmp_frente, template_frente, ""),
        (tmp_verso, template_verso, template_frente),
    ]
    for content_pdf, template_pdf, fallback in page_configs:
        content_reader = PdfReader(content_pdf)
        content_page = content_reader.pages[0]
        template_page, _ = _pick_template_page(template_pdf, fallback)
        if template_page is not None:
            final_page = deepcopy(template_page)
            final_page.merge_page(content_page)
            writer.add_page(final_page)
        else:
            writer.add_page(content_page)
    with open(output_path, "wb") as f:
        writer.write(f)

    for tmp in [tmp_frente, tmp_verso]:
        try:
            os.remove(tmp)
        except Exception:
            pass


# ─── Preview Engine ──────────────────────────────────────────────────────────
class PreviewEngine:
    """Gera prévia do PDF em thread separada com debounce."""

    def __init__(self, on_update, debounce_ms=900):
        self.on_update = on_update
        self.debounce_ms = debounce_ms
        self._lock = threading.Lock()
        self._timer = None
        self._params = {}
        self._temp_dir = tempfile.mkdtemp()
        self._running = True

    def schedule(self, **params):
        with self._lock:
            self._params = params
            if self._timer:
                self._timer.cancel()
            self._timer = threading.Timer(self.debounce_ms / 1000.0, self._generate)
            self._timer.daemon = True
            self._timer.start()

    def _generate(self):
        if not self._running:
            return
        with self._lock:
            params = dict(self._params)
        tmp = os.path.join(self._temp_dir, f"prev_{uuid.uuid4().hex[:8]}.pdf")
        try:
            gerar_certificado_pdf(output_path=tmp, **params)
            doc = fitz.open(tmp)
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
                os.remove(tmp)
            except Exception:
                pass

    def stop(self):
        self._running = False
        if self._timer:
            self._timer.cancel()


# ─── Painel de Prévia ────────────────────────────────────────────────────────
class PreviewPanel(ctk.CTkFrame):
    def __init__(self, master, label="", **kwargs):
        super().__init__(master, **kwargs)
        self._label_txt = label
        self._pages = []
        self._tk_images = []
        self._page_idx = 0
        self._zoom = 1.0
        self._panel_w = 400

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=6, pady=(6, 2))
        ctk.CTkLabel(header, text=label,
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color="#7EB3E8").pack(side="left")

        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.pack(pady=(2, 2))
        self._btn_prev = ctk.CTkButton(nav, text="◀", width=32, command=self._prev)
        self._lbl_page = ctk.CTkLabel(nav, text="", width=55)
        self._btn_next = ctk.CTkButton(nav, text="▶", width=32, command=self._next)
        self._btn_zm = ctk.CTkButton(nav, text="−", width=28, command=self._zoom_out)
        self._lbl_zm = ctk.CTkLabel(nav, text="100%", width=42)
        self._btn_zp = ctk.CTkButton(nav, text="+", width=28, command=self._zoom_in)
        for w in [self._btn_prev, self._lbl_page, self._btn_next,
                  ctk.CTkLabel(nav, text="|", text_color="#555"),
                  self._btn_zm, self._lbl_zm, self._btn_zp]:
            w.pack(side="left", padx=1)

        cont = ctk.CTkFrame(self, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=4, pady=4)
        self._canvas = tk.Canvas(cont, bg="#D8E4EE", highlightthickness=0)
        vsb = ctk.CTkScrollbar(cont, orientation="vertical", command=self._canvas.yview)
        hsb = ctk.CTkScrollbar(cont, orientation="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._canvas.bind("<MouseWheel>", self._on_scroll)
        self._canvas.bind("<Button-4>", self._on_scroll)
        self._canvas.bind("<Button-5>", self._on_scroll)

        self._status_lbl = ctk.CTkLabel(self, text="Prévia aparecerá aqui...",
                                         text_color="#8A9BAD")
        self._status_lbl.pack(expand=True)

    def _on_scroll(self, event):
        if event.state & 0x0004:
            if event.delta > 0 or event.num == 4:
                self._zoom_in()
            else:
                self._zoom_out()
        else:
            d = -1 if (event.delta > 0 or event.num == 4) else 1
            self._canvas.yview_scroll(d, "units")

    def update_pages(self, images):
        self._status_lbl.pack_forget()
        self._pages = images
        if self._page_idx >= len(images):
            self._page_idx = 0
        self._render()

    def show_status(self, msg):
        self._pages = []
        self._canvas.delete("all")
        self._status_lbl.configure(text=msg)
        self._status_lbl.pack(expand=True)

    def set_panel_width(self, w):
        if abs(self._panel_w - w) > 15:
            self._panel_w = w
            if self._pages:
                self._render()

    def _render(self):
        if not self._pages:
            return
        img = self._pages[self._page_idx]
        base_w = max(self._panel_w - 30, 200)
        ratio = (base_w / img.width) * self._zoom
        tw, th = int(img.width * ratio), int(img.height * ratio)
        resized = img.resize((tw, th), PILImage.LANCZOS)
        tk_img = ImageTk.PhotoImage(resized)
        self._tk_images = [tk_img]
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, anchor="nw", image=tk_img)
        self._canvas.config(scrollregion=(0, 0, tw, th))
        total = len(self._pages)
        self._lbl_page.configure(text=f"{self._page_idx + 1}/{total}")
        self._lbl_zm.configure(text=f"{int(self._zoom * 100)}%")
        self._btn_prev.configure(state="normal" if self._page_idx > 0 else "disabled")
        self._btn_next.configure(state="normal" if self._page_idx < total - 1 else "disabled")

    def _prev(self):
        if self._page_idx > 0:
            self._page_idx -= 1
            self._render()

    def _next(self):
        if self._page_idx < len(self._pages) - 1:
            self._page_idx += 1
            self._render()

    def _zoom_in(self):
        self._zoom = min(4.0, self._zoom + 0.20)
        self._render()

    def _zoom_out(self):
        self._zoom = max(0.3, self._zoom - 0.20)
        self._render()


# ─── App Principal ───────────────────────────────────────────────────────────
class CertificadoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Certificados")
        self.cfg = load_config()
        self.geometry(self.cfg.get("window_geometry", "1500x860"))
        self.minsize(1200, 720)

        self._suspend_preview = False
        self._last_pdf = ""
        self._lista_nomes = []

        self._build_ui()
        self._load_cfg_to_ui()

        self._engine = PreviewEngine(on_update=self._on_preview_ready, debounce_ms=800)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_resize)
        self.bind("<Control-g>", lambda e: self._gerar())
        self.bind("<Control-G>", lambda e: self._gerar())

        if self.cfg.get("preview_auto", True):
            self.after(400, self._schedule_preview)

    # ── Construção da UI ─────────────────────────────────────────────────────
    def _build_ui(self):
        # Barra superior
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(top, text="🎓 Gerador de Certificados",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="#7EB3E8").pack(side="left")

        # Navegação entre módulos
        if HAS_REGISTRY:
            self._module_targets = {label: name for name, label in discover_modules()
                                    if name != "gerar_certificado.py"}
            if self._module_targets:
                ctk.CTkLabel(top, text="Ir para:").pack(side="left", padx=(20, 4))
                self._mod_var = tk.StringVar(value=list(self._module_targets.keys())[0])
                ctk.CTkOptionMenu(top, variable=self._mod_var,
                                  values=list(self._module_targets.keys()),
                                  width=220).pack(side="left", padx=(0, 4))
                ctk.CTkButton(top, text="Abrir", width=72,
                              command=self._switch_module).pack(side="left")

        self._btn_abrir = ctk.CTkButton(top, text="📂 Abrir PDF", width=110,
                                         command=self._abrir_pdf,
                                         fg_color="#2A6040", hover_color="#1E4A30")
        # não pack ainda

        ctk.CTkButton(top, text="✔ Gerar PDF  (Ctrl+G)", width=160,
                      command=self._gerar,
                      fg_color="#1A4A7A", hover_color="#133A62").pack(side="right", padx=4)
        ctk.CTkButton(top, text="Limpar nome", width=110,
                      command=self._limpar_aprendiz).pack(side="right", padx=4)

        self._chk_auto = tk.BooleanVar(value=bool(self.cfg.get("preview_auto", True)))
        ctk.CTkCheckBox(top, text="Prévia automática",
                        variable=self._chk_auto).pack(side="right", padx=10)
        ctk.CTkButton(top, text="Atualizar prévia", width=130,
                      command=self._schedule_preview).pack(side="right", padx=4)

        # Área principal: esquerda (formulário) | direita (prévia)
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        # ── Painel esquerdo ──────────────────────────────────────────────────
        left = ctk.CTkScrollableFrame(main, width=520, label_text="")
        left.pack(side="left", fill="both", padx=(0, 10))

        # Templates
        tpl_frame = ctk.CTkFrame(left, fg_color="transparent")
        tpl_frame.pack(fill="x", pady=(4, 0))
        ctk.CTkLabel(tpl_frame, text="TEMPLATES DE FUNDO",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7D93A8").pack(anchor="w", padx=4, pady=(6, 2))

        self._lbl_frente = self._template_row(left, "Frente:", "frente")
        self._lbl_verso = self._template_row(left, "Verso: ", "verso")

        ctk.CTkFrame(left, height=1, fg_color="#2A3A4A").pack(fill="x", pady=8)

        # Nome do Aprendiz
        ctk.CTkLabel(left, text="NOME DO PARTICIPANTE",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7D93A8").pack(anchor="w", padx=4, pady=(4, 2))
        self._ent_aprendiz = ctk.CTkEntry(left, placeholder_text="Nome completo do participante",
                                           font=ctk.CTkFont(size=14))
        self._ent_aprendiz.pack(fill="x", padx=4, pady=(0, 6))
        self._ent_aprendiz.bind("<KeyRelease>", self._on_field_change)
        self._lbl_lista_nomes = ctk.CTkLabel(
            left,
            text="Lista: não carregada",
            text_color="#6A7D8D",
            font=ctk.CTkFont(size=10),
        )
        self._lbl_lista_nomes.pack(anchor="w", padx=4, pady=(0, 4))
        row_lista = ctk.CTkFrame(left, fg_color="transparent")
        row_lista.pack(fill="x", padx=4, pady=(0, 6))
        ctk.CTkButton(row_lista, text="Importar lista TXT", width=130,
                      command=self._importar_lista_nomes).pack(side="left")
        ctk.CTkButton(row_lista, text="Limpar lista", width=96,
                      command=self._limpar_lista_nomes).pack(side="left", padx=6)

        ctk.CTkFrame(left, height=1, fg_color="#2A3A4A").pack(fill="x", pady=4)

        # Informações salvas
        ctk.CTkLabel(left, text="INFORMAÇÕES DO TREINAMENTO  (salvas automaticamente)",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7D93A8").pack(anchor="w", padx=4, pady=(6, 2))

        self._ent_instrutor = self._labeled_entry(left, "Instrutor responsável:")
        self._ent_carga = self._labeled_entry(left, "Carga horária:")
        self._ent_data = self._labeled_entry(left, "Data de realização:")
        self._ent_local = self._labeled_entry(left, "Local:")

        ctk.CTkFrame(left, height=1, fg_color="#2A3A4A").pack(fill="x", pady=8)

        # Texto do certificado
        ctk.CTkLabel(left, text="TEXTO DO CERTIFICADO  (frente)",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7D93A8").pack(anchor="w", padx=4, pady=(4, 2))
        ctk.CTkLabel(left, text="Cole ou edite o texto que aparecerá na frente do certificado.",
                     text_color="#6A7D8D", font=ctk.CTkFont(size=10)).pack(anchor="w", padx=4)
        self._txt_texto = ctk.CTkTextbox(left, height=130,
                                          font=ctk.CTkFont(family="Courier New", size=11),
                                          wrap="word")
        self._txt_texto.pack(fill="x", padx=4, pady=(4, 6))
        self._txt_texto.bind("<<Modified>>", self._on_textbox_change)

        ctk.CTkFrame(left, height=1, fg_color="#2A3A4A").pack(fill="x", pady=4)

        # Tópicos (verso)
        ctk.CTkLabel(left, text="TÓPICOS ABORDADOS  (verso)",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7D93A8").pack(anchor="w", padx=4, pady=(6, 2))
        ctk.CTkLabel(left, text="Um tópico por linha. Aparecerão como lista no verso do certificado.",
                     text_color="#6A7D8D", font=ctk.CTkFont(size=10)).pack(anchor="w", padx=4)
        self._txt_topicos = ctk.CTkTextbox(left, height=160,
                                            font=ctk.CTkFont(family="Courier New", size=11),
                                            wrap="word")
        self._txt_topicos.pack(fill="x", padx=4, pady=(4, 6))
        self._txt_topicos.bind("<<Modified>>", self._on_textbox_change)

        # ── Painel direito: prévia frente + verso lado a lado ───────────────
        right = ctk.CTkFrame(main, fg_color="transparent")
        right.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(right, text="PRÉ-VISUALIZAÇÃO EM TEMPO REAL",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#7D93A8").pack(anchor="w", padx=4, pady=(4, 2))

        previews = ctk.CTkFrame(right, fg_color="transparent")
        previews.pack(fill="both", expand=True)
        previews.columnconfigure(0, weight=1)
        previews.columnconfigure(1, weight=1)
        previews.rowconfigure(0, weight=1)

        self._prev_frente = PreviewPanel(
            previews,
            label="✦ FRENTE",
            fg_color=("#D8E4EE", "#1A2A38"),
            corner_radius=8,
        )
        self._prev_frente.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        self._prev_verso = PreviewPanel(
            previews,
            label="✦ VERSO",
            fg_color=("#D8E4EE", "#1A2A38"),
            corner_radius=8,
        )
        self._prev_verso.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        # Status
        self._status = ctk.CTkLabel(self, text="Pronto.", anchor="w")
        self._status.pack(fill="x", padx=14, pady=(0, 8))

    def _template_row(self, parent, label, key):
        """Cria linha de seleção de template com label de caminho."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=4, pady=2)
        ctk.CTkLabel(row, text=label, width=52).pack(side="left")
        lbl = ctk.CTkLabel(row, text="Não selecionado",
                            text_color="#5B7A8D",
                            font=ctk.CTkFont(size=10),
                            anchor="w")
        lbl.pack(side="left", fill="x", expand=True, padx=(4, 8))
        ctk.CTkButton(row, text="Selecionar", width=96,
                      command=lambda: self._pick_template(key, lbl)).pack(side="left", padx=2)
        ctk.CTkButton(row, text="Limpar", width=68,
                      command=lambda: self._clear_template(key, lbl)).pack(side="left")
        return lbl

    def _labeled_entry(self, parent, label):
        ctk.CTkLabel(parent, text=label, text_color="#A0B4C4",
                     font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4, pady=(4, 0))
        ent = ctk.CTkEntry(parent)
        ent.pack(fill="x", padx=4, pady=(0, 4))
        ent.bind("<KeyRelease>", self._on_field_change)
        return ent

    # ── Carrega config na UI ──────────────────────────────────────────────────
    def _load_cfg_to_ui(self):
        self._suspend_preview = True

        for key, lbl in [("template_frente", self._lbl_frente),
                          ("template_verso", self._lbl_verso)]:
            path = self.cfg.get(key, "")
            if path and os.path.exists(path):
                lbl.configure(text=os.path.basename(path))
            else:
                lbl.configure(text="Não selecionado")

        for ent, key in [
            (self._ent_instrutor, "instrutor"),
            (self._ent_carga, "carga_horaria"),
            (self._ent_data, "data_realizacao"),
            (self._ent_local, "local"),
        ]:
            ent.delete(0, "end")
            ent.insert(0, self.cfg.get(key, ""))

        self._txt_texto.delete("1.0", "end")
        self._txt_texto.insert("1.0", self.cfg.get("texto_certificado", ""))
        self._txt_texto.edit_modified(False)

        self._txt_topicos.delete("1.0", "end")
        self._txt_topicos.insert("1.0", self.cfg.get("topicos", ""))
        self._txt_topicos.edit_modified(False)

        lista_path = self.cfg.get("lista_nomes_path", "")
        if lista_path and os.path.exists(lista_path):
            self._carregar_lista_nomes(lista_path, atualizar_preview=False)
        else:
            self._atualizar_label_lista_nomes()

        self._suspend_preview = False

    # ── Coleta dados da UI ────────────────────────────────────────────────────
    def _collect_params(self) -> dict:
        aprendiz_preview = self._ent_aprendiz.get().strip()
        if self._lista_nomes:
            aprendiz_preview = self._lista_nomes[0]
        return dict(
            aprendiz=aprendiz_preview,
            texto=self._txt_texto.get("1.0", "end").strip(),
            topicos=self._txt_topicos.get("1.0", "end").strip(),
            instrutor=self._ent_instrutor.get().strip(),
            carga=self._ent_carga.get().strip(),
            data=self._ent_data.get().strip(),
            local=self._ent_local.get().strip(),
            template_frente=self.cfg.get("template_frente", ""),
            template_verso=self.cfg.get("template_verso", ""),
        )

    # ── Salvar dados não-aprendiz no config ────────────────────────────────────
    def _sync_cfg(self):
        self.cfg["texto_certificado"] = self._txt_texto.get("1.0", "end").strip()
        self.cfg["topicos"] = self._txt_topicos.get("1.0", "end").strip()
        self.cfg["instrutor"] = self._ent_instrutor.get().strip()
        self.cfg["carga_horaria"] = self._ent_carga.get().strip()
        self.cfg["data_realizacao"] = self._ent_data.get().strip()
        self.cfg["local"] = self._ent_local.get().strip()
        self.cfg["preview_auto"] = bool(self._chk_auto.get())
        self.cfg["lista_nomes_path"] = self.cfg.get("lista_nomes_path", "")
        save_config(self.cfg)

    # ── Templates ───────────────────────────────────────────────────────────
    def _pick_template(self, key: str, lbl):
        path = filedialog.askopenfilename(
            initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
            filetypes=[("PDF / Imagem", "*.pdf *.png *.jpg *.jpeg *.bmp")],
        )
        if not path:
            return
        self.cfg[f"template_{key}"] = path
        self.cfg["last_dir"] = os.path.dirname(path)
        lbl.configure(text=os.path.basename(path))
        save_config(self.cfg)
        self._schedule_preview()

    def _clear_template(self, key: str, lbl):
        self.cfg[f"template_{key}"] = ""
        lbl.configure(text="Não selecionado")
        save_config(self.cfg)
        self._schedule_preview()

    # ── Prévia ────────────────────────────────────────────────────────────────
    def _on_field_change(self, event=None):
        if self._suspend_preview:
            return
        if self._chk_auto.get():
            self._schedule_preview()

    def _on_textbox_change(self, event=None):
        if event:
            event.widget.edit_modified(False)
        if self._suspend_preview:
            return
        if self._chk_auto.get():
            self._schedule_preview()

    def _schedule_preview(self):
        params = self._collect_params()
        self._prev_frente.show_status("⏳ Atualizando...")
        self._prev_verso.show_status("⏳ Atualizando...")
        self._status.configure(text="Gerando prévia...")
        self._engine.schedule(**params)

    def _on_preview_ready(self, images, error):
        def _update():
            if error:
                self._prev_frente.show_status(f"⚠ {error[:100]}")
                self._prev_verso.show_status(f"⚠ {error[:100]}")
                self._status.configure(text="Erro na prévia.")
            elif images:
                # Página 0 = frente, página 1 = verso
                if len(images) >= 1:
                    self._prev_frente.update_pages([images[0]])
                if len(images) >= 2:
                    self._prev_verso.update_pages([images[1]])
                self._status.configure(text="Prévia atualizada.")
        self.after(0, _update)

    def _on_resize(self, event=None):
        for panel in [self._prev_frente, self._prev_verso]:
            w = panel.winfo_width()
            if w > 50:
                panel.set_panel_width(w)

    # ── Gerar PDF ─────────────────────────────────────────────────────────────
    def _gerar(self):
        aprendiz_digitado = self._ent_aprendiz.get().strip()
        nomes_geracao = list(self._lista_nomes) if self._lista_nomes else []
        if not nomes_geracao and not aprendiz_digitado:
            messagebox.showerror("Erro", "Informe o nome do participante.")
            return
        if not nomes_geracao:
            nomes_geracao = [aprendiz_digitado]

        if len(nomes_geracao) == 1:
            save = filedialog.asksaveasfilename(
                initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"Certificado - {nomes_geracao[0]}.pdf",
            )
            if not save:
                return
            output_paths = [(nomes_geracao[0], save)]
            self.cfg["last_dir"] = os.path.dirname(save)
        else:
            out_dir = filedialog.askdirectory(
                initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
                title="Selecione a pasta de saída para os certificados",
            )
            if not out_dir:
                return
            output_paths = []
            for nome in nomes_geracao:
                filename = f"Certificado - {_sanitize_filename(nome)}.pdf"
                output_paths.append((nome, os.path.join(out_dir, filename)))
            self.cfg["last_dir"] = out_dir
        self._sync_cfg()

        self._status.configure(
            text="Gerando PDF..." if len(output_paths) == 1
            else f"Gerando {len(output_paths)} PDFs..."
        )
        params_base = self._collect_params()

        def _worker():
            try:
                ultimo_arquivo = ""
                for nome, path in output_paths:
                    params = dict(params_base)
                    params["aprendiz"] = nome
                    gerar_certificado_pdf(output_path=path, **params)
                    ultimo_arquivo = path
                self.after(
                    0,
                    lambda p=ultimo_arquivo, total=len(output_paths): self._on_gerar_ok(p, total),
                )
            except Exception as e:
                self.after(0, lambda err=str(e): self._on_gerar_err(err))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_gerar_ok(self, path, total_arquivos=1):
        self._last_pdf = path
        if total_arquivos == 1:
            self._status.configure(text=f"PDF gerado: {os.path.basename(path)}")
        else:
            self._status.configure(text=f"{total_arquivos} PDFs gerados com sucesso.")
        self._btn_abrir.pack(side="left", padx=6)
        if total_arquivos == 1:
            msg = f"Certificado gerado!\n{os.path.basename(path)}\n\nAbrir agora?"
        else:
            msg = (f"Foram gerados {total_arquivos} certificados.\n"
                   f"Último arquivo: {os.path.basename(path)}\n\nAbrir último PDF agora?")
        if messagebox.askyesno("Sucesso", msg):
            self._abrir_pdf()

    def _on_gerar_err(self, err):
        self._status.configure(text="Erro ao gerar certificado.")
        messagebox.showerror("Erro", err)

    def _abrir_pdf(self):
        if self._last_pdf and os.path.exists(self._last_pdf):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(self._last_pdf)
                elif sys.platform.startswith("darwin"):
                    subprocess.run(["open", self._last_pdf], check=False)
                else:
                    subprocess.run(["xdg-open", self._last_pdf], check=False)
            except Exception:
                pass
        else:
            self._status.configure(text="Nenhum PDF gerado nesta sessão.")

    def _limpar_aprendiz(self):
        self._ent_aprendiz.delete(0, "end")
        if self._chk_auto.get():
            self._schedule_preview()

    def _atualizar_label_lista_nomes(self):
        if self._lista_nomes:
            origem = os.path.basename(self.cfg.get("lista_nomes_path", ""))
            self._lbl_lista_nomes.configure(
                text=f"Lista: {len(self._lista_nomes)} nomes ({origem})"
            )
        else:
            self._lbl_lista_nomes.configure(text="Lista: não carregada")

    def _carregar_lista_nomes(self, path: str, atualizar_preview=True):
        with open(path, "r", encoding="utf-8") as f:
            nomes = [linha.strip() for linha in f if linha.strip()]
        if not nomes:
            raise ValueError("O arquivo TXT não contém nomes válidos.")
        self._lista_nomes = nomes
        self.cfg["lista_nomes_path"] = path
        self.cfg["last_dir"] = os.path.dirname(path)
        self._atualizar_label_lista_nomes()
        save_config(self.cfg)
        if atualizar_preview:
            self._schedule_preview()

    def _importar_lista_nomes(self):
        path = filedialog.askopenfilename(
            initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
            filetypes=[("Arquivo TXT", "*.txt")],
        )
        if not path:
            return
        try:
            self._carregar_lista_nomes(path)
        except Exception as exc:
            messagebox.showerror("Erro ao importar lista", str(exc))

    def _limpar_lista_nomes(self):
        self._lista_nomes = []
        self.cfg["lista_nomes_path"] = ""
        self._atualizar_label_lista_nomes()
        save_config(self.cfg)
        if self._chk_auto.get():
            self._schedule_preview()

    # ── Módulos ───────────────────────────────────────────────────────────────
    def _switch_module(self):
        if not HAS_REGISTRY:
            return
        label = self._mod_var.get()
        module_filename = self._module_targets.get(label)
        if not module_filename:
            return
        module_path = os.path.join(os.path.dirname(__file__), module_filename)
        if not os.path.exists(module_path):
            messagebox.showerror("Erro", f"Módulo não encontrado: {module_filename}")
            return
        try:
            save_last_module(module_filename)
            subprocess.Popen([sys.executable, module_path])
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Erro", f"Falha ao abrir módulo: {exc}")

    # ── Fechar ────────────────────────────────────────────────────────────────
    def _on_close(self):
        self.cfg["window_geometry"] = self.geometry()
        self._sync_cfg()
        self._engine.stop()
        self.destroy()


if __name__ == "__main__":
    save_last_module("gerar_certificado.py")
    ctk.set_appearance_mode("dark")
    app = CertificadoApp()
    app.mainloop()
