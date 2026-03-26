"""
Módulo: Gerador de Certificados — Grupo BAW
Reproduz fielmente o layout do template (borda azul dupla + título CERTIFICADO).
Usa reportlab canvas direto para posicionamento pixel-perfect sobre o template PDF.
A prévia renderiza o template + conteúdo exatamente como o PDF final.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import json, os, re, threading, tempfile, subprocess, sys, uuid, shutil

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import cm, mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics

from pypdf import PdfReader, PdfWriter

try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

from PIL import Image as PILImage, ImageTk

try:
    from module_registry import discover_modules
    HAS_REGISTRY = True
except ImportError:
    HAS_REGISTRY = False

# ─── Dimensões da página ─────────────────────────────────────────────────────
PAGE_W, PAGE_H = landscape(A4)   # 841.9 x 595.3 pt

# ─── Cores do template BAW ───────────────────────────────────────────────────
C_AZUL   = colors.HexColor("#1B2A6B")
C_TEXTO  = colors.HexColor("#111111")
C_SUB    = colors.HexColor("#222222")
C_CINZA  = colors.HexColor("#444444")

# ─── Arquivos ────────────────────────────────────────────────────────────────
_DIR                = os.path.dirname(os.path.abspath(__file__))
LAUNCHER_STATE_FILE = os.path.join(_DIR, "launcher_state.json")
CONFIG_FILE         = os.path.join(_DIR, "config_certificado.json")


def save_last_module(name: str):
    try:
        with open(LAUNCHER_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_module": name}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ─── Config ──────────────────────────────────────────────────────────────────
_DEFAULTS = {
    "template_frente":  "",
    "empresa":          "A BAW Brasil Ind. e Com. Ltda",
    "nome_treinamento": "",
    "duracao":          "",
    "data_certificado": "",
    "local":            "",
    "supervisor_nome":  "",
    "supervisor_cargo": "Supervisor técnico",
    "instrutor_nome":   "",
    "instrutor_cargo":  "Instrutor técnico",
    "topicos":          "",
    "last_dir":         "",
    "window_geometry":  "1560x900",
    "preview_auto":     True,
}


def load_config() -> dict:
    cfg = dict(_DEFAULTS)
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    return cfg


def save_config(cfg: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ─── Geração do conteúdo do certificado ─────────────────────────────────────
def _draw_certificate(c: rl_canvas.Canvas, data: dict):
    """
    Desenha o conteúdo textual do certificado sobre o canvas.
    Layout baseado fielmente no exemplo_certificado.pdf fornecido:
      - "empresa certifica que:"  (centralizado, 14pt)
      - NOME DO PARTICIPANTE      (caixa alta, grande, negrito)
      - "concluiu satisfatoriamente o:"
      - Nome do treinamento       (negrito, pode quebrar linha)
      - Duração / Local / Data
      - Duas assinaturas na base
    """
    W, H = PAGE_W, PAGE_H
    cx = W / 2

    empresa          = data.get("empresa", "").strip()
    aprendiz         = data.get("aprendiz", "").strip()
    nome_treinamento = data.get("nome_treinamento", "").strip()
    duracao          = data.get("duracao", "").strip()
    data_cert        = data.get("data_certificado", "").strip()
    local_cert       = data.get("local", "").strip()
    sup_nome         = data.get("supervisor_nome", "").strip()
    sup_cargo        = data.get("supervisor_cargo", "Supervisor técnico").strip()
    ins_nome         = data.get("instrutor_nome", "").strip()
    ins_cargo        = data.get("instrutor_cargo", "Instrutor técnico").strip()

    # Área de conteúdo: começa abaixo do título "CERTIFICADO" do template
    # O título do template ocupa ~3.4 cm do topo
    y = H - 3.55 * cm

    # ── "Empresa certifica que:" ─────────────────────────────────────────────
    c.setFont("Helvetica", 13)
    c.setFillColor(C_SUB)
    txt_emp = f"{empresa} certifica que:" if empresa else "Certifica que:"
    c.drawCentredString(cx, y, txt_emp)
    y -= 0.95 * cm

    # ── Nome do participante — caixa alta, tamanho adaptativo ───────────────
    nome_display = aprendiz.upper() if aprendiz else "[NOME DO PARTICIPANTE]"
    max_w = PAGE_W - 5.5 * cm
    fsize = 52
    while fsize > 16:
        if pdfmetrics.stringWidth(nome_display, "Helvetica-Bold", fsize) <= max_w:
            break
        fsize -= 1

    c.setFont("Helvetica-Bold", fsize)
    c.setFillColor(C_TEXTO)
    c.drawCentredString(cx, y, nome_display)
    y -= (fsize * 0.75) / 72 * 2.54 * cm + 0.25 * cm

    # ── "concluiu satisfatoriamente o:" ─────────────────────────────────────
    c.setFont("Helvetica", 13)
    c.setFillColor(C_SUB)
    c.drawCentredString(cx, y, "concluiu satisfatoriamente o:")
    y -= 0.80 * cm

    # ── Nome do treinamento (negrito, quebra de linha automática) ────────────
    if nome_treinamento:
        fsize_t = 15
        max_w_t = PAGE_W - 5.0 * cm
        linhas = _wrap_text(nome_treinamento, "Helvetica-Bold", fsize_t, max_w_t)
        c.setFont("Helvetica-Bold", fsize_t)
        c.setFillColor(C_TEXTO)
        for linha in linhas:
            c.drawCentredString(cx, y, linha)
            y -= 0.62 * cm
    else:
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(C_TEXTO)
        c.drawCentredString(cx, y, "[Nome do Treinamento]")
        y -= 0.62 * cm

    y -= 0.22 * cm

    # ── Duração ───────────────────────────────────────────────────────────────
    if duracao:
        c.setFont("Helvetica", 13)
        c.setFillColor(C_SUB)
        c.drawCentredString(cx, y, f"Duração: {duracao}")
        y -= 0.64 * cm

    # ── Local ─────────────────────────────────────────────────────────────────
    if local_cert:
        c.setFont("Helvetica", 13)
        c.setFillColor(C_SUB)
        c.drawCentredString(cx, y, f"Local: {local_cert}")
        y -= 0.64 * cm

    # ── Data ──────────────────────────────────────────────────────────────────
    if data_cert:
        c.setFont("Helvetica", 13)
        c.setFillColor(C_SUB)
        c.drawCentredString(cx, y, data_cert)

    # ── Assinaturas ──────────────────────────────────────────────────────────
    # Posição fixa próxima ao rodapé (acima da borda inferior do template)
    sig_y   = 2.42 * cm    # Y da linha de assinatura
    sig_len = 6.0 * cm
    col_esq = W * 0.27
    col_dir = W * 0.73

    c.setStrokeColor(C_TEXTO)
    c.setLineWidth(0.75)

    for cx_sig, nome, cargo in [
        (col_esq, sup_nome, sup_cargo),
        (col_dir, ins_nome, ins_cargo),
    ]:
        c.line(cx_sig - sig_len / 2, sig_y, cx_sig + sig_len / 2, sig_y)
        c.setFont("Helvetica", 11)
        c.setFillColor(C_TEXTO)
        c.drawCentredString(cx_sig, sig_y - 0.42 * cm,
                            nome if nome else "___________________")
        c.setFont("Helvetica", 10)
        c.setFillColor(C_CINZA)
        c.drawCentredString(cx_sig, sig_y - 0.42 * cm - 0.42 * cm, cargo)


def _wrap_text(text: str, font: str, size: float, max_w: float) -> list:
    """Quebra texto em linhas que caibam em max_w pontos."""
    words = text.split()
    linhas, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if pdfmetrics.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                linhas.append(cur)
            cur = w
    if cur:
        linhas.append(cur)
    return linhas or [text]


# ─── Renderiza template como imagem PNG temporária ───────────────────────────
def _template_to_png(template_path: str, dpi_scale: float = 3.0) -> str | None:
    """Converte pág. 0 do PDF template em PNG temporário. Retorna path ou None."""
    if not HAS_FITZ or not template_path or not os.path.exists(template_path):
        return None
    try:
        doc = fitz.open(template_path)
        mat = fitz.Matrix(dpi_scale, dpi_scale)
        pix = doc[0].get_pixmap(matrix=mat, alpha=False)
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        pix.save(tmp.name)
        tmp.close()
        doc.close()
        return tmp.name
    except Exception:
        return None


# ─── Gerador principal do PDF ────────────────────────────────────────────────
def gerar_certificado_pdf(output_path: str, data: dict, template_frente: str = ""):
    """
    Gera o PDF do certificado:
    1. Cria PDF de conteúdo com reportlab canvas
    2. Se houver template, mescla template (fundo) + conteúdo via pypdf
    """
    tmp_content = output_path + f"._c_{uuid.uuid4().hex[:8]}.pdf"

    # ── Passo 1: gera conteúdo ─────────────────────────────────────────────
    c = rl_canvas.Canvas(tmp_content, pagesize=(PAGE_W, PAGE_H))
    _draw_certificate(c, data)
    c.save()

    # ── Passo 2: mescla com template ───────────────────────────────────────
    tpl = template_frente
    if tpl and os.path.exists(tpl):
        tpl_reader = PdfReader(tpl)
        cnt_reader = PdfReader(tmp_content)
        writer = PdfWriter()

        tpl_page = tpl_reader.pages[0]
        cnt_page = cnt_reader.pages[0]
        tpl_page.merge_page(cnt_page)
        writer.add_page(tpl_page)

        with open(output_path, "wb") as f:
            writer.write(f)
    else:
        shutil.copy(tmp_content, output_path)

    try:
        os.remove(tmp_content)
    except Exception:
        pass


# ─── Renderiza prévia exatamente igual ao PDF final ─────────────────────────
def render_preview(data: dict, template_frente: str) -> PILImage.Image | None:
    """
    Gera o PDF final em tmp, converte pág 0 em imagem PIL e retorna.
    A imagem é idêntica ao que será gerado.
    """
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.close()
    try:
        gerar_certificado_pdf(output_path=tmp.name, data=data,
                               template_frente=template_frente)
        if HAS_FITZ:
            doc = fitz.open(tmp.name)
            mat = fitz.Matrix(2.4, 2.4)
            pix = doc[0].get_pixmap(matrix=mat, alpha=False)
            img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            return img
        return None
    except Exception:
        return None
    finally:
        try:
            os.remove(tmp.name)
        except Exception:
            pass


# ─── Preview Engine (thread + debounce) ─────────────────────────────────────
class PreviewEngine:
    def __init__(self, on_update, debounce_ms: int = 650):
        self.on_update = on_update
        self.debounce_ms = debounce_ms
        self._lock = threading.Lock()
        self._timer = None
        self._params: dict = {}
        self._running = True

    def schedule(self, data: dict, template_frente: str):
        with self._lock:
            self._params = {"data": data, "template_frente": template_frente}
            if self._timer:
                self._timer.cancel()
            self._timer = threading.Timer(self.debounce_ms / 1000.0, self._run)
            self._timer.daemon = True
            self._timer.start()

    def _run(self):
        if not self._running:
            return
        with self._lock:
            params = dict(self._params)
        try:
            img = render_preview(**params)
            self.on_update(img, None)
        except Exception as e:
            self.on_update(None, str(e))

    def stop(self):
        self._running = False
        if self._timer:
            self._timer.cancel()


# ─── Painel de prévia ────────────────────────────────────────────────────────
class PreviewPanel(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._img_pil: PILImage.Image | None = None
        self._tk_img = None
        self._zoom = 1.0
        self._panel_w = 600

        # controles de zoom
        ctrl = ctk.CTkFrame(self, fg_color="transparent")
        ctrl.pack(fill="x", padx=8, pady=(6, 2))
        ctk.CTkButton(ctrl, text="−", width=30, command=self._zoom_out).pack(side="left", padx=1)
        self._lbl_zoom = ctk.CTkLabel(ctrl, text="100%", width=46)
        self._lbl_zoom.pack(side="left", padx=2)
        ctk.CTkButton(ctrl, text="+", width=30, command=self._zoom_in).pack(side="left", padx=1)
        self._lbl_status = ctk.CTkLabel(ctrl, text="", text_color="#6A8FAD",
                                         font=ctk.CTkFont(size=10))
        self._lbl_status.pack(side="left", padx=(12, 0))

        # canvas
        cont = ctk.CTkFrame(self, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self._canvas = tk.Canvas(cont, bg="#A8BCCF", highlightthickness=0)
        vsb = ctk.CTkScrollbar(cont, orientation="vertical",   command=self._canvas.yview)
        hsb = ctk.CTkScrollbar(cont, orientation="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right",  fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._canvas.bind("<MouseWheel>", self._on_mw)
        self._canvas.bind("<Button-4>",   self._on_mw)
        self._canvas.bind("<Button-5>",   self._on_mw)
        self._canvas.bind("<Button-1>",
                          lambda e: self._canvas.scan_mark(e.x, e.y))
        self._canvas.bind("<B1-Motion>",
                          lambda e: self._canvas.scan_dragto(e.x, e.y, gain=1))

        # placeholder inicial
        self._ph_id = self._canvas.create_text(
            300, 250, text="A prévia aparecerá aqui\nautomaticamente.",
            fill="#6A8FAD", font=("Helvetica", 13), justify="center"
        )

    # ── API pública ──────────────────────────────────────────────────────────
    def update_image(self, img: PILImage.Image):
        if self._ph_id:
            self._canvas.delete(self._ph_id)
            self._ph_id = None
        self._img_pil = img
        self._render()

    def show_status(self, msg: str):
        self._lbl_status.configure(text=msg)

    def set_panel_width(self, w: int):
        if abs(self._panel_w - w) > 15:
            self._panel_w = w
            if self._img_pil:
                self._render()

    # ── Internos ─────────────────────────────────────────────────────────────
    def _render(self):
        if not self._img_pil:
            return
        img = self._img_pil
        avail = max(self._panel_w - 20, 200)
        ratio = (avail / img.width) * self._zoom
        tw, th = int(img.width * ratio), int(img.height * ratio)
        resized = img.resize((tw, th), PILImage.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(resized)
        self._canvas.delete("img")
        self._canvas.create_image(0, 0, anchor="nw", image=self._tk_img, tags="img")
        self._canvas.config(scrollregion=(0, 0, tw, th))
        self._lbl_zoom.configure(text=f"{int(self._zoom * 100)}%")

    def _on_mw(self, event):
        if event.state & 0x0004:
            if event.delta > 0 or event.num == 4:
                self._zoom_in()
            else:
                self._zoom_out()
        else:
            d = -1 if (event.delta > 0 or event.num == 4) else 1
            self._canvas.yview_scroll(d, "units")

    def _zoom_in(self):
        self._zoom = min(4.0, self._zoom + 0.15)
        self._render()

    def _zoom_out(self):
        self._zoom = max(0.15, self._zoom - 0.15)
        self._render()


# ─── Aplicação principal ─────────────────────────────────────────────────────
class CertificadoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.title("Gerador de Certificados — Grupo BAW")
        self.geometry(self.cfg.get("window_geometry", "1560x900"))
        self.minsize(1200, 700)

        self._suspend = False
        self._last_pdf = ""

        self._build_ui()
        self._load_cfg_to_ui()

        self._engine = PreviewEngine(on_update=self._on_preview_ready, debounce_ms=600)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_resize)
        self.bind("<Control-g>", lambda e: self._gerar())
        self.bind("<Control-G>", lambda e: self._gerar())

        self.after(500, self._schedule_preview)

    # ═══════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        # ── Topbar ──────────────────────────────────────────────────────────
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(top, text="🎓  Gerador de Certificados — Grupo BAW",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color="#7EB3E8").pack(side="left")

        if HAS_REGISTRY:
            mods = {lbl: nm for nm, lbl in discover_modules()
                    if nm != "gerar_certificado.py"}
            if mods:
                self._mod_targets = mods
                ctk.CTkLabel(top, text="Ir para:").pack(side="left", padx=(22, 4))
                self._mod_var = tk.StringVar(value=list(mods.keys())[0])
                ctk.CTkOptionMenu(top, variable=self._mod_var,
                                  values=list(mods.keys()), width=210).pack(side="left")
                ctk.CTkButton(top, text="Abrir", width=68,
                              command=self._switch_module).pack(side="left", padx=4)

        self._btn_open = ctk.CTkButton(
            top, text="📂 Abrir PDF", width=120,
            command=self._abrir_pdf, fg_color="#2A6040", hover_color="#1E4A30")

        ctk.CTkButton(top, text="✔  Gerar PDF  (Ctrl+G)", width=175,
                      command=self._gerar,
                      fg_color="#1A4A7A", hover_color="#133A62").pack(side="right", padx=4)

        self._chk_auto = tk.BooleanVar(value=bool(self.cfg.get("preview_auto", True)))
        ctk.CTkCheckBox(top, text="Prévia automática",
                        variable=self._chk_auto,
                        command=self._schedule_preview).pack(side="right", padx=10)
        ctk.CTkButton(top, text="↺ Atualizar prévia", width=138,
                      command=self._schedule_preview).pack(side="right", padx=4)

        # ── Corpo ────────────────────────────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        # ── Formulário (esquerda, scroll) ─────────────────────────────────
        left = ctk.CTkScrollableFrame(
            body, width=500,
            fg_color=("#1A2835", "#1A2835"), corner_radius=10)
        left.pack(side="left", fill="y", padx=(0, 10))

        def section_label(txt):
            ctk.CTkFrame(left, height=1, fg_color="#2C4055").pack(
                fill="x", padx=8, pady=(12, 4))
            ctk.CTkLabel(left, text=txt,
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color="#4E88B8").pack(anchor="w", padx=10, pady=(0, 4))

        def field(parent, label, placeholder="", bold=False):
            ctk.CTkLabel(parent, text=label, text_color="#90AABB",
                         font=ctk.CTkFont(size=11)).pack(anchor="w", padx=10, pady=(4, 0))
            kw = {}
            if bold:
                kw["font"] = ctk.CTkFont(size=13, weight="bold")
                kw["border_color"] = "#4E88B8"
                kw["border_width"] = 2
            ent = ctk.CTkEntry(parent, placeholder_text=placeholder, **kw)
            ent.pack(fill="x", padx=10, pady=(2, 0))
            ent.bind("<KeyRelease>", self._on_field_change)
            return ent

        # Template
        section_label("TEMPLATE DO CERTIFICADO  (PDF de fundo)")
        trow = ctk.CTkFrame(left, fg_color="transparent")
        trow.pack(fill="x", padx=10, pady=(0, 4))
        self._lbl_tpl = ctk.CTkLabel(trow, text="Não selecionado",
                                      text_color="#4A6A80",
                                      font=ctk.CTkFont(size=10), anchor="w")
        self._lbl_tpl.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(trow, text="Selecionar", width=100,
                      command=self._pick_template).pack(side="left", padx=(6, 2))
        ctk.CTkButton(trow, text="Limpar", width=70,
                      command=self._clear_template).pack(side="left")

        # Participante
        section_label("PARTICIPANTE")
        self._ent_aprendiz = field(left, "★  Nome do funcionário / participante treinado:",
                                    "Nome completo", bold=True)

        # Empresa
        section_label("EMPRESA")
        self._ent_empresa = field(left, "Nome da empresa:", "A BAW Brasil Ind. e Com. Ltda")

        # Treinamento
        section_label("TREINAMENTO")
        self._ent_treinamento = field(left, "Nome do treinamento / curso:",
                                       "Ex.: Treinamento de Consumíveis e Processos...")
        self._ent_duracao     = field(left, "Duração:", "Ex.: 4 horas")
        self._ent_data        = field(left, "Data:", "Ex.: Março, 2023")
        self._ent_local       = field(left, "Local:", "Ex.: Caxias do Sul – RS")

        # Assinaturas lado a lado
        section_label("ASSINATURAS")
        sig_outer = ctk.CTkFrame(left, fg_color="transparent")
        sig_outer.pack(fill="x", padx=10, pady=(0, 4))
        sig_outer.columnconfigure(0, weight=1)
        sig_outer.columnconfigure(1, weight=1)

        def sig_block(parent, col, titulo, ph_nome, ph_cargo):
            g = ctk.CTkFrame(parent, fg_color="#122030", corner_radius=8)
            g.grid(row=0, column=col, sticky="nsew",
                   padx=(0, 4) if col == 0 else (4, 0))
            ctk.CTkLabel(g, text=titulo,
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color="#5E9DC8").pack(anchor="w", padx=8, pady=(8, 2))
            ctk.CTkLabel(g, text="Nome:", text_color="#90AABB",
                         font=ctk.CTkFont(size=10)).pack(anchor="w", padx=8)
            en = ctk.CTkEntry(g, placeholder_text=ph_nome)
            en.pack(fill="x", padx=8, pady=(0, 4))
            en.bind("<KeyRelease>", self._on_field_change)
            ctk.CTkLabel(g, text="Cargo:", text_color="#90AABB",
                         font=ctk.CTkFont(size=10)).pack(anchor="w", padx=8)
            ec = ctk.CTkEntry(g, placeholder_text=ph_cargo)
            ec.pack(fill="x", padx=8, pady=(0, 10))
            ec.bind("<KeyRelease>", self._on_field_change)
            return en, ec

        self._ent_sup_nome, self._ent_sup_cargo = sig_block(
            sig_outer, 0, "◀  Assinatura Esquerda",
            "Ex.: Renato Castelan", "Ex.: Supervisor técnico")
        self._ent_ins_nome, self._ent_ins_cargo = sig_block(
            sig_outer, 1, "▶  Assinatura Direita",
            "Ex.: Marco Silva", "Ex.: Instrutor técnico")

        # Tópicos
        section_label("TÓPICOS DO TREINAMENTO  (informativo / verso)")
        ctk.CTkLabel(left, text="Um tópico por linha.",
                     text_color="#4A6A80", font=ctk.CTkFont(size=9)).pack(
            anchor="w", padx=10)
        self._txt_topicos = ctk.CTkTextbox(
            left, height=110,
            font=ctk.CTkFont(family="Courier New", size=11), wrap="word")
        self._txt_topicos.pack(fill="x", padx=10, pady=(2, 4))
        self._txt_topicos.bind("<<Modified>>", self._on_tb_change)

        # ── Prévia (direita) ─────────────────────────────────────────────────
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side="right", fill="both", expand=True)

        hdr = ctk.CTkFrame(right, fg_color="transparent")
        hdr.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(hdr,
                     text="PRÉ-VISUALIZAÇÃO  —  idêntica ao PDF que será gerado",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#4E88B8").pack(side="left", padx=4)

        self._preview = PreviewPanel(
            right,
            fg_color=("#B8CCE0", "#111E2A"),
            corner_radius=10,
        )
        self._preview.pack(fill="both", expand=True)

        # ── Status bar ───────────────────────────────────────────────────────
        self._status = ctk.CTkLabel(self, text="Pronto.", anchor="w",
                                     font=ctk.CTkFont(size=11))
        self._status.pack(fill="x", padx=14, pady=(0, 8))

    # ═══════════════════════════════════════════════════════════════════════
    def _load_cfg_to_ui(self):
        self._suspend = True
        try:
            tpl = self.cfg.get("template_frente", "")
            self._lbl_tpl.configure(
                text=os.path.basename(tpl) if tpl and os.path.exists(tpl)
                else "Não selecionado"
            )
            pairs = [
                (self._ent_empresa,     "empresa"),
                (self._ent_treinamento, "nome_treinamento"),
                (self._ent_duracao,     "duracao"),
                (self._ent_data,        "data_certificado"),
                (self._ent_local,       "local"),
                (self._ent_sup_nome,    "supervisor_nome"),
                (self._ent_sup_cargo,   "supervisor_cargo"),
                (self._ent_ins_nome,    "instrutor_nome"),
                (self._ent_ins_cargo,   "instrutor_cargo"),
            ]
            for ent, key in pairs:
                ent.delete(0, "end")
                ent.insert(0, self.cfg.get(key, ""))

            self._txt_topicos.delete("1.0", "end")
            self._txt_topicos.insert("1.0", self.cfg.get("topicos", ""))
            self._txt_topicos.edit_modified(False)
        finally:
            self._suspend = False

    def _collect_data(self) -> dict:
        return {
            "empresa":          self._ent_empresa.get().strip(),
            "aprendiz":         self._ent_aprendiz.get().strip(),
            "nome_treinamento": self._ent_treinamento.get().strip(),
            "duracao":          self._ent_duracao.get().strip(),
            "data_certificado": self._ent_data.get().strip(),
            "local":            self._ent_local.get().strip(),
            "supervisor_nome":  self._ent_sup_nome.get().strip(),
            "supervisor_cargo": self._ent_sup_cargo.get().strip(),
            "instrutor_nome":   self._ent_ins_nome.get().strip(),
            "instrutor_cargo":  self._ent_ins_cargo.get().strip(),
            "topicos":          self._txt_topicos.get("1.0", "end").strip(),
        }

    def _sync_cfg(self):
        d = self._collect_data()
        for k in ["empresa", "nome_treinamento", "duracao", "data_certificado",
                  "local", "supervisor_nome", "supervisor_cargo",
                  "instrutor_nome", "instrutor_cargo", "topicos"]:
            self.cfg[k] = d.get(k, "")
        self.cfg["preview_auto"] = bool(self._chk_auto.get())
        save_config(self.cfg)

    # ─── Template ─────────────────────────────────────────────────────────
    def _pick_template(self):
        path = filedialog.askopenfilename(
            initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
            filetypes=[("PDF", "*.pdf"), ("Imagem", "*.png *.jpg *.jpeg")],
        )
        if not path:
            return
        self.cfg["template_frente"] = path
        self.cfg["last_dir"] = os.path.dirname(path)
        self._lbl_tpl.configure(text=os.path.basename(path))
        save_config(self.cfg)
        self._schedule_preview()

    def _clear_template(self):
        self.cfg["template_frente"] = ""
        self._lbl_tpl.configure(text="Não selecionado")
        save_config(self.cfg)
        self._schedule_preview()

    # ─── Prévia ───────────────────────────────────────────────────────────
    def _on_field_change(self, event=None):
        if not self._suspend and self._chk_auto.get():
            self._schedule_preview()

    def _on_tb_change(self, event=None):
        if event:
            event.widget.edit_modified(False)
        if not self._suspend and self._chk_auto.get():
            self._schedule_preview()

    def _schedule_preview(self):
        self._preview.show_status("⏳ atualizando…")
        self._status.configure(text="Gerando prévia…")
        self._engine.schedule(
            data=self._collect_data(),
            template_frente=self.cfg.get("template_frente", ""),
        )

    def _on_preview_ready(self, img, error):
        def _upd():
            if error:
                self._preview.show_status(f"⚠ {error[:110]}")
                self._status.configure(text="Erro na prévia.")
            elif img:
                self._preview.update_image(img)
                self._preview.show_status("")
                self._status.configure(text="Prévia atualizada.")
        self.after(0, _upd)

    def _on_resize(self, event=None):
        w = self._preview.winfo_width()
        if w > 50:
            self._preview.set_panel_width(w)

    # ─── Gerar ────────────────────────────────────────────────────────────
    def _gerar(self):
        aprendiz = self._ent_aprendiz.get().strip()
        if not aprendiz:
            messagebox.showerror("Erro", "Informe o nome do participante antes de gerar.")
            return

        save = filedialog.asksaveasfilename(
            initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"Certificado - {aprendiz}.pdf",
        )
        if not save:
            return
        self.cfg["last_dir"] = os.path.dirname(save)
        self._sync_cfg()
        self._status.configure(text="Gerando PDF…")

        data = self._collect_data()
        tpl  = self.cfg.get("template_frente", "")

        def _worker():
            try:
                gerar_certificado_pdf(output_path=save, data=data, template_frente=tpl)
                self.after(0, lambda: self._on_ok(save))
            except Exception as e:
                self.after(0, lambda err=str(e): self._on_err(err))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_ok(self, path):
        self._last_pdf = path
        self._status.configure(text=f"✔ PDF gerado: {os.path.basename(path)}")
        self._btn_open.pack(side="left", padx=8)
        if messagebox.askyesno("Sucesso",
                                f"Certificado gerado!\n\n{os.path.basename(path)}"
                                f"\n\nAbrir agora?"):
            self._abrir_pdf()

    def _on_err(self, err):
        self._status.configure(text="Erro ao gerar.")
        messagebox.showerror("Erro ao gerar certificado", err)

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

    # ─── Módulos ──────────────────────────────────────────────────────────
    def _switch_module(self):
        if not HAS_REGISTRY:
            return
        label = self._mod_var.get()
        fname = self._mod_targets.get(label)
        if not fname:
            return
        path = os.path.join(_DIR, fname)
        if not os.path.exists(path):
            messagebox.showerror("Erro", f"Módulo não encontrado: {fname}")
            return
        try:
            save_last_module(fname)
            subprocess.Popen([sys.executable, path])
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    # ─── Fechar ───────────────────────────────────────────────────────────
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
