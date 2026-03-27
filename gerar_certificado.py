"""
Gerador de Certificados — BAW Brasil
Versão refatorada: template como camada de fundo, prévia fiel ao PDF final.

Dependências:
    pip install reportlab pypdf pymupdf pillow customtkinter
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import json, os, re, threading, tempfile, subprocess, sys, uuid, io

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image as PILImage, ImageTk

# ─── Constantes ──────────────────────────────────────────────────────────────
PAGE_SIZE   = landscape(A4)
W, H        = PAGE_SIZE          # 841.89 x 595.28 pts
MARGIN      = 2.4 * cm

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_certificado.json")

# Paleta de cores do template BAW (azul escuro / laranja)
AZUL_ESCURO  = colors.HexColor("#0D1F5C")
AZUL_MEDIO   = colors.HexColor("#1A3C8F")
AZUL_CLARO   = colors.HexColor("#4A6FA5")
CINZA_TEXTO  = colors.HexColor("#2D2D2D")
CINZA_SUB    = colors.HexColor("#555555")

# ─── Config ──────────────────────────────────────────────────────────────────
DEFAULTS = {
    "template_frente": "",
    "template_verso":  "",
    "lista_nomes_path": "",
    "texto_corpo": "concluiu satisfatoriamente o:",
    "nome_curso": "",
    "supervisor": "",
    "supervisor_cargo": "Supervisor técnico",
    "instrutor": "",
    "instrutor_cargo": "Instrutor técnico",
    "carga_horaria": "",
    "mes_ano": "",
    "topicos": "",
    "last_dir": "",
    "window_geometry": "1600x900",
    "preview_auto": True,
}


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _sanitize(value: str) -> str:
    nome = re.sub(r'[\\/:*?"<>|]+', "_", (value or "").strip())
    return re.sub(r"\s+", " ", nome).strip() or "Participante"


# ─── Geração do PDF ───────────────────────────────────────────────────────────

def _draw_frente(c: rl_canvas.Canvas, params: dict):
    """Desenha a frente do certificado sobre a canvas."""
    aprendiz     = params.get("aprendiz", "")
    texto_corpo  = params.get("texto_corpo", "concluiu satisfatoriamente o:")
    nome_curso   = params.get("nome_curso", "")
    supervisor   = params.get("supervisor", "")
    sup_cargo    = params.get("supervisor_cargo", "Supervisor técnico")
    instrutor    = params.get("instrutor", "")
    inst_cargo   = params.get("instrutor_cargo", "Instrutor técnico")
    carga        = params.get("carga_horaria", "")
    mes_ano      = params.get("mes_ano", "")

    cx = W / 2  # centro horizontal

    # ── Cabeçalho: "A BAW Brasil Ind. e Com. Ltda certifica que:" ────────────
    y = H - MARGIN - 1.0 * cm
    c.setFont("Helvetica", 13)
    c.setFillColor(CINZA_TEXTO)
    cabecalho = "A BAW Brasil Ind. e Com. Ltda certifica que:"
    c.drawCentredString(cx, y, cabecalho)

    # ── Nome do aprendiz ─────────────────────────────────────────────────────
    y -= 1.6 * cm
    c.setFont("Helvetica-Bold", 38)
    c.setFillColor(AZUL_ESCURO)
    nome_display = (aprendiz or "[NOME DO PARTICIPANTE]").upper()
    # Auto-shrink se muito longo
    fs = 38
    while c.stringWidth(nome_display, "Helvetica-Bold", fs) > W - 4 * MARGIN and fs > 18:
        fs -= 1
    c.setFont("Helvetica-Bold", fs)
    c.drawCentredString(cx, y, nome_display)

    # ── Linha separadora sob o nome ───────────────────────────────────────────
    y -= 0.45 * cm
    lw = min(c.stringWidth(nome_display, "Helvetica-Bold", fs) + 1.5 * cm, W - 4 * MARGIN)
    c.setStrokeColor(AZUL_CLARO)
    c.setLineWidth(0.8)
    c.line(cx - lw/2, y, cx + lw/2, y)

    # ── Texto corpo ──────────────────────────────────────────────────────────
    y -= 0.8 * cm
    c.setFont("Helvetica", 12)
    c.setFillColor(CINZA_SUB)
    c.drawCentredString(cx, y, texto_corpo)

    # ── Nome do curso ────────────────────────────────────────────────────────
    if nome_curso and nome_curso.strip():
        y -= 0.55 * cm
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(AZUL_ESCURO)
        # Quebra linha se necessário
        palavras = nome_curso.split()
        linhas, linha_atual = [], ""
        for p in palavras:
            teste = (linha_atual + " " + p).strip()
            if c.stringWidth(teste, "Helvetica-Bold", 14) > W - 4 * MARGIN:
                linhas.append(linha_atual)
                linha_atual = p
            else:
                linha_atual = teste
        if linha_atual:
            linhas.append(linha_atual)
        for ln in linhas:
            y -= 0.5 * cm
            c.drawCentredString(cx, y, ln)

    # ── Duração e Mês/Ano ────────────────────────────────────────────────────
    y -= 1.0 * cm
    c.setFont("Helvetica", 11)
    c.setFillColor(CINZA_TEXTO)
    if carga:
        c.drawCentredString(cx, y, f"Duração: {carga}")
        y -= 0.55 * cm
    if mes_ano:
        c.drawCentredString(cx, y, mes_ano)
        y -= 0.55 * cm

    # ── Assinaturas ──────────────────────────────────────────────────────────
    y_assin = MARGIN + 1.8 * cm
    col_w   = (W - 2 * MARGIN) * 0.28
    x_left  = MARGIN + col_w * 0.5
    x_right = W - MARGIN - col_w * 0.5

    for x_col, nome_sig, cargo_sig in [
        (x_left,  supervisor, sup_cargo),
        (x_right, instrutor,  inst_cargo),
    ]:
        # Linha de assinatura
        c.setStrokeColor(AZUL_ESCURO)
        c.setLineWidth(0.8)
        c.line(x_col - col_w * 0.5, y_assin + 0.5 * cm,
               x_col + col_w * 0.5, y_assin + 0.5 * cm)
        c.setFont("Helvetica-Bold", 11)
        c.setFillColor(AZUL_ESCURO)
        c.drawCentredString(x_col, y_assin + 0.15 * cm, nome_sig or "")
        c.setFont("Helvetica", 10)
        c.setFillColor(CINZA_SUB)
        c.drawCentredString(x_col, y_assin - 0.35 * cm, cargo_sig or "")


def _draw_verso(c: rl_canvas.Canvas, params: dict):
    """Desenha o verso (conteúdo programático) sobre a canvas."""
    topicos  = params.get("topicos", "")
    aprendiz = params.get("aprendiz", "")
    instrutor = params.get("instrutor", "")
    carga    = params.get("carga_horaria", "")
    data     = params.get("mes_ano", "")

    cx = W / 2

    # ── Título ────────────────────────────────────────────────────────────────
    y = H - MARGIN - 1.2 * cm
    c.setFont("Helvetica-Bold", 20)
    c.setFillColor(AZUL_ESCURO)
    c.drawCentredString(cx, y, "Conteúdo Treinamento de Processos e Operação Baw")

    # Linha decorativa
    y -= 0.45 * cm
    c.setStrokeColor(AZUL_CLARO)
    c.setLineWidth(1.0)
    c.line(MARGIN + 1 * cm, y, W - MARGIN - 1 * cm, y)

    # ── Lista de tópicos ──────────────────────────────────────────────────────
    y -= 0.7 * cm
    c.setFont("Helvetica", 11)
    c.setFillColor(CINZA_TEXTO)

    linhas = [l.strip() for l in (topicos or "").strip().splitlines() if l.strip()]
    x_txt  = MARGIN + 1.5 * cm
    line_h = 0.52 * cm

    for ln in linhas:
        if y < MARGIN + 2 * cm:
            break
        # Detecta se é sub-item (começa com espaço ou "-" ou tab)
        if ln.startswith(("-", "–", "•")) or ln[:1] == " ":
            c.setFont("Helvetica", 10)
            c.setFillColor(CINZA_SUB)
            bullet = "▸"
            c.drawString(x_txt + 0.6 * cm, y, f"{bullet} {ln.lstrip('-–• ')}")
        else:
            c.setFont("Helvetica-Bold", 11)
            c.setFillColor(AZUL_ESCURO)
            c.drawString(x_txt, y, f"▶  {ln}")
            c.setFont("Helvetica", 11)
            c.setFillColor(CINZA_TEXTO)
        y -= line_h

    # ── Rodapé verso ─────────────────────────────────────────────────────────
    y_rod = MARGIN + 0.8 * cm
    partes = []
    if aprendiz: partes.append(f"Participante: {aprendiz}")
    if instrutor: partes.append(f"Instrutor: {instrutor}")
    if carga:    partes.append(f"Carga: {carga}")
    if data:     partes.append(f"Data: {data}")
    if partes:
        c.setFont("Helvetica", 9)
        c.setFillColor(CINZA_SUB)
        c.drawCentredString(cx, y_rod, "  ·  ".join(partes))


def _pick_template(template_path: str):
    """Retorna página PDF do template ou None."""
    if template_path and os.path.exists(template_path):
        try:
            r = PdfReader(template_path)
            if r.pages:
                return r.pages[0], r
        except Exception:
            pass
    return None, None


def gerar_certificado_pdf(output_path: str, params: dict,
                          template_frente: str = "",
                          template_verso: str = ""):
    """Gera PDF (frente + verso) sobrepondo conteúdo ao template."""
    tmp = tempfile.mkdtemp()
    tmp_f = os.path.join(tmp, "frente.pdf")
    tmp_v = os.path.join(tmp, "verso.pdf")

    # ── Frente ────────────────────────────────────────────────────────────────
    c = rl_canvas.Canvas(tmp_f, pagesize=PAGE_SIZE)
    _draw_frente(c, params)
    c.save()

    # ── Verso ─────────────────────────────────────────────────────────────────
    c = rl_canvas.Canvas(tmp_v, pagesize=PAGE_SIZE)
    _draw_verso(c, params)
    c.save()

    # ── Merge: template (fundo) + conteúdo (cima) ─────────────────────────────
    writer = PdfWriter()
    for content_pdf, tpl_path, fallback in [
        (tmp_f, template_frente, ""),
        (tmp_v, template_verso,  template_frente),
    ]:
        content_page = PdfReader(content_pdf).pages[0]
        tpl_page, _  = _pick_template(tpl_path)
        if tpl_page is None and fallback:
            tpl_page, _ = _pick_template(fallback)

        if tpl_page is not None:
            final = tpl_page.copy()
            final.merge_page(content_page)
            writer.add_page(final)
        else:
            writer.add_page(content_page)

    with open(output_path, "wb") as f:
        writer.write(f)

    # limpeza
    for p in [tmp_f, tmp_v]:
        try: os.remove(p)
        except: pass
    try: os.rmdir(tmp)
    except: pass


# ─── Preview Engine ──────────────────────────────────────────────────────────

class PreviewEngine:
    def __init__(self, on_update, debounce_ms=700):
        self.on_update   = on_update
        self.debounce_ms = debounce_ms
        self._lock   = threading.Lock()
        self._timer  = None
        self._params = {}
        self._tmpdir = tempfile.mkdtemp()
        self._alive  = True

    def schedule(self, params: dict, tpl_f: str, tpl_v: str):
        with self._lock:
            self._job = (dict(params), tpl_f, tpl_v)
            if self._timer:
                self._timer.cancel()
            self._timer = threading.Timer(self.debounce_ms / 1000.0, self._run)
            self._timer.daemon = True
            self._timer.start()

    def _run(self):
        if not self._alive:
            return
        with self._lock:
            params, tpl_f, tpl_v = self._job
        tmp = os.path.join(self._tmpdir, f"prev_{uuid.uuid4().hex[:6]}.pdf")
        try:
            gerar_certificado_pdf(tmp, params, tpl_f, tpl_v)
            doc = fitz.open(tmp)
            imgs = []
            mat  = fitz.Matrix(2.0, 2.0)
            for pg in doc:
                pix = pg.get_pixmap(matrix=mat, alpha=False)
                img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                imgs.append(img)
            doc.close()
            self.on_update(imgs, None)
        except Exception as e:
            self.on_update(None, str(e))
        finally:
            try: os.remove(tmp)
            except: pass

    def stop(self):
        self._alive = False
        if self._timer:
            self._timer.cancel()


# ─── Widget de Prévia ─────────────────────────────────────────────────────────

class PreviewPanel(ctk.CTkFrame):
    def __init__(self, master, label="", **kw):
        super().__init__(master, **kw)
        self._pages   = []
        self._tk_imgs = []
        self._idx     = 0
        self._zoom    = 1.0
        self._panel_w = 420

        # Header
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=6, pady=(6, 2))
        ctk.CTkLabel(hdr, text=label,
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color="#4A90D9").pack(side="left")

        # Nav
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.pack(pady=(2, 2))
        self._b_prev  = ctk.CTkButton(nav, text="◀", width=30, command=self._prev)
        self._lbl_pg  = ctk.CTkLabel(nav, text="", width=55)
        self._b_next  = ctk.CTkButton(nav, text="▶", width=30, command=self._next)
        self._b_zm    = ctk.CTkButton(nav, text="−", width=26, command=self._zoom_out)
        self._lbl_zm  = ctk.CTkLabel(nav, text="100%", width=42)
        self._b_zp    = ctk.CTkButton(nav, text="+", width=26, command=self._zoom_in)
        for w in [self._b_prev, self._lbl_pg, self._b_next,
                  ctk.CTkLabel(nav, text="|", text_color="#444"),
                  self._b_zm, self._lbl_zm, self._b_zp]:
            w.pack(side="left", padx=1)

        # Canvas
        cont = ctk.CTkFrame(self, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=4, pady=4)
        self._cv  = tk.Canvas(cont, bg="#C8D8E8", highlightthickness=0)
        vsb = ctk.CTkScrollbar(cont, orientation="vertical",   command=self._cv.yview)
        hsb = ctk.CTkScrollbar(cont, orientation="horizontal", command=self._cv.xview)
        self._cv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right",  fill="y")
        self._cv.pack(side="left", fill="both", expand=True)
        self._cv.bind("<MouseWheel>", self._scroll)
        self._cv.bind("<Button-4>",   self._scroll)
        self._cv.bind("<Button-5>",   self._scroll)

        self._status = ctk.CTkLabel(self, text="Prévia aparecerá aqui…",
                                     text_color="#7A9AB8")
        self._status.pack(expand=True)

    def _scroll(self, e):
        if e.state & 0x0004:
            self._zoom_in() if (e.delta > 0 or e.num == 4) else self._zoom_out()
        else:
            d = -1 if (e.delta > 0 or e.num == 4) else 1
            self._cv.yview_scroll(d, "units")

    def update_pages(self, images):
        self._status.pack_forget()
        self._pages = images
        if self._idx >= len(images):
            self._idx = 0
        self._render()

    def show_status(self, msg):
        self._pages = []
        self._cv.delete("all")
        self._status.configure(text=msg)
        self._status.pack(expand=True)

    def set_width(self, w):
        if abs(self._panel_w - w) > 15:
            self._panel_w = w
            if self._pages:
                self._render()

    def _render(self):
        if not self._pages:
            return
        img    = self._pages[self._idx]
        base_w = max(self._panel_w - 30, 200)
        ratio  = (base_w / img.width) * self._zoom
        tw, th = int(img.width * ratio), int(img.height * ratio)
        res    = img.resize((tw, th), PILImage.LANCZOS)
        tk_img = ImageTk.PhotoImage(res)
        self._tk_imgs = [tk_img]
        self._cv.delete("all")
        self._cv.create_image(0, 0, anchor="nw", image=tk_img)
        self._cv.config(scrollregion=(0, 0, tw, th))
        total = len(self._pages)
        self._lbl_pg.configure(text=f"{self._idx+1}/{total}")
        self._lbl_zm.configure(text=f"{int(self._zoom*100)}%")
        self._b_prev.configure(state="normal" if self._idx > 0       else "disabled")
        self._b_next.configure(state="normal" if self._idx < total-1 else "disabled")

    def _prev(self):
        if self._idx > 0: self._idx -= 1; self._render()
    def _next(self):
        if self._idx < len(self._pages)-1: self._idx += 1; self._render()
    def _zoom_in(self):  self._zoom = min(4.0, self._zoom+0.20); self._render()
    def _zoom_out(self): self._zoom = max(0.25, self._zoom-0.20); self._render()


# ─── App Principal ────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("🎓 Gerador de Certificados — BAW Brasil")
        self.cfg = load_config()
        geo = self.cfg.get("window_geometry", "1600x900")
        self.geometry(geo)
        self.minsize(1280, 760)

        self._lista_nomes: list[str] = []
        self._last_pdf = ""
        self._suspend  = False

        self._build_ui()
        self._load_ui()

        self._engine = PreviewEngine(on_update=self._on_preview, debounce_ms=600)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.bind("<Configure>",   self._on_resize)
        self.bind("<Control-g>",   lambda e: self._gerar())
        self.bind("<Control-G>",   lambda e: self._gerar())

        if self.cfg.get("preview_auto", True):
            self.after(500, self._atualizar_preview)

    # ──────────────────────────────────────────────────────────────────────────
    # UI
    # ──────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Barra superior ────────────────────────────────────────────────────
        bar = ctk.CTkFrame(self, fg_color=("#1A2A4A", "#0D1A30"), height=50)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        ctk.CTkLabel(bar, text="🎓  Gerador de Certificados — BAW Brasil",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color="#5BA4E8").pack(side="left", padx=16)

        right_bar = ctk.CTkFrame(bar, fg_color="transparent")
        right_bar.pack(side="right", padx=10)

        self._chk_auto = tk.BooleanVar(value=bool(self.cfg.get("preview_auto", True)))
        ctk.CTkCheckBox(right_bar, text="Prévia auto",
                        variable=self._chk_auto).pack(side="right", padx=6)
        ctk.CTkButton(right_bar, text="↺ Atualizar prévia", width=140,
                      command=self._atualizar_preview,
                      fg_color="#1A3A6A", hover_color="#0F2850").pack(side="right", padx=4)
        ctk.CTkButton(right_bar, text="✔ Gerar PDF  (Ctrl+G)", width=160,
                      command=self._gerar,
                      fg_color="#0D4A1A", hover_color="#083610").pack(side="right", padx=4)

        # ── Área principal ────────────────────────────────────────────────────
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=(6, 4))

        # ── Painel formulário (esquerda) ──────────────────────────────────────
        left = ctk.CTkScrollableFrame(main, width=500,
                                       fg_color=("#1C2C44", "#111C2E"),
                                       scrollbar_button_color="#2A4A7A")
        left.pack(side="left", fill="y", padx=(0, 8))

        def sec(txt):
            ctk.CTkLabel(left, text=txt,
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color="#5BA4E8").pack(anchor="w", padx=8, pady=(10, 2))
            ctk.CTkFrame(left, height=1, fg_color="#2A4A7A").pack(fill="x", padx=8)

        def entry(parent, label, placeholder=""):
            ctk.CTkLabel(parent, text=label, text_color="#8AAAC8",
                         font=ctk.CTkFont(size=11)).pack(anchor="w", padx=8, pady=(6, 1))
            e = ctk.CTkEntry(parent, placeholder_text=placeholder)
            e.pack(fill="x", padx=8, pady=(0, 2))
            e.bind("<KeyRelease>", self._field_changed)
            return e

        # Templates
        sec("TEMPLATES DE FUNDO  (PDF com borda/logo)")
        self._lbl_tpl_f = self._tpl_row(left, "Frente:", "frente")
        self._lbl_tpl_v = self._tpl_row(left, "Verso: ", "verso")

        # Participante
        sec("PARTICIPANTE")
        ctk.CTkLabel(left, text="Nome completo:", text_color="#8AAAC8",
                     font=ctk.CTkFont(size=11)).pack(anchor="w", padx=8, pady=(6, 1))
        self._ent_aprendiz = ctk.CTkEntry(left, font=ctk.CTkFont(size=14, weight="bold"),
                                           placeholder_text="Nome do participante")
        self._ent_aprendiz.pack(fill="x", padx=8, pady=(0, 4))
        self._ent_aprendiz.bind("<KeyRelease>", self._field_changed)

        # Lista
        self._lbl_lista = ctk.CTkLabel(left, text="Lista: não carregada",
                                        text_color="#5B7A9D",
                                        font=ctk.CTkFont(size=10))
        self._lbl_lista.pack(anchor="w", padx=8)
        row_lst = ctk.CTkFrame(left, fg_color="transparent")
        row_lst.pack(fill="x", padx=8, pady=4)
        ctk.CTkButton(row_lst, text="📋 Importar lista TXT", width=148,
                      command=self._importar_lista).pack(side="left")
        ctk.CTkButton(row_lst, text="Limpar", width=80,
                      command=self._limpar_lista,
                      fg_color="#3A1A1A", hover_color="#5A2A2A").pack(side="left", padx=6)

        # Curso
        sec("CURSO / TREINAMENTO")
        ctk.CTkLabel(left, text="Texto intro (ex: 'concluiu satisfatoriamente o:'):",
                     text_color="#8AAAC8", font=ctk.CTkFont(size=11)).pack(anchor="w", padx=8, pady=(6,1))
        self._ent_corpo = ctk.CTkEntry(left)
        self._ent_corpo.pack(fill="x", padx=8, pady=(0, 2))
        self._ent_corpo.bind("<KeyRelease>", self._field_changed)

        ctk.CTkLabel(left, text="Nome do curso/treinamento:",
                     text_color="#8AAAC8", font=ctk.CTkFont(size=11)).pack(anchor="w", padx=8, pady=(4,1))
        self._ent_curso = ctk.CTkTextbox(left, height=60,
                                          font=ctk.CTkFont(family="Courier New", size=11),
                                          wrap="word")
        self._ent_curso.pack(fill="x", padx=8, pady=(0, 2))
        self._ent_curso.bind("<<Modified>>", self._textbox_changed)

        # Informações
        sec("INFORMAÇÕES DO CERTIFICADO  (salvas automaticamente)")
        self._ent_carga   = entry(left, "Carga horária:",    "Ex: 4 horas")
        self._ent_mes_ano = entry(left, "Mês / Ano:",        "Ex: Março, 2025")

        # Assinaturas
        sec("ASSINATURAS")
        self._ent_supervisor      = entry(left, "Supervisor — Nome:",  "Ex: Renato Castelan")
        self._ent_supervisor_cargo = entry(left, "Supervisor — Cargo:", "Supervisor técnico")
        self._ent_instrutor        = entry(left, "Instrutor — Nome:",   "Ex: Marco Silva")
        self._ent_instrutor_cargo  = entry(left, "Instrutor — Cargo:",  "Instrutor técnico")

        # Tópicos verso
        sec("TÓPICOS — VERSO DO CERTIFICADO")
        ctk.CTkLabel(left,
                     text="Uma linha por tópico. Linhas com '-' ou espaço serão sub-itens.",
                     text_color="#5B7A9D", font=ctk.CTkFont(size=10)).pack(anchor="w", padx=8)
        self._txt_topicos = ctk.CTkTextbox(left, height=200,
                                            font=ctk.CTkFont(family="Courier New", size=11),
                                            wrap="word")
        self._txt_topicos.pack(fill="x", padx=8, pady=(4, 8))
        self._txt_topicos.bind("<<Modified>>", self._textbox_changed)

        # ── Painel prévia (direita) ───────────────────────────────────────────
        right = ctk.CTkFrame(main, fg_color="transparent")
        right.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(right, text="PRÉ-VISUALIZAÇÃO  (fiel ao PDF final com template)",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#5BA4E8").pack(anchor="w", padx=4, pady=(4, 2))

        pv_frame = ctk.CTkFrame(right, fg_color="transparent")
        pv_frame.pack(fill="both", expand=True)
        pv_frame.columnconfigure(0, weight=1)
        pv_frame.columnconfigure(1, weight=1)
        pv_frame.rowconfigure(0, weight=1)

        self._pv_f = PreviewPanel(pv_frame, label="✦ FRENTE",
                                   fg_color=("#C8D8E8", "#15253A"),
                                   corner_radius=8)
        self._pv_f.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        self._pv_v = PreviewPanel(pv_frame, label="✦ VERSO",
                                   fg_color=("#C8D8E8", "#15253A"),
                                   corner_radius=8)
        self._pv_v.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        # ── Status bar ───────────────────────────────────────────────────────
        self._status = ctk.CTkLabel(self, text="Pronto.", anchor="w",
                                     text_color="#6A8AAD")
        self._status.pack(fill="x", padx=14, pady=(0, 6))

    def _tpl_row(self, parent, label, key):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=3)
        ctk.CTkLabel(row, text=label, width=52, text_color="#8AAAC8").pack(side="left")
        lbl = ctk.CTkLabel(row, text="Não selecionado",
                            text_color="#4A7A9D",
                            font=ctk.CTkFont(size=10), anchor="w")
        lbl.pack(side="left", fill="x", expand=True, padx=(4, 8))
        ctk.CTkButton(row, text="Selecionar", width=90,
                      command=lambda: self._pick_tpl(key, lbl)).pack(side="left", padx=2)
        ctk.CTkButton(row, text="Limpar", width=66,
                      command=lambda: self._clear_tpl(key, lbl),
                      fg_color="#3A1A1A", hover_color="#5A2A2A").pack(side="left")
        return lbl

    # ──────────────────────────────────────────────────────────────────────────
    # Carregar / coletar dados
    # ──────────────────────────────────────────────────────────────────────────
    def _load_ui(self):
        self._suspend = True
        for key, lbl in [("template_frente", self._lbl_tpl_f),
                          ("template_verso",  self._lbl_tpl_v)]:
            p = self.cfg.get(key, "")
            lbl.configure(text=os.path.basename(p) if p and os.path.exists(p) else "Não selecionado")

        pairs = [
            (self._ent_corpo,           "texto_corpo"),
            (self._ent_carga,           "carga_horaria"),
            (self._ent_mes_ano,         "mes_ano"),
            (self._ent_supervisor,      "supervisor"),
            (self._ent_supervisor_cargo,"supervisor_cargo"),
            (self._ent_instrutor,       "instrutor"),
            (self._ent_instrutor_cargo, "instrutor_cargo"),
        ]
        for ent, key in pairs:
            ent.delete(0, "end")
            ent.insert(0, self.cfg.get(key, ""))

        self._ent_curso.delete("1.0", "end")
        self._ent_curso.insert("1.0", self.cfg.get("nome_curso", ""))
        self._ent_curso.edit_modified(False)

        self._txt_topicos.delete("1.0", "end")
        self._txt_topicos.insert("1.0", self.cfg.get("topicos", ""))
        self._txt_topicos.edit_modified(False)

        lista_path = self.cfg.get("lista_nomes_path", "")
        if lista_path and os.path.exists(lista_path):
            self._load_lista_file(lista_path, preview=False)
        self._update_lista_label()

        self._suspend = False

    def _collect(self) -> dict:
        aprendiz = self._ent_aprendiz.get().strip()
        if self._lista_nomes:
            aprendiz = self._lista_nomes[0]
        return dict(
            aprendiz          = aprendiz,
            texto_corpo       = self._ent_corpo.get().strip(),
            nome_curso        = self._ent_curso.get("1.0", "end").strip(),
            carga_horaria     = self._ent_carga.get().strip(),
            mes_ano           = self._ent_mes_ano.get().strip(),
            supervisor        = self._ent_supervisor.get().strip(),
            supervisor_cargo  = self._ent_supervisor_cargo.get().strip(),
            instrutor         = self._ent_instrutor.get().strip(),
            instrutor_cargo   = self._ent_instrutor_cargo.get().strip(),
            topicos           = self._txt_topicos.get("1.0", "end").strip(),
        )

    def _sync_cfg(self):
        c = self._collect()
        for k in ["texto_corpo","nome_curso","carga_horaria","mes_ano",
                  "supervisor","supervisor_cargo","instrutor","instrutor_cargo","topicos"]:
            self.cfg[k] = c.get(k, "")
        self.cfg["preview_auto"] = bool(self._chk_auto.get())
        save_config(self.cfg)

    # ──────────────────────────────────────────────────────────────────────────
    # Templates
    # ──────────────────────────────────────────────────────────────────────────
    def _pick_tpl(self, key, lbl):
        path = filedialog.askopenfilename(
            initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
            filetypes=[("PDF", "*.pdf *.png *.jpg *.jpeg")])
        if not path:
            return
        self.cfg[f"template_{key}"] = path
        self.cfg["last_dir"] = os.path.dirname(path)
        lbl.configure(text=os.path.basename(path))
        save_config(self.cfg)
        self._atualizar_preview()

    def _clear_tpl(self, key, lbl):
        self.cfg[f"template_{key}"] = ""
        lbl.configure(text="Não selecionado")
        save_config(self.cfg)
        self._atualizar_preview()

    # ──────────────────────────────────────────────────────────────────────────
    # Preview
    # ──────────────────────────────────────────────────────────────────────────
    def _field_changed(self, event=None):
        if self._suspend or not self._chk_auto.get():
            return
        self._atualizar_preview()

    def _textbox_changed(self, event=None):
        if event:
            event.widget.edit_modified(False)
        if self._suspend or not self._chk_auto.get():
            return
        self._atualizar_preview()

    def _atualizar_preview(self):
        params = self._collect()
        self._pv_f.show_status("⏳ Renderizando…")
        self._pv_v.show_status("⏳ Renderizando…")
        self._status.configure(text="Gerando prévia…")
        self._engine.schedule(
            params,
            self.cfg.get("template_frente", ""),
            self.cfg.get("template_verso",  ""),
        )

    def _on_preview(self, images, error):
        def _upd():
            if error:
                self._pv_f.show_status(f"⚠ {error[:120]}")
                self._pv_v.show_status(f"⚠ {error[:120]}")
                self._status.configure(text="Erro na prévia.")
            elif images:
                if len(images) >= 1: self._pv_f.update_pages([images[0]])
                if len(images) >= 2: self._pv_v.update_pages([images[1]])
                self._status.configure(text="✔ Prévia atualizada.")
        self.after(0, _upd)

    def _on_resize(self, event=None):
        for pv in [self._pv_f, self._pv_v]:
            w = pv.winfo_width()
            if w > 50:
                pv.set_width(w)

    # ──────────────────────────────────────────────────────────────────────────
    # Geração
    # ──────────────────────────────────────────────────────────────────────────
    def _gerar(self):
        aprendiz_manual = self._ent_aprendiz.get().strip()
        nomes = list(self._lista_nomes) if self._lista_nomes else []
        if not nomes and not aprendiz_manual:
            messagebox.showerror("Erro", "Informe o nome do participante ou importe uma lista.")
            return
        if not nomes:
            nomes = [aprendiz_manual]

        if len(nomes) == 1:
            save = filedialog.asksaveasfilename(
                initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"Certificado - {nomes[0]}.pdf",
            )
            if not save:
                return
            jobs = [(nomes[0], save)]
            self.cfg["last_dir"] = os.path.dirname(save)
        else:
            out_dir = filedialog.askdirectory(
                initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
                title="Pasta de saída para os certificados",
            )
            if not out_dir:
                return
            jobs = [(n, os.path.join(out_dir, f"Certificado - {_sanitize(n)}.pdf")) for n in nomes]
            self.cfg["last_dir"] = out_dir

        self._sync_cfg()
        params_base = self._collect()
        tpl_f = self.cfg.get("template_frente", "")
        tpl_v = self.cfg.get("template_verso",  "")

        self._status.configure(
            text=f"Gerando {'PDF' if len(jobs)==1 else str(len(jobs))+' PDFs'}…")

        def _worker():
            try:
                ultimo = ""
                for nome, path in jobs:
                    p = dict(params_base); p["aprendiz"] = nome
                    gerar_certificado_pdf(path, p, tpl_f, tpl_v)
                    ultimo = path
                self.after(0, lambda u=ultimo, t=len(jobs): self._ok(u, t))
            except Exception as e:
                self.after(0, lambda err=str(e): self._err(err))

        threading.Thread(target=_worker, daemon=True).start()

    def _ok(self, path, total):
        self._last_pdf = path
        txt = (f"✔ PDF gerado: {os.path.basename(path)}"
               if total == 1 else f"✔ {total} PDFs gerados.")
        self._status.configure(text=txt)
        msg = (f"Certificado gerado!\n{os.path.basename(path)}\n\nAbrir agora?"
               if total == 1
               else f"{total} certificados gerados.\nÚltimo: {os.path.basename(path)}\n\nAbrir último?")
        if messagebox.askyesno("Sucesso", msg):
            self._abrir(path)

    def _err(self, err):
        self._status.configure(text="Erro ao gerar certificado.")
        messagebox.showerror("Erro", err)

    def _abrir(self, path):
        if path and os.path.exists(path):
            try:
                if sys.platform.startswith("win"):   os.startfile(path)
                elif sys.platform.startswith("darwin"): subprocess.run(["open", path])
                else: subprocess.run(["xdg-open", path])
            except Exception: pass

    # ──────────────────────────────────────────────────────────────────────────
    # Lista de nomes
    # ──────────────────────────────────────────────────────────────────────────
    def _update_lista_label(self):
        if self._lista_nomes:
            origem = os.path.basename(self.cfg.get("lista_nomes_path", ""))
            self._lbl_lista.configure(
                text=f"Lista: {len(self._lista_nomes)} nomes  ({origem})")
        else:
            self._lbl_lista.configure(text="Lista: não carregada")

    def _load_lista_file(self, path: str, preview=True):
        with open(path, "r", encoding="utf-8") as f:
            nomes = [l.strip() for l in f if l.strip()]
        if not nomes:
            raise ValueError("Arquivo TXT sem nomes válidos.")
        self._lista_nomes = nomes
        self.cfg["lista_nomes_path"] = path
        self.cfg["last_dir"] = os.path.dirname(path)
        self._update_lista_label()
        save_config(self.cfg)
        if preview:
            self._atualizar_preview()

    def _importar_lista(self):
        path = filedialog.askopenfilename(
            initialdir=self.cfg.get("last_dir") or os.path.expanduser("~"),
            filetypes=[("Texto", "*.txt")])
        if not path:
            return
        try:
            self._load_lista_file(path)
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _limpar_lista(self):
        self._lista_nomes = []
        self.cfg["lista_nomes_path"] = ""
        self._update_lista_label()
        save_config(self.cfg)
        if self._chk_auto.get():
            self._atualizar_preview()

    # ──────────────────────────────────────────────────────────────────────────
    # Fechar
    # ──────────────────────────────────────────────────────────────────────────
    def _close(self):
        self.cfg["window_geometry"] = self.geometry()
        self._sync_cfg()
        self._engine.stop()
        self.destroy()


# ─── Entry point ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
