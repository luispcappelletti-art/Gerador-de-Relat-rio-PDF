import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import json, os, re, unicodedata, threading, tempfile

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import *
from reportlab.lib.styles import *
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

from pypdf import PdfReader, PdfWriter
import fitz  # pymupdf — instale com: pip install pymupdf
from PIL import Image as PILImage, ImageTk

CONFIG_FILE = "config.json"


# =========================
# CONFIG
# =========================
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"template_path": ""}


def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)


# =========================
# PARSER
# =========================
def normalize(text):
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c)).lower().strip()


def parse_text(text):
    sections = {"info": {}}

    aliases = {
        "equipamento": "equipamento",
        "fonte": "fonte",
        "cliente": "cliente",
        "cnc": "cnc",
        "thc": "thc",
        "fabricante": "fabricante",
        "contato cliente": "contato_cliente",
        "data": "data",
        "tecnico": "tecnico",
        "técnico": "tecnico",
        "acompanhamento remoto": "acompanhamento",
        "horario de inicio": "inicio",
        "horario de termino": "fim",
        "tempo do atendimento": "tempo_atendimento",
        "tempo em espera": "tempo_espera",
    }

    for line in text.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            key = normalize(k)
            if key in aliases:
                sections["info"][aliases[key]] = v.strip()

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
def gerar_pdf(sections, template_path, output_path, fotos=None):
    import uuid
    temp_pdf = output_path + f".tmp_{uuid.uuid4().hex[:8]}.pdf"
    fotos = fotos or []

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(name="Titulo",
                              fontSize=18, alignment=1, spaceAfter=6,
                              textColor=colors.HexColor("#0E2A44"), leading=22))

    styles.add(ParagraphStyle(name="Secao",
                              fontSize=12.5, spaceBefore=14, spaceAfter=7,
                              textColor=colors.HexColor("#123A5A"), leading=15))

    styles.add(ParagraphStyle(name="Body",
                              fontSize=10.5, leading=15, textColor=colors.HexColor("#1F2B37")))

    story = []

    info = sections.get("info", {})
    dados = []

    campos = [
        ("Equipamento", "equipamento"),
        ("Fonte", "fonte"),
        ("Cliente", "cliente"),
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

    for label, key in campos:
        valor = info.get(key)
        if _valor_info_preenchido(valor):
            dados.append([label, str(valor).strip()])

    if dados:
        tabela = Table(dados, colWidths=[5 * cm, 10 * cm])
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FCFDFE")),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#C7D4DF")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DCE5EC")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E9F0F6")),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(tabela)
        story.append(Spacer(1, 15))

    def add_secao(titulo, conteudo):
        story.append(Paragraph(f"<b>{titulo}</b>", styles["Secao"]))
        story.extend(processar_lista(conteudo, styles))

    if sections.get("descricao"):
        add_secao("1 – ESCOPO DO ATENDIMENTO", sections["descricao"])
    if sections.get("detalhamento"):
        add_secao("2 – DETALHAMENTO DO PROBLEMA", sections["detalhamento"])
    if sections.get("diagnostico"):
        add_secao("3 – DIAGNÓSTICO", sections["diagnostico"])
    if sections.get("acoes"):
        add_secao("4 – AÇÕES CORRETIVAS", sections["acoes"])
    if sections.get("resultado"):
        add_secao("5 – RESULTADO", sections["resultado"])
    if sections.get("estado"):
        add_secao("6 – ESTADO FINAL", sections["estado"])

    if fotos:
        story.append(PageBreak())
        story.append(Paragraph("<b>ANEXOS FOTOGRÁFICOS</b>", styles["Secao"]))
        story.append(Spacer(1, 6))

        largura_util = A4[0] - (5 * cm)
        largura_max = largura_util - (1 * cm)
        altura_max = 8.1 * cm

        for bloco in range(0, len(fotos), 2):
            fotos_bloco = fotos[bloco:bloco + 2]
            if bloco > 0:
                story.append(PageBreak())

            for posicao, foto in enumerate(fotos_bloco, start=1):
                idx = bloco + posicao
                caminho = foto.get("path")
                titulo = foto.get("title") or f"Foto {idx}"

                story.append(Paragraph(f"<b>{idx}. {titulo}</b>", styles["Body"]))
                story.append(Spacer(1, 4))

                try:
                    img_reader = ImageReader(caminho)
                    largura_original, altura_original = img_reader.getSize()
                    escala = min(largura_max / largura_original, altura_max / altura_original)
                    largura = largura_original * escala
                    altura = altura_original * escala

                    imagem = Image(caminho, width=largura, height=altura)
                    imagem.hAlign = "CENTER"
                    story.append(imagem)
                except Exception:
                    story.append(Paragraph("Não foi possível carregar esta imagem.", styles["Body"]))

                if posicao != len(fotos_bloco):
                    story.append(Spacer(1, 14))

    def page_chrome(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#F5F8FB"))
        canvas.rect(2.3 * cm, A4[1] - 2.65 * cm, A4[0] - (4.6 * cm), 1.2 * cm, stroke=0, fill=1)
        canvas.setStrokeColor(colors.HexColor("#D1DCE6"))
        canvas.rect(2.3 * cm, A4[1] - 2.65 * cm, A4[0] - (4.6 * cm), 1.2 * cm, stroke=1, fill=0)
        canvas.setFont("Helvetica-Bold", 12)
        canvas.setFillColor(colors.HexColor("#0E2A44"))
        canvas.drawCentredString(A4[0] / 2, A4[1] - 1.95 * cm, "RELATÓRIO TÉCNICO DE ATENDIMENTO")

        canvas.setFont("Helvetica", 8.8)
        canvas.setFillColor(colors.HexColor("#5B6E7D"))
        canvas.drawString(2.5 * cm, 1.4 * cm, "Relatório técnico")
        canvas.drawRightString(18.5 * cm, 1.4 * cm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(temp_pdf, pagesize=A4,
                            leftMargin=2.5 * cm, rightMargin=2.5 * cm,
                            topMargin=3.4 * cm, bottomMargin=2.5 * cm)

    doc.build(story, onFirstPage=page_chrome, onLaterPages=page_chrome)

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
        self._temp_dir = tempfile.mkdtemp()
        self._temp_pdf = os.path.join(self._temp_dir, "preview.pdf")
        self._running = True

    def schedule_update(self, text):
        with self._lock:
            self._current_text = text
            if self._timer is not None:
                self._timer.cancel()
            self._timer = threading.Timer(
                self.debounce_ms / 1000.0, self._generate
            )
            self._timer.daemon = True
            self._timer.start()

    def _generate(self):
        if not self._running:
            return
        with self._lock:
            text = self._current_text

        import uuid
        temp_pdf = os.path.join(self._temp_dir, f"preview_{uuid.uuid4().hex[:8]}.pdf")
        try:
            texto = limpar_texto(text)
            sections = parse_text(texto)
            gerar_pdf(sections, "", temp_pdf)

            doc = fitz.open(temp_pdf)
            images = []
            # Matriz aumentada para 2.0 para garantir nitidez ao fazer zoom
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


# =========================
# PREVIEW PANEL
# =========================
class PreviewPanel(ctk.CTkFrame):
    """Painel com suporte a Zoom, Scroll Vertical/Horizontal e Pan (arrastar)."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self._label_status = ctk.CTkLabel(
            self, text="A pré-visualização aparecerá aqui...",
            text_color="#8A9BAD", font=ctk.CTkFont(size=12)
        )
        self._label_status.pack(expand=True)

        # Barra de Navegação e Zoom
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

        # Área do Canvas
        self._canvas_container = ctk.CTkFrame(self, fg_color="transparent")

        self._canvas = tk.Canvas(self._canvas_container, bg="#E8EEF3", highlightthickness=0)
        self._vsb = ctk.CTkScrollbar(self._canvas_container, orientation="vertical", command=self._canvas.yview)
        self._hsb = ctk.CTkScrollbar(self._canvas_container, orientation="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=self._vsb.set, xscrollcommand=self._hsb.set)

        self._hsb.pack(side="bottom", fill="x", padx=(0, 16))
        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        # Configuração de Interação (Scroll, Pan e Zoom via Teclado/Mouse)
        self._canvas.bind("<MouseWheel>", self._on_mousewheel)  # Windows
        self._canvas.bind("<Button-4>", self._on_mousewheel)  # Linux scroll up
        self._canvas.bind("<Button-5>", self._on_mousewheel)  # Linux scroll down

        # Pan (Arrastar com botão esquerdo)
        self._canvas.bind("<Button-1>", self._on_pan_start)
        self._canvas.bind("<B1-Motion>", self._on_pan_move)

        self._pages = []
        self._tk_images = []
        self._current_page = 0
        self._panel_width = 380
        self._zoom_factor = 1.0

    def _on_mousewheel(self, event):
        # Verifica se Ctrl está pressionado para Zoom
        if event.state & 0x0004:  # Control mask
            if event.delta > 0 or event.num == 4:
                self._zoom_in()
            else:
                self._zoom_out()
            return "break"

        # Verifica se Shift está pressionado para Scroll Horizontal
        if event.state & 0x0001:  # Shift mask
            if event.delta:
                self._canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self._canvas.xview_scroll(-1, "units")
            elif event.num == 5:
                self._canvas.xview_scroll(1, "units")
        else:
            # Scroll Vertical padrão
            if event.delta:
                self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self._canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self._canvas.yview_scroll(1, "units")

    def _on_pan_start(self, event):
        """Prepara o canvas para arrastar."""
        self._canvas.scan_mark(event.x, event.y)

    def _on_pan_move(self, event):
        """Arrasta o conteúdo do canvas conforme o movimento do mouse."""
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
        if self._zoom_factor < 4.0:
            self._zoom_factor += 0.25
            self._show_page()

    def _zoom_out(self):
        if self._zoom_factor > 0.4:
            self._zoom_factor -= 0.25
            self._show_page()

    def set_width(self, w):
        if abs(self._panel_width - w) > 15:
            self._panel_width = w
            if self._pages:
                self._show_page()


# =========================
# UI PRINCIPAL
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Relatórios")
        self.geometry("1280x750")
        self.minsize(900, 500)

        self.config_data = load_config()

        # ── Barra Superior ────────────────────────────
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(10, 4))

        self.label = ctk.CTkLabel(top, text="Template: não selecionado", font=ctk.CTkFont(size=12))
        self.label.pack(side="left")

        ctk.CTkButton(top, text="Selecionar Template", width=160,
                      command=self.select_template).pack(side="left", padx=10)

        ctk.CTkButton(top, text="Gerar PDF", width=120,
                      command=self.generate).pack(side="right", padx=4)

        ctk.CTkButton(top, text="Limpar", width=90,
                      command=self._limpar).pack(side="right", padx=4)

        self._preview_visible = True
        self._toggle_btn = ctk.CTkButton(
            top, text="◀ Ocultar prévia", width=130, command=self._toggle_preview
        )
        self._toggle_btn.pack(side="right", padx=8)

        # ── Divisão Principal ──────────────────────────
        self._main = ctk.CTkFrame(self, fg_color="transparent")
        self._main.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        # Esquerda: Área de texto
        left = ctk.CTkFrame(self._main, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(left, text="Cole ou digite o texto do relatório:",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(0, 4))

        self.text = ctk.CTkTextbox(left, wrap="word",
                                   font=ctk.CTkFont(family="Courier New", size=12))
        self.text.pack(fill="both", expand=True)
        self.text.bind("<<Modified>>", self._on_text_change)

        # Direita: Pré-visualização
        self._preview_panel = PreviewPanel(
            self._main,
            width=420,
            fg_color=("#E8EEF3", "#1E2C38"),
            corner_radius=8
        )
        self._preview_panel.pack(side="right", fill="both", padx=(10, 0))

        # Motor de pré-visualização
        self._engine = PreviewEngine(
            on_update=self._on_preview_ready,
            debounce_ms=900
        )

        self.update_label()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Configure>", self._on_resize)

    def _limpar(self):
        self.text.delete("1.0", "end")
        self._preview_panel.show_status("A pré-visualização aparecerá aqui...")

    def _on_text_change(self, event=None):
        self.text.edit_modified(False)
        texto = self.text.get("1.0", "end").strip()
        if not texto:
            self._preview_panel.show_status("A pré-visualização aparecerá aqui...")
            return
        self._preview_panel.show_generating()
        self._engine.schedule_update(texto)

    def _on_preview_ready(self, images, error):
        def _update():
            if error:
                self._preview_panel.show_status(f"⚠ {error[:80]}")
            elif images:
                self._preview_panel.update_pages(images)

        self.after(0, _update)

    def _toggle_preview(self):
        if self._preview_visible:
            self._preview_panel.pack_forget()
            self._toggle_btn.configure(text="▶ Mostrar prévia")
            self._preview_visible = False
        else:
            self._preview_panel.pack(side="right", fill="both", padx=(10, 0))
            self._toggle_btn.configure(text="◀ Ocultar prévia")
            self._preview_visible = True

    def _on_resize(self, event=None):
        w = self._preview_panel.winfo_width()
        if w > 50:
            self._preview_panel.set_width(w)

    def _on_close(self):
        self._engine.stop()
        self.destroy()

    def update_label(self):
        path = self.config_data.get("template_path", "")
        self.label.configure(text=f"Template: {path}" if path else "Template: não selecionado")

    def select_template(self):
        file = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if file:
            self.config_data["template_path"] = file
            save_config(self.config_data)
            self.update_label()

    def generate(self):
        texto = self.text.get("1.0", "end").strip()
        if not texto:
            messagebox.showerror("Erro", "Cole o texto primeiro")
            return

        texto = limpar_texto(texto)
        sections = parse_text(texto)

        save = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not save:
            return

        fotos = []
        if messagebox.askyesno("Anexar fotos", "Deseja anexar fotos no PDF?"):
            arquivos = filedialog.askopenfilenames(
                title="Selecione as fotos",
                filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")],
            )
            if arquivos:
                usar_titulos = messagebox.askyesno("Títulos", "Definir título para cada foto?")
                for idx, arquivo in enumerate(arquivos, start=1):
                    titulo = ""
                    if usar_titulos:
                        titulo = simpledialog.askstring("Título", f"Foto {idx} ({os.path.basename(arquivo)}):") or ""
                    fotos.append({"path": arquivo, "title": titulo.strip()})

        try:
            gerar_pdf(sections, self.config_data.get("template_path", ""), save, fotos=fotos)
            messagebox.showinfo("Sucesso", "PDF gerado!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
