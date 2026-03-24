import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import json, os, re, unicodedata

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import *
from reportlab.lib.styles import *
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader

from pypdf import PdfReader, PdfWriter

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
        "cliente": "cliente",
        "cnc": "cnc",
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
        content = parts[i+1].strip()

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


# =========================
# LISTAS AUTOMÁTICAS
# =========================
def processar_lista(texto, styles):
    elementos = []
    linhas = texto.split("\n")

    lista = []
    item_atual = ""

    def adicionar_item():
        if item_atual.strip():
            lista.append(
                ListItem(
                    Paragraph(item_atual.strip(), styles["Body"]),
                    leftIndent=12,
                    bulletText="°"
                )
            )

    for linha in linhas:
        linha = linha.strip()

        # NOVO ITEM
        if re.match(r"^(\d+[\.\)]|-|•|\u2022)", linha):
            if item_atual:
                adicionar_item()
                item_atual = ""

            linha_limpa = re.sub(r"^(\d+[\.\)]|-|•|\u2022)\s*", "", linha)
            item_atual = linha_limpa

        else:
            # CONTINUAÇÃO DO ITEM
            if item_atual:
                item_atual += " " + linha
            else:
                # TEXTO NORMAL (fora de lista)
                if linha:
                    elementos.append(Paragraph(linha, styles["Body"]))

    # último item
    if item_atual:
        adicionar_item()

    if lista:
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
    temp_pdf = "temp.pdf"
    fotos = fotos or []

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(name="Titulo",
        fontSize=18, alignment=1, spaceAfter=6, textColor=colors.HexColor("#0E2A44"), leading=22))

    styles.add(ParagraphStyle(name="Secao",
        fontSize=12.5, spaceBefore=14, spaceAfter=7, textColor=colors.HexColor("#123A5A"), leading=15))

    styles.add(ParagraphStyle(name="Body",
        fontSize=10.5, leading=15, textColor=colors.HexColor("#1F2B37")))

    story = []

    # TABELA INFO
    info = sections.get("info", {})
    dados = []

    campos = [
        ("Cliente", "cliente"),
        ("CNC", "cnc"),
        ("Data", "data"),
        ("Técnico", "tecnico"),
        ("Acompanhamento", "acompanhamento"),
        ("Início", "inicio"),
        ("Fim", "fim"),
        ("Tempo Atendimento", "tempo_atendimento"),
        ("Tempo Espera", "tempo_espera"),
    ]

    for label, key in campos:
        if info.get(key):
            dados.append([label, info[key]])

    if dados:
        tabela = Table(dados, colWidths=[5*cm, 10*cm])
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FCFDFE")),
            ("BOX", (0,0), (-1,-1), 1, colors.HexColor("#C7D4DF")),
            ("INNERGRID", (0,0), (-1,-1), 0.5, colors.HexColor("#DCE5EC")),
            ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#E9F0F6")),
            ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("RIGHTPADDING", (0,0), (-1,-1), 8),
            ("TOPPADDING", (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(tabela)
        story.append(Spacer(1, 15))

    # SEÇÕES
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

        largura_max = 14.5 * cm
        altura_max = 11 * cm

        for idx, foto in enumerate(fotos, start=1):
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

            story.append(Spacer(1, 12))
            if idx != len(fotos):
                story.append(PageBreak())

    # CABEÇALHO + RODAPÉ COM NUMERAÇÃO
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
        canvas.drawString(2.5*cm, 1.4*cm, "Relatório técnico")
        canvas.drawRightString(18.5*cm, 1.4*cm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(temp_pdf, pagesize=A4,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        topMargin=3.4*cm, bottomMargin=2.5*cm)

    doc.build(story, onFirstPage=page_chrome, onLaterPages=page_chrome)

    # TEMPLATE EM TODAS AS PÁGINAS
    if template_path and os.path.exists(template_path):
        template_reader = PdfReader(template_path)
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
        os.rename(temp_pdf, output_path)


# =========================
# UI
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Relatórios")
        self.geometry("900x600")

        self.config_data = load_config()

        self.label = ctk.CTkLabel(self, text="Template: não selecionado")
        self.label.pack(pady=10)

        ctk.CTkButton(self, text="Selecionar Template",
                      command=self.select_template).pack()

        self.text = ctk.CTkTextbox(self, width=800, height=350)
        self.text.pack(pady=10)

        frame = ctk.CTkFrame(self)
        frame.pack(pady=10)

        ctk.CTkButton(frame, text="Gerar PDF",
                      command=self.generate).pack(side="left", padx=10)

        ctk.CTkButton(frame, text="Limpar",
                      command=lambda: self.text.delete("1.0", "end")
                      ).pack(side="left", padx=10)

        self.update_label()

    def update_label(self):
        path = self.config_data.get("template_path", "")
        if path:
            self.label.configure(text=f"Template: {path}")

    def select_template(self):
        file = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if file:
            self.config_data["template_path"] = file
            save_config(self.config_data)
            self.update_label()

    def generate(self):
        texto = self.text.get("1.0", "end").strip()

        if not texto:
            messagebox.showerror("Erro", "Cole o texto")
            return

        sections = parse_text(texto)

        save = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not save:
            return

        fotos = []
        anexar_fotos = messagebox.askyesno("Anexar fotos", "Deseja anexar fotos no PDF?")
        if anexar_fotos:
            arquivos = filedialog.askopenfilenames(
                title="Selecione uma ou mais fotos",
                filetypes=[
                    ("Imagens", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"),
                    ("Todos os arquivos", "*.*"),
                ],
            )
            if arquivos:
                usar_titulos = messagebox.askyesno(
                    "Título das fotos",
                    "Deseja definir um título para cada foto?",
                )
                for idx, arquivo in enumerate(arquivos, start=1):
                    titulo = ""
                    if usar_titulos:
                        titulo = simpledialog.askstring(
                            "Título da foto",
                            f"Digite o título da foto {idx}:",
                            parent=self,
                        ) or ""
                    fotos.append({"path": arquivo, "title": titulo.strip()})

        try:
            gerar_pdf(sections, self.config_data.get("template_path", ""), save, fotos=fotos)
            messagebox.showinfo("Sucesso", "PDF gerado!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))


# =========================
# RUN
# =========================
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
