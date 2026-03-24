import customtkinter as ctk
from tkinter import filedialog, messagebox
import json, os, re, unicodedata
from copy import deepcopy
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import *
from reportlab.lib.styles import *
from reportlab.lib.units import cm

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
    for linha in linhas:
        linha = linha.strip()

        if re.match(r"^(\d+[\.\)]|-|\u2022)", linha):
            item = re.sub(r"^(\d+[\.\)]|-|\u2022)\s*", "", linha)
            lista.append(ListItem(Paragraph(item, styles["Body"])))
        else:
            if lista:
                elementos.append(ListFlowable(lista, bulletType='1'))
                lista = []
            if linha:
                elementos.append(Paragraph(linha, styles["Body"]))

    if lista:
        elementos.append(ListFlowable(lista, bulletType='1'))

    elementos.append(Spacer(1, 10))
    return elementos


# =========================
# PDF
# =========================
def gerar_pdf(sections, template_path, output_path):
    temp_pdf = "temp.pdf"

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(name="Titulo",
        fontSize=18, alignment=1, spaceAfter=10, textColor=colors.HexColor("#0E2A44")))

    styles.add(ParagraphStyle(name="Subtitulo",
        fontSize=9, alignment=1, textColor=colors.grey, spaceAfter=15))

    styles.add(ParagraphStyle(name="Secao",
        fontSize=12, spaceBefore=12, spaceAfter=6, textColor=colors.HexColor("#123A5A")))

    styles.add(ParagraphStyle(name="Body",
        fontSize=10.5, leading=15))

    story = []

    # TÍTULO
    story.append(Paragraph("<b>RELATÓRIO DE ATENDIMENTO TÉCNICO</b>", styles["Titulo"]))
    story.append(Paragraph("Documento técnico padronizado", styles["Subtitulo"]))

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
            ("BACKGROUND", (0,0), (-1,-1), colors.white),
            ("BOX", (0,0), (-1,-1), 1, colors.grey),
            ("INNERGRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
            ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#ECEFF1")),
            ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ]))
        story.append(tabela)
        story.append(Spacer(1, 15))

    # SEÇÕES
    def add_secao(titulo, conteudo):
        story.append(Paragraph(f"<b>{titulo}</b>", styles["Secao"]))
        story.extend(processar_lista(conteudo, styles))

    if sections.get("descricao"):
        add_secao("1 – DESCRIÇÃO BREVE", sections["descricao"])

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

    # RODAPÉ COM NUMERAÇÃO
    def footer(canvas, doc):
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(20*cm, 1.5*cm, f"Página {doc.page}")
        canvas.drawString(2*cm, 1.5*cm,
            f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    doc = SimpleDocTemplate(temp_pdf, pagesize=A4,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        topMargin=3*cm, bottomMargin=2.5*cm)

    doc.build(story, onFirstPage=footer, onLaterPages=footer)

    # TEMPLATE EM TODAS AS PÁGINAS
    if template_path and os.path.exists(template_path):
        template = PdfReader(template_path)
        content = PdfReader(temp_pdf)
        writer = PdfWriter()

        for i in range(len(content.pages)):
            base = deepcopy(template.pages[0])
            base.merge_page(content.pages[i])
            writer.add_page(base)

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

        try:
            gerar_pdf(sections, self.config_data.get("template_path", ""), save)
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
