import customtkinter as ctk
from tkinter import filedialog, messagebox
import json
import os
import re
import unicodedata

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from copy import deepcopy
from datetime import datetime

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
def parse_text(text):
    def normalize_label(value):
        value = unicodedata.normalize("NFKD", value)
        value = "".join(ch for ch in value if not unicodedata.combining(ch))
        value = value.lower()
        value = re.sub(r"\s+", " ", value).strip()
        return value

    info_aliases = {
        "cliente": "cliente",
        "equipamento": "equipamento",
        "fonte": "fonte",
        "cnc": "cnc",
        "thc": "thc",
        "fabricante": "fabricante",
        "contato cliente": "contato_cliente",
        "data": "data",
        "tecnico": "tecnico",
        "técnico": "tecnico",
        "acompanhamento remoto": "acompanhamento_remoto",
        "horario de inicio": "inicio",
        "horario de termino": "fim",
        "tempo do atendimento": "tempo_atendimento",
        "tempo em espera": "tempo_espera",
    }

    section_aliases = {
        "descricao breve do problema": "descricao",
        "descricao breve": "descricao",
        "detalhamento completo do problema": "detalhamento",
        "detalhamento do problema": "detalhamento",
        "detalhamento": "detalhamento",
        "diagnostico": "diagnostico",
        "acoes corretivas": "acoes",
        "acoes corretivas aplicadas": "acoes",
        "resultado": "resultado",
        "estado final do sistema": "estado",
        "estado final": "estado",
    }

    sections = {"info": {}}

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        raw_key, raw_value = line.split(":", 1)
        key = normalize_label(raw_key)
        if key in info_aliases:
            sections["info"][info_aliases[key]] = raw_value.strip()

    split_sections = re.split(r"\n\s*(\d+\s*[–-]\s*.+)", text)

    for i in range(1, len(split_sections), 2):
        raw_title = split_sections[i].strip()
        content = split_sections[i + 1].strip()
        clean_title = re.sub(r"^\d+\s*[–-]\s*", "", raw_title)
        section_key = section_aliases.get(normalize_label(clean_title))
        if section_key:
            sections[section_key] = content

    return sections


# =========================
# PDF
# =========================
def gerar_pdf(sections, template_path, output_path):
    temp_pdf = "temp.pdf"

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="ReportTitle",
        parent=styles["Heading1"],
        fontSize=20,
        textColor=colors.HexColor("#0E2A44"),
        alignment=1,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="ReportSubtitle",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#4B647A"),
        alignment=1,
        spaceAfter=18,
    ))
    styles.add(ParagraphStyle(
        name="SectionTitle",
        parent=styles["Heading2"],
        fontSize=13,
        textColor=colors.HexColor("#123A5A"),
        spaceBefore=12,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name="SectionBody",
        parent=styles["BodyText"],
        fontSize=10.5,
        leading=16,
        textColor=colors.HexColor("#1F2F3D"),
    ))

    story = []

    def linha():
        story.append(Spacer(1, 8))

    def titulo_principal(texto):
        story.append(Paragraph(f"<b>{texto}</b>", styles["ReportTitle"]))
        linha()

    def titulo_secao(texto):
        story.append(Paragraph(f"<b>{texto}</b>", styles["SectionTitle"]))
        linha()

    def texto(texto):
        story.append(Paragraph(texto.replace("\n", "<br/>"), styles["SectionBody"]))
        story.append(Spacer(1, 12))

    # =========================
    # CONTEÚDO
    # =========================

    titulo_principal("RELATÓRIO DE ATENDIMENTO TÉCNICO")
    story.append(Paragraph("Documento técnico padronizado", styles["ReportSubtitle"]))

    info = sections.get("info", {})
    info_rows = []
    table_map = [
        ("Cliente", info.get("cliente", "-")),
        ("Equipamento", info.get("equipamento", "-")),
        ("Fonte", info.get("fonte", "-")),
        ("CNC", info.get("cnc", "-")),
        ("THC", info.get("thc", "-")),
        ("Fabricante", info.get("fabricante", "-")),
        ("Contato Cliente", info.get("contato_cliente", "-")),
        ("Data", info.get("data", "-")),
        ("Técnico", info.get("tecnico", "-")),
        ("Acompanhamento remoto", info.get("acompanhamento_remoto", "-")),
        ("Horário de início", info.get("inicio", "-")),
        ("Horário de término", info.get("fim", "-")),
        ("Tempo do Atendimento", info.get("tempo_atendimento", "-")),
        ("Tempo em espera", info.get("tempo_espera", "-")),
    ]

    for label, value in table_map:
        if value and value != "-":
            info_rows.append([f"<b>{label}</b>", value])

    if info_rows:
        info_table = Table(info_rows, colWidths=[5.5 * cm, 9.5 * cm])
        info_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F8FB")),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#C8D2DD")),
            ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D5DEE7")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#243848")),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(info_table)
        story.append(Spacer(1, 14))

    if sections.get("descricao"):
        titulo_secao("1 – DESCRIÇÃO BREVE")
        texto(sections["descricao"])

    if sections.get("detalhamento"):
        titulo_secao("2 – DETALHAMENTO DO PROBLEMA")
        texto(sections["detalhamento"])

    if sections.get("diagnostico"):
        titulo_secao("3 – DIAGNÓSTICO")
        texto(sections["diagnostico"])

    if sections.get("acoes"):
        titulo_secao("4 – AÇÕES CORRETIVAS")
        texto(sections["acoes"])

    if sections.get("resultado"):
        titulo_secao("5 – RESULTADO")
        texto(sections["resultado"])

    if sections.get("estado"):
        titulo_secao("6 – ESTADO FINAL DO SISTEMA")
        texto(sections["estado"])

    story.append(Spacer(1, 8))
    story.append(Paragraph(
        f"<font size='9' color='#4B647A'>Relatório gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</font>",
        styles["Normal"]
    ))

    doc = SimpleDocTemplate(
        temp_pdf,
        pagesize=A4,
        leftMargin=2.5*cm,
        rightMargin=2.5*cm,
        topMargin=3*cm,
        bottomMargin=2.5*cm
    )

    doc.build(story)

    # =========================
    # MESCLAGEM CORRIGIDA
    # =========================
    if template_path and os.path.exists(template_path):
        template = PdfReader(template_path)
        content = PdfReader(temp_pdf)
        writer = PdfWriter()

        base_page = template.pages[0]

        for i in range(len(content.pages)):
            new_page = deepcopy(base_page)
            new_page.merge_page(content.pages[i])
            writer.add_page(new_page)

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

        self.title("Gerador de Relatório PDF")
        self.geometry("900x600")

        self.config_data = load_config()

        # Template
        self.template_label = ctk.CTkLabel(self, text="Template: não selecionado")
        self.template_label.pack(pady=10)

        self.btn_template = ctk.CTkButton(self, text="Selecionar Template", command=self.select_template)
        self.btn_template.pack(pady=5)

        # Texto
        self.textbox = ctk.CTkTextbox(self, width=800, height=350)
        self.textbox.pack(pady=10)

        # Botões
        frame = ctk.CTkFrame(self)
        frame.pack(pady=10)

        self.btn_generate = ctk.CTkButton(frame, text="Gerar PDF", command=self.generate_pdf)
        self.btn_generate.pack(side="left", padx=10)

        self.btn_clear = ctk.CTkButton(frame, text="Limpar", command=self.clear_text)
        self.btn_clear.pack(side="left", padx=10)

        self.load_template_label()

    def load_template_label(self):
        path = self.config_data.get("template_path", "")
        if path:
            self.template_label.configure(text=f"Template: {path}")

    def select_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.config_data["template_path"] = file_path
            save_config(self.config_data)
            self.template_label.configure(text=f"Template: {file_path}")

    def clear_text(self):
        self.textbox.delete("1.0", "end")

    def generate_pdf(self):
        text = self.textbox.get("1.0", "end").strip()

        if not text:
            messagebox.showerror("Erro", "Cole o texto primeiro")
            return

        sections = parse_text(text)

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")

        if not save_path:
            return

        try:
            gerar_pdf(sections, self.config_data.get("template_path", ""), save_path)
            messagebox.showinfo("Sucesso", "PDF gerado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))


# =========================
# RUN
# =========================
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()
