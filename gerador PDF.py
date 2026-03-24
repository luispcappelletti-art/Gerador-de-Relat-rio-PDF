import customtkinter as ctk
from tkinter import filedialog, messagebox
import json
import os
import re

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from copy import deepcopy

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
    sections = {}

    patterns = {
        "data": r"Data:\s*(.*)",
        "inicio": r"Horário de início:\s*(.*)",
        "fim": r"Horário de término:\s*(.*)",
        "tempo_atendimento": r"Tempo do Atendimento:\s*(.*)",
        "tempo_espera": r"Tempo em espera:\s*(.*)"
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        sections[key] = match.group(1).strip() if match else ""

    split_sections = re.split(r"\n\s*(\d+\s*[–-]\s*.+)", text)

    for i in range(1, len(split_sections), 2):
        title = split_sections[i].strip()
        content = split_sections[i+1].strip()

        if "Descrição breve" in title:
            sections["descricao"] = content
        elif "Detalhamento" in title:
            sections["detalhamento"] = content
        elif "Diagnóstico" in title:
            sections["diagnostico"] = content
        elif "Ações corretivas" in title:
            sections["acoes"] = content
        elif "Resultado" in title:
            sections["resultado"] = content
        elif "Estado final" in title:
            sections["estado"] = content

    return sections


# =========================
# PDF
# =========================
def gerar_pdf(sections, template_path, output_path):
    temp_pdf = "temp.pdf"

    styles = getSampleStyleSheet()

    # Estilos personalizados
    styles["Heading1"].fontSize = 16
    styles["Heading1"].spaceAfter = 14

    styles["Heading2"].fontSize = 13
    styles["Heading2"].spaceAfter = 10

    styles["Normal"].fontSize = 10
    styles["Normal"].leading = 14

    story = []

    def linha():
        story.append(Spacer(1, 8))

    def titulo_principal(texto):
        story.append(Paragraph(f"<b>{texto}</b>", styles["Heading1"]))
        linha()

    def titulo_secao(texto):
        story.append(Paragraph(f"<b>{texto}</b>", styles["Heading2"]))
        linha()

    def texto(texto):
        story.append(Paragraph(texto.replace("\n", "<br/>"), styles["Normal"]))
        story.append(Spacer(1, 12))

    # =========================
    # CONTEÚDO
    # =========================

    titulo_principal("RELATÓRIO TÉCNICO DE ATENDIMENTO")

    texto(f"""
    <b>Data:</b> {sections.get('data', '')}<br/>
    <b>Início:</b> {sections.get('inicio', '')}<br/>
    <b>Fim:</b> {sections.get('fim', '')}<br/>
    <b>Tempo Atendimento:</b> {sections.get('tempo_atendimento', '')}<br/>
    <b>Tempo Espera:</b> {sections.get('tempo_espera', '')}
    """)

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