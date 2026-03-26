import customtkinter as ctk
from tkinter import messagebox


class CertificadoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gerar Certificados - Exemplo")
        self.geometry("620x420")
        self.minsize(560, 360)

        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            frame,
            text="Tela de exemplo - Gerar Certificado",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(10, 12))

        ctk.CTkLabel(frame, text="Nome do participante:").pack(anchor="w", padx=10)
        self.nome_entry = ctk.CTkEntry(frame, placeholder_text="Digite o nome")
        self.nome_entry.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(frame, text="Curso/Evento:").pack(anchor="w", padx=10)
        self.curso_entry = ctk.CTkEntry(frame, placeholder_text="Digite o curso/evento")
        self.curso_entry.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(frame, text="Texto prévia:").pack(anchor="w", padx=10)
        self.preview = ctk.CTkTextbox(frame, height=120)
        self.preview.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        actions = ctk.CTkFrame(frame, fg_color="transparent")
        actions.pack(fill="x", padx=8, pady=(0, 8))

        ctk.CTkButton(actions, text="Atualizar prévia", command=self.atualizar_previa).pack(side="left", padx=4)
        ctk.CTkButton(actions, text="Gerar certificado (exemplo)", command=self.gerar_certificado_exemplo).pack(side="left", padx=4)

        self.status = ctk.CTkLabel(frame, text="Módulo de exemplo pronto.", anchor="w")
        self.status.pack(fill="x", padx=10)

    def atualizar_previa(self):
        nome = self.nome_entry.get().strip() or "[Nome do participante]"
        curso = self.curso_entry.get().strip() or "[Curso/Evento]"

        texto = (
            "CERTIFICADO\n\n"
            f"Certificamos que {nome} participou de {curso}.\n\n"
            "(Função de exemplo - ajuste o layout depois.)"
        )

        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", texto)
        self.status.configure(text="Prévia atualizada.")

    def gerar_certificado_exemplo(self):
        self.atualizar_previa()
        messagebox.showinfo(
            "Exemplo",
            "Função de exemplo executada.\n"
            "Aqui você poderá implementar a geração real do certificado.",
        )
        self.status.configure(text="Função de geração (exemplo) executada.")


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = CertificadoApp()
    app.mainloop()
