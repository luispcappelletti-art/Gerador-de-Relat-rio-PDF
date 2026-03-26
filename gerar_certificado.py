import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import json
import os
import subprocess
import sys
from module_registry import discover_modules


LAUNCHER_STATE_FILE = os.path.join(os.path.dirname(__file__), "launcher_state.json")


def save_last_module(module_filename):
    try:
        with open(LAUNCHER_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_module": module_filename}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


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

        self.current_module = "gerar_certificado.py"
        self._module_targets = {label: name for name, label in discover_modules() if name != self.current_module}
        self._module_label_var = tk.StringVar(value=self._default_module_target_label())

        switch_row = ctk.CTkFrame(frame, fg_color="transparent")
        switch_row.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(switch_row, text="Ir para:").pack(side="left")
        self._module_menu = ctk.CTkOptionMenu(
            switch_row,
            variable=self._module_label_var,
            values=self._menu_values(),
            width=250,
        )
        self._module_menu.pack(side="left", padx=(8, 6))
        ctk.CTkButton(switch_row, text="Abrir", command=self._switch_to_selected_module).pack(side="left")

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

    def _menu_values(self):
        return list(self._module_targets.keys()) or ["Nenhuma outra tela disponível"]

    def _default_module_target_label(self):
        values = self._menu_values()
        return values[0]

    def _switch_to_selected_module(self):
        module_label = self._module_label_var.get()
        module_filename = self._module_targets.get(module_label)
        if not module_filename:
            messagebox.showwarning("Aviso", "Nenhuma outra tela disponível no momento.")
            return
        self._switch_module(module_filename)

    def _switch_module(self, module_filename):
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


if __name__ == "__main__":
    save_last_module("gerar_certificado.py")
    ctk.set_appearance_mode("dark")
    app = CertificadoApp()
    app.mainloop()
