from tkinter import messagebox


def handle_certificate_generation(*, persist_certificate_fields, save_config_fn, config_data):
    persist_certificate_fields()
    save_config_fn(config_data)
    messagebox.showinfo(
        "Modelo em preparação",
        "A tela de Certificado de treinamento foi criada.\n"
        "A geração de PDF deste modelo será implementada no próximo passo.",
    )
