import os
import threading
from tkinter import filedialog, messagebox


def generate_report_async(
    *,
    texto,
    sections,
    fotos,
    config_data,
    foto_cols_var,
    foto_h_var,
    watermark_opacity_var,
    watermark_scale_var,
    watermark_mode_var,
    signature_var,
    get_watermark_random_count,
    get_section_offsets_cm,
    get_horarios_from_ui,
    pick_pdf_save_dir,
    remember_pdf_save_dir,
    set_status,
    hide_open_pdf_button,
    on_success,
    on_error,
    validar_sections_fn,
    gerar_pdf_fn,
):
    if not texto:
        messagebox.showerror("Erro", "Cole o texto primeiro.")
        return

    avisos = validar_sections_fn(sections, fotos)
    if avisos:
        msg = "Atenção antes de gerar:\n\n" + "\n".join(f"• {a}" for a in avisos)
        msg += "\n\nDeseja continuar mesmo assim?"
        if not messagebox.askyesno("Validação", msg):
            return

    save = filedialog.asksaveasfilename(
        initialdir=pick_pdf_save_dir(),
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
    )
    if not save:
        return

    remember_pdf_save_dir(save)
    set_status("Gerando PDF...")
    hide_open_pdf_button()

    def _worker():
        try:
            gerar_pdf_fn(
                sections,
                config_data.get("template_path", ""),
                save,
                fotos=fotos,
                foto_cols=int(foto_cols_var.get()),
                foto_max_height_cm=float(foto_h_var.get()),
                section_offsets_cm=get_section_offsets_cm(),
                horarios=get_horarios_from_ui(),
                watermark_path=config_data.get("watermark_path", ""),
                watermark_opacity=float(watermark_opacity_var.get()),
                watermark_scale=float(watermark_scale_var.get()),
                watermark_mode=str(watermark_mode_var.get()),
                watermark_random_count=get_watermark_random_count(),
                cover_header_scale=float(config_data.get("cover_header_scale", "1.8")),
                include_signature_page=bool(signature_var.get()),
            )
            on_success(save)
        except Exception as exc:
            on_error(str(exc))

    threading.Thread(target=_worker, daemon=True).start()
