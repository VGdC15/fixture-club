import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

# Importa tu motor (main.py)
import main as engine

DEFAULT_INPUT = "Plantilla_Clubes_REAL.xlsx"
DEFAULT_OUTPUT = "Fixture_2026.xlsx"


def open_file(path: str):
    if not os.path.exists(path):
        messagebox.showerror("No encontrado", f"No existe el archivo:\n{path}")
        return

    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception as ex:
        messagebox.showerror("Error", f"No pude abrir el archivo.\n\nDetalle:\n{ex}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Fixture")
        self.geometry("620x320")
        self.resizable(False, False)

        self.input_path = tk.StringVar(value=os.path.abspath(DEFAULT_INPUT))
        self.output_path = tk.StringVar(value=os.path.abspath(DEFAULT_OUTPUT))
        self.status = tk.StringVar(value="Listo. Elegí el Excel base o generá directo 😉")

        self._build_ui()

    def _build_ui(self):
        pad = 12

        title = tk.Label(self, text="Generador de Fixture (Todos contra Todos)", font=("Segoe UI", 14, "bold"))
        title.pack(pady=(pad, 6))

        frm = tk.Frame(self)
        frm.pack(fill="x", padx=pad, pady=6)

        # Input
        tk.Label(frm, text="Excel base (Plantilla):", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.input_path, width=60).grid(row=1, column=0, sticky="w", pady=(4, 10))
        tk.Button(frm, text="Buscar…", command=self.pick_input, width=12).grid(row=1, column=1, padx=(10, 0))

        # Output
        tk.Label(frm, text="Salida (Fixture):", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.output_path, width=60).grid(row=3, column=0, sticky="w", pady=(4, 10))
        tk.Button(frm, text="Guardar como…", command=self.pick_output, width=12).grid(row=3, column=1, padx=(10, 0))

        # Buttons
        btns = tk.Frame(self)
        btns.pack(fill="x", padx=pad, pady=(6, 6))

        self.btn_generate = tk.Button(btns, text="✅ Generar Fixture", command=self.generate, height=2)
        self.btn_generate.pack(side="left", fill="x", expand=True)

        tk.Button(btns, text="📂 Abrir Plantilla", command=lambda: open_file(self.input_path.get()), height=2).pack(
            side="left", fill="x", expand=True, padx=(10, 0)
        )

        self.btn_open_out = tk.Button(btns, text="📄 Abrir Fixture", command=lambda: open_file(self.output_path.get()), height=2)
        self.btn_open_out.pack(side="left", fill="x", expand=True, padx=(10, 0))

        # Status bar
        status_frame = tk.Frame(self, bd=1, relief="sunken")
        status_frame.pack(side="bottom", fill="x")
        tk.Label(status_frame, textvariable=self.status, anchor="w", padx=10).pack(fill="x")

        # Tip
        tip = tk.Label(
            self,
            text="Tip: marcá con X las categorías por club. El programa detecta categorías nuevas automáticamente.",
            fg="#444",
        )
        tip.pack(pady=(8, 0))

    def pick_input(self):
        path = filedialog.askopenfilename(
            title="Seleccionar plantilla",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=os.getcwd(),
        )
        if path:
            self.input_path.set(path)
            self.status.set("Plantilla seleccionada. Listo para generar.")

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Guardar fixture como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=os.getcwd(),
            initialfile=DEFAULT_OUTPUT,
        )
        if path:
            self.output_path.set(path)
            self.status.set("Ruta de salida elegida. Listo para generar.")

    def generate(self):
        in_path = self.input_path.get()
        out_path = self.output_path.get()

        if not os.path.exists(in_path):
            messagebox.showerror("Falta plantilla", "No encuentro el Excel base. Elegilo con 'Buscar…'.")
            return

        # Bloquear botones mientras corre
        self.btn_generate.config(state="disabled")
        self.btn_open_out.config(state="disabled")
        self.status.set("Generando fixture… (no cierres la ventana)")

        def run():
            try:
                engine.generate_fixture(in_path, out_path)
            except Exception as e:
                # ✅ FIX: capturamos la excepción en el lambda
                self.after(0, lambda err=e: self._on_error(err))
                return
            self.after(0, self._on_success)

        threading.Thread(target=run, daemon=True).start()

    def _on_success(self):
        self.btn_generate.config(state="normal")
        self.btn_open_out.config(state="normal")
        self.status.set("🎉 Listo. Fixture generado. Podés abrirlo con 'Abrir Fixture'.")
        messagebox.showinfo("Éxito", "Fixture generado correctamente ✅")

    def _on_error(self, err: Exception):
        self.btn_generate.config(state="normal")
        self.btn_open_out.config(state="normal")
        self.status.set("⚠️ Hubo un error. Revisá el mensaje.")
        messagebox.showerror("Error al generar", f"Ocurrió un error:\n\n{err}")


if __name__ == "__main__":
    App().mainloop()
