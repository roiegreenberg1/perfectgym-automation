import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(2)

import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from formatter import format_report

# Load config
with open("config.json") as f:
    config = json.load(f)

# ── Drag-and-drop support (requires tkinterdnd2) ──────────────────────────────
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

class App(TkinterDnD.Tk if DND_AVAILABLE else tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("LTS Afternoon Sheets Formatting Tool")
        self.resizable(False, False)
        self.configure(bg="#ffffff")

        width, height = 1100, 900
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

        self.file_path = tk.StringVar(value="")
        self._load_logo()
        self._build_ui()

    def _load_logo(self):
        self.logo_image = None
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images", "GESAC_logo.png")
        if os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                target_h = 150
                ratio = target_h / img.height
                target_w = int(img.width * ratio)
                img = img.resize((target_w, target_h), Image.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(img)
            except ImportError:
                pass

    def _build_ui(self):

        # Header
        header = tk.Frame(self, bg="#ffffff", height=200)
        header.pack(fill="x")
        header.pack_propagate(False)

        header_inner = tk.Frame(header, bg="#ffffff")
        header_inner.place(relx=0, rely=0.5, anchor="w", x=40)

        if self.logo_image:
            tk.Label(
                header_inner, image=self.logo_image,
                bg="#ffffff"
            ).pack(side="left", padx=(0, 30))

        tk.Label(
            header_inner,
            text="LTS Afternoon Sheets Formatting Tool",
            font=("Segoe UI", 16, "bold"),
            bg="#ffffff", fg="#1a2e4a"
        ).pack(side="left")

        # Top divider
        tk.Frame(self, bg="#1a2e4a", height=3).pack(fill="x")

        # Main content
        content = tk.Frame(self, bg="#ffffff")
        content.pack(fill="both", expand=True, padx=50, pady=30)

        tk.Label(
            content,
            text="Group Enrollment Details",
            font=("Segoe UI", 14, "bold"),
            bg="#ffffff", fg="#1a2e4a"
        ).pack(anchor="w")

        tk.Label(
            content,
            text="Select or drop your unformatted .xlsx file below, then click 'Format Report'.",
            font=("Segoe UI", 10),
            bg="#ffffff", fg="#888888",
            wraplength=750,
            justify="left"
        ).pack(anchor="w", pady=(6, 20))

        # File drop zone
        self.drop_frame = tk.Frame(
            content, bg="#f8fafc",
            highlightbackground="#c8d8e8",
            highlightthickness=2,
        )
        self.drop_frame.pack(fill="x", ipady=50)

        inner = tk.Frame(self.drop_frame, bg="#f8fafc")
        inner.pack(expand=True)

        tk.Label(
            inner, text="📂",
            font=("Segoe UI", 30),
            bg="#f8fafc"
        ).pack()

        self.drop_label = tk.Label(
            inner,
            text="Drop .xlsx file here  or  click to browse",
            font=("Segoe UI", 11),
            bg="#f8fafc", fg="#999999"
        )
        self.drop_label.pack(pady=(8, 0))

        for widget in [self.drop_frame, inner, self.drop_label]:
            widget.bind("<Button-1>", lambda e: self._browse())

        if DND_AVAILABLE:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)

        # Format button
        self.format_btn = tk.Button(
            content,
            text="Format Report",
            font=("Segoe UI", 13, "bold"),
            bg="#1a2e4a", fg="white",
            activebackground="#2a4a70",
            activeforeground="white",
            relief="flat", cursor="hand2",
            pady=18,
            command=self._run_formatter
        )
        self.format_btn.pack(fill="x", pady=(28, 0))

        # Status label
        self.status_label = tk.Label(
            content, text="",
            font=("Segoe UI", 10),
            bg="#ffffff", fg="#888888"
        )
        self.status_label.pack(pady=(14, 0))

    def _browse(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self._set_file(path)

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        if path.endswith(".xlsx"):
            self._set_file(path)
        else:
            messagebox.showerror("Invalid file", "Please drop an .xlsx file.")

    def _set_file(self, path):
        self.file_path.set(path)
        filename = os.path.basename(path)
        self.drop_label.config(
            text=f"✓  {filename}",
            fg="#1a2e4a",
            font=("Segoe UI", 11, "bold")
        )
        self.status_label.config(text="")

    def _run_formatter(self):
        path = self.file_path.get()

        if not path:
            messagebox.showwarning("No file", "Please select an .xlsx file first.")
            return

        if not os.path.exists(path):
            messagebox.showerror("File not found", "The selected file no longer exists.")
            return

        try:
            self.format_btn.config(state="disabled", text="Formatting...")
            self.update()

            output_file = format_report(path, config)

            self.format_btn.config(state="normal", text="Format Report")
            self.status_label.config(text="✓  Done! The formatted report has been opened in Excel.", fg="#1a2e4a")
            self.update()

            os.startfile(os.path.abspath(output_file))

        except Exception as e:
            self.format_btn.config(state="normal", text="Format Report")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()