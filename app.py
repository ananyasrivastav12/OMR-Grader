from __future__ import annotations

import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from code import grade_batch


class OMRGraderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("OMR Grader")
        self.root.geometry("860x450")
        self.root.minsize(780, 420)
        self.root.configure(bg="#fff7fb")

        self.answer_sheet_var = tk.StringVar()
        self.omr_folder_var = tk.StringVar()
        self.save_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Pick files and click Generate ✨")
        self.result_var = tk.StringVar(value="")
        self.loader_var = tk.StringVar(value="")
        self.save_debug_var = tk.BooleanVar(value=False)

        self._loader_frames = ["🐇", "🐇💨", "🐇✨", "🐇💨✨"]
        self._loader_idx = 0
        self._is_busy = False

        self._build_style()
        self._build_ui()

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#ffffff", foreground="#4a2f43", font=("Segoe UI", 18, "bold"))
        style.configure("Hint.TLabel", background="#ffffff", foreground="#6d4e62", font=("Segoe UI", 10))
        style.configure("Status.TLabel", background="#ffffff", foreground="#3e2d3b", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=(16, 9))
        style.configure("Cute.TButton", font=("Segoe UI", 10, "bold"), padding=(10, 6))

        style.map(
            "Primary.TButton",
            background=[("active", "#f7cfe5"), ("!disabled", "#f2b8d6")],
            foreground=[("!disabled", "#42263a")],
        )
        style.map(
            "Cute.TButton",
            background=[("active", "#f8dcee"), ("!disabled", "#f4d0e6")],
            foreground=[("!disabled", "#4a2f43")],
        )

    def _build_ui(self) -> None:
        card = ttk.Frame(self.root, style="Card.TFrame", padding=24)
        card.pack(fill="both", expand=True, padx=22, pady=22)

        ttk.Label(card, text="OMR Result Studio", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            card,
            text="Select the input answer sheet, a folder of .bmp files, and where to save the results.",
            style="Hint.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 18))

        ttk.Label(card, text="📄 Answer Sheet (.xlsx)", style="Hint.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(card, textvariable=self.answer_sheet_var).grid(row=3, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(card, text="Pick File", style="Cute.TButton", command=self._pick_answer_sheet).grid(row=3, column=1, sticky="ew")

        ttk.Label(card, text="📁 OMR Images Folder (.bmp)", style="Hint.TLabel").grid(row=4, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(card, textvariable=self.omr_folder_var).grid(row=5, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(card, text="Pick Folder", style="Cute.TButton", command=self._pick_omr_folder).grid(row=5, column=1, sticky="ew")

        ttk.Label(card, text="💾 Select the folder where the Result file should be saved.", style="Hint.TLabel").grid(row=6, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(card, textvariable=self.save_path_var).grid(row=7, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(card, text="Choose Save Path", style="Cute.TButton", command=self._pick_save_path).grid(row=7, column=1, sticky="ew")

        ttk.Checkbutton(card, text="Save debug files (don't turn on.)", variable=self.save_debug_var).grid(
            row=8, column=0, sticky="w", pady=(16, 12)
        )

        self.run_button = ttk.Button(card, text="✨ Generate Result", style="Primary.TButton", command=self._run)
        self.run_button.grid(row=9, column=0, sticky="w", pady=(2, 8))

        self.progress = ttk.Progressbar(card, mode="indeterminate", length=220)
        self.progress.grid(row=9, column=1, sticky="w", padx=(8, 0))

        self.loader_label = ttk.Label(card, textvariable=self.loader_var, style="Status.TLabel")
        self.loader_label.grid(row=9, column=2, sticky="w", padx=(10, 0))

        ttk.Label(card, textvariable=self.status_var, style="Status.TLabel").grid(row=10, column=0, columnspan=3, sticky="w", pady=(8, 4))
        ttk.Label(card, textvariable=self.result_var, style="Hint.TLabel", wraplength=760).grid(row=11, column=0, columnspan=3, sticky="w")

        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=0)
        card.columnconfigure(2, weight=0)

    def _pick_answer_sheet(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Answer Sheet",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.answer_sheet_var.set(path)
            if not self.save_path_var.get().strip():
                self._prefill_save_path()

    def _pick_omr_folder(self) -> None:
        path = filedialog.askdirectory(title="Select OMR Folder")
        if path:
            self.omr_folder_var.set(path)
            if not self.save_path_var.get().strip():
                self._prefill_save_path()

    def _prefill_save_path(self) -> None:
        folder = self.omr_folder_var.get().strip()
        if not folder:
            return
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_name = f"{stamp}-result.xlsx"
        self.save_path_var.set(str(Path(folder) / default_name))

    def _pick_save_path(self) -> None:
        folder = self.omr_folder_var.get().strip()
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_name = f"{stamp}-result.xlsx"

        path = filedialog.asksaveasfilename(
            title="Save Result Excel",
            defaultextension=".xlsx",
            initialdir=folder if folder else None,
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.save_path_var.set(path)

    def _set_busy(self, busy: bool) -> None:
        self._is_busy = busy
        if busy:
            self.run_button.configure(state="disabled")
            self.progress.start(12)
            self._loader_idx = 0
            self._tick_loader()
        else:
            self.progress.stop()
            self.run_button.configure(state="normal")
            self.loader_var.set("")

    def _tick_loader(self) -> None:
        if not self._is_busy:
            return
        self.loader_var.set(self._loader_frames[self._loader_idx])
        self._loader_idx = (self._loader_idx + 1) % len(self._loader_frames)
        self.root.after(160, self._tick_loader)

    def _run(self) -> None:
        answer_sheet = Path(self.answer_sheet_var.get().strip())
        omr_folder = Path(self.omr_folder_var.get().strip())
        save_path_raw = self.save_path_var.get().strip()
        save_path = Path(save_path_raw) if save_path_raw else None

        if not answer_sheet.exists() or answer_sheet.suffix.lower() != ".xlsx":
            messagebox.showerror("Invalid Answer Sheet", "Please select a valid .xlsx answer sheet file.")
            return
        if not omr_folder.exists() or not omr_folder.is_dir():
            messagebox.showerror("Invalid Folder", "Please select a valid folder containing .bmp OMR images.")
            return
        if save_path is None:
            messagebox.showerror("Missing Save Path", "Please choose where to save the result Excel file.")
            return
        if save_path.suffix.lower() != ".xlsx":
            messagebox.showerror("Invalid Save Path", "Result file must end with .xlsx")
            return

        self.status_var.set("Working on it... grading sheets now 🌸")
        self.result_var.set("")
        self._set_busy(True)

        worker = threading.Thread(
            target=self._run_worker,
            args=(answer_sheet, omr_folder, save_path, bool(self.save_debug_var.get())),
            daemon=True,
        )
        worker.start()

    def _run_worker(self, answer_sheet: Path, omr_folder: Path, save_path: Path, save_debug: bool) -> None:
        try:
            _, excel_path = grade_batch(
                answer_sheet_path=answer_sheet,
                omr_folder=omr_folder,
                output_dir=save_path.parent,
                output_excel_path=save_path,
                save_debug=save_debug,
            )
            self.root.after(0, self._on_success, excel_path)
        except Exception as exc:
            self.root.after(0, self._on_error, str(exc))

    def _on_success(self, excel_path: Path) -> None:
        self._set_busy(False)
        self.status_var.set("Done! Results are ready 💾")
        self.result_var.set(f"Saved: {excel_path}")
        messagebox.showinfo("Success", f"Result file created:\n{excel_path}")

    def _on_error(self, error: str) -> None:
        self._set_busy(False)
        self.status_var.set("Oops, something went wrong.")
        self.result_var.set(error)
        messagebox.showerror("Error", error)


def main() -> None:
    root = tk.Tk()
    OMRGraderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
