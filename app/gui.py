from __future__ import annotations

import asyncio
import json
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .excel_service import load_excel_plan, save_results
from .transpofrete_bot import BotOptions, run_consultas


COLORS = {
    "white": "#ffffff",
    "navy": "#212a57",
    "blue": "#079ae1",
    "orange": "#f36229",
    "light": "#f4f7fb",
    "muted": "#6b7280",
    "text": "#111827",
    "soft": "#eef4fb",
    "log_bg": "#0f172a",
    "log_fg": "#e5e7eb",
}


SETTINGS_FILE = Path("transpofrete_settings.json")


class TranspofreteApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Automação Transpofrete")
        self.root.geometry("940x700")
        self.root.minsize(860, 640)
        self.root.configure(bg=COLORS["light"])

        self.file_path_var = tk.StringVar()
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.max_tabs_var = tk.IntVar(value=5)

        self.headless_var = tk.BooleanVar(value=True)
        self.fast_mode_var = tk.BooleanVar(value=True)
        self.settle_ms_var = tk.IntVar(value=900)

        self.event_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.is_running = False

        self.password_entry: ttk.Entry | None = None
        self.password_toggle_button: ttk.Button | None = None
        self.start_button: ttk.Button | None = None

        self._load_settings()
        self._configure_style()
        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._process_queue)

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists():
            return

        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as file:
                data = json.load(file)

            self.file_path_var.set(str(data.get("file_path", "") or ""))
            self.username_var.set(str(data.get("username", "") or ""))
            self.password_var.set(str(data.get("password", "") or ""))

            self.max_tabs_var.set(int(data.get("max_tabs", 5) or 5))
            self.headless_var.set(bool(data.get("headless", True)))
            self.fast_mode_var.set(bool(data.get("fast_mode", True)))
            self.settle_ms_var.set(int(data.get("settle_ms", 900) or 900))

        except Exception:
            self.file_path_var.set("")
            self.username_var.set("")
            self.password_var.set("")
            self.max_tabs_var.set(5)
            self.headless_var.set(True)
            self.fast_mode_var.set(True)
            self.settle_ms_var.set(900)

    def _save_settings(self) -> None:
        try:
            max_tabs = int(self.max_tabs_var.get())
        except Exception:
            max_tabs = 5

        try:
            settle_ms = int(self.settle_ms_var.get())
        except Exception:
            settle_ms = 900

        data = {
            "file_path": self.file_path_var.get().strip(),
            "username": self.username_var.get().strip(),
            "password": self.password_var.get().strip(),
            "max_tabs": max_tabs,
            "headless": bool(self.headless_var.get()),
            "fast_mode": bool(self.fast_mode_var.get()),
            "settle_ms": settle_ms,
        }

        try:
            with SETTINGS_FILE.open("w", encoding="utf-8") as file:
                json.dump(data, file, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def _configure_style(self) -> None:
        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Modern.TEntry",
            fieldbackground="#f9fafb",
            borderwidth=0,
            relief="flat",
            padding=8,
            font=("Segoe UI", 10),
        )

        style.configure(
            "Modern.TSpinbox",
            fieldbackground="#f9fafb",
            borderwidth=0,
            relief="flat",
            padding=8,
            font=("Segoe UI", 10),
        )

        style.configure(
            "Primary.TButton",
            background=COLORS["blue"],
            foreground=COLORS["white"],
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
            focusthickness=0,
            padding=(14, 9),
        )

        style.map(
            "Primary.TButton",
            background=[("active", "#087fbd"), ("disabled", "#9ca3af")],
            foreground=[("active", COLORS["white"]), ("disabled", COLORS["white"])],
        )

        style.configure(
            "Ghost.TButton",
            background="#e5edf6",
            foreground=COLORS["navy"],
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            focusthickness=0,
            padding=(10, 8),
        )

        style.map(
            "Ghost.TButton",
            background=[("active", "#d7e6f4")],
            foreground=[("active", COLORS["navy"])],
        )

        style.configure(
            "Start.TButton",
            background=COLORS["orange"],
            foreground=COLORS["white"],
            font=("Segoe UI", 11, "bold"),
            borderwidth=0,
            focusthickness=0,
            padding=(18, 12),
        )

        style.map(
            "Start.TButton",
            background=[("active", "#d84f1f"), ("disabled", "#9ca3af")],
            foreground=[("active", COLORS["white"]), ("disabled", COLORS["white"])],
        )

        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#e5e7eb",
            background=COLORS["blue"],
            bordercolor="#e5e7eb",
            lightcolor=COLORS["blue"],
            darkcolor=COLORS["blue"],
            thickness=12,
        )

    def _build_ui(self) -> None:
        main = tk.Frame(self.root, bg=COLORS["light"], bd=0, highlightthickness=0)
        main.pack(fill="both", expand=True)

        self._build_header(main)

        body = tk.Frame(main, bg=COLORS["light"], bd=0, highlightthickness=0)
        body.pack(fill="both", expand=True, padx=24, pady=20)

        top = tk.Frame(body, bg=COLORS["light"], bd=0, highlightthickness=0)
        top.pack(fill="x")

        left_card = self._card(top)
        left_card.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_card = self._card(top)
        right_card.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self._build_file_access(left_card)
        self._build_execution(right_card)

        progress_card = self._card(body)
        progress_card.pack(fill="x", pady=(18, 0))
        self._build_progress(progress_card)

        log_card = self._card(body)
        log_card.pack(fill="both", expand=True, pady=(18, 0))
        self._build_log(log_card)

    def _build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=COLORS["navy"], bd=0, highlightthickness=0)
        header.pack(fill="x")

        inner = tk.Frame(header, bg=COLORS["navy"], bd=0, highlightthickness=0)
        inner.pack(fill="x", padx=30, pady=24)

        tk.Label(
            inner,
            text="Automação Transpofrete",
            bg=COLORS["navy"],
            fg=COLORS["white"],
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")

        tk.Label(
            inner,
            text="Consulta faturas das sheets Vencidos e A Vencer, atualiza status e captura observação quando estiver rejeitada.",
            bg=COLORS["navy"],
            fg="#dbeafe",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

    def _card(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["white"],
            bd=0,
            highlightthickness=0,
            padx=20,
            pady=18,
        )

    def _section_title(self, parent: tk.Widget, title: str, subtitle: str | None = None) -> None:
        tk.Label(
            parent,
            text=title,
            bg=COLORS["white"],
            fg=COLORS["navy"],
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")

        if subtitle:
            tk.Label(
                parent,
                text=subtitle,
                bg=COLORS["white"],
                fg=COLORS["muted"],
                font=("Segoe UI", 9),
                wraplength=390,
                justify="left",
            ).pack(anchor="w", pady=(3, 14))

    def _field_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(12, 4))

    def _build_file_access(self, parent: tk.Frame) -> None:
        self._section_title(
            parent,
            "Arquivo e acesso",
            "Os últimos dados usados serão carregados automaticamente na próxima execução.",
        )

        file_row = tk.Frame(parent, bg=COLORS["white"], bd=0, highlightthickness=0)
        file_row.pack(fill="x")

        self.file_entry = ttk.Entry(
            file_row,
            textvariable=self.file_path_var,
            style="Modern.TEntry",
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=3)

        ttk.Button(
            file_row,
            text="Selecionar",
            style="Primary.TButton",
            command=self._select_file,
        ).pack(side="left")

        self._field_label(parent, "Usuário")
        ttk.Entry(
            parent,
            textvariable=self.username_var,
            style="Modern.TEntry",
        ).pack(fill="x", ipady=3)

        self._field_label(parent, "Senha")

        password_row = tk.Frame(parent, bg=COLORS["white"], bd=0, highlightthickness=0)
        password_row.pack(fill="x")

        self.password_entry = ttk.Entry(
            password_row,
            textvariable=self.password_var,
            show="*",
            style="Modern.TEntry",
        )
        self.password_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=3)

        self.password_toggle_button = ttk.Button(
            password_row,
            text="Mostrar",
            style="Ghost.TButton",
            command=self._toggle_password,
        )
        self.password_toggle_button.pack(side="left")

        tk.Label(
            parent,
            text="Obs.: a senha fica salva localmente.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
            wraplength=390,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

    def _toggle_password(self) -> None:
        if self.password_entry is None:
            return

        current_show = self.password_entry.cget("show")

        if current_show == "":
            self.password_entry.configure(show="*")

            if self.password_toggle_button is not None:
                self.password_toggle_button.configure(text="Mostrar")
        else:
            self.password_entry.configure(show="")

            if self.password_toggle_button is not None:
                self.password_toggle_button.configure(text="Ocultar")

    def _build_execution(self, parent: tk.Frame) -> None:
        self._section_title(
            parent,
            "Execução",
            "Por padrão, o navegador roda oculto. Desmarque essa opção somente se quiser acompanhar o processo.",
        )

        row = tk.Frame(parent, bg=COLORS["white"], bd=0, highlightthickness=0)
        row.pack(fill="x")

        left = tk.Frame(row, bg=COLORS["white"], bd=0, highlightthickness=0)
        left.pack(side="left", fill="x", expand=True, padx=(0, 8))

        right = tk.Frame(row, bg=COLORS["white"], bd=0, highlightthickness=0)
        right.pack(side="left", fill="x", expand=True, padx=(8, 0))

        tk.Label(
            left,
            text="Abas simultâneas",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 4))

        ttk.Spinbox(
            left,
            from_=1,
            to=5,
            textvariable=self.max_tabs_var,
            style="Modern.TSpinbox",
            width=8,
        ).pack(anchor="w", ipady=3)

        tk.Label(
            right,
            text="Pausa entre ações (ms)",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 4))

        ttk.Spinbox(
            right,
            from_=300,
            to=3000,
            increment=100,
            textvariable=self.settle_ms_var,
            style="Modern.TSpinbox",
            width=10,
        ).pack(anchor="w", ipady=3)

        self._build_checkbox(
            parent,
            "Rodar navegador oculto",
            self.headless_var,
            pady=(18, 6),
        )

        self._build_checkbox(
            parent,
            "Modo rápido: bloquear imagens, fontes e mídia",
            self.fast_mode_var,
            pady=(2, 12),
        )

        self.start_button = ttk.Button(
            parent,
            text="Iniciar automação",
            style="Start.TButton",
            command=self._start,
        )
        self.start_button.pack(fill="x", pady=(8, 0))

    def _build_checkbox(
        self,
        parent: tk.Widget,
        text: str,
        variable: tk.BooleanVar,
        pady: tuple[int, int],
    ) -> None:
        wrapper = tk.Frame(parent, bg=COLORS["white"], bd=0, highlightthickness=0)
        wrapper.pack(fill="x", pady=pady)

        check = tk.Checkbutton(
            wrapper,
            text=text,
            variable=variable,
            bg=COLORS["white"],
            fg=COLORS["text"],
            activebackground=COLORS["white"],
            activeforeground=COLORS["text"],
            selectcolor=COLORS["white"],
            font=("Segoe UI", 10),
            bd=0,
            highlightthickness=0,
            relief="flat",
            anchor="w",
        )
        check.pack(anchor="w")

    def _build_progress(self, parent: tk.Frame) -> None:
        top = tk.Frame(parent, bg=COLORS["white"], bd=0, highlightthickness=0)
        top.pack(fill="x")

        self.progress_label = tk.Label(
            top,
            text="Pronto para iniciar",
            bg=COLORS["white"],
            fg=COLORS["navy"],
            font=("Segoe UI", 13, "bold"),
        )
        self.progress_label.pack(side="left")

        self.counter_label = tk.Label(
            top,
            text="0/0",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10, "bold"),
        )
        self.counter_label.pack(side="right")

        self.progressbar = ttk.Progressbar(
            parent,
            mode="determinate",
            style="Modern.Horizontal.TProgressbar",
        )
        self.progressbar.pack(fill="x", pady=(12, 0))

    def _build_log(self, parent: tk.Frame) -> None:
        tk.Label(
            parent,
            text="Log da execução",
            bg=COLORS["white"],
            fg=COLORS["navy"],
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")

        tk.Label(
            parent,
            text="Acompanhe cada aba, fatura consultada e eventuais erros diagnosticados.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(3, 10))

        self.log_text = ScrolledText(
            parent,
            height=14,
            font=("Consolas", 9),
            bg=COLORS["log_bg"],
            fg=COLORS["log_fg"],
            insertbackground=COLORS["log_fg"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            padx=14,
            pady=12,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

    def _select_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[
                ("Planilhas Excel", "*.xlsx *.xlsm"),
                ("Todos os arquivos", "*.*"),
            ],
        )

        if file_path:
            self.file_path_var.set(file_path)
            self._save_settings()

    def _validate(self) -> bool:
        if not self.file_path_var.get().strip():
            messagebox.showwarning("Atenção", "Selecione a planilha.")
            return False

        if not self.username_var.get().strip():
            messagebox.showwarning("Atenção", "Informe o usuário.")
            return False

        if not self.password_var.get().strip():
            messagebox.showwarning("Atenção", "Informe a senha.")
            return False

        try:
            max_tabs = int(self.max_tabs_var.get())
        except Exception:
            max_tabs = 5

        if max_tabs < 1 or max_tabs > 5:
            messagebox.showwarning("Atenção", "A quantidade de abas deve ser entre 1 e 5.")
            return False

        try:
            settle_ms = int(self.settle_ms_var.get())
        except Exception:
            settle_ms = 900

        if settle_ms < 300 or settle_ms > 3000:
            messagebox.showwarning("Atenção", "A pausa entre ações deve ser entre 300 e 3000 ms.")
            return False

        return True

    def _start(self) -> None:
        if self.is_running:
            return

        if not self._validate():
            return

        self._save_settings()

        self.is_running = True

        if self.start_button is not None:
            self.start_button.configure(state="disabled")

        self.progressbar["value"] = 0
        self.progressbar["maximum"] = 100
        self.progress_label.configure(text="Executando...")
        self.counter_label.configure(text="0/0")

        self._clear_log()

        self.worker_thread = threading.Thread(
            target=self._run_worker,
            daemon=True,
        )
        self.worker_thread.start()

    def _run_worker(self) -> None:
        try:
            file_path = self.file_path_var.get().strip()

            self._thread_log("Iniciando automação...")
            self._thread_log("Lendo planilha...")

            plan = load_excel_plan(file_path)
            total = len(plan.tasks)

            self._thread_log(f"Sheets consideradas: {', '.join(plan.sheet_names)}")
            self._thread_log(f"Faturas únicas pendentes encontradas: {total}")
            self._thread_log(
                f"Data inicial usada no filtro: {plan.dt_from.strftime('%d/%m/%Y 00:00:00')}"
            )
            self._thread_log("A data inicial é a mais antiga entre Vencidos e A Vencer.")
            self._thread_log("Data final será mantida como estiver no portal.")
            self._thread_log("A automação atualizará somente Status/Situação e OBS./Motivo.")

            options = BotOptions(
                username=self.username_var.get().strip(),
                password=self.password_var.get().strip(),
                headless=bool(self.headless_var.get()),
                max_tabs=min(5, max(1, int(self.max_tabs_var.get()))),
                slow_mo_ms=0,
                timeout_ms=45000,
                fast_mode=bool(self.fast_mode_var.get()),
                settle_ms=int(self.settle_ms_var.get()),
            )

            results = asyncio.run(
                run_consultas(
                    tasks=plan.tasks,
                    dt_from=plan.dt_from,
                    options=options,
                    log=self._thread_log,
                    progress=self._thread_progress,
                )
            )

            self._thread_log("Gravando resultados na planilha original...")
            save_results(plan, results)

            self.event_queue.put(("done", "Automação finalizada com sucesso."))

        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def _thread_log(self, message: str) -> None:
        self.event_queue.put(("log", message))

    def _thread_progress(self, done: int, total: int) -> None:
        self.event_queue.put(("progress", done, total))

    def _process_queue(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()

                if not event:
                    continue

                kind = event[0]

                if kind == "log":
                    self._append_log(event[1])

                elif kind == "progress":
                    done, total = event[1], event[2]
                    self._set_progress(done, total)

                elif kind == "done":
                    self._append_log(event[1])
                    self.progress_label.configure(text="Finalizado")
                    self._finish_running()
                    self._save_settings()
                    messagebox.showinfo("Concluído", event[1])

                elif kind == "error":
                    self._append_log(f"Erro inesperado: {event[1]}")
                    self.progress_label.configure(text="Erro")
                    self._finish_running()
                    self._save_settings()
                    messagebox.showerror("Erro", event[1])

        except queue.Empty:
            pass

        self.root.after(100, self._process_queue)

    def _set_progress(self, done: int, total: int) -> None:
        total = max(total, 1)
        self.progressbar["maximum"] = total
        self.progressbar["value"] = done
        self.counter_label.configure(text=f"{done}/{total}")

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _finish_running(self) -> None:
        self.is_running = False

        if self.start_button is not None:
            self.start_button.configure(state="normal")

    def _on_close(self) -> None:
        self._save_settings()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    TranspofreteApp(root)
    root.mainloop()