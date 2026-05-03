"""
Excel → PDF 일괄 변환기 메인 GUI (tkinter + ttk)
tkinterdnd2 가 설치되어 있으면 드래그 앤 드롭을 활성화한다.
"""

import datetime
import os
import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List

from app.converter import (
    ConversionOptions,
    ConversionProgress,
    Converter,
    OverwritePolicy,
    SheetScope,
    collect_excel_files,
)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


# ---------------------------------------------------------------------------
# 색상 / 스타일 상수
# ---------------------------------------------------------------------------
CLR_BG = "#F5F5F5"
CLR_FRAME = "#FFFFFF"
CLR_ACCENT = "#2E86AB"
CLR_SUCCESS = "#27AE60"
CLR_FAIL = "#E74C3C"
CLR_WARN = "#F39C12"
CLR_TEXT = "#2C3E50"
CLR_SUBTEXT = "#7F8C8D"
FONT_BODY = ("Malgun Gothic", 9)
FONT_BOLD = ("Malgun Gothic", 9, "bold")
FONT_TITLE = ("Malgun Gothic", 13, "bold")
FONT_LOG = ("Consolas", 8)


class App:
    """메인 애플리케이션 클래스."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self._files: List[str] = []
        self._converter: Converter | None = None
        self._worker_thread: threading.Thread | None = None
        self._queue: queue.Queue = queue.Queue()
        self._start_time: float = 0.0
        self._last_results: list = []

        self._setup_root()
        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # 초기 설정
    # ------------------------------------------------------------------

    def _setup_root(self):
        self.root.title("Excel → PDF 일괄 변환기")
        self.root.geometry("820x680")
        self.root.minsize(700, 580)
        self.root.configure(bg=CLR_BG)
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=CLR_BG)
        style.configure("Card.TFrame", background=CLR_FRAME, relief="flat")
        style.configure("TLabel", background=CLR_BG, foreground=CLR_TEXT, font=FONT_BODY)
        style.configure("Card.TLabel", background=CLR_FRAME, foreground=CLR_TEXT, font=FONT_BODY)
        style.configure("Title.TLabel", background=CLR_BG, foreground=CLR_TEXT, font=FONT_TITLE)
        style.configure(
            "Accent.TButton",
            font=FONT_BOLD,
            foreground="white",
            background=CLR_ACCENT,
            padding=(12, 6),
        )
        style.map("Accent.TButton", background=[("active", "#1A6B8A"), ("disabled", "#BDC3C7")])
        style.configure("TButton", font=FONT_BODY, padding=(8, 5))
        style.configure(
            "TProgressbar",
            troughcolor="#DDE1E7",
            background=CLR_ACCENT,
            thickness=16,
        )
        style.configure("TCheckbutton", background=CLR_FRAME, foreground=CLR_TEXT, font=FONT_BODY)
        style.configure("TRadiobutton", background=CLR_FRAME, foreground=CLR_TEXT, font=FONT_BODY)
        style.configure("TLabelframe", background=CLR_FRAME)
        style.configure("TLabelframe.Label", background=CLR_FRAME, foreground=CLR_TEXT, font=FONT_BOLD)

    # ------------------------------------------------------------------
    # UI 구성
    # ------------------------------------------------------------------

    def _build_ui(self):
        title_bar = ttk.Frame(self.root)
        title_bar.pack(fill="x", padx=16, pady=(12, 4))
        ttk.Label(title_bar, text="Excel → PDF 일괄 변환기", style="Title.TLabel").pack(side="left")

        content = ttk.Frame(self.root)
        content.pack(fill="both", expand=True, padx=16, pady=4)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        self._build_file_panel(content)
        self._build_option_panel(content)

        bottom = ttk.Frame(self.root)
        bottom.pack(fill="both", expand=False, padx=16, pady=(0, 12))
        self._build_progress_panel(bottom)
        self._build_log_panel(bottom)
        self._build_action_bar()

    # ── 파일 목록 패널 ──────────────────────────────────────────────

    def _build_file_panel(self, parent):
        frame = ttk.LabelFrame(parent, text=" 변환 파일 목록 ", style="TLabelframe")
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=2)
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        folder_row = ttk.Frame(frame, style="Card.TFrame")
        folder_row.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        folder_row.columnconfigure(1, weight=1)

        ttk.Label(folder_row, text="폴더:", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 4))
        self._folder_var = tk.StringVar()
        folder_entry = ttk.Entry(folder_row, textvariable=self._folder_var, state="readonly", font=FONT_BODY)
        folder_entry.grid(row=0, column=1, sticky="ew", padx=4)
        self._folder_btn = ttk.Button(folder_row, text="폴더 선택…", command=self._browse_folder)
        self._folder_btn.grid(row=0, column=2, padx=(4, 0))

        self._recursive_var = tk.BooleanVar(value=True)
        self._recursive_chk = ttk.Checkbutton(
            folder_row,
            text="하위 폴더 포함",
            variable=self._recursive_var,
            style="TCheckbutton",
            command=self._on_recursive_changed,
        )
        self._recursive_chk.grid(row=0, column=3, padx=(8, 0))

        list_frame = ttk.Frame(frame, style="Card.TFrame")
        list_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        self._listbox = tk.Listbox(
            list_frame,
            selectmode="extended",
            font=FONT_BODY,
            bg=CLR_FRAME,
            fg=CLR_TEXT,
            selectbackground=CLR_ACCENT,
            selectforeground="white",
            borderwidth=1,
            relief="solid",
            activestyle="none",
        )
        self._listbox.grid(row=0, column=0, sticky="nsew")

        sb_y = ttk.Scrollbar(list_frame, orient="vertical", command=self._listbox.yview)
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self._listbox.xview)
        sb_x.grid(row=1, column=0, sticky="ew")
        self._listbox.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        if DND_AVAILABLE:
            self._listbox.drop_target_register(DND_FILES)
            self._listbox.dnd_bind("<<Drop>>", self._on_drop)
            ttk.Label(
                frame,
                text="파일/폴더를 여기로 드래그하거나 '파일 추가' 버튼을 이용하세요.",
                style="Card.TLabel",
                foreground=CLR_SUBTEXT,
            ).grid(row=2, column=0, sticky="w", padx=8, pady=(0, 2))
        else:
            ttk.Label(
                frame,
                text="'파일 추가' 또는 '폴더 선택' 버튼으로 파일을 추가하세요.",
                style="Card.TLabel",
                foreground=CLR_SUBTEXT,
            ).grid(row=2, column=0, sticky="w", padx=8, pady=(0, 2))

        btn_row = ttk.Frame(frame, style="Card.TFrame")
        btn_row.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 8))

        self._add_btn = ttk.Button(btn_row, text="+ 파일 추가", command=self._add_files)
        self._add_btn.pack(side="left", padx=(0, 4))
        self._remove_btn = ttk.Button(btn_row, text="선택 제거", command=self._remove_selected)
        self._remove_btn.pack(side="left", padx=4)
        self._clear_btn = ttk.Button(btn_row, text="전체 지우기", command=self._clear_files)
        self._clear_btn.pack(side="left", padx=4)

        self._file_count_label = ttk.Label(btn_row, text="0개 파일", style="Card.TLabel", foreground=CLR_SUBTEXT)
        self._file_count_label.pack(side="right")

    # ── 옵션 패널 ────────────────────────────────────────────────────

    def _build_option_panel(self, parent):
        frame = ttk.LabelFrame(parent, text=" 변환 옵션 ", style="TLabelframe")
        frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=2)
        frame.columnconfigure(0, weight=1)

        pad = {"padx": 12, "pady": 4, "sticky": "w"}

        # 시트 범위
        ttk.Label(frame, text="시트 범위", style="Card.TLabel", font=FONT_BOLD).grid(row=0, column=0, **pad, pady=(12, 2))
        self._sheet_scope = tk.StringVar(value=SheetScope.ALL.value)
        ttk.Radiobutton(
            frame, text="모든 시트 변환", variable=self._sheet_scope,
            value=SheetScope.ALL.value, style="TRadiobutton",
        ).grid(row=1, column=0, **pad, pady=1)
        ttk.Radiobutton(
            frame, text="활성 시트만 변환", variable=self._sheet_scope,
            value=SheetScope.ACTIVE.value, style="TRadiobutton",
        ).grid(row=2, column=0, **pad, pady=1)

        ttk.Separator(frame, orient="horizontal").grid(row=3, column=0, sticky="ew", padx=8, pady=8)

        # 인쇄 옵션
        ttk.Label(frame, text="인쇄 옵션", style="Card.TLabel", font=FONT_BOLD).grid(row=4, column=0, **pad, pady=(0, 2))
        self._fit_to_page = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame, text="열 너비를 한 페이지에\n맞추기 (FitToPage)",
            variable=self._fit_to_page, style="TCheckbutton",
        ).grid(row=5, column=0, **pad, pady=1)

        self._skip_hidden = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame, text="숨겨진 시트 건너뛰기",
            variable=self._skip_hidden, style="TCheckbutton",
        ).grid(row=6, column=0, **pad, pady=1)

        ttk.Separator(frame, orient="horizontal").grid(row=7, column=0, sticky="ew", padx=8, pady=8)

        # PDF 저장 위치
        ttk.Label(frame, text="PDF 저장 위치", style="Card.TLabel", font=FONT_BOLD).grid(row=8, column=0, **pad, pady=(0, 2))
        self._save_same = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame, text="원본 파일과 같은 폴더",
            variable=self._save_same, style="TCheckbutton",
            command=self._on_save_same_changed,
        ).grid(row=9, column=0, **pad, pady=1)

        out_row = ttk.Frame(frame, style="Card.TFrame")
        out_row.grid(row=10, column=0, sticky="ew", padx=12, pady=(0, 4))
        out_row.columnconfigure(0, weight=1)

        self._out_var = tk.StringVar()
        self._out_entry = ttk.Entry(out_row, textvariable=self._out_var, state="disabled", font=FONT_BODY)
        self._out_entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self._out_btn = ttk.Button(out_row, text="…", width=3, command=self._browse_output, state="disabled")
        self._out_btn.grid(row=0, column=1)

        ttk.Separator(frame, orient="horizontal").grid(row=11, column=0, sticky="ew", padx=8, pady=8)

        # 기존 PDF 처리
        ttk.Label(frame, text="기존 PDF 처리", style="Card.TLabel", font=FONT_BOLD).grid(row=12, column=0, **pad, pady=(0, 2))
        self._overwrite_policy = tk.StringVar(value=OverwritePolicy.OVERWRITE.value)
        ttk.Radiobutton(
            frame, text="덮어쓰기", variable=self._overwrite_policy,
            value=OverwritePolicy.OVERWRITE.value, style="TRadiobutton",
        ).grid(row=13, column=0, **pad, pady=1)
        ttk.Radiobutton(
            frame, text="건너뛰기", variable=self._overwrite_policy,
            value=OverwritePolicy.SKIP.value, style="TRadiobutton",
        ).grid(row=14, column=0, **pad, pady=1)
        ttk.Radiobutton(
            frame, text="이름 뒤에 번호 붙이기", variable=self._overwrite_policy,
            value=OverwritePolicy.RENAME.value, style="TRadiobutton",
        ).grid(row=15, column=0, **pad, pady=1)

        frame.rowconfigure(16, weight=1)

    # ── 진행 상태 패널 ───────────────────────────────────────────────

    def _build_progress_panel(self, parent):
        prog_frame = ttk.Frame(parent)
        prog_frame.pack(fill="x", pady=(6, 2))

        self._status_label = ttk.Label(prog_frame, text="대기 중", foreground=CLR_SUBTEXT, font=FONT_BODY)
        self._status_label.pack(side="left")
        self._percent_label = ttk.Label(prog_frame, text="", foreground=CLR_ACCENT, font=FONT_BOLD)
        self._percent_label.pack(side="right")

        self._progress_var = tk.DoubleVar()
        self._progress_bar = ttk.Progressbar(
            parent, variable=self._progress_var, maximum=100, style="TProgressbar"
        )
        self._progress_bar.pack(fill="x", pady=(0, 4))

    # ── 로그 패널 ────────────────────────────────────────────────────

    def _build_log_panel(self, parent):
        log_frame = ttk.LabelFrame(parent, text=" 변환 로그 ", style="TLabelframe")
        log_frame.pack(fill="both", expand=True)

        inner = ttk.Frame(log_frame)
        inner.pack(fill="both", expand=True, padx=4, pady=4)
        inner.rowconfigure(0, weight=1)
        inner.columnconfigure(0, weight=1)

        self._log = tk.Text(
            inner,
            height=7,
            state="disabled",
            font=FONT_LOG,
            bg="#1E1E1E",
            fg="#DCDCDC",
            insertbackground="white",
            wrap="none",
            borderwidth=0,
        )
        self._log.grid(row=0, column=0, sticky="nsew")
        self._log.tag_configure("ok", foreground="#4EC9B0")
        self._log.tag_configure("skip", foreground="#DCDCAA")
        self._log.tag_configure("fail", foreground="#F44747")
        self._log.tag_configure("info", foreground="#9CDCFE")
        self._log.tag_configure("warn", foreground="#CE9178")

        sb_y = ttk.Scrollbar(inner, orient="vertical", command=self._log.yview)
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x = ttk.Scrollbar(inner, orient="horizontal", command=self._log.xview)
        sb_x.grid(row=1, column=0, sticky="ew")
        self._log.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        log_btn_row = ttk.Frame(log_frame)
        log_btn_row.pack(fill="x", padx=4, pady=(0, 4))
        ttk.Button(log_btn_row, text="로그 저장…", command=self._save_log).pack(side="right", padx=(4, 0))
        ttk.Button(log_btn_row, text="로그 지우기", command=self._clear_log).pack(side="right")

    # ── 액션 버튼 바 ─────────────────────────────────────────────────

    def _build_action_bar(self):
        bar = ttk.Frame(self.root)
        bar.pack(fill="x", padx=16, pady=(0, 12))

        self._start_btn = ttk.Button(
            bar, text="▶  변환 시작", style="Accent.TButton", command=self._start_conversion
        )
        self._start_btn.pack(side="left")

        self._stop_btn = ttk.Button(
            bar, text="■  중지", command=self._stop_conversion, state="disabled"
        )
        self._stop_btn.pack(side="left", padx=(8, 0))

        self._open_folder_btn = ttk.Button(
            bar, text="📂 출력 폴더 열기", command=self._open_output_folder, state="disabled"
        )
        self._open_folder_btn.pack(side="left", padx=(8, 0))

        self._summary_label = ttk.Label(bar, text="", font=FONT_BODY, foreground=CLR_SUBTEXT)
        self._summary_label.pack(side="right")

    # ------------------------------------------------------------------
    # 이벤트 핸들러
    # ------------------------------------------------------------------

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="엑셀 파일이 있는 폴더를 선택하세요")
        if not folder:
            return
        self._folder_var.set(folder)
        self._load_files_from_folder(folder)

    def _load_files_from_folder(self, folder: str):
        files = collect_excel_files(folder, recursive=self._recursive_var.get())
        self._set_files(files)
        self._log_info(f"폴더 스캔 완료: {len(files)}개 파일 발견 ({folder})")

    def _on_recursive_changed(self):
        folder = self._folder_var.get()
        if folder:
            self._load_files_from_folder(folder)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="엑셀 파일 선택",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All Files", "*.*"),
            ],
        )
        if not paths:
            return
        new = [p for p in paths if p not in self._files]
        self._files.extend(new)
        self._refresh_listbox()
        self._log_info(f"{len(new)}개 파일 추가됨.")

    def _remove_selected(self):
        selected = list(self._listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            del self._files[idx]
        self._refresh_listbox()

    def _clear_files(self):
        self._files.clear()
        self._refresh_listbox()

    def _on_drop(self, event):
        raw = event.data
        paths = self.root.tk.splitlist(raw)
        added = 0
        for p in paths:
            p = p.strip()
            if os.path.isdir(p):
                files = collect_excel_files(p, recursive=self._recursive_var.get())
                new = [f for f in files if f not in self._files]
                self._files.extend(new)
                added += len(new)
            elif os.path.isfile(p) and p not in self._files:
                if Path(p).suffix.lower() in {".xls", ".xlsx", ".xlsm", ".xlsb"}:
                    self._files.append(p)
                    added += 1
        self._refresh_listbox()
        if added:
            self._log_info(f"드래그 앤 드롭: {added}개 파일 추가됨.")

    def _browse_output(self):
        folder = filedialog.askdirectory(title="PDF 저장 폴더 선택")
        if folder:
            self._out_var.set(folder)

    def _on_save_same_changed(self):
        same = self._save_same.get()
        state = "disabled" if same else "normal"
        self._out_entry.configure(state=state)
        self._out_btn.configure(state=state)

    def _open_output_folder(self):
        """완료된 PDF 가 저장된 폴더를 탐색기로 연다."""
        if self._save_same.get():
            # 파일마다 폴더가 다를 수 있으므로 첫 번째 성공 파일 기준
            for r in (self._last_results or []):
                if r.success:
                    folder = str(Path(r.path).parent)
                    break
            else:
                folder = self._folder_var.get() or os.getcwd()
        else:
            folder = self._out_var.get() or os.getcwd()
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("폴더 열기 실패", str(e))

    # ------------------------------------------------------------------
    # 변환 제어
    # ------------------------------------------------------------------

    def _start_conversion(self):
        if not self._files:
            messagebox.showwarning("파일 없음", "변환할 엑셀 파일을 먼저 추가해 주세요.")
            return

        if not self._save_same.get() and not self._out_var.get():
            messagebox.showwarning("저장 폴더 없음", "PDF 를 저장할 폴더를 선택해 주세요.")
            return

        options = ConversionOptions(
            sheet_scope=SheetScope(self._sheet_scope.get()),
            fit_to_page=self._fit_to_page.get(),
            include_subfolders=self._recursive_var.get(),
            output_same_folder=self._save_same.get(),
            output_folder=self._out_var.get(),
            skip_hidden_sheets=self._skip_hidden.get(),
            overwrite_policy=OverwritePolicy(self._overwrite_policy.get()),
        )

        self._converter = Converter(
            files=list(self._files),
            options=options,
            on_progress=self._on_progress,
            on_finished=self._on_finished,
        )

        self._set_converting(True)
        self._progress_var.set(0)
        self._summary_label.configure(text="")
        self._open_folder_btn.configure(state="disabled")
        self._last_results = []
        self._start_time = time.monotonic()
        self._log_info(f"변환 시작: {len(self._files)}개 파일")

        self._worker_thread = threading.Thread(target=self._converter.run, daemon=True)
        self._worker_thread.start()

    def _stop_conversion(self):
        if self._converter:
            self._converter.stop()
            self._log_warn("사용자에 의해 변환이 중지되었습니다.")
        self._stop_btn.configure(state="disabled")

    def _set_converting(self, converting: bool):
        """변환 중에는 파일 목록 조작 UI 를 비활성화한다."""
        state = "disabled" if converting else "normal"
        for widget in (
            self._start_btn,
            self._add_btn,
            self._remove_btn,
            self._clear_btn,
            self._folder_btn,
            self._recursive_chk,
        ):
            widget.configure(state=state)
        self._stop_btn.configure(state="normal" if converting else "disabled")

    # ------------------------------------------------------------------
    # 워커 → UI 큐
    # ------------------------------------------------------------------

    def _on_progress(self, progress: ConversionProgress):
        # Converter 가 이미 copy.copy 한 스냅샷을 전달한다
        self._queue.put(("progress", progress))

    def _on_finished(self, progress: ConversionProgress):
        self._queue.put(("finished", progress))

    def _poll_queue(self):
        try:
            while True:
                event, data = self._queue.get_nowait()
                if event == "progress":
                    self._handle_progress(data)
                elif event == "finished":
                    self._handle_finished(data)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _handle_progress(self, progress: ConversionProgress):
        self._progress_var.set(progress.percent)
        self._percent_label.configure(text=f"{progress.percent:.0f}%")
        self._status_label.configure(text=f"변환 중: {os.path.basename(progress.current_file)}")

        if progress.results:
            last = progress.results[-1]
            name = os.path.basename(last.path)
            if last.skipped:
                self._log_skip(f"[건너뜀] {name}  — {last.message}")
            elif last.success:
                self._log_ok(f"[완료] {name}  ({last.duration:.1f}s)")
            else:
                self._log_fail(f"[실패] {name}  → {last.message}")

    def _handle_finished(self, progress: ConversionProgress):
        self._last_results = progress.results
        self._set_converting(False)
        self._progress_var.set(progress.percent)
        self._percent_label.configure(text=f"{progress.percent:.0f}%")
        self._open_folder_btn.configure(state="normal")

        elapsed = time.monotonic() - self._start_time
        ok = sum(1 for r in progress.results if r.success and not r.skipped)
        skipped = sum(1 for r in progress.results if r.skipped)
        fail = sum(1 for r in progress.results if not r.success)

        parts = [f"성공 {ok}개"]
        if skipped:
            parts.append(f"건너뜀 {skipped}개")
        if fail:
            parts.append(f"실패 {fail}개")
        summary = "완료: " + " / ".join(parts) + f"  ({elapsed:.1f}s)"

        self._status_label.configure(text=summary)
        self._summary_label.configure(text=summary)
        self._log_info(f"─── {summary} ───")

        if fail > 0:
            messagebox.showwarning(
                "변환 완료 (일부 실패)",
                f"총 {progress.total}개 파일 중\n"
                f"성공: {ok}개 / 건너뜀: {skipped}개 / 실패: {fail}개\n\n"
                "실패 파일 목록은 로그를 확인하세요.",
            )
        else:
            messagebox.showinfo(
                "변환 완료",
                f"변환이 완료되었습니다.\n성공: {ok}개"
                + (f" / 건너뜀: {skipped}개" if skipped else "")
                + f"\n총 소요 시간: {elapsed:.1f}초",
            )

    # ------------------------------------------------------------------
    # 로그 유틸
    # ------------------------------------------------------------------

    def _log_write(self, text: str, tag: str):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log.configure(state="normal")
        self._log.insert("end", f"[{ts}] {text}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_ok(self, msg: str):
        self._log_write(msg, "ok")

    def _log_skip(self, msg: str):
        self._log_write(msg, "skip")

    def _log_fail(self, msg: str):
        self._log_write(msg, "fail")

    def _log_info(self, msg: str):
        self._log_write(msg, "info")

    def _log_warn(self, msg: str):
        self._log_write(msg, "warn")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _save_log(self):
        path = filedialog.asksaveasfilename(
            title="로그 저장",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            initialfile=f"excel2pdf_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if not path:
            return
        content = self._log.get("1.0", "end")
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self._log_info(f"로그 저장됨: {path}")
        except Exception as e:
            messagebox.showerror("저장 실패", str(e))

    # ------------------------------------------------------------------
    # 파일 목록 유틸
    # ------------------------------------------------------------------

    def _set_files(self, files: List[str]):
        self._files = files
        self._refresh_listbox()

    def _refresh_listbox(self):
        self._listbox.delete(0, "end")
        for f in self._files:
            self._listbox.insert("end", f)
        self._file_count_label.configure(text=f"{len(self._files)}개 파일")

    # ------------------------------------------------------------------
    # 종료
    # ------------------------------------------------------------------

    def _on_close(self):
        if self._worker_thread and self._worker_thread.is_alive():
            if not messagebox.askyesno("종료 확인", "변환이 진행 중입니다.\n정말 종료하시겠습니까?"):
                return
            if self._converter:
                self._converter.stop()
        self.root.destroy()
