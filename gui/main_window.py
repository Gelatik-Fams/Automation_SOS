import os
import queue
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import customtkinter as ctk

from gui.config import load_config, save_config, get_version, BASE_DIR
from gui.logger import logger
from gui.utils import validate_workbook, user_friendly_error, fmt_elapsed, detect_clusters, detect_cluster_dirs
from gui.widgets import (
    C_BLUE_DARK, C_BLUE_MED, C_BLUE_LIGHT, C_GREY_BG,
    C_GREEN, C_RED, C_WHITE, C_TEXT_DARK, C_TEXT_MUTED,
    SectionCard, ClusterRow,
)


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        self._cfg = load_config()
        version = get_version()

        self.title(f'Gelatik Automation  v{version}')
        w = self._cfg.get('window_width', 860)
        h = self._cfg.get('window_height', 620)
        self.geometry(f'{w}x{h}')
        self.minsize(760, 560)
        self.configure(fg_color=C_GREY_BG)

        self._folder: str = ''
        self._clusters: list[str] = []
        self._log_queue: queue.Queue = queue.Queue()
        self._engine_thread: threading.Thread | None = None
        self._running = False
        self._start_time: float = 0.0
        self._cluster_results: dict[str, tuple[bool, str]] = {}

        self._build_ui()
        self._poll_log()

        logger.info(f'Application started — v{version}')

        last = self._cfg.get('last_folder', '')
        if last and os.path.isdir(last):
            self._apply_folder(last)

        self.bind('<Configure>', self._on_resize)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        version = get_version()

        topbar = ctk.CTkFrame(self, fg_color=C_BLUE_DARK, height=52, corner_radius=0)
        topbar.pack(fill='x', side='top')
        topbar.pack_propagate(False)

        ctk.CTkLabel(
            topbar,
            text='Gelatik Automation',
            font=ctk.CTkFont(family='Helvetica', size=15, weight='bold'),
            text_color=C_WHITE,
        ).pack(side='left', padx=16, pady=10)

        ctk.CTkLabel(
            topbar,
            text=f'v{version}',
            font=ctk.CTkFont(size=10),
            text_color='#6E90AE',
        ).pack(side='left', padx=(0, 8), pady=10)

        ctk.CTkButton(
            topbar,
            text='?',
            width=28, height=28,
            fg_color='transparent',
            hover_color='#2E6DA4',
            text_color='#A3C2E3',
            font=ctk.CTkFont(size=14),
            command=self._show_about,
        ).pack(side='right', padx=12)

        body = ctk.CTkFrame(self, fg_color='transparent')
        body.pack(fill='both', expand=True, padx=12, pady=10)

        left = ctk.CTkFrame(body, fg_color='transparent', width=320)
        left.pack(side='left', fill='y', padx=(0, 8))
        left.pack_propagate(False)

        right = ctk.CTkFrame(body, fg_color='transparent')
        right.pack(side='left', fill='both', expand=True)

        self._build_left(left)
        self._build_right(right)

    def _build_left(self, parent):
        folder_card = SectionCard(parent, 'Folder Automation')
        folder_card.pack(fill='x', pady=(0, 8))

        self._folder_label = ctk.CTkLabel(
            folder_card.content,
            text='Belum dipilih',
            font=ctk.CTkFont(size=11),
            text_color=C_TEXT_MUTED,
            wraplength=260, justify='left', anchor='w',
        )
        self._folder_label.pack(fill='x', pady=(0, 6))

        ctk.CTkButton(
            folder_card.content,
            text='Browse Folder…',
            command=self._browse_folder,
            fg_color=C_BLUE_MED, hover_color=C_BLUE_DARK,
            font=ctk.CTkFont(size=12), height=32,
        ).pack(fill='x')

        cluster_card = SectionCard(parent, 'Cluster Terdeteksi')
        cluster_card.pack(fill='x', pady=(0, 8))
        self._cluster_frame = cluster_card.content

        ctk.CTkLabel(
            self._cluster_frame,
            text='Pilih folder terlebih dahulu',
            font=ctk.CTkFont(size=11),
            text_color=C_TEXT_MUTED,
        ).pack(pady=4)

        output_card = SectionCard(parent, 'File Output')
        output_card.pack(fill='x', pady=(0, 8))
        self._output_label = ctk.CTkLabel(
            output_card.content,
            text='—',
            font=ctk.CTkFont(size=11),
            text_color=C_TEXT_MUTED,
            wraplength=260, justify='left', anchor='w',
        )
        self._output_label.pack(fill='x')

        self._status_var = ctk.StringVar(value='Menunggu')
        self._status_label = ctk.CTkLabel(
            parent,
            textvariable=self._status_var,
            font=ctk.CTkFont(size=11, weight='bold'),
            text_color=C_TEXT_MUTED,
        )
        self._status_label.pack(anchor='w', pady=(4, 2))

        self._progress = ctk.CTkProgressBar(
            parent, mode='determinate',
            fg_color=C_BLUE_LIGHT, progress_color=C_BLUE_MED,
        )
        self._progress.set(0)
        self._progress.pack(fill='x', pady=(0, 10))

        self._gen_btn = ctk.CTkButton(
            parent,
            text='Generate Summary',
            command=self._start_engine,
            fg_color=C_BLUE_DARK, hover_color=C_BLUE_MED,
            font=ctk.CTkFont(family='Helvetica', size=14, weight='bold'),
            height=44, state='disabled',
        )
        self._gen_btn.pack(fill='x')

    def _build_right(self, parent):
        log_card = SectionCard(parent, 'Log Proses')
        log_card.pack(fill='both', expand=True)
        self._log_box = ctk.CTkTextbox(
            log_card.content,
            font=ctk.CTkFont(family='Courier', size=11),
            fg_color='#0D1117', text_color='#C9D1D9',
            corner_radius=4, wrap='word', state='disabled',
        )
        self._log_box.pack(fill='both', expand=True)

    # ------------------------------------------------------------------
    # Folder browsing
    # ------------------------------------------------------------------

    def _browse_folder(self):
        from tkinter import filedialog
        folder = filedialog.askdirectory(title='Pilih Folder Automation')
        if folder:
            self._apply_folder(folder)

    def _apply_folder(self, folder: str):
        self._folder = folder
        self._folder_label.configure(text=folder, text_color=C_TEXT_DARK)
        self._clusters = detect_clusters(folder)
        self._refresh_clusters()
        self._refresh_output_label()
        self._gen_btn.configure(state='normal' if self._clusters else 'disabled')

        self._cfg['last_folder'] = folder
        save_config(self._cfg)
        logger.info(f'Folder dipilih: {folder} — cluster terdeteksi: {self._clusters}')

    def _refresh_clusters(self):
        for w in self._cluster_frame.winfo_children():
            w.destroy()
        if not self._clusters:
            ctk.CTkLabel(
                self._cluster_frame,
                text='Tidak ada cluster CSV ditemukan',
                font=ctk.CTkFont(size=11), text_color=C_RED,
            ).pack(pady=4)
            return
        for name in self._clusters:
            ClusterRow(self._cluster_frame, name).pack(fill='x', pady=2)

    def _refresh_output_label(self):
        if not self._folder or not self._clusters:
            self._output_label.configure(text='—')
            return
        files = [f'Summary SOS_{c}.xlsx' for c in self._clusters]
        self._output_label.configure(text='\n'.join(files), text_color=C_TEXT_DARK)

    # ------------------------------------------------------------------
    # Engine runner
    # ------------------------------------------------------------------

    def _start_engine(self):
        if self._running:
            return
        self._running = True
        self._cluster_results = {}
        self._start_time = time.time()
        self._gen_btn.configure(state='disabled', text='Sedang Berjalan…')
        self._progress.set(0)
        self._set_status('Memulai…', C_BLUE_MED)
        self._log_append('[INFO] Memulai proses generate…\n')
        logger.info(f'Generate dimulai — folder: {self._folder}')
        self._engine_thread = threading.Thread(target=self._run_engine_worker, daemon=True)
        self._engine_thread.start()

    def _run_engine_worker(self):
        sys.path.insert(0, str(BASE_DIR))
        import tes  # noqa: engine import

        orig_dir = os.getcwd()
        try:
            dirs_to_process = detect_cluster_dirs(self._folder)

            if not dirs_to_process:
                self._q_log('[ERROR] Tidak ada CSV ditemukan di folder yang dipilih')
                self._q_done(success=False)
                return

            total = len(dirs_to_process)

            for i, (cluster_dir, cluster_name) in enumerate(dirs_to_process):
                self._q_log(f'[INFO] Memproses cluster: {cluster_name}')
                logger.info(f'Cluster {cluster_name} — mulai')
                base_prog = i / total
                step = 1 / total
                ok = False
                err_msg = ''

                try:
                    os.chdir(cluster_dir)

                    summary_path = tes.get_summary_output_path('.', cluster_name=cluster_name, extension='.xlsx')
                    targets_local, target_rows = tes.load_or_create_summary_targets(summary_path)
                    self._q_progress(base_prog + step * 0.1)

                    df_raw = tes.baca_semua_csv()
                    if df_raw is None or df_raw.empty:
                        raise ValueError(f'Tidak ada data CSV di {cluster_name}')
                    csv_count = len(df_raw)
                    self._q_log(f'[INFO] {csv_count:,} baris CSV')
                    logger.info(f'Cluster {cluster_name} — {csv_count:,} baris CSV')
                    self._q_progress(base_prog + step * 0.30)

                    df, removed = tes.validasi_data(df_raw)
                    self._q_log(f'[INFO] Valid: {len(df):,} baris ({len(removed)} duplikat dihapus)')
                    self._q_progress(base_prog + step * 0.50)

                    target_rows = tes._enrich_targets_with_df_values(df, target_rows)
                    if target_rows is not None:
                        targets_local = tes._target_rows_to_dict(target_rows)
                    self._q_progress(base_prog + step * 0.65)

                    self._q_log(f'[INFO] Membuat Summary SOS_{cluster_name}.xlsx…')
                    output_path = tes.export_summary_excel(
                        df, targets_local, output_dir='.', cluster_name=cluster_name,
                        target_rows=target_rows,
                    )
                    self._q_progress(base_prog + step * 0.90)

                    validation_err = validate_workbook(output_path)
                    if validation_err:
                        try:
                            os.remove(output_path)
                        except OSError:
                            pass
                        raise ValueError(f'Validasi gagal: {validation_err}')

                    size_kb = os.path.getsize(output_path) / 1024
                    self._q_log(f'[OK] {os.path.basename(output_path)} ({size_kb:.0f} KB)')
                    logger.info(f'Cluster {cluster_name} — berhasil: {output_path} ({size_kb:.0f} KB)')
                    ok = True

                except Exception as exc:
                    friendly = user_friendly_error(exc)
                    err_msg = friendly
                    self._q_log(f'[ERROR] {friendly}')
                    logger.error(f'Cluster {cluster_name} — {exc}')
                    logger.debug(traceback.format_exc())

                finally:
                    os.chdir(orig_dir)

                self._cluster_results[cluster_name] = (ok, err_msg)
                self._q_progress(base_prog + step)

            elapsed = time.time() - self._start_time
            logger.info(f'Generate selesai — {fmt_elapsed(elapsed)} — hasil: {self._cluster_results}')
            self._q_progress(1.0)
            any_ok = any(ok for ok, _ in self._cluster_results.values())
            self._q_done(success=any_ok)

        except Exception as exc:
            logger.error(f'Unexpected error: {exc}')
            logger.debug(traceback.format_exc())
            self._q_log(f'[ERROR] {user_friendly_error(exc)}')
            self._q_done(success=False)
        finally:
            os.chdir(orig_dir)

    def _q_log(self, msg: str):
        self._log_queue.put(('log', msg))

    def _q_progress(self, value: float):
        self._log_queue.put(('progress', min(1.0, max(0.0, value))))

    def _q_done(self, success: bool):
        self._log_queue.put(('done', success))

    # ------------------------------------------------------------------
    # Log polling (main thread)
    # ------------------------------------------------------------------

    def _poll_log(self):
        try:
            while True:
                kind, payload = self._log_queue.get_nowait()
                if kind == 'log':
                    self._log_append(payload + '\n')
                elif kind == 'progress':
                    self._progress.set(payload)
                elif kind == 'done':
                    self._on_engine_done(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_log)

    def _log_append(self, text: str):
        self._log_box.configure(state='normal')
        self._log_box.insert('end', text)
        self._log_box.see('end')
        self._log_box.configure(state='disabled')

    # ------------------------------------------------------------------
    # Engine completion
    # ------------------------------------------------------------------

    def _on_engine_done(self, success: bool):
        self._running = False
        self._gen_btn.configure(state='normal', text='Generate Summary')
        elapsed = time.time() - self._start_time
        if success:
            self._set_status('Selesai', C_GREEN)
            self._show_done_dialog(elapsed)
        else:
            self._set_status('Gagal', C_RED)

    def _set_status(self, text: str, color: str):
        self._status_var.set(text)
        self._status_label.configure(text_color=color)

    # ------------------------------------------------------------------
    # Success dialog
    # ------------------------------------------------------------------

    def _show_done_dialog(self, elapsed: float):
        dialog = ctk.CTkToplevel(self)
        dialog.title('Generate Complete')
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.configure(fg_color=C_WHITE)

        n_rows = max(1, len(self._cluster_results))
        h = 180 + n_rows * 32
        w = 400
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        dialog.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')

        ctk.CTkLabel(
            dialog,
            text='Generation Complete',
            font=ctk.CTkFont(family='Helvetica', size=16, weight='bold'),
            text_color=C_BLUE_DARK,
        ).pack(pady=(18, 10))

        # Per-cluster result grid
        grid = ctk.CTkFrame(dialog, fg_color='transparent')
        grid.pack(fill='x', padx=24)

        for cluster, (ok, msg) in self._cluster_results.items():
            row = ctk.CTkFrame(grid, fg_color='transparent')
            row.pack(fill='x', pady=3)
            ctk.CTkLabel(
                row,
                text=cluster,
                font=ctk.CTkFont(size=12, weight='bold'),
                text_color=C_TEXT_DARK,
                anchor='w', width=200,
            ).pack(side='left')
            if ok:
                ctk.CTkLabel(row, text='✓ Success', font=ctk.CTkFont(size=12),
                             text_color=C_GREEN).pack(side='right')
            else:
                ctk.CTkLabel(row, text='✗ Failed', font=ctk.CTkFont(size=12),
                             text_color=C_RED).pack(side='right')

        # Divider
        ctk.CTkFrame(dialog, fg_color='#D0D0D0', height=1).pack(fill='x', padx=24, pady=10)

        # Elapsed time
        time_row = ctk.CTkFrame(dialog, fg_color='transparent')
        time_row.pack(fill='x', padx=24, pady=(0, 14))
        ctk.CTkLabel(time_row, text='Elapsed Time',
                     font=ctk.CTkFont(size=12), text_color=C_TEXT_MUTED).pack(side='left')
        ctk.CTkLabel(time_row, text=fmt_elapsed(elapsed),
                     font=ctk.CTkFont(size=12, weight='bold'), text_color=C_TEXT_DARK).pack(side='right')

        # Buttons
        btn_row = ctk.CTkFrame(dialog, fg_color='transparent')
        btn_row.pack(pady=(0, 16))

        ctk.CTkButton(
            btn_row,
            text='Buka Folder Output',
            command=lambda: (self._open_output_folder(), dialog.destroy()),
            fg_color=C_BLUE_MED, hover_color=C_BLUE_DARK,
            width=160, height=36,
        ).pack(side='left', padx=6)

        ctk.CTkButton(
            btn_row,
            text='Tutup',
            command=dialog.destroy,
            fg_color='#E0E0E0', text_color=C_TEXT_DARK, hover_color='#C8C8C8',
            width=80, height=36,
        ).pack(side='left', padx=6)

    # ------------------------------------------------------------------
    # Utilities
    # ------------------------------------------------------------------

    def _open_output_folder(self):
        if self._folder and os.path.isdir(self._folder):
            if sys.platform == 'win32':
                os.startfile(self._folder)  # type: ignore[attr-defined]
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', self._folder])
            else:
                subprocess.Popen(['xdg-open', self._folder])

    def _on_resize(self, event):
        if event.widget is self:
            self._cfg['window_width'] = self.winfo_width()
            self._cfg['window_height'] = self.winfo_height()

    def _show_about(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title('About')
        dialog.geometry('300x220')
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.configure(fg_color=C_WHITE)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        dialog.geometry(f'300x220+{(sw - 300) // 2}+{(sh - 220) // 2}')

        ctk.CTkLabel(dialog, text='Gelatik Automation',
                     font=ctk.CTkFont(size=16, weight='bold'),
                     text_color=C_BLUE_DARK).pack(pady=(24, 4))
        ctk.CTkLabel(dialog, text=f'v{get_version()}',
                     font=ctk.CTkFont(size=12),
                     text_color=C_TEXT_MUTED).pack()
        ctk.CTkLabel(dialog, text='SOS Dashboard Generator',
                     font=ctk.CTkFont(size=12),
                     text_color=C_TEXT_DARK).pack(pady=(8, 4))
        ctk.CTkLabel(dialog, text='© 2025 Gelatik',
                     font=ctk.CTkFont(size=11),
                     text_color=C_TEXT_MUTED).pack()

        log_path = BASE_DIR / 'logs' / 'automation.log'
        ctk.CTkLabel(dialog, text=f'Log: {log_path}',
                     font=ctk.CTkFont(size=9),
                     text_color=C_TEXT_MUTED,
                     wraplength=270).pack(pady=(12, 0))

        ctk.CTkButton(dialog, text='Tutup', command=dialog.destroy,
                      fg_color=C_BLUE_MED, hover_color=C_BLUE_DARK,
                      width=80, height=32).pack(pady=16)
