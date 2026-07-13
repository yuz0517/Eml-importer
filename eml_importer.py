"""
EML to Outlook Importer v2
pywin32 COM 대신 PowerShell COM 호출 방식
Outlook 보안 설정에 영향받지 않음
"""

import os
import sys
import time
import json
import logging
import threading
import queue
import subprocess
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def ps_run(code: str, timeout: int = 30) -> str:
    """PowerShell 코드 실행 후 stdout 반환"""
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-Command", code],
            capture_output=True, text=True, timeout=timeout,
            encoding='utf-8', errors='replace'
        )
        return r.stdout.strip()
    except subprocess.TimeoutExpired:
        return "ERR:타임아웃"
    except FileNotFoundError:
        return "ERR:PowerShell을 찾을 수 없습니다"
    except Exception as e:
        return f"ERR:{e}"


def get_folders() -> list[dict]:
    """Outlook 폴더 전체 목록 반환"""
    code = """
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    $list = @()
    foreach ($store in $ns.Stores) {
        $acc = $store.DisplayName
        $q = New-Object System.Collections.Queue
        $q.Enqueue(@{f=$store.GetRootFolder();d=0})
        while ($q.Count -gt 0) {
            $item = $q.Dequeue()
            $folder = $item.f
            $list += [PSCustomObject]@{account=$acc;name=$folder.Name;depth=$item.d}
            foreach ($sub in $folder.Folders) {
                $q.Enqueue(@{f=$sub;d=$item.d+1})
            }
        }
    }
    $list | ConvertTo-Json -Compress
} catch { Write-Output "ERR:$($_.Exception.Message)" }
"""
    out = ps_run(code, timeout=15)
    if not out or out.startswith("ERR:"):
        return []
    try:
        data = json.loads(out)
        return data if isinstance(data, list) else [data]
    except Exception:
        return []


def import_eml(eml_path: str, folder_name: str) -> tuple[bool, str]:
    """단일 EML → Outlook 지정 폴더 삽입 (PowerShell)"""
    safe_path = eml_path.replace("'", "''")
    safe_folder = folder_name.replace("'", "''")
    code = f"""
$ErrorActionPreference = 'Stop'
try {{
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    $target = $null
    foreach ($store in $ns.Stores) {{
        $q = New-Object System.Collections.Queue
        $q.Enqueue($store.GetRootFolder())
        while ($q.Count -gt 0) {{
            $f = $q.Dequeue()
            if ($f.Name -eq '{safe_folder}') {{ $target = $f; break }}
            foreach ($sub in $f.Folders) {{ $q.Enqueue($sub) }}
        }}
        if ($target) {{ break }}
    }}
    if (-not $target) {{
        $target = $ns.GetDefaultFolder(5)
    }}
    $mail = $ol.Session.OpenSharedItem('{safe_path}')
    $mail.Move($target) | Out-Null
    Write-Output 'OK'
}} catch {{
    Write-Output "ERR:$($_.Exception.Message)"
}}
"""
    out = ps_run(code, timeout=20)
    if out == "OK":
        return True, ""
    elif out.startswith("ERR:"):
        return False, out[4:]
    else:
        return False, out or "알 수 없는 오류"


def batch_import(eml_dir: str, folder_name: str,
                 recursive: bool = False,
                 progress_cb=None, cancel_event=None) -> tuple[int, int]:
    pattern = '**/*.eml' if recursive else '*.eml'
    files = list(Path(eml_dir).glob(pattern))
    if not files:
        return 0, 0

    total = len(files)
    success = fail = 0

    for i, f in enumerate(files):
        if cancel_event and cancel_event.is_set():
            break
        if progress_cb:
            progress_cb('progress', (i + 1, total, f.name))

        ok, err = import_eml(str(f), folder_name)
        if ok:
            success += 1
            if progress_cb:
                progress_cb('success', f.name)
        else:
            fail += 1
            if progress_cb:
                progress_cb('error', f"{f.name}: {err}")

        # 10개마다 잠깐 쉬어서 Outlook 과부하 방지
        if (i + 1) % 10 == 0:
            time.sleep(0.5)

    return success, fail


# ─────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EML → Outlook 일괄 가져오기 v2")
        self.geometry("740x650")
        self.configure(bg="#f5f5f5")
        self.folders_data = []
        self.cancel_event = threading.Event()
        self.msg_queue = queue.Queue()
        self.folder_name_var = tk.StringVar(value="보낸 편지함")
        self._build_ui()
        self.after(300, self._load_folders)
        self.after(100, self._poll_queue)

    def _build_ui(self):
        p = dict(padx=12, pady=5)

        tk.Label(self, text="EML → Outlook 일괄 가져오기",
                 font=("맑은 고딕", 14, "bold"), bg="#f5f5f5").pack(pady=(14, 2))
        tk.Label(self, text="PowerShell 방식 v2 — Outlook 보안 설정 우회",
                 font=("맑은 고딕", 9), fg="#888", bg="#f5f5f5").pack()
        ttk.Separator(self).pack(fill='x', pady=8)

        # ① EML 폴더
        f1 = tk.LabelFrame(self, text=" ① EML 파일 폴더 ",
                            font=("맑은 고딕", 9, "bold"), bg="#f5f5f5")
        f1.pack(fill='x', **p)
        r1 = tk.Frame(f1, bg="#f5f5f5")
        r1.pack(fill='x', padx=8, pady=6)
        self.eml_dir_var = tk.StringVar()
        tk.Entry(r1, textvariable=self.eml_dir_var,
                 font=("맑은 고딕", 9), width=52).pack(side='left', padx=(0, 6))
        tk.Button(r1, text="폴더 선택", command=self._browse,
                  font=("맑은 고딕", 9)).pack(side='left')
        self.recursive_var = tk.BooleanVar()
        tk.Checkbutton(f1, text="하위 폴더 포함", variable=self.recursive_var,
                       font=("맑은 고딕", 9), bg="#f5f5f5").pack(
            anchor='w', padx=8, pady=(0, 5))

        # ② Outlook 폴더
        f2 = tk.LabelFrame(self, text=" ② Outlook 대상 폴더 ",
                            font=("맑은 고딕", 9, "bold"), bg="#f5f5f5")
        f2.pack(fill='x', **p)

        r2a = tk.Frame(f2, bg="#f5f5f5")
        r2a.pack(fill='x', padx=8, pady=(6, 2))
        tk.Label(r2a, text="폴더명:", font=("맑은 고딕", 9), bg="#f5f5f5").pack(side='left')
        tk.Entry(r2a, textvariable=self.folder_name_var,
                 font=("맑은 고딕", 9), width=28).pack(side='left', padx=6)
        tk.Label(r2a, text="← 직접 입력하거나 아래 목록에서 선택",
                 font=("맑은 고딕", 8), fg="#888", bg="#f5f5f5").pack(side='left')

        r2b = tk.Frame(f2, bg="#f5f5f5")
        r2b.pack(fill='x', padx=8, pady=(2, 8))
        self.folder_combo = ttk.Combobox(r2b, state='readonly',
                                          width=52, font=("맑은 고딕", 9))
        self.folder_combo.pack(side='left', padx=(0, 6))
        self.folder_combo.bind("<<ComboboxSelected>>", self._on_select)
        tk.Button(r2b, text="↻ 새로고침", command=self._load_folders,
                  font=("맑은 고딕", 9)).pack(side='left')

        # 버튼
        bf = tk.Frame(self, bg="#f5f5f5")
        bf.pack(pady=8)
        self.run_btn = tk.Button(bf, text="▶  가져오기 시작", command=self._start,
                                  font=("맑은 고딕", 10, "bold"),
                                  bg="#0078d4", fg="white",
                                  width=18, height=2, relief='flat', cursor='hand2')
        self.run_btn.pack(side='left', padx=6)
        self.cancel_btn = tk.Button(bf, text="■  취소", command=self._cancel,
                                     state='disabled', font=("맑은 고딕", 10),
                                     width=10, height=2, relief='flat')
        self.cancel_btn.pack(side='left', padx=6)

        # 진행
        pf = tk.Frame(self, bg="#f5f5f5")
        pf.pack(fill='x', padx=12)
        self.prog_var = tk.DoubleVar()
        ttk.Progressbar(pf, variable=self.prog_var, maximum=100).pack(fill='x', pady=(0, 3))
        self.status_var = tk.StringVar(value="대기 중...")
        tk.Label(pf, textvariable=self.status_var, font=("맑은 고딕", 9),
                 bg="#f5f5f5", fg="#555").pack(anchor='w')

        # 로그
        lf = tk.LabelFrame(self, text=" 처리 로그 ",
                            font=("맑은 고딕", 9, "bold"), bg="#f5f5f5")
        lf.pack(fill='both', expand=True, padx=12, pady=(4, 12))
        self.log_box = scrolledtext.ScrolledText(lf, height=10,
                                                  font=("Consolas", 8),
                                                  state='disabled',
                                                  bg="#1e1e1e", fg="#d4d4d4")
        self.log_box.pack(fill='both', expand=True, padx=4, pady=4)
        self.log_box.tag_config('ok', foreground='#4ec9b0')
        self.log_box.tag_config('err', foreground='#f48771')
        self.log_box.tag_config('info', foreground='#9cdcfe')

    def _load_folders(self):
        self._log("Outlook 폴더 목록 불러오는 중...", 'info')
        threading.Thread(target=self._fetch_folders, daemon=True).start()

    def _fetch_folders(self):
        data = get_folders()
        self.msg_queue.put(('folders', data))

    def _on_select(self, _=None):
        idx = self.folder_combo.current()
        if 0 <= idx < len(self.folders_data):
            self.folder_name_var.set(self.folders_data[idx].get('name', ''))

    def _browse(self):
        d = filedialog.askdirectory(title="EML 폴더 선택")
        if d:
            self.eml_dir_var.set(d)
            pat = '**/*.eml' if self.recursive_var.get() else '*.eml'
            cnt = len(list(Path(d).glob(pat)))
            self._log(f"선택: {d}  ({cnt}개 EML)", 'info')

    def _start(self):
        eml_dir = self.eml_dir_var.get().strip()
        folder = self.folder_name_var.get().strip()
        if not eml_dir or not os.path.isdir(eml_dir):
            messagebox.showwarning("경고", "EML 폴더를 선택하세요.")
            return
        if not folder:
            messagebox.showwarning("경고", "Outlook 대상 폴더명을 입력하세요.")
            return
        self.cancel_event.clear()
        self.run_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.prog_var.set(0)
        self._log(f"시작 → [{folder}]", 'info')
        threading.Thread(
            target=self._do_import,
            args=(eml_dir, folder, self.recursive_var.get()),
            daemon=True).start()

    def _do_import(self, eml_dir, folder, recursive):
        def cb(t, d): self.msg_queue.put((t, d))
        s, f = batch_import(eml_dir, folder, recursive, cb, self.cancel_event)
        self.msg_queue.put(('done', (s, f)))

    def _cancel(self):
        self.cancel_event.set()
        self._log("취소 요청...", 'err')

    def _poll_queue(self):
        try:
            while True:
                t, d = self.msg_queue.get_nowait()
                if t == 'progress':
                    cur, total, name = d
                    self.prog_var.set(cur / total * 100)
                    self.status_var.set(f"처리 중 ({cur}/{total}): {name}")
                elif t == 'success':
                    self._log(f"✓ {d}", 'ok')
                elif t == 'error':
                    self._log(f"✗ {d}", 'err')
                elif t == 'folders':
                    self.folders_data = d
                    labels = []
                    for x in d:
                        acc = x.get('account', '')
                        name = x.get('name', '')
                        indent = "  " * x.get('depth', 0)
                        labels.append(f"{indent}{name}  ({acc})")
                    self.folder_combo['values'] = labels
                    cnt = len(d)
                    if cnt:
                        self._log(f"폴더 {cnt}개 로드 완료", 'info')
                    else:
                        self._log("⚠ 폴더 목록 로드 실패 — Outlook이 실행 중인지 확인하세요", 'err')
                elif t == 'done':
                    s, f = d
                    self.prog_var.set(100)
                    self.status_var.set(f"완료 — 성공: {s}개  실패: {f}개")
                    self._log(f"✅ 완료 — 성공: {s}개  실패: {f}개", 'ok')
                    self.run_btn.config(state='normal')
                    self.cancel_btn.config(state='disabled')
                    messagebox.showinfo("완료", f"성공: {s}개\n실패: {f}개")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _log(self, text, tag='info'):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_box.config(state='normal')
        self.log_box.insert('end', f"[{ts}] {text}\n", tag)
        self.log_box.see('end')
        self.log_box.config(state='disabled')


if __name__ == '__main__':
    app = App()
    app.mainloop()
