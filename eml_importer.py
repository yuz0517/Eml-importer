"""
EML to Outlook PST Importer
그룹웨어에서 내보낸 .eml 파일을 Outlook PST에 일괄 삽입하는 도구
"""

import os
import sys
import glob
import time
import logging
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
import threading
import queue

# Windows COM 자동화 (win32com)
try:
    import win32com.client
    import pywintypes
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# 로깅 설정
def setup_logging(log_path=None):
    handlers = [logging.StreamHandler()]
    if log_path:
        handlers.append(logging.FileHandler(log_path, encoding='utf-8'))
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=handlers
    )
    return logging.getLogger(__name__)

logger = logging.getLogger(__name__)


class EmlImporter:
    """EML 파일을 Outlook에 삽입하는 핵심 클래스"""

    def __init__(self):
        self.outlook = None
        self.namespace = None

    def connect_outlook(self):
        """Outlook COM 연결"""
        if not WIN32_AVAILABLE:
            raise RuntimeError("win32com 모듈이 없습니다. 'pip install pywin32' 로 설치하세요.")
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            logger.info("Outlook 연결 성공")
            return True
        except Exception as e:
            raise RuntimeError(f"Outlook 연결 실패: {e}\nOutlook이 설치되어 있고 실행 중인지 확인하세요.")

    def get_pst_accounts(self):
        """연결된 PST/계정 폴더 목록 반환"""
        if not self.namespace:
            self.connect_outlook()
        accounts = []
        for store in self.namespace.Stores:
            try:
                root = store.GetRootFolder()
                accounts.append({
                    'name': store.DisplayName,
                    'root': root,
                    'store': store
                })
            except Exception:
                pass
        return accounts

    def get_folders(self, parent_folder, prefix=""):
        """폴더 트리 재귀 탐색"""
        folders = []
        try:
            folders.append({'name': prefix + parent_folder.Name, 'folder': parent_folder})
            for sub in parent_folder.Folders:
                folders.extend(self.get_folders(sub, prefix + "  "))
        except Exception:
            pass
        return folders

    def import_eml_to_folder(self, eml_path, target_folder, progress_callback=None):
        """
        단일 EML 파일을 Outlook 폴더에 삽입
        전략: MailItem 생성 후 EML 내용을 파싱하여 필드 매핑
        """
        try:
            eml_path = Path(eml_path)
            if not eml_path.exists():
                raise FileNotFoundError(f"파일 없음: {eml_path}")

            # EML 파일 읽기 및 파싱
            import email
            from email import policy as email_policy

            with open(eml_path, 'rb') as f:
                raw = f.read()

            msg = email.message_from_bytes(raw, policy=email_policy.default)

            # Outlook MailItem 생성
            mail_item = self.outlook.CreateItem(0)  # olMailItem = 0

            # 기본 헤더 매핑
            subject = msg.get('Subject', '(제목 없음)')
            mail_item.Subject = subject

            # 발신자
            from_addr = msg.get('From', '')
            mail_item.SentOnBehalfOfName = from_addr

            # 수신자
            to_addr = msg.get('To', '')
            cc_addr = msg.get('Cc', '')
            if to_addr:
                mail_item.To = to_addr
            if cc_addr:
                mail_item.CC = cc_addr

            # 날짜
            date_str = msg.get('Date', '')
            if date_str:
                try:
                    from email.utils import parsedate_to_datetime
                    dt = parsedate_to_datetime(date_str)
                    mail_item.ReceivedTime = dt.strftime('%Y-%m-%d %H:%M:%S')
                    mail_item.SentOn = dt.strftime('%Y-%m-%d %H:%M:%S')
                except Exception:
                    pass

            # 본문 추출 (HTML 우선, 없으면 텍스트)
            body_html = None
            body_text = None

            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == 'text/html' and not body_html:
                        try:
                            body_html = part.get_content()
                        except Exception:
                            body_html = part.get_payload(decode=True).decode('utf-8', errors='replace')
                    elif content_type == 'text/plain' and not body_text:
                        try:
                            body_text = part.get_content()
                        except Exception:
                            body_text = part.get_payload(decode=True).decode('utf-8', errors='replace')
            else:
                content_type = msg.get_content_type()
                if content_type == 'text/html':
                    body_html = msg.get_content()
                else:
                    body_text = msg.get_content()

            if body_html:
                mail_item.HTMLBody = body_html
            elif body_text:
                mail_item.Body = body_text

            # 첨부파일 처리
            attachment_count = 0
            if msg.is_multipart():
                import tempfile
                temp_dir = tempfile.mkdtemp()
                for part in msg.walk():
                    disposition = str(part.get('Content-Disposition', ''))
                    if 'attachment' in disposition:
                        filename = part.get_filename()
                        if filename:
                            # RFC2231 디코딩
                            from email.header import decode_header
                            decoded = decode_header(filename)
                            fname_parts = []
                            for fpart, fenc in decoded:
                                if isinstance(fpart, bytes):
                                    fname_parts.append(fpart.decode(fenc or 'utf-8', errors='replace'))
                                else:
                                    fname_parts.append(fpart)
                            filename = ''.join(fname_parts)

                            temp_path = os.path.join(temp_dir, filename)
                            with open(temp_path, 'wb') as tf:
                                tf.write(part.get_payload(decode=True))
                            mail_item.Attachments.Add(temp_path)
                            attachment_count += 1

            # 지정 폴더로 이동 후 저장
            mail_item.Move(target_folder)

            logger.info(f"✓ 삽입 완료: {eml_path.name} (첨부: {attachment_count}개)")
            if progress_callback:
                progress_callback('success', eml_path.name)
            return True

        except Exception as e:
            logger.error(f"✗ 실패: {eml_path} → {e}")
            if progress_callback:
                progress_callback('error', f"{Path(eml_path).name}: {e}")
            return False

    def import_folder(self, eml_dir, target_folder, recursive=False,
                      progress_callback=None, cancel_event=None):
        """폴더 내 모든 EML 파일 일괄 삽입"""
        pattern = '**/*.eml' if recursive else '*.eml'
        eml_files = list(Path(eml_dir).glob(pattern))

        if not eml_files:
            logger.warning(f"EML 파일을 찾을 수 없습니다: {eml_dir}")
            return 0, 0

        total = len(eml_files)
        success = 0
        fail = 0

        logger.info(f"총 {total}개 EML 파일 처리 시작")

        for i, eml_path in enumerate(eml_files):
            if cancel_event and cancel_event.is_set():
                logger.info("사용자 취소")
                break

            if progress_callback:
                progress_callback('progress', (i + 1, total, str(eml_path.name)))

            ok = self.import_eml_to_folder(eml_path, target_folder, progress_callback)
            if ok:
                success += 1
            else:
                fail += 1

            # Outlook 과부하 방지
            if (i + 1) % 20 == 0:
                time.sleep(0.5)

        logger.info(f"완료 — 성공: {success}, 실패: {fail}, 전체: {total}")
        return success, fail


# ───────────────────────────────────────────
#  GUI
# ───────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EML → Outlook PST 일괄 가져오기")
        self.geometry("720x620")
        self.resizable(True, True)
        self.configure(bg="#f5f5f5")

        self.importer = EmlImporter()
        self.folders = []
        self.cancel_event = threading.Event()
        self.msg_queue = queue.Queue()

        self._build_ui()
        self._try_connect_outlook()
        self.after(100, self._poll_queue)

    # ── UI 구성 ──────────────────────────────
    def _build_ui(self):
        pad = dict(padx=12, pady=6)

        # 타이틀
        tk.Label(self, text="EML → Outlook PST 일괄 가져오기",
                 font=("맑은 고딕", 14, "bold"), bg="#f5f5f5").pack(pady=(16, 4))
        tk.Label(self, text="그룹웨어에서 다운로드한 .eml 파일을 Outlook 폴더에 일괄 삽입합니다.",
                 font=("맑은 고딕", 9), fg="#666", bg="#f5f5f5").pack()

        ttk.Separator(self).pack(fill='x', pady=10)

        # ① EML 폴더 선택
        frm1 = tk.LabelFrame(self, text=" ① EML 파일 폴더 ", font=("맑은 고딕", 9, "bold"),
                              bg="#f5f5f5", fg="#333")
        frm1.pack(fill='x', **pad)

        row1 = tk.Frame(frm1, bg="#f5f5f5")
        row1.pack(fill='x', padx=8, pady=6)
        self.eml_dir_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.eml_dir_var, font=("맑은 고딕", 9),
                 width=55).pack(side='left', padx=(0, 6))
        tk.Button(row1, text="폴더 선택", command=self._browse_dir,
                  font=("맑은 고딕", 9)).pack(side='left')

        self.recursive_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frm1, text="하위 폴더 포함", variable=self.recursive_var,
                       font=("맑은 고딕", 9), bg="#f5f5f5").pack(anchor='w', padx=8, pady=(0, 6))

        # ② Outlook 폴더 선택
        frm2 = tk.LabelFrame(self, text=" ② Outlook 대상 폴더 ", font=("맑은 고딕", 9, "bold"),
                              bg="#f5f5f5", fg="#333")
        frm2.pack(fill='x', **pad)

        row2 = tk.Frame(frm2, bg="#f5f5f5")
        row2.pack(fill='x', padx=8, pady=6)

        self.account_var = tk.StringVar()
        self.account_combo = ttk.Combobox(row2, textvariable=self.account_var,
                                          state='readonly', width=28, font=("맑은 고딕", 9))
        self.account_combo.pack(side='left', padx=(0, 6))
        self.account_combo.bind("<<ComboboxSelected>>", self._on_account_change)

        tk.Button(row2, text="↻ 새로고침", command=self._try_connect_outlook,
                  font=("맑은 고딕", 9)).pack(side='left')

        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(frm2, textvariable=self.folder_var,
                                         state='readonly', width=62, font=("맑은 고딕", 9))
        self.folder_combo.pack(fill='x', padx=8, pady=(0, 8))

        # ③ 실행 버튼
        btn_frame = tk.Frame(self, bg="#f5f5f5")
        btn_frame.pack(pady=8)
        self.run_btn = tk.Button(btn_frame, text="▶  가져오기 시작", command=self._start_import,
                                 font=("맑은 고딕", 10, "bold"), bg="#0078d4", fg="white",
                                 width=20, height=2, relief='flat', cursor='hand2')
        self.run_btn.pack(side='left', padx=6)
        self.cancel_btn = tk.Button(btn_frame, text="■  취소", command=self._cancel,
                                    font=("맑은 고딕", 10), state='disabled',
                                    width=10, height=2, relief='flat')
        self.cancel_btn.pack(side='left', padx=6)

        # 진행률
        prog_frame = tk.Frame(self, bg="#f5f5f5")
        prog_frame.pack(fill='x', padx=12)
        self.prog_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(prog_frame, variable=self.prog_var,
                                        maximum=100, length=400)
        self.progress.pack(fill='x', pady=(0, 4))
        self.status_var = tk.StringVar(value="대기 중...")
        tk.Label(prog_frame, textvariable=self.status_var, font=("맑은 고딕", 9),
                 bg="#f5f5f5", fg="#555").pack(anchor='w')

        # 로그
        log_frame = tk.LabelFrame(self, text=" 처리 로그 ", font=("맑은 고딕", 9, "bold"),
                                  bg="#f5f5f5", fg="#333")
        log_frame.pack(fill='both', expand=True, padx=12, pady=(6, 12))
        self.log_box = scrolledtext.ScrolledText(log_frame, height=10,
                                                  font=("Consolas", 8), state='disabled',
                                                  bg="#1e1e1e", fg="#d4d4d4")
        self.log_box.pack(fill='both', expand=True, padx=4, pady=4)
        self.log_box.tag_config('ok', foreground='#4ec9b0')
        self.log_box.tag_config('err', foreground='#f48771')
        self.log_box.tag_config('info', foreground='#9cdcfe')

    # ── Outlook 연결 ─────────────────────────
    def _try_connect_outlook(self):
        if not WIN32_AVAILABLE:
            self._log("⚠ win32com 없음 — pywin32를 설치하세요", 'err')
            return
        try:
            self.importer.connect_outlook()
            accounts = self.importer.get_pst_accounts()
            names = [a['name'] for a in accounts]
            self.account_combo['values'] = names
            self._accounts_data = accounts
            if names:
                self.account_combo.current(0)
                self._on_account_change()
            self._log(f"Outlook 연결 완료 — {len(accounts)}개 계정/PST 발견", 'info')
        except Exception as e:
            self._log(f"Outlook 연결 오류: {e}", 'err')
            messagebox.showerror("연결 오류", str(e))

    def _on_account_change(self, event=None):
        idx = self.account_combo.current()
        if idx < 0 or not hasattr(self, '_accounts_data'):
            return
        account = self._accounts_data[idx]
        self.folders = self.importer.get_folders(account['root'])
        names = [f['name'] for f in self.folders]
        self.folder_combo['values'] = names
        if names:
            self.folder_combo.current(0)

    def _browse_dir(self):
        d = filedialog.askdirectory(title="EML 파일이 있는 폴더 선택")
        if d:
            self.eml_dir_var.set(d)
            count = len(list(Path(d).glob('**/*.eml' if self.recursive_var.get() else '*.eml')))
            self._log(f"폴더 선택: {d}  ({count}개 EML)", 'info')

    # ── 가져오기 실행 ────────────────────────
    def _start_import(self):
        eml_dir = self.eml_dir_var.get().strip()
        if not eml_dir or not os.path.isdir(eml_dir):
            messagebox.showwarning("경고", "EML 폴더를 먼저 선택하세요.")
            return

        folder_idx = self.folder_combo.current()
        if folder_idx < 0:
            messagebox.showwarning("경고", "Outlook 대상 폴더를 선택하세요.")
            return

        target_folder = self.folders[folder_idx]['folder']

        self.cancel_event.clear()
        self.run_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.prog_var.set(0)
        self._log("─" * 60, 'info')
        self._log(f"가져오기 시작: {eml_dir}", 'info')

        thread = threading.Thread(
            target=self._run_import,
            args=(eml_dir, target_folder, self.recursive_var.get()),
            daemon=True
        )
        thread.start()

    def _run_import(self, eml_dir, target_folder, recursive):
        """백그라운드 스레드에서 실행"""
        def cb(event_type, data):
            self.msg_queue.put((event_type, data))

        success, fail = self.importer.import_folder(
            eml_dir, target_folder, recursive,
            progress_callback=cb,
            cancel_event=self.cancel_event
        )
        self.msg_queue.put(('done', (success, fail)))

    def _cancel(self):
        self.cancel_event.set()
        self._log("취소 요청...", 'err')

    # ── 큐 폴링 ─────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                event_type, data = self.msg_queue.get_nowait()
                if event_type == 'progress':
                    cur, total, name = data
                    pct = cur / total * 100
                    self.prog_var.set(pct)
                    self.status_var.set(f"처리 중 ({cur}/{total}): {name}")
                elif event_type == 'success':
                    self._log(f"✓ {data}", 'ok')
                elif event_type == 'error':
                    self._log(f"✗ {data}", 'err')
                elif event_type == 'done':
                    success, fail = data
                    self.prog_var.set(100)
                    self.status_var.set(f"완료 — 성공: {success}개  실패: {fail}개")
                    self._log(f"✅ 완료 — 성공: {success}개  실패: {fail}개", 'ok')
                    self.run_btn.config(state='normal')
                    self.cancel_btn.config(state='disabled')
                    messagebox.showinfo("완료", f"가져오기 완료\n성공: {success}개\n실패: {fail}개")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _log(self, text, tag='info'):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_box.config(state='normal')
        self.log_box.insert('end', f"[{ts}] {text}\n", tag)
        self.log_box.see('end')
        self.log_box.config(state='disabled')


# ───────────────────────────────────────────
#  CLI 모드
# ───────────────────────────────────────────

def cli_main():
    parser = argparse.ArgumentParser(description='EML → Outlook PST 일괄 가져오기')
    parser.add_argument('eml_dir', help='EML 파일이 있는 폴더 경로')
    parser.add_argument('--folder', default='받은 편지함', help='Outlook 대상 폴더명 (기본: 받은 편지함)')
    parser.add_argument('--recursive', action='store_true', help='하위 폴더 포함')
    parser.add_argument('--log', help='로그 파일 경로')
    parser.add_argument('--list-folders', action='store_true', help='Outlook 폴더 목록 출력')
    args = parser.parse_args()

    setup_logging(args.log)
    importer = EmlImporter()
    importer.connect_outlook()

    if args.list_folders:
        accounts = importer.get_pst_accounts()
        for acc in accounts:
            print(f"\n[{acc['name']}]")
            for f in importer.get_folders(acc['root']):
                print(f"  {f['name']}")
        return

    # 폴더명으로 대상 찾기
    target = None
    accounts = importer.get_pst_accounts()
    for acc in accounts:
        for f in importer.get_folders(acc['root']):
            if args.folder in f['name']:
                target = f['folder']
                logger.info(f"대상 폴더: {f['name']}")
                break
        if target:
            break

    if not target:
        logger.error(f"폴더를 찾을 수 없음: {args.folder}")
        sys.exit(1)

    importer.import_folder(args.eml_dir, target, args.recursive)


# ───────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) > 1:
        cli_main()
    else:
        app = App()
        app.mainloop()
