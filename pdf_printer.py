"""
문서 PDF 인쇄 도구
- 프린터를 "Microsoft Print to PDF"로 강제 지정
- 열린 PPT/Excel/한글 문서를 PDF로 인쇄 저장
- DRM 우회 가능
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time

# COM 관련
try:
    import win32com.client
    import win32print
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


class PDFPrinter:
    """PDF 인쇄 도구"""

    PDF_PRINTER_NAME = "Microsoft Print to PDF"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("문서 PDF 인쇄 도구")
        self.root.geometry("550x450")
        self.root.resizable(False, False)

        # 상태 변수
        self.current_doc_name = tk.StringVar(value="감지 중...")
        self.current_doc_type = tk.StringVar(value="-")
        self.save_path = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="프로그램 시작됨")
        self.progress_var = tk.DoubleVar(value=0)
        self.original_printer = None

        self.setup_ui()
        self.check_pdf_printer()
        self.detect_open_document()

    def setup_ui(self):
        """UI 구성"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(main_frame, text="문서 PDF 인쇄 도구",
                                font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # 설명
        desc_label = ttk.Label(main_frame,
                               text="암호화된 문서를 PDF로 인쇄하여 보안 해제",
                               font=("맑은 고딕", 9))
        desc_label.pack(pady=(0, 15))

        # 프린터 상태 프레임
        printer_frame = ttk.LabelFrame(main_frame, text="PDF 프린터 상태", padding="10")
        printer_frame.pack(fill=tk.X, pady=(0, 10))

        self.printer_status = ttk.Label(printer_frame, text="확인 중...",
                                         font=("맑은 고딕", 10))
        self.printer_status.pack()

        # 문서 정보 프레임
        info_frame = ttk.LabelFrame(main_frame, text="현재 열린 문서", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))

        # 문서명
        doc_frame = ttk.Frame(info_frame)
        doc_frame.pack(fill=tk.X, pady=2)
        ttk.Label(doc_frame, text="문서명:", width=10).pack(side=tk.LEFT)
        ttk.Label(doc_frame, textvariable=self.current_doc_name,
                  font=("맑은 고딕", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 문서 종류
        type_frame = ttk.Frame(info_frame)
        type_frame.pack(fill=tk.X, pady=2)
        ttk.Label(type_frame, text="종류:", width=10).pack(side=tk.LEFT)
        ttk.Label(type_frame, textvariable=self.current_doc_type,
                  font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_document).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(main_frame, text="PDF 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=(0, 15))

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_save_path).pack(side=tk.LEFT)

        # 인쇄 버튼
        self.print_button = ttk.Button(main_frame, text="PDF로 인쇄",
                                        command=self.start_print,
                                        style="Accent.TButton")
        self.print_button.pack(pady=10)

        # 진행바
        self.progress = ttk.Progressbar(main_frame, variable=self.progress_var,
                                         maximum=100, length=450)
        self.progress.pack(pady=(0, 10))

        # 상태 표시
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X)
        ttk.Label(status_frame, text="상태:").pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.status_text,
                  font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(5, 0))

        # 스타일 설정
        style = ttk.Style()
        style.configure("Accent.TButton", font=("맑은 고딕", 11, "bold"))

    def check_pdf_printer(self):
        """PDF 프린터 존재 확인"""
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            if self.PDF_PRINTER_NAME in printers:
                self.printer_status.config(text=f"✅ {self.PDF_PRINTER_NAME} 사용 가능",
                                           foreground="green")
            else:
                self.printer_status.config(text=f"❌ {self.PDF_PRINTER_NAME} 없음",
                                           foreground="red")
                messagebox.showwarning("경고",
                    f"{self.PDF_PRINTER_NAME}가 설치되어 있지 않습니다.\n"
                    "Windows 설정에서 기능을 활성화해주세요.")
        except Exception as e:
            self.printer_status.config(text=f"확인 실패: {str(e)}", foreground="red")

    def browse_save_path(self):
        """저장 경로 선택"""
        # 기본 파일명 제안
        doc_name = self.current_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 문서 없음":
            default_name = os.path.splitext(doc_name)[0] + ".pdf"
        else:
            default_name = "output.pdf"

        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf")],
            initialfile=default_name,
            title="PDF 저장 위치 선택"
        )
        if path:
            self.save_path.set(path)

    def detect_open_document(self):
        """열려있는 문서 감지"""
        self.status_text.set("문서 감지 중...")
        self.current_doc_name.set("감지 중...")
        self.current_doc_type.set("-")

        thread = threading.Thread(target=self._detect_documents)
        thread.daemon = True
        thread.start()

    def _detect_documents(self):
        """문서 감지 (백그라운드)"""
        pythoncom.CoInitialize()

        detected = False

        # PowerPoint 확인
        try:
            ppt = win32com.client.GetObject(Class="PowerPoint.Application")
            if ppt.Presentations.Count > 0:
                presentation = ppt.ActivePresentation
                name = presentation.Name
                self.root.after(0, lambda: self.current_doc_name.set(name))
                self.root.after(0, lambda: self.current_doc_type.set("PowerPoint"))
                detected = True
        except:
            pass

        # Excel 확인
        if not detected:
            try:
                excel = win32com.client.GetObject(Class="Excel.Application")
                if excel.Workbooks.Count > 0:
                    workbook = excel.ActiveWorkbook
                    name = workbook.Name
                    self.root.after(0, lambda: self.current_doc_name.set(name))
                    self.root.after(0, lambda: self.current_doc_type.set("Excel"))
                    detected = True
            except:
                pass

        # 한글 확인
        if not detected:
            try:
                hwp = win32com.client.GetObject(Class="HWPFrame.HwpObject")
                path = hwp.Path
                name = os.path.basename(path) if path else "제목 없음"
                self.root.after(0, lambda: self.current_doc_name.set(name))
                self.root.after(0, lambda: self.current_doc_type.set("한글"))
                detected = True
            except:
                pass

        if not detected:
            self.root.after(0, lambda: self.current_doc_name.set("열린 문서 없음"))
            self.root.after(0, lambda: self.current_doc_type.set("-"))
            self.root.after(0, lambda: self.status_text.set("문서를 먼저 열어주세요"))
        else:
            self.root.after(0, lambda: self.status_text.set("문서 감지 완료"))

        pythoncom.CoUninitialize()

    def get_default_printer(self):
        """현재 기본 프린터 가져오기"""
        try:
            return win32print.GetDefaultPrinter()
        except:
            return None

    def set_default_printer(self, printer_name):
        """기본 프린터 설정"""
        try:
            win32print.SetDefaultPrinter(printer_name)
            return True
        except:
            # CMD로 시도
            try:
                subprocess.run(
                    f'rundll32 printui.dll,PrintUIEntry /y /n "{printer_name}"',
                    shell=True, check=True
                )
                return True
            except:
                return False

    def start_print(self):
        """인쇄 시작"""
        if not self.save_path.get():
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        doc_type = self.current_doc_type.get()
        if doc_type == "-":
            messagebox.showwarning("경고", "열린 문서가 없습니다.")
            return

        self.print_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._print_to_pdf)
        thread.daemon = True
        thread.start()

    def _print_to_pdf(self):
        """PDF로 인쇄 (백그라운드)"""
        pythoncom.CoInitialize()

        try:
            doc_type = self.current_doc_type.get()
            save_path = self.save_path.get()

            # 1. 현재 기본 프린터 저장
            self.root.after(0, lambda: self.status_text.set("기본 프린터 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(10))
            self.original_printer = self.get_default_printer()

            # 2. PDF 프린터로 변경
            self.root.after(0, lambda: self.status_text.set("PDF 프린터로 변경 중..."))
            self.root.after(0, lambda: self.progress_var.set(20))
            if not self.set_default_printer(self.PDF_PRINTER_NAME):
                raise Exception("PDF 프린터 설정 실패")

            time.sleep(0.5)  # 프린터 변경 대기

            # 3. 문서 인쇄
            self.root.after(0, lambda: self.status_text.set("PDF로 인쇄 중..."))
            self.root.after(0, lambda: self.progress_var.set(40))

            if doc_type == "PowerPoint":
                self._print_powerpoint(save_path)
            elif doc_type == "Excel":
                self._print_excel(save_path)
            elif doc_type == "한글":
                self._print_hwp(save_path)

            self.root.after(0, lambda: self.progress_var.set(80))

            # 4. 원래 프린터로 복원
            self.root.after(0, lambda: self.status_text.set("기본 프린터 복원 중..."))
            if self.original_printer:
                self.set_default_printer(self.original_printer)

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("PDF 인쇄 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"PDF 저장 완료!\n{save_path}\n\n"
                "※ 저장된 PDF를 이 PC에서 열지 마세요!\n"
                "USB 등으로 복사 후 다른 PC에서 사용하세요."))

        except Exception as e:
            # 에러 시에도 프린터 복원
            if self.original_printer:
                self.set_default_printer(self.original_printer)

            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"인쇄 중 오류 발생:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.print_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    def _print_powerpoint(self, save_path):
        """PowerPoint 인쇄"""
        ppt_app = win32com.client.GetObject(Class="PowerPoint.Application")
        presentation = ppt_app.ActivePresentation

        # PrintToFile 옵션으로 PDF 출력
        # PrintOut 메서드: PrintOut(From, To, PrintToFile, Copies, Collate)
        presentation.PrintOut(
            PrintToFile=save_path,
            Copies=1
        )

        # 인쇄 완료 대기
        time.sleep(2)

    def _print_excel(self, save_path):
        """Excel 인쇄"""
        excel_app = win32com.client.GetObject(Class="Excel.Application")
        workbook = excel_app.ActiveWorkbook

        # ActiveSheet 또는 전체 Workbook 인쇄
        # ExportAsFixedFormat으로 PDF 직접 생성 시도
        try:
            # Type=0: PDF, Type=1: XPS
            workbook.ExportAsFixedFormat(
                Type=0,
                Filename=save_path,
                Quality=0,  # 표준 품질
                IncludeDocProperties=True,
                OpenAfterPublish=False
            )
        except:
            # ExportAsFixedFormat 실패 시 PrintOut 사용
            workbook.PrintOut(
                PrintToFile=True,
                PrToFileName=save_path,
                Copies=1
            )

        time.sleep(2)

    def _print_hwp(self, save_path):
        """한글 인쇄"""
        try:
            hwp = win32com.client.GetObject(Class="HWPFrame.HwpObject")
        except:
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

        # 한글 PDF 저장
        try:
            # SaveAs로 PDF 저장 시도
            hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave)
            hwp.HParameterSet.HFileOpenSave.filename = save_path
            hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave)
        except:
            # PrintOut으로 시도
            hwp.HAction.GetDefault("Print", hwp.HParameterSet.HPrint)
            hwp.HParameterSet.HPrint.PrintToFile = 1
            hwp.HParameterSet.HPrint.FileName = save_path
            hwp.HAction.Execute("Print", hwp.HParameterSet.HPrint)

        time.sleep(2)

    def run(self):
        """프로그램 실행"""
        self.root.mainloop()


def check_dependencies():
    """의존성 확인"""
    if not HAS_WIN32:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("의존성 오류",
            "pywin32 패키지가 필요합니다.\n\n"
            "설치 명령어:\npip install pywin32")
        sys.exit(1)


if __name__ == "__main__":
    check_dependencies()
    app = PDFPrinter()
    app.run()
