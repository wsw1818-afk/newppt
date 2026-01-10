"""
문서 보안 해제 도구 v3
- PPT, Excel, 한글(HWP) 지원
- COM으로 데이터 읽기
- python-pptx, openpyxl, python-hwp로 직접 파일 생성 (DRM 우회)
- 슬라이드 전체 이미지 캡처 + 도형 속성 직접 재생성
- 바탕화면에 상세 로그 저장
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import tempfile
import shutil
import datetime
import traceback
import io
import time

# python-pptx 관련
try:
    from pptx import Presentation
    from pptx.util import Pt, Emu, Inches
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn
    from pptx.oxml import parse_xml
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# openpyxl 관련 (Excel)
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# COM 관련
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


class Logger:
    """바탕화면에 로그 파일 저장"""

    def __init__(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = os.path.join(desktop, f"DocExtractor_Log_{timestamp}.txt")
        self.log_file = open(self.log_path, "w", encoding="utf-8")
        self.log(f"=== 문서 추출기 v3 로그 시작 ===")
        self.log(f"로그 파일: {self.log_path}")
        self.log(f"시작 시간: {datetime.datetime.now()}")
        self.log(f"지원 문서: PPT, Excel, 한글(HWP)")
        self.log("")

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{timestamp}] {message}"
        print(line)
        self.log_file.write(line + "\n")
        self.log_file.flush()

    def error(self, message, exception=None):
        self.log(f"[ERROR] {message}")
        if exception:
            self.log(f"[ERROR] 예외 타입: {type(exception).__name__}")
            self.log(f"[ERROR] 예외 메시지: {str(exception)}")
            self.log(f"[ERROR] 상세 트레이스백:")
            for line in traceback.format_exc().split("\n"):
                self.log(f"        {line}")

    def close(self):
        self.log("")
        self.log(f"=== 로그 종료: {datetime.datetime.now()} ===")
        self.log_file.close()


# AutoShape Type 매핑 (COM -> python-pptx)
AUTOSHAPE_MAPPING = {
    1: MSO_SHAPE.RECTANGLE,
    2: MSO_SHAPE.PARALLELOGRAM,
    3: MSO_SHAPE.TRAPEZOID,
    4: MSO_SHAPE.DIAMOND,
    5: MSO_SHAPE.ROUNDED_RECTANGLE,
    6: MSO_SHAPE.OCTAGON,
    9: MSO_SHAPE.ISOSCELES_TRIANGLE,
    10: MSO_SHAPE.RIGHT_TRIANGLE,
    11: MSO_SHAPE.OVAL,
    12: MSO_SHAPE.HEXAGON,
    13: MSO_SHAPE.CROSS,
    16: MSO_SHAPE.CUBE,
    # ... 필요시 더 추가
}


class DocumentExtractorV3:
    """문서 추출기 v3 - PPT, Excel, 한글 지원"""

    def __init__(self):
        self.logger = Logger()
        self.logger.log("DocumentExtractor v3 초기화 시작")

        self.root = tk.Tk()
        self.root.title("문서 보안 해제 도구 v3")
        self.root.geometry("600x600")
        self.root.resizable(False, False)

        # 상태 변수 (공통)
        self.status_text = tk.StringVar(value="프로그램 시작됨")
        self.progress_var = tk.DoubleVar(value=0)

        # PPT 상태 변수
        self.ppt_doc_name = tk.StringVar(value="감지 중...")
        self.ppt_slide_count = tk.StringVar(value="-")
        self.ppt_save_path = tk.StringVar(value="")
        self.ppt_list = []
        self.selected_ppt_index = tk.IntVar(value=0)

        # Excel 상태 변수
        self.excel_doc_name = tk.StringVar(value="감지 중...")
        self.excel_sheet_count = tk.StringVar(value="-")
        self.excel_save_path = tk.StringVar(value="")
        self.excel_list = []
        self.selected_excel_index = tk.IntVar(value=0)

        # 한글 상태 변수
        self.hwp_doc_name = tk.StringVar(value="감지 중...")
        self.hwp_page_count = tk.StringVar(value="-")
        self.hwp_save_path = tk.StringVar(value="")
        self.hwp_list = []
        self.selected_hwp_index = tk.IntVar(value=0)

        # 탭 변경 추적 (중복 감지 방지)
        self.last_tab_index = -1
        self.tab_detected = [False, False, False]  # PPT, Excel, HWP 각각 감지 완료 여부

        self.setup_ui()

        self.logger.log("DocumentExtractor v3 초기화 완료")

    def setup_ui(self):
        """UI 구성"""
        self.logger.log("UI 구성 시작")

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(main_frame, text="문서 보안 해제 도구 v3",
                                font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 5))

        # 설명
        desc_label = ttk.Label(main_frame,
                               text="PPT, Excel, 한글(HWP) 문서의 DRM을 우회하여 새 파일로 저장합니다",
                               font=("맑은 고딕", 9), justify=tk.CENTER)
        desc_label.pack(pady=(0, 10))

        # 탭 노트북
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # PPT 탭
        self.ppt_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.ppt_tab, text="  PowerPoint  ")
        self._setup_ppt_tab()

        # Excel 탭
        self.excel_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.excel_tab, text="  Excel  ")
        self._setup_excel_tab()

        # 한글 탭
        self.hwp_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.hwp_tab, text="  한글(HWP)  ")
        self._setup_hwp_tab()

        # 진행바 (공통)
        self.progress = ttk.Progressbar(main_frame, variable=self.progress_var,
                                         maximum=100, length=550)
        self.progress.pack(pady=(0, 5))

        # 상태 표시 (공통)
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X)
        ttk.Label(status_frame, text="상태:").pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.status_text,
                  font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(5, 0))

        # 스타일 설정
        style = ttk.Style()
        style.configure("Accent.TButton", font=("맑은 고딕", 11, "bold"))

        self.logger.log("UI 구성 완료")

    def _setup_ppt_tab(self):
        """PPT 탭 설정"""
        tab = self.ppt_tab

        # 문서 정보 프레임
        info_frame = ttk.LabelFrame(tab, text="열린 PPT 선택", padding="10")
        info_frame.pack(fill=tk.X, pady=5, padx=5)

        # PPT 선택 콤보박스
        select_frame = ttk.Frame(info_frame)
        select_frame.pack(fill=tk.X, pady=2)
        ttk.Label(select_frame, text="PPT 선택:", width=12).pack(side=tk.LEFT)
        self.ppt_combo = ttk.Combobox(select_frame, state="readonly", width=40)
        self.ppt_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.ppt_combo.bind("<<ComboboxSelected>>", self.on_ppt_selected)

        # 슬라이드 수
        slide_frame = ttk.Frame(info_frame)
        slide_frame.pack(fill=tk.X, pady=2)
        ttk.Label(slide_frame, text="슬라이드 수:", width=12).pack(side=tk.LEFT)
        ttk.Label(slide_frame, textvariable=self.ppt_slide_count,
                  font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_ppt).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(tab, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.ppt_save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_ppt_save_path).pack(side=tk.LEFT)

        # 추출 모드 프레임
        mode_frame = ttk.LabelFrame(tab, text="추출 모드", padding="10")
        mode_frame.pack(fill=tk.X, pady=5, padx=5)

        self.ppt_extract_mode = tk.StringVar(value="hybrid")

        ttk.Radiobutton(mode_frame, text="하이브리드 (도형 속성 재생성 + 실패시 슬라이드 이미지)",
                        variable=self.ppt_extract_mode, value="hybrid").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="슬라이드 이미지만 (각 슬라이드를 이미지로 저장)",
                        variable=self.ppt_extract_mode, value="image_only").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="텍스트만 (텍스트 내용만 추출)",
                        variable=self.ppt_extract_mode, value="text_only").pack(anchor=tk.W)

        # 추출 버튼
        self.ppt_extract_button = ttk.Button(tab, text="새 PPT로 추출 (DRM 우회)",
                                              command=self.start_ppt_extraction,
                                              style="Accent.TButton")
        self.ppt_extract_button.pack(pady=10)

        # 탭 활성화 시 자동 감지 (초기 감지도 이 이벤트로 처리됨)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _setup_excel_tab(self):
        """Excel 탭 설정"""
        tab = self.excel_tab

        # 문서 정보 프레임
        info_frame = ttk.LabelFrame(tab, text="열린 Excel 선택", padding="10")
        info_frame.pack(fill=tk.X, pady=5, padx=5)

        # Excel 선택 콤보박스
        select_frame = ttk.Frame(info_frame)
        select_frame.pack(fill=tk.X, pady=2)
        ttk.Label(select_frame, text="Excel 선택:", width=12).pack(side=tk.LEFT)
        self.excel_combo = ttk.Combobox(select_frame, state="readonly", width=40)
        self.excel_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.excel_combo.bind("<<ComboboxSelected>>", self.on_excel_selected)

        # 시트 수
        sheet_frame = ttk.Frame(info_frame)
        sheet_frame.pack(fill=tk.X, pady=2)
        ttk.Label(sheet_frame, text="시트 수:", width=12).pack(side=tk.LEFT)
        ttk.Label(sheet_frame, textvariable=self.excel_sheet_count,
                  font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_excel).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(tab, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.excel_save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_excel_save_path).pack(side=tk.LEFT)

        # 추출 옵션 프레임
        option_frame = ttk.LabelFrame(tab, text="추출 옵션", padding="10")
        option_frame.pack(fill=tk.X, pady=5, padx=5)

        self.excel_include_format = tk.BooleanVar(value=True)
        self.excel_include_formulas = tk.BooleanVar(value=False)

        ttk.Checkbutton(option_frame, text="서식 포함 (글꼴, 색상, 테두리)",
                        variable=self.excel_include_format).pack(anchor=tk.W)
        ttk.Checkbutton(option_frame, text="수식 대신 값만 저장",
                        variable=self.excel_include_formulas).pack(anchor=tk.W)

        # 추출 버튼
        self.excel_extract_button = ttk.Button(tab, text="새 Excel로 추출 (DRM 우회)",
                                                command=self.start_excel_extraction,
                                                style="Accent.TButton")
        self.excel_extract_button.pack(pady=10)

    def _setup_hwp_tab(self):
        """한글 탭 설정"""
        tab = self.hwp_tab

        # 문서 정보 프레임
        info_frame = ttk.LabelFrame(tab, text="열린 한글 문서 선택", padding="10")
        info_frame.pack(fill=tk.X, pady=5, padx=5)

        # 한글 선택 콤보박스
        select_frame = ttk.Frame(info_frame)
        select_frame.pack(fill=tk.X, pady=2)
        ttk.Label(select_frame, text="한글 선택:", width=12).pack(side=tk.LEFT)
        self.hwp_combo = ttk.Combobox(select_frame, state="readonly", width=40)
        self.hwp_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.hwp_combo.bind("<<ComboboxSelected>>", self.on_hwp_selected)

        # 페이지 수
        page_frame = ttk.Frame(info_frame)
        page_frame.pack(fill=tk.X, pady=2)
        ttk.Label(page_frame, text="페이지 수:", width=12).pack(side=tk.LEFT)
        ttk.Label(page_frame, textvariable=self.hwp_page_count,
                  font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_hwp).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(tab, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.hwp_save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_hwp_save_path).pack(side=tk.LEFT)

        # 추출 옵션 프레임
        option_frame = ttk.LabelFrame(tab, text="추출 옵션", padding="10")
        option_frame.pack(fill=tk.X, pady=5, padx=5)

        self.hwp_extract_mode = tk.StringVar(value="text")

        ttk.Radiobutton(option_frame, text="텍스트만 추출 (서식 없음)",
                        variable=self.hwp_extract_mode, value="text").pack(anchor=tk.W)
        ttk.Radiobutton(option_frame, text="텍스트 + 표 (표 구조 유지)",
                        variable=self.hwp_extract_mode, value="text_table").pack(anchor=tk.W)

        # 추출 버튼
        self.hwp_extract_button = ttk.Button(tab, text="새 한글 파일로 추출 (DRM 우회)",
                                              command=self.start_hwp_extraction,
                                              style="Accent.TButton")
        self.hwp_extract_button.pack(pady=10)

        # 안내 라벨
        notice_label = ttk.Label(tab,
                                 text="※ 한글(HWP)은 텍스트/표만 추출 가능합니다.\n   이미지와 복잡한 서식은 지원되지 않습니다.",
                                 font=("맑은 고딕", 9), foreground="gray")
        notice_label.pack(pady=5)

    def _on_tab_changed(self, event):
        """탭 변경 시 해당 문서 감지 (중복 감지 방지, debounce 적용)"""
        # 기존 예약된 감지가 있으면 취소
        if hasattr(self, '_pending_detect') and self._pending_detect:
            self.root.after_cancel(self._pending_detect)
            self._pending_detect = None

        # 50ms 후에 감지 실행 (debounce)
        self._pending_detect = self.root.after(50, self._do_detect)

    def _do_detect(self):
        """실제 감지 실행"""
        self._pending_detect = None
        current_tab = self.notebook.index(self.notebook.select())

        # 이미 감지 완료된 탭이면 스킵
        if self.tab_detected[current_tab]:
            return

        # 플래그를 먼저 설정하여 중복 호출 방지
        self.tab_detected[current_tab] = True

        # 해당 탭 감지 실행
        if current_tab == 0:  # PPT
            self.detect_open_ppt()
        elif current_tab == 1:  # Excel
            self.detect_open_excel()
        elif current_tab == 2:  # 한글
            self.detect_open_hwp()

    # ========== PPT 관련 메서드 ==========

    def browse_ppt_save_path(self):
        """PPT 저장 경로 선택"""
        self.logger.log("PPT 저장 경로 선택 대화상자 열기")

        doc_name = self.ppt_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 PPT 없음":
            default_name = os.path.splitext(doc_name)[0] + "_복사본.pptx"
        else:
            default_name = "새문서.pptx"

        path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint 파일", "*.pptx")],
            initialfile=default_name,
            title="저장할 위치 선택"
        )
        if path:
            self.ppt_save_path.set(path)
            self.logger.log(f"PPT 저장 경로 선택됨: {path}")

    def detect_open_ppt(self):
        """열려있는 PPT 감지"""
        self.logger.log("PPT 감지 시작")
        self.status_text.set("PPT 감지 중...")
        self.ppt_doc_name.set("감지 중...")
        self.ppt_slide_count.set("-")

        thread = threading.Thread(target=self._detect_ppt)
        thread.daemon = True
        thread.start()

    def _detect_ppt(self):
        """PPT 감지 (백그라운드)"""
        self.logger.log("백그라운드 PPT 감지 스레드 시작")
        pythoncom.CoInitialize()

        try:
            self.logger.log("PowerPoint.Application 객체 연결 시도")
            ppt = win32com.client.GetObject(Class="PowerPoint.Application")
            ppt_count = ppt.Presentations.Count
            self.logger.log(f"PowerPoint 연결 성공, 열린 프레젠테이션 수: {ppt_count}")

            if ppt_count > 0:
                ppt_names = []
                ppt_info = []

                for i in range(1, ppt_count + 1):
                    try:
                        presentation = ppt.Presentations(i)
                        name = presentation.Name
                        slide_count = presentation.Slides.Count
                        ppt_names.append(f"{name} ({slide_count}장)")
                        ppt_info.append((name, slide_count, i))
                        self.logger.log(f"  PPT {i}: {name}, {slide_count}장")
                    except Exception as e:
                        self.logger.log(f"  PPT {i} 정보 가져오기 실패: {str(e)}")

                self.ppt_list = ppt_info

                def update_combo():
                    self.ppt_combo['values'] = ppt_names
                    if ppt_names:
                        self.ppt_combo.current(0)
                        self.selected_ppt_index.set(1)
                        self.ppt_doc_name.set(ppt_info[0][0])
                        self.ppt_slide_count.set(f"{ppt_info[0][1]}장")
                    self.status_text.set(f"PPT {len(ppt_names)}개 감지됨")

                self.root.after(0, update_combo)
            else:
                self.logger.log("열린 프레젠테이션 없음")
                self.ppt_list = []
                def clear_combo():
                    self.ppt_combo.set("")
                    self.ppt_combo['values'] = []
                    self.ppt_doc_name.set("열린 PPT 없음")
                    self.ppt_slide_count.set("-")
                    self.status_text.set("PPT를 먼저 열어주세요")
                self.root.after(0, clear_combo)

        except Exception as e:
            self.logger.error("PPT 감지 실패", e)
            self.ppt_list = []
            err_msg = str(e)[:30]
            def show_error():
                self.ppt_combo.set("")
                self.ppt_doc_name.set("열린 PPT 없음")
                self.ppt_slide_count.set("-")
                self.status_text.set(f"PPT 감지 실패: {err_msg}")
            self.root.after(0, show_error)

        pythoncom.CoUninitialize()

    def on_ppt_selected(self, event):
        """PPT 콤보박스 선택 이벤트"""
        selected_idx = self.ppt_combo.current()
        if selected_idx >= 0 and selected_idx < len(self.ppt_list):
            name, slide_count, ppt_index = self.ppt_list[selected_idx]
            self.selected_ppt_index.set(ppt_index)
            self.ppt_doc_name.set(name)
            self.ppt_slide_count.set(f"{slide_count}장")
            self.logger.log(f"PPT 선택: {name} (인덱스 {ppt_index})")

    def start_ppt_extraction(self):
        """PPT 추출 시작"""
        self.logger.log("PPT 추출 시작 버튼 클릭")

        if not self.ppt_save_path.get():
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.ppt_doc_name.get() == "열린 PPT 없음":
            messagebox.showwarning("경고", "열린 PPT가 없습니다.")
            return

        self.ppt_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_ppt)
        thread.daemon = True
        thread.start()

    def _extract_ppt(self):
        """PPT 추출 (백그라운드)"""
        self.logger.log("=== PPT 추출 프로세스 시작 ===")
        pythoncom.CoInitialize()

        temp_dir = None
        mode = self.ppt_extract_mode.get()

        try:
            save_path = self.ppt_save_path.get()
            self.root.after(0, lambda: self.status_text.set("원본 PPT 연결 중..."))

            ppt_app = win32com.client.GetObject(Class="PowerPoint.Application")

            ppt_index = self.selected_ppt_index.get()
            if ppt_index > 0 and ppt_index <= ppt_app.Presentations.Count:
                source_pres = ppt_app.Presentations(ppt_index)
            else:
                source_pres = ppt_app.ActivePresentation

            self.logger.log(f"원본 프레젠테이션: {source_pres.Name}")

            self.root.after(0, lambda: self.status_text.set("새 PPT 생성 중..."))
            self.root.after(0, lambda: self.progress_var.set(5))

            new_pres = Presentation()

            slide_width = source_pres.PageSetup.SlideWidth
            slide_height = source_pres.PageSetup.SlideHeight
            new_pres.slide_width = Emu(int(slide_width * 12700))
            new_pres.slide_height = Emu(int(slide_height * 12700))

            total_slides = source_pres.Slides.Count
            temp_dir = tempfile.mkdtemp()
            blank_layout = new_pres.slide_layouts[6]

            for i in range(1, total_slides + 1):
                self.logger.log(f"--- 슬라이드 {i}/{total_slides} 처리 ---")
                progress = 5 + (i / total_slides) * 85
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                self.root.after(0, lambda n=i, t=total_slides: self.status_text.set(f"슬라이드 {n}/{t} 처리 중..."))

                source_slide = source_pres.Slides(i)
                new_slide = new_pres.slides.add_slide(blank_layout)

                if mode == "image_only":
                    self._export_slide_as_image(source_slide, new_slide, temp_dir, i, new_pres)
                elif mode == "text_only":
                    self._extract_text_only(source_slide, new_slide)
                else:
                    self._extract_hybrid(source_slide, new_slide, temp_dir, i, new_pres)

            self.root.after(0, lambda: self.status_text.set("파일 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(95))

            new_pres.save(save_path)
            self.logger.log(f"저장 완료: {save_path}")

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("PPT 추출 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"PPT 추출 완료!\n{save_path}\n\n총 {total_slides}장"))

        except Exception as e:
            self.logger.error("PPT 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.root.after(0, lambda: self.ppt_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    # ========== Excel 관련 메서드 ==========

    def browse_excel_save_path(self):
        """Excel 저장 경로 선택"""
        self.logger.log("Excel 저장 경로 선택")

        doc_name = self.excel_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 Excel 없음":
            default_name = os.path.splitext(doc_name)[0] + "_복사본.xlsx"
        else:
            default_name = "새문서.xlsx"

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile=default_name,
            title="저장할 위치 선택"
        )
        if path:
            self.excel_save_path.set(path)
            self.logger.log(f"Excel 저장 경로: {path}")

    def detect_open_excel(self):
        """열려있는 Excel 감지"""
        self.logger.log("Excel 감지 시작")
        self.status_text.set("Excel 감지 중...")
        self.excel_doc_name.set("감지 중...")
        self.excel_sheet_count.set("-")

        thread = threading.Thread(target=self._detect_excel)
        thread.daemon = True
        thread.start()

    def _detect_excel(self):
        """Excel 감지 (백그라운드)"""
        pythoncom.CoInitialize()

        try:
            excel = win32com.client.GetObject(Class="Excel.Application")
            wb_count = excel.Workbooks.Count
            self.logger.log(f"Excel 연결 성공, 열린 통합문서 수: {wb_count}")

            if wb_count > 0:
                excel_names = []
                excel_info = []

                for i in range(1, wb_count + 1):
                    try:
                        workbook = excel.Workbooks(i)
                        name = workbook.Name
                        sheet_count = workbook.Sheets.Count
                        excel_names.append(f"{name} ({sheet_count}시트)")
                        excel_info.append((name, sheet_count, i))
                        self.logger.log(f"  Excel {i}: {name}, {sheet_count}시트")
                    except Exception as e:
                        self.logger.log(f"  Excel {i} 정보 실패: {str(e)}")

                self.excel_list = excel_info

                def update_combo():
                    self.excel_combo['values'] = excel_names
                    if excel_names:
                        self.excel_combo.current(0)
                        self.selected_excel_index.set(1)
                        self.excel_doc_name.set(excel_info[0][0])
                        self.excel_sheet_count.set(f"{excel_info[0][1]}시트")
                    self.status_text.set(f"Excel {len(excel_names)}개 감지됨")

                self.root.after(0, update_combo)
            else:
                self.excel_list = []
                def clear_combo():
                    self.excel_combo.set("")
                    self.excel_combo['values'] = []
                    self.excel_doc_name.set("열린 Excel 없음")
                    self.excel_sheet_count.set("-")
                    self.status_text.set("Excel을 먼저 열어주세요")
                self.root.after(0, clear_combo)

        except Exception as e:
            self.logger.error("Excel 감지 실패", e)
            self.excel_list = []
            def show_error():
                self.excel_combo.set("")
                self.excel_doc_name.set("열린 Excel 없음")
                self.excel_sheet_count.set("-")
                self.status_text.set("Excel 감지 실패")
            self.root.after(0, show_error)

        pythoncom.CoUninitialize()

    def on_excel_selected(self, event):
        """Excel 콤보박스 선택 이벤트"""
        selected_idx = self.excel_combo.current()
        if selected_idx >= 0 and selected_idx < len(self.excel_list):
            name, sheet_count, excel_index = self.excel_list[selected_idx]
            self.selected_excel_index.set(excel_index)
            self.excel_doc_name.set(name)
            self.excel_sheet_count.set(f"{sheet_count}시트")
            self.logger.log(f"Excel 선택: {name}")

    def start_excel_extraction(self):
        """Excel 추출 시작"""
        self.logger.log("Excel 추출 시작")

        if not HAS_OPENPYXL:
            messagebox.showerror("오류", "openpyxl 패키지가 필요합니다.\npip install openpyxl")
            return

        if not self.excel_save_path.get():
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.excel_doc_name.get() == "열린 Excel 없음":
            messagebox.showwarning("경고", "열린 Excel이 없습니다.")
            return

        self.excel_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_excel)
        thread.daemon = True
        thread.start()

    def _extract_excel(self):
        """Excel 추출 (백그라운드)"""
        self.logger.log("=== Excel 추출 시작 ===")
        pythoncom.CoInitialize()

        try:
            save_path = self.excel_save_path.get()
            include_format = self.excel_include_format.get()
            values_only = self.excel_include_formulas.get()

            self.root.after(0, lambda: self.status_text.set("Excel 연결 중..."))

            excel_app = win32com.client.GetObject(Class="Excel.Application")
            excel_index = self.selected_excel_index.get()

            if excel_index > 0 and excel_index <= excel_app.Workbooks.Count:
                source_wb = excel_app.Workbooks(excel_index)
            else:
                source_wb = excel_app.ActiveWorkbook

            self.logger.log(f"원본 통합문서: {source_wb.Name}")

            # openpyxl로 새 통합문서 생성
            new_wb = Workbook()
            # 기본 시트 제거 (나중에 추가할 것이므로)
            default_sheet = new_wb.active

            total_sheets = source_wb.Sheets.Count
            self.root.after(0, lambda: self.progress_var.set(5))

            for sheet_idx in range(1, total_sheets + 1):
                source_sheet = source_wb.Sheets(sheet_idx)
                sheet_name = source_sheet.Name
                self.logger.log(f"  시트 처리: {sheet_name}")

                progress = 5 + (sheet_idx / total_sheets) * 85
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                self.root.after(0, lambda n=sheet_name: self.status_text.set(f"시트 '{n}' 처리 중..."))

                # 새 시트 생성
                if sheet_idx == 1:
                    new_sheet = default_sheet
                    new_sheet.title = sheet_name
                else:
                    new_sheet = new_wb.create_sheet(title=sheet_name)

                # 사용 범위 가져오기
                try:
                    used_range = source_sheet.UsedRange
                    if used_range is None:
                        continue

                    row_count = used_range.Rows.Count
                    col_count = used_range.Columns.Count
                    start_row = used_range.Row
                    start_col = used_range.Column

                    self.logger.log(f"    범위: {row_count}행 x {col_count}열 (시작: {start_row},{start_col})")

                    # 데이터 복사
                    for r in range(row_count):
                        for c in range(col_count):
                            try:
                                cell = source_sheet.Cells(start_row + r, start_col + c)
                                new_cell = new_sheet.cell(row=start_row + r, column=start_col + c)

                                # 값 복사 (수식 또는 값)
                                if values_only:
                                    new_cell.value = cell.Value
                                else:
                                    try:
                                        formula = cell.Formula
                                        if formula and str(formula).startswith('='):
                                            new_cell.value = formula
                                        else:
                                            new_cell.value = cell.Value
                                    except:
                                        new_cell.value = cell.Value

                                # 서식 복사
                                if include_format:
                                    try:
                                        # 폰트
                                        src_font = cell.Font
                                        new_cell.font = Font(
                                            name=src_font.Name if src_font.Name else 'Calibri',
                                            size=src_font.Size if src_font.Size else 11,
                                            bold=bool(src_font.Bold),
                                            italic=bool(src_font.Italic),
                                            color=self._excel_color_to_hex(src_font.Color)
                                        )
                                    except:
                                        pass

                                    try:
                                        # 배경색
                                        interior_color = cell.Interior.Color
                                        if interior_color and interior_color != 16777215:  # 흰색이 아니면
                                            hex_color = self._excel_color_to_hex(interior_color)
                                            if hex_color:
                                                new_cell.fill = PatternFill(start_color=hex_color,
                                                                            end_color=hex_color,
                                                                            fill_type='solid')
                                    except:
                                        pass

                            except Exception as cell_err:
                                pass  # 개별 셀 오류는 무시

                    # 열 너비 복사
                    try:
                        for c in range(1, col_count + 1):
                            width = source_sheet.Columns(start_col + c - 1).ColumnWidth
                            new_sheet.column_dimensions[get_column_letter(start_col + c - 1)].width = width
                    except:
                        pass

                except Exception as sheet_err:
                    self.logger.error(f"시트 '{sheet_name}' 처리 오류", sheet_err)

            # 저장
            self.root.after(0, lambda: self.status_text.set("파일 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(95))

            new_wb.save(save_path)
            self.logger.log(f"저장 완료: {save_path}")

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("Excel 추출 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"Excel 추출 완료!\n{save_path}\n\n총 {total_sheets}시트"))

        except Exception as e:
            self.logger.error("Excel 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.excel_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    def _excel_color_to_hex(self, color):
        """Excel 색상을 hex 문자열로 변환"""
        try:
            if color is None or color == 0:
                return None
            # Excel 색상은 BGR 형식
            b = (color >> 16) & 0xFF
            g = (color >> 8) & 0xFF
            r = color & 0xFF
            return f"{r:02X}{g:02X}{b:02X}"
        except:
            return None

    # ========== 한글(HWP) 관련 메서드 ==========

    def browse_hwp_save_path(self):
        """한글 저장 경로 선택"""
        self.logger.log("한글 저장 경로 선택")

        doc_name = self.hwp_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 한글 없음":
            default_name = os.path.splitext(doc_name)[0] + "_복사본.hwp"
        else:
            default_name = "새문서.hwp"

        path = filedialog.asksaveasfilename(
            defaultextension=".hwp",
            filetypes=[("한글 파일", "*.hwp"), ("텍스트 파일", "*.txt")],
            initialfile=default_name,
            title="저장할 위치 선택"
        )
        if path:
            self.hwp_save_path.set(path)
            self.logger.log(f"한글 저장 경로: {path}")

    def detect_open_hwp(self):
        """열려있는 한글 감지"""
        self.logger.log("한글 감지 시작")
        self.status_text.set("한글 감지 중...")
        self.hwp_doc_name.set("감지 중...")
        self.hwp_page_count.set("-")

        thread = threading.Thread(target=self._detect_hwp)
        thread.daemon = True
        thread.start()

    def _detect_hwp(self):
        """한글 감지 (백그라운드)"""
        pythoncom.CoInitialize()

        try:
            # HWPFrame.HwpObject 또는 HWPFrame.HwpCtrl 시도
            hwp = None
            try:
                hwp = win32com.client.GetObject(Class="HWPFrame.HwpObject")
            except:
                try:
                    hwp = win32com.client.GetObject(Class="HWPFrame.HwpCtrl")
                except:
                    pass

            if hwp is None:
                raise Exception("한글 프로그램을 찾을 수 없습니다")

            self.logger.log("한글 연결 성공")

            # 현재 열린 문서 정보
            try:
                doc_count = 1  # 한글은 보통 하나의 문서만 활성화
                hwp_names = []
                hwp_info = []

                # 현재 문서 정보 가져오기
                try:
                    doc_path = hwp.Path
                    doc_name = os.path.basename(doc_path) if doc_path else "새 문서"
                except:
                    doc_name = "열린 문서"

                try:
                    # 페이지 수 가져오기 시도
                    page_count = hwp.PageCount
                except:
                    page_count = "-"

                hwp_names.append(f"{doc_name}")
                hwp_info.append((doc_name, page_count, 1))
                self.logger.log(f"  한글: {doc_name}, {page_count}페이지")

                self.hwp_list = hwp_info

                def update_combo():
                    self.hwp_combo['values'] = hwp_names
                    if hwp_names:
                        self.hwp_combo.current(0)
                        self.selected_hwp_index.set(1)
                        self.hwp_doc_name.set(hwp_info[0][0])
                        self.hwp_page_count.set(f"{hwp_info[0][1]}페이지" if hwp_info[0][1] != "-" else "-")
                    self.status_text.set("한글 문서 감지됨")

                self.root.after(0, update_combo)

            except Exception as e:
                self.logger.error("한글 문서 정보 가져오기 실패", e)
                raise

        except Exception as e:
            self.logger.error("한글 감지 실패", e)
            self.hwp_list = []
            def show_error():
                self.hwp_combo.set("")
                self.hwp_doc_name.set("열린 한글 없음")
                self.hwp_page_count.set("-")
                self.status_text.set("한글을 먼저 열어주세요")
            self.root.after(0, show_error)

        pythoncom.CoUninitialize()

    def on_hwp_selected(self, event):
        """한글 콤보박스 선택 이벤트"""
        selected_idx = self.hwp_combo.current()
        if selected_idx >= 0 and selected_idx < len(self.hwp_list):
            name, page_count, hwp_index = self.hwp_list[selected_idx]
            self.selected_hwp_index.set(hwp_index)
            self.hwp_doc_name.set(name)
            self.hwp_page_count.set(f"{page_count}페이지" if page_count != "-" else "-")
            self.logger.log(f"한글 선택: {name}")

    def start_hwp_extraction(self):
        """한글 추출 시작"""
        self.logger.log("한글 추출 시작")

        if not self.hwp_save_path.get():
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.hwp_doc_name.get() == "열린 한글 없음":
            messagebox.showwarning("경고", "열린 한글 문서가 없습니다.")
            return

        self.hwp_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_hwp)
        thread.daemon = True
        thread.start()

    def _extract_hwp(self):
        """한글 추출 (백그라운드)"""
        self.logger.log("=== 한글 추출 시작 ===")
        pythoncom.CoInitialize()

        try:
            save_path = self.hwp_save_path.get()
            extract_mode = self.hwp_extract_mode.get()

            self.root.after(0, lambda: self.status_text.set("한글 연결 중..."))

            # 한글 COM 연결
            hwp = None
            try:
                hwp = win32com.client.GetObject(Class="HWPFrame.HwpObject")
            except:
                hwp = win32com.client.GetObject(Class="HWPFrame.HwpCtrl")

            self.logger.log("한글 연결 성공")
            self.root.after(0, lambda: self.progress_var.set(10))

            # 텍스트 추출 방식
            self.root.after(0, lambda: self.status_text.set("텍스트 추출 중..."))

            extracted_text = ""

            try:
                # 방법 1: GetTextFile 사용 (가장 확실)
                temp_txt = os.path.join(tempfile.gettempdir(), "hwp_temp_extract.txt")

                # 전체 선택
                hwp.Run("SelectAll")
                time.sleep(0.1)

                # 클립보드로 복사
                hwp.Run("Copy")
                time.sleep(0.1)

                # 클립보드에서 텍스트 가져오기
                import win32clipboard
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        extracted_text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    elif win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
                        extracted_text = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT).decode('cp949', errors='ignore')
                finally:
                    win32clipboard.CloseClipboard()

                # 선택 해제
                hwp.Run("Cancel")

                self.logger.log(f"텍스트 추출 완료: {len(extracted_text)}자")

            except Exception as e:
                self.logger.error("텍스트 추출 오류", e)
                # 대안: 직접 텍스트 가져오기 시도
                try:
                    extracted_text = hwp.GetTextFile("TEXT", "")
                except:
                    pass

            self.root.after(0, lambda: self.progress_var.set(70))

            # 파일로 저장
            self.root.after(0, lambda: self.status_text.set("파일 저장 중..."))

            if save_path.lower().endswith('.hwp'):
                # HWP로 저장하려면 새 한글 문서 생성 필요
                try:
                    # 새 한글 인스턴스로 저장 시도
                    new_hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                    new_hwp.XHwpWindows.Item(0).Visible = False

                    # 새 문서 생성
                    new_hwp.Run("FileNew")

                    # 텍스트 삽입
                    new_hwp.HAction.GetDefault("InsertText", new_hwp.HParameterSet.HInsertText.HSet)
                    new_hwp.HParameterSet.HInsertText.Text = extracted_text
                    new_hwp.HAction.Execute("InsertText", new_hwp.HParameterSet.HInsertText.HSet)

                    # 저장
                    new_hwp.SaveAs(save_path)
                    new_hwp.Quit()

                    self.logger.log(f"HWP 저장 완료: {save_path}")

                except Exception as hwp_save_err:
                    self.logger.error("HWP 저장 실패, TXT로 저장", hwp_save_err)
                    # TXT로 대체 저장
                    txt_path = save_path.replace('.hwp', '.txt')
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write(extracted_text)
                    save_path = txt_path
            else:
                # TXT로 저장
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(extracted_text)

            self.logger.log(f"저장 완료: {save_path}")

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("한글 추출 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"한글 추출 완료!\n{save_path}\n\n추출된 텍스트: {len(extracted_text)}자"))

        except Exception as e:
            self.logger.error("한글 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.hwp_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    # ========== PPT 슬라이드 처리 메서드 (기존 코드) ==========

    def _export_slide_as_image(self, source_slide, target_slide, temp_dir, slide_num, new_pres):
        """슬라이드 전체를 이미지로 내보내기"""
        self.logger.log(f"  슬라이드 {slide_num}를 이미지로 내보내기...")

        try:
            img_path = os.path.join(temp_dir, f"slide_{slide_num}.png")
            width = int(source_slide.Parent.PageSetup.SlideWidth * 2)
            height = int(source_slide.Parent.PageSetup.SlideHeight * 2)

            source_slide.Export(img_path, "PNG", width, height)

            target_slide.shapes.add_picture(
                img_path, Emu(0), Emu(0),
                new_pres.slide_width, new_pres.slide_height
            )
            self.logger.log(f"  슬라이드 이미지 추가 성공")

        except Exception as e:
            self.logger.error(f"  슬라이드 이미지 내보내기 실패", e)
            self._extract_text_only(source_slide, target_slide)

    def _extract_text_only(self, source_slide, target_slide):
        """텍스트만 추출"""
        for shape in source_slide.Shapes:
            try:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    if text.strip():
                        left = Emu(int(shape.Left * 12700))
                        top = Emu(int(shape.Top * 12700))
                        width = Emu(int(shape.Width * 12700))
                        height = Emu(int(shape.Height * 12700))

                        textbox = target_slide.shapes.add_textbox(left, top, width, height)
                        tf = textbox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = text

                        try:
                            src_font = shape.TextFrame.TextRange.Font
                            p.font.size = Pt(src_font.Size)
                            p.font.bold = src_font.Bold
                            p.font.italic = src_font.Italic
                        except:
                            pass

            except Exception as e:
                pass

    def _extract_hybrid(self, source_slide, target_slide, temp_dir, slide_num, new_pres):
        """하이브리드 모드"""
        shapes_list = []
        for shape in source_slide.Shapes:
            try:
                z_order = shape.ZOrderPosition
                shapes_list.append((z_order, shape))
            except:
                shapes_list.append((0, shape))

        shapes_list.sort(key=lambda x: x[0])

        for z_order, shape in shapes_list:
            try:
                self._recreate_shape(shape, target_slide, temp_dir)
            except:
                pass

    def _recreate_shape(self, source_shape, target_slide, temp_dir):
        """도형 재생성"""
        shape_type = source_shape.Type

        try:
            left = Emu(int(source_shape.Left * 12700))
            top = Emu(int(source_shape.Top * 12700))
            width = Emu(int(source_shape.Width * 12700))
            height = Emu(int(source_shape.Height * 12700))

            # 이미지
            if shape_type in [13, 11]:
                return self._handle_image_shape(source_shape, target_slide, temp_dir, left, top, width, height)

            # 텍스트박스
            if shape_type == 17:
                textbox = target_slide.shapes.add_textbox(left, top, width, height)
                self._copy_text_frame(source_shape, textbox)
                return True

            # AutoShape
            if shape_type == 1:
                return self._handle_autoshape(source_shape, target_slide, left, top, width, height)

            # Placeholder
            if shape_type == 14:
                textbox = target_slide.shapes.add_textbox(left, top, width, height)
                self._copy_text_frame(source_shape, textbox)
                return True

            # Table
            if shape_type == 19:
                return self._handle_table(source_shape, target_slide, left, top, width, height)

            # Connector
            if shape_type == 9:
                return self._handle_connector(source_shape, target_slide, left, top)

            # Group
            if shape_type == 6:
                for item in source_shape.GroupItems:
                    self._recreate_shape(item, target_slide, temp_dir)
                return True

            # Freeform
            if shape_type == 5:
                return self._handle_freeform(source_shape, target_slide, temp_dir, left, top, width, height)

            # 기타 - 텍스트만
            if source_shape.HasTextFrame and source_shape.TextFrame.HasText:
                text = source_shape.TextFrame.TextRange.Text
                if text.strip():
                    textbox = target_slide.shapes.add_textbox(left, top, width, height)
                    self._copy_text_frame(source_shape, textbox)
                    return True

            return False

        except:
            return False

    def _handle_image_shape(self, source_shape, target_slide, temp_dir, left, top, width, height):
        """이미지 도형 처리"""
        # Export 시도
        try:
            img_path = os.path.join(temp_dir, f"img_{id(source_shape)}.png")
            source_shape.Export(img_path, 2)
            target_slide.shapes.add_picture(img_path, left, top, width, height)
            return True
        except:
            pass

        # 클립보드 시도
        for retry in range(3):
            try:
                source_shape.Copy()
                time.sleep(0.15)
                img_data = self._get_image_from_clipboard(temp_dir)
                if img_data:
                    target_slide.shapes.add_picture(img_data, left, top, width, height)
                    return True
            except:
                time.sleep(0.2)

        # Placeholder
        try:
            placeholder = target_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = RGBColor(220, 220, 220)
            placeholder.text_frame.paragraphs[0].text = "[이미지]"
            return True
        except:
            return False

    def _handle_autoshape(self, source_shape, target_slide, left, top, width, height):
        """AutoShape 처리"""
        try:
            auto_shape_type = source_shape.AutoShapeType
            pptx_shape_type = AUTOSHAPE_MAPPING.get(auto_shape_type, MSO_SHAPE.RECTANGLE)
            new_shape = target_slide.shapes.add_shape(pptx_shape_type, left, top, width, height)

            # 채우기
            try:
                if source_shape.Fill.Visible:
                    fill_rgb = source_shape.Fill.ForeColor.RGB
                    r, g, b = fill_rgb & 0xFF, (fill_rgb >> 8) & 0xFF, (fill_rgb >> 16) & 0xFF
                    new_shape.fill.solid()
                    new_shape.fill.fore_color.rgb = RGBColor(r, g, b)
            except:
                pass

            # 테두리
            try:
                if source_shape.Line.Visible:
                    line_rgb = source_shape.Line.ForeColor.RGB
                    r, g, b = line_rgb & 0xFF, (line_rgb >> 8) & 0xFF, (line_rgb >> 16) & 0xFF
                    new_shape.line.color.rgb = RGBColor(r, g, b)
                    new_shape.line.width = Pt(source_shape.Line.Weight)
            except:
                pass

            self._copy_text_frame(source_shape, new_shape)
            return True
        except:
            return False

    def _handle_table(self, source_shape, target_slide, left, top, width, height):
        """테이블 처리"""
        try:
            rows = source_shape.Table.Rows.Count
            cols = source_shape.Table.Columns.Count
            table = target_slide.shapes.add_table(rows, cols, left, top, width, height).table

            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    try:
                        cell_text = source_shape.Table.Cell(r, c).Shape.TextFrame.TextRange.Text
                        table.cell(r-1, c-1).text = cell_text
                    except:
                        pass
            return True
        except:
            return False

    def _handle_connector(self, source_shape, target_slide, left, top):
        """연결선 처리"""
        try:
            end_x = Emu(int((source_shape.Left + source_shape.Width) * 12700))
            end_y = Emu(int((source_shape.Top + source_shape.Height) * 12700))

            connector = target_slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, left, top, end_x, end_y
            )

            try:
                line_rgb = source_shape.Line.ForeColor.RGB
                r, g, b = line_rgb & 0xFF, (line_rgb >> 8) & 0xFF, (line_rgb >> 16) & 0xFF
                connector.line.color.rgb = RGBColor(r, g, b)
                connector.line.width = Pt(source_shape.Line.Weight)
            except:
                pass
            return True
        except:
            return False

    def _handle_freeform(self, source_shape, target_slide, temp_dir, left, top, width, height):
        """Freeform 처리"""
        for retry in range(3):
            try:
                source_shape.Copy()
                time.sleep(0.15)
                img_data = self._get_image_from_clipboard(temp_dir)
                if img_data:
                    target_slide.shapes.add_picture(img_data, left, top, width, height)
                    return True
            except:
                time.sleep(0.2)

        # 텍스트만
        if source_shape.HasTextFrame and source_shape.TextFrame.HasText:
            textbox = target_slide.shapes.add_textbox(left, top, width, height)
            self._copy_text_frame(source_shape, textbox)
            return True
        return False

    def _copy_text_frame(self, source_shape, target_shape):
        """텍스트 프레임 복사"""
        try:
            if not source_shape.HasTextFrame or not source_shape.TextFrame.HasText:
                return

            text = source_shape.TextFrame.TextRange.Text
            if not text.strip():
                return

            tf = target_shape.text_frame
            tf.word_wrap = True
            tf.paragraphs[0].text = text

            try:
                src_font = source_shape.TextFrame.TextRange.Font
                tf.paragraphs[0].font.size = Pt(src_font.Size) if src_font.Size else Pt(12)
            except:
                pass
        except:
            pass

    def _get_image_from_clipboard(self, temp_dir):
        """클립보드에서 이미지 추출"""
        try:
            import win32clipboard
            from PIL import Image
            import struct

            win32clipboard.OpenClipboard()
            try:
                # DIB
                if win32clipboard.IsClipboardFormatAvailable(8):
                    data = win32clipboard.GetClipboardData(8)
                    img = self._dib_to_image(data)
                    if img:
                        img_path = os.path.join(temp_dir, f"clipboard_{id(data)}.png")
                        img.save(img_path, "PNG")
                        return img_path

                # EMF
                if win32clipboard.IsClipboardFormatAvailable(14):
                    import ctypes
                    hemf = win32clipboard.GetClipboardData(14)
                    gdi32 = ctypes.windll.gdi32
                    size = gdi32.GetEnhMetaFileBits(hemf, 0, None)
                    if size > 0:
                        buffer = ctypes.create_string_buffer(size)
                        gdi32.GetEnhMetaFileBits(hemf, size, buffer)
                        emf_path = os.path.join(temp_dir, f"clipboard_{id(hemf)}.emf")
                        with open(emf_path, 'wb') as f:
                            f.write(buffer.raw)
                        return emf_path

                # PNG
                png_format = win32clipboard.RegisterClipboardFormat("PNG")
                if win32clipboard.IsClipboardFormatAvailable(png_format):
                    data = win32clipboard.GetClipboardData(png_format)
                    img_path = os.path.join(temp_dir, f"clipboard_{id(data)}.png")
                    with open(img_path, 'wb') as f:
                        f.write(data)
                    return img_path

            finally:
                win32clipboard.CloseClipboard()
        except:
            pass
        return None

    def _dib_to_image(self, dib_data):
        """DIB 데이터를 PIL Image로 변환"""
        try:
            from PIL import Image
            import struct

            header_size = struct.unpack('<I', dib_data[0:4])[0]
            width = struct.unpack('<i', dib_data[4:8])[0]
            height = struct.unpack('<i', dib_data[8:12])[0]
            bit_count = struct.unpack('<H', dib_data[14:16])[0]

            if bit_count == 32:
                pixel_data = dib_data[header_size:]
                img = Image.frombytes('RGBA', (width, abs(height)), pixel_data, 'raw', 'BGRA')
                if height > 0:
                    img = img.transpose(Image.FLIP_TOP_BOTTOM)
                return img

            elif bit_count == 24:
                row_size = ((width * 3 + 3) // 4) * 4
                pixel_data = dib_data[header_size:]
                img = Image.new('RGB', (width, abs(height)))
                pixels = img.load()

                for y in range(abs(height)):
                    src_y = abs(height) - 1 - y if height > 0 else y
                    row_start = src_y * row_size
                    for x in range(width):
                        b = pixel_data[row_start + x * 3]
                        g = pixel_data[row_start + x * 3 + 1]
                        r = pixel_data[row_start + x * 3 + 2]
                        pixels[x, y] = (r, g, b)
                return img
        except:
            pass
        return None

    def run(self):
        """프로그램 실행"""
        self.logger.log("메인 루프 시작")
        self.root.mainloop()
        self.logger.log("메인 루프 종료")
        self.logger.close()


def check_dependencies():
    """의존성 확인"""
    errors = []

    if not HAS_WIN32COM:
        errors.append("pywin32 패키지 필요 (pip install pywin32)")

    if not HAS_PPTX:
        errors.append("python-pptx 패키지 필요 (pip install python-pptx)")

    if not HAS_OPENPYXL:
        errors.append("openpyxl 패키지 필요 (pip install openpyxl)")

    if errors:
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning("패키지 필요",
            "다음 패키지를 설치하면 더 많은 기능을 사용할 수 있습니다:\n\n" + "\n".join(errors))

    # pywin32는 필수
    if not HAS_WIN32COM:
        messagebox.showerror("필수 패키지 누락",
            "pywin32 패키지가 필요합니다.\npip install pywin32")
        sys.exit(1)


if __name__ == "__main__":
    check_dependencies()
    app = DocumentExtractorV3()
    app.run()
