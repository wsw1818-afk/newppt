"""
문서 보안 해제 도구 v3
- PPT, Excel 지원
- COM으로 데이터 읽기
- python-pptx, openpyxl로 직접 파일 생성 (DRM 우회)
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
        self.log(f"지원 문서: PPT, Excel")
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
# COM AutoShapeType 상수: https://docs.microsoft.com/en-us/office/vba/api/office.msoautoshapetype
AUTOSHAPE_MAPPING = {
    # 기본 도형
    1: MSO_SHAPE.RECTANGLE,           # 사각형
    2: MSO_SHAPE.PARALLELOGRAM,       # 평행사변형
    3: MSO_SHAPE.TRAPEZOID,           # 사다리꼴
    4: MSO_SHAPE.DIAMOND,             # 마름모
    5: MSO_SHAPE.ROUNDED_RECTANGLE,   # 둥근 사각형
    6: MSO_SHAPE.OCTAGON,             # 팔각형
    7: MSO_SHAPE.ISOSCELES_TRIANGLE,  # 이등변 삼각형 (별칭)
    8: MSO_SHAPE.RIGHT_TRIANGLE,      # 직각 삼각형 (별칭)
    9: MSO_SHAPE.ISOSCELES_TRIANGLE,  # 이등변 삼각형
    10: MSO_SHAPE.RIGHT_TRIANGLE,     # 직각 삼각형
    11: MSO_SHAPE.OVAL,               # 타원/원
    12: MSO_SHAPE.HEXAGON,            # 육각형
    13: MSO_SHAPE.CROSS,              # 십자가
    14: MSO_SHAPE.REGULAR_PENTAGON,   # 오각형 (별칭)
    15: MSO_SHAPE.STAR_4_POINT,       # 4각 별 (별칭)
    16: MSO_SHAPE.CUBE,               # 정육면체

    # 화살표
    13: MSO_SHAPE.RIGHT_ARROW,        # 오른쪽 화살표
    14: MSO_SHAPE.LEFT_ARROW,         # 왼쪽 화살표
    15: MSO_SHAPE.UP_ARROW,           # 위쪽 화살표
    16: MSO_SHAPE.DOWN_ARROW,         # 아래쪽 화살표
    17: MSO_SHAPE.LEFT_RIGHT_ARROW,   # 좌우 화살표
    18: MSO_SHAPE.UP_DOWN_ARROW,      # 상하 화살표
    19: MSO_SHAPE.QUAD_ARROW,         # 4방향 화살표
    20: MSO_SHAPE.CHEVRON,            # 갈매기형
    21: MSO_SHAPE.NOTCHED_RIGHT_ARROW,# 홈이 있는 화살표
    22: MSO_SHAPE.PENTAGON,           # 오각형 (집 모양)
    23: MSO_SHAPE.CHEVRON,            # 갈매기형 (별칭)

    # 별
    12: MSO_SHAPE.STAR_5_POINT,       # 5각 별
    37: MSO_SHAPE.STAR_6_POINT,       # 6각 별
    38: MSO_SHAPE.STAR_8_POINT,       # 8각 별
    39: MSO_SHAPE.STAR_16_POINT,      # 16각 별
    40: MSO_SHAPE.STAR_24_POINT,      # 24각 별
    41: MSO_SHAPE.STAR_32_POINT,      # 32각 별

    # 블록 화살표
    24: MSO_SHAPE.RIGHT_ARROW_CALLOUT,    # 설명선 화살표
    25: MSO_SHAPE.LEFT_ARROW_CALLOUT,     # 설명선 화살표
    26: MSO_SHAPE.UP_ARROW_CALLOUT,       # 설명선 화살표
    27: MSO_SHAPE.DOWN_ARROW_CALLOUT,     # 설명선 화살표
    28: MSO_SHAPE.LEFT_RIGHT_ARROW_CALLOUT,
    29: MSO_SHAPE.UP_DOWN_ARROW_CALLOUT,

    # 설명선/말풍선
    30: MSO_SHAPE.ROUNDED_RECTANGLE,  # 둥근 사각형 설명선
    31: MSO_SHAPE.OVAL_CALLOUT,       # 타원 설명선
    32: MSO_SHAPE.CLOUD_CALLOUT,      # 구름 설명선

    # 기호
    33: MSO_SHAPE.HEART,              # 하트
    34: MSO_SHAPE.LIGHTNING_BOLT,     # 번개
    35: MSO_SHAPE.SUN,                # 태양
    36: MSO_SHAPE.MOON,               # 달

    # 도형 (추가)
    42: MSO_SHAPE.FOLDED_CORNER,      # 접힌 모서리
    43: MSO_SHAPE.SMILEY_FACE,        # 스마일
    44: MSO_SHAPE.NO_SYMBOL,          # 금지 표시
    45: MSO_SHAPE.BLOCK_ARC,          # 호
    46: MSO_SHAPE.DONUT,              # 도넛
    47: MSO_SHAPE.RECTANGLE,           # 기울어진 텍스트 (TEXT_SLANT 미지원 → 폴백)
    48: MSO_SHAPE.RECTANGLE,           # 아치형 텍스트 (TEXT_ARCH_DOWN_CURVE 미지원 → 폴백)
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
        self.hwp_save_path = tk.StringVar(value="")
        self.hwp_list = []
        self.selected_hwp_index = tk.IntVar(value=0)

        # 탭 변경 추적 (중복 감지 방지)
        self.last_tab_index = -1
        self.tab_detected = [False, False, False]  # PPT, Excel, 한글 각각 감지 완료 여부

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
                               text="PPT, Excel 문서의 DRM을 우회하여 새 파일로 저장합니다",
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
        self.notebook.add(self.hwp_tab, text="  한글  ")
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
        ttk.Label(select_frame, text="문서 선택:", width=12).pack(side=tk.LEFT)
        self.hwp_combo = ttk.Combobox(select_frame, state="readonly", width=40)
        self.hwp_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.hwp_combo.bind("<<ComboboxSelected>>", self.on_hwp_selected)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_hwp).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(tab, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.hwp_save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_hwp_save_path).pack(side=tk.LEFT)

        # 저장 형식 프레임
        format_frame = ttk.LabelFrame(tab, text="저장 형식", padding="10")
        format_frame.pack(fill=tk.X, pady=5, padx=5)

        self.hwp_save_format = tk.StringVar(value="hwp")
        ttk.Radiobutton(format_frame, text="HWP (한글 문서)",
                        variable=self.hwp_save_format, value="hwp").pack(anchor=tk.W)
        ttk.Radiobutton(format_frame, text="HWPX (한글 2014 이상)",
                        variable=self.hwp_save_format, value="hwpx").pack(anchor=tk.W)

        # 추출 버튼
        self.hwp_extract_button = ttk.Button(tab, text="새 한글 문서로 추출 (DRM 우회)",
                                             command=self.start_hwp_extraction,
                                             style="Accent.TButton")
        self.hwp_extract_button.pack(pady=10)

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
            ppt = None

            # 방법 1: GetObject (이미 실행 중인 인스턴스에 연결)
            try:
                ppt = win32com.client.GetObject(Class="PowerPoint.Application")
                self.logger.log("GetObject 연결 성공")
            except Exception as e1:
                self.logger.log(f"GetObject 실패: {str(e1)[:50]}")

                # 방법 2: GetActiveObject (COM Moniker 사용)
                try:
                    ppt = win32com.client.GetActiveObject("PowerPoint.Application")
                    self.logger.log("GetActiveObject 연결 성공")
                except Exception as e2:
                    self.logger.log(f"GetActiveObject 실패: {str(e2)[:50]}")

                    # 방법 3: Dispatch로 연결 시도 (이미 실행 중이면 기존 인스턴스에 연결됨)
                    try:
                        ppt = win32com.client.Dispatch("PowerPoint.Application")
                        self.logger.log("Dispatch 연결 성공")
                    except Exception as e3:
                        self.logger.log(f"Dispatch 실패: {str(e3)[:50]}")
                        raise Exception("PowerPoint에 연결할 수 없습니다. PowerPoint를 먼저 실행해주세요.")

            if ppt is None:
                raise Exception("PowerPoint 연결 실패")
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

            # 다양한 방법으로 PPT 연결 시도
            ppt_app = None
            try:
                ppt_app = win32com.client.GetObject(Class="PowerPoint.Application")
            except:
                try:
                    ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
                except:
                    ppt_app = win32com.client.Dispatch("PowerPoint.Application")

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

                # 슬라이드 배경색 복사
                self._copy_slide_background(source_slide, new_slide)

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

            # 다양한 방법으로 Excel 연결 시도
            excel_app = None
            try:
                excel_app = win32com.client.GetObject(Class="Excel.Application")
            except:
                try:
                    excel_app = win32com.client.GetActiveObject("Excel.Application")
                except:
                    excel_app = win32com.client.Dispatch("Excel.Application")

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
            # color가 None인 경우만 None 반환 (검정색 #000000은 유효한 색상)
            if color is None:
                return None
            # Excel 색상은 BGR 형식
            b = (color >> 16) & 0xFF
            g = (color >> 8) & 0xFF
            r = color & 0xFF
            return f"{r:02X}{g:02X}{b:02X}"
        except Exception as e:
            self.logger.error(f"색상 변환 실패: {color}", e)
            return None

    # ========== PPT 슬라이드 처리 메서드 ==========

    def _copy_slide_background(self, source_slide, target_slide):
        """슬라이드 배경색 복사"""
        try:
            bg = source_slide.Background
            fill = bg.Fill
            if fill.Visible:
                fill_type = fill.Type  # 1=solid
                if fill_type == 1:
                    fill_rgb = fill.ForeColor.RGB
                    r = fill_rgb & 0xFF
                    g = (fill_rgb >> 8) & 0xFF
                    b = (fill_rgb >> 16) & 0xFF
                    background = target_slide.background
                    background.fill.solid()
                    background.fill.fore_color.rgb = RGBColor(r, g, b)
        except:
            pass

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

        except Exception as e:
            self.logger.log(f"도형 재생성 실패 (타입 {shape_type}): {str(e)[:50]}")
            return False

    def _handle_image_shape(self, source_shape, target_slide, temp_dir, left, top, width, height):
        """이미지 도형 처리"""
        img_path = None

        # Export 시도
        try:
            img_path = os.path.join(temp_dir, f"img_{id(source_shape)}.png")
            source_shape.Export(img_path, 2)
            target_slide.shapes.add_picture(img_path, left, top, width, height)
            return True
        except Exception as e:
            self.logger.log(f"이미지 Export 실패: {str(e)[:50]}")
            # 실패 시 임시 파일 정리
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass

        # 클립보드 시도
        for retry in range(3):
            clipboard_img = None
            try:
                source_shape.Copy()
                time.sleep(0.15)
                clipboard_img = self._get_image_from_clipboard(temp_dir)
                if clipboard_img:
                    target_slide.shapes.add_picture(clipboard_img, left, top, width, height)
                    return True
            except Exception as e:
                self.logger.log(f"클립보드 이미지 추출 실패 (시도 {retry+1}/3): {str(e)[:50]}")
                time.sleep(0.2)
            finally:
                # 클립보드에서 생성된 임시 파일 정리
                if clipboard_img and os.path.exists(clipboard_img):
                    try:
                        os.remove(clipboard_img)
                    except Exception:
                        pass

        # Placeholder
        try:
            placeholder = target_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = RGBColor(220, 220, 220)
            placeholder.text_frame.paragraphs[0].text = "[이미지]"
            return True
        except Exception as e:
            self.logger.error("이미지 Placeholder 생성 실패", e)
            return False

    def _handle_autoshape(self, source_shape, target_slide, left, top, width, height):
        """AutoShape 처리"""
        try:
            auto_shape_type = source_shape.AutoShapeType
            pptx_shape_type = AUTOSHAPE_MAPPING.get(auto_shape_type, MSO_SHAPE.RECTANGLE)
            new_shape = target_slide.shapes.add_shape(pptx_shape_type, left, top, width, height)

            # 회전
            try:
                rotation = source_shape.Rotation
                if rotation and rotation != 0:
                    new_shape.rotation = rotation
            except:
                pass

            # 채우기
            try:
                fill = source_shape.Fill
                if fill.Visible:
                    fill_type = fill.Type  # 1=solid,2=gradient,3=texture,4=pattern,5=background
                    if fill_type == 1:  # solid
                        fill_rgb = fill.ForeColor.RGB
                        r, g, b = fill_rgb & 0xFF, (fill_rgb >> 8) & 0xFF, (fill_rgb >> 16) & 0xFF
                        new_shape.fill.solid()
                        new_shape.fill.fore_color.rgb = RGBColor(r, g, b)
                        # 투명도
                        try:
                            transparency = fill.Transparency
                            if transparency and transparency > 0:
                                new_shape.fill.fore_color.theme_color  # 접근 확인
                                # XML로 직접 투명도 설정
                                from lxml import etree
                                from pptx.oxml.ns import qn as _qn
                                spPr = new_shape._element.spPr
                                solidFill = spPr.find('.//' + _qn('a:solidFill'))
                                if solidFill is not None:
                                    srgbClr = solidFill.find(_qn('a:srgbClr'))
                                    if srgbClr is not None:
                                        alpha = etree.SubElement(srgbClr, _qn('a:alpha'))
                                        alpha.set('val', str(int((1 - transparency) * 100000)))
                        except:
                            pass
                    else:
                        # gradient/pattern → 단색 근사
                        try:
                            fill_rgb = fill.ForeColor.RGB
                            r, g, b = fill_rgb & 0xFF, (fill_rgb >> 8) & 0xFF, (fill_rgb >> 16) & 0xFF
                            new_shape.fill.solid()
                            new_shape.fill.fore_color.rgb = RGBColor(r, g, b)
                        except:
                            pass
                else:
                    new_shape.fill.background()  # 투명
            except:
                pass

            # 테두리
            try:
                line = source_shape.Line
                if line.Visible:
                    line_rgb = line.ForeColor.RGB
                    r, g, b = line_rgb & 0xFF, (line_rgb >> 8) & 0xFF, (line_rgb >> 16) & 0xFF
                    new_shape.line.color.rgb = RGBColor(r, g, b)
                    new_shape.line.width = Pt(line.Weight)
                else:
                    new_shape.line.fill.background()  # 테두리 없음
            except:
                pass

            self._copy_text_frame(source_shape, new_shape)
            return True
        except Exception as e:
            self.logger.log(f"AutoShape 처리 실패: {str(e)[:60]}")
            return False

    def _handle_table(self, source_shape, target_slide, left, top, width, height):
        """테이블 처리 — 셀 텍스트·배경색·폰트 복사"""
        try:
            src_table = source_shape.Table
            rows = src_table.Rows.Count
            cols = src_table.Columns.Count
            tbl_shape = target_slide.shapes.add_table(rows, cols, left, top, width, height)
            table = tbl_shape.table

            # 열 너비 복사
            try:
                for c in range(1, cols + 1):
                    col_width = src_table.Columns(c).Width
                    table.columns[c-1].width = Emu(int(col_width * 12700))
            except:
                pass

            # 행 높이 복사
            try:
                for r in range(1, rows + 1):
                    row_height = src_table.Rows(r).Height
                    table.rows[r-1].height = Emu(int(row_height * 12700))
            except:
                pass

            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    try:
                        src_cell = src_table.Cell(r, c)
                        dst_cell = table.cell(r-1, c-1)
                        src_cell_shape = src_cell.Shape

                        # 텍스트+서식 복사
                        self._copy_text_frame(src_cell_shape, dst_cell)

                        # 셀 배경색
                        try:
                            fill = src_cell_shape.Fill
                            if fill.Visible and fill.Type == 1:
                                fill_rgb = fill.ForeColor.RGB
                                r_c = fill_rgb & 0xFF
                                g_c = (fill_rgb >> 8) & 0xFF
                                b_c = (fill_rgb >> 16) & 0xFF
                                dst_cell.fill.solid()
                                dst_cell.fill.fore_color.rgb = RGBColor(r_c, g_c, b_c)
                        except:
                            pass

                    except:
                        pass
            return True
        except Exception as e:
            self.logger.log(f"테이블 처리 실패: {str(e)[:60]}")
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
        img_path = None
        for retry in range(3):
            try:
                source_shape.Copy()
                time.sleep(0.15)
                img_path = self._get_image_from_clipboard(temp_dir)
                if img_path:
                    target_slide.shapes.add_picture(img_path, left, top, width, height)
                    return True
            except Exception as e:
                self.logger.log(f"Freeform 클립보드 추출 실패 (시도 {retry+1}/3): {str(e)[:50]}")
                time.sleep(0.2)
            finally:
                # 임시 파일 정리
                if img_path and os.path.exists(img_path):
                    try:
                        os.remove(img_path)
                    except Exception:
                        pass

        # 텍스트만
        if source_shape.HasTextFrame and source_shape.TextFrame.HasText:
            textbox = target_slide.shapes.add_textbox(left, top, width, height)
            self._copy_text_frame(source_shape, textbox)
            return True
        return False

    def _copy_text_frame(self, source_shape, target_shape):
        """텍스트 프레임 복사 — 단락/런 단위 서식 보존"""
        try:
            if not source_shape.HasTextFrame or not source_shape.TextFrame.HasText:
                return

            src_tf = source_shape.TextFrame
            dst_tf = target_shape.text_frame
            dst_tf.word_wrap = True

            # 수직 정렬
            try:
                va = src_tf.TextRange.ParagraphFormat.Alignment  # COM 수직은 별도
            except:
                pass
            try:
                anchor_map = {1: MSO_ANCHOR.TOP, 3: MSO_ANCHOR.MIDDLE, 4: MSO_ANCHOR.BOTTOM}
                va_val = src_tf.VerticalAnchor  # 1=top,2=middle(없음),3=middle,4=bottom
                if va_val in anchor_map:
                    dst_tf.vertical_anchor = anchor_map[va_val]
            except:
                pass

            para_count = src_tf.TextRange.Paragraphs().Count

            # 기존 단락 초기화
            from pptx.oxml.ns import qn as _qn
            from lxml import etree
            txBody = dst_tf._txBody
            for old_p in txBody.findall(_qn('a:p')):
                txBody.remove(old_p)

            for pi in range(1, para_count + 1):
                src_para = src_tf.TextRange.Paragraphs(pi)

                # 새 단락 XML 엘리먼트 생성
                new_p = etree.SubElement(txBody, _qn('a:p'))

                # 단락 서식 (pPr)
                try:
                    pPr = etree.SubElement(new_p, _qn('a:pPr'))
                    align_map = {1: 'l', 2: 'ctr', 3: 'r', 4: 'dist', 5: 'just'}
                    align_val = src_para.ParagraphFormat.Alignment
                    if align_val in align_map:
                        pPr.set('algn', align_map[align_val])
                    indent = src_para.ParagraphFormat.Indent
                    if indent and indent != 0:
                        pPr.set('indent', str(int(indent * 12700)))
                    space_before = src_para.ParagraphFormat.SpaceBefore
                    if space_before and space_before > 0:
                        spcBef = etree.SubElement(pPr, _qn('a:spcBef'))
                        spcPts = etree.SubElement(spcBef, _qn('a:spcPts'))
                        spcPts.set('val', str(int(space_before * 100)))
                except:
                    pass

                # 런(Run) 단위 텍스트+서식
                try:
                    run_count = src_para.Runs.Count
                    if run_count == 0:
                        # 빈 단락
                        etree.SubElement(new_p, _qn('a:endParaRPr'), attrib={'lang': 'ko-KR'})
                        continue

                    for ri in range(1, run_count + 1):
                        src_run = src_para.Runs(ri)
                        run_text = src_run.Text
                        src_font = src_run.Font

                        new_r = etree.SubElement(new_p, _qn('a:r'))
                        rPr = etree.SubElement(new_r, _qn('a:rPr'), attrib={'lang': 'ko-KR', 'dirty': '0'})

                        # 굵게/기울임/밑줄
                        try:
                            if src_font.Bold:
                                rPr.set('b', '1')
                        except: pass
                        try:
                            if src_font.Italic:
                                rPr.set('i', '1')
                        except: pass
                        try:
                            if src_font.Underline:
                                rPr.set('u', 'sng')
                        except: pass
                        # 글꼴 크기
                        try:
                            if src_font.Size:
                                rPr.set('sz', str(int(src_font.Size * 100)))
                        except: pass
                        # 글꼴 이름
                        try:
                            if src_font.Name:
                                latin = etree.SubElement(rPr, _qn('a:latin'))
                                latin.set('typeface', src_font.Name)
                        except: pass
                        # 글자 색상
                        try:
                            rgb_val = src_font.Color.RGB
                            r = rgb_val & 0xFF
                            g = (rgb_val >> 8) & 0xFF
                            b = (rgb_val >> 16) & 0xFF
                            solidFill = etree.SubElement(rPr, _qn('a:solidFill'))
                            srgbClr = etree.SubElement(solidFill, _qn('a:srgbClr'))
                            srgbClr.set('val', f'{r:02X}{g:02X}{b:02X}')
                        except: pass

                        # 텍스트
                        t_elem = etree.SubElement(new_r, _qn('a:t'))
                        t_elem.text = run_text if run_text else ''

                except Exception as re:
                    # 런 읽기 실패 시 전체 단락 텍스트로 폴백
                    try:
                        para_text = src_para.Text
                        new_r = etree.SubElement(new_p, _qn('a:r'))
                        rPr = etree.SubElement(new_r, _qn('a:rPr'), attrib={'lang': 'ko-KR'})
                        try:
                            sz = src_tf.TextRange.Font.Size
                            if sz:
                                rPr.set('sz', str(int(sz * 100)))
                        except: pass
                        t_elem = etree.SubElement(new_r, _qn('a:t'))
                        t_elem.text = para_text if para_text else ''
                    except:
                        pass

        except Exception as e:
            self.logger.log(f"텍스트 프레임 복사 실패: {str(e)[:80]}")

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

    # ========== 한글 관련 메서드 ==========

    def browse_hwp_save_path(self):
        """한글 저장 경로 선택"""
        self.logger.log("한글 저장 경로 선택 대화상자 열기")

        doc_name = self.hwp_doc_name.get()
        save_format = self.hwp_save_format.get()

        if doc_name and doc_name != "감지 중..." and doc_name != "열린 한글 없음":
            default_name = os.path.splitext(doc_name)[0] + "_복사본"
        else:
            default_name = "새문서"

        if save_format == "hwpx":
            ext = ".hwpx"
            filetypes = [("한글 2014+ 파일", "*.hwpx")]
        else:
            ext = ".hwp"
            filetypes = [("한글 파일", "*.hwp")]

        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=filetypes,
            initialfile=default_name,
            title="저장할 위치 선택"
        )
        if path:
            self.hwp_save_path.set(path)
            self.logger.log(f"한글 저장 경로 선택됨: {path}")

    def detect_open_hwp(self):
        """열려있는 한글 감지"""
        self.logger.log("한글 감지 시작")
        self.status_text.set("한글 감지 중...")
        self.hwp_doc_name.set("감지 중...")

        thread = threading.Thread(target=self._detect_hwp)
        thread.daemon = True
        thread.start()

    def _detect_hwp(self):
        """한글 감지 (백그라운드)"""
        self.logger.log("백그라운드 한글 감지 스레드 시작")
        pythoncom.CoInitialize()

        try:
            self.logger.log("HWPFrame.HwpObject 객체 연결 시도")
            hwp = None

            # 방법 1: GetObject (이미 실행 중인 인스턴스에 연결)
            try:
                hwp = win32com.client.GetObject(Class="HWPFrame.HwpObject")
                self.logger.log("GetObject 연결 성공")
            except Exception as e1:
                self.logger.log(f"GetObject 실패: {str(e1)[:50]}")

                # 방법 2: Dispatch로 연결 시도
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                    self.logger.log("Dispatch 연결 성공")
                except Exception as e2:
                    self.logger.log(f"Dispatch 실패: {str(e2)[:50]}")
                    raise Exception("한글에 연결할 수 없습니다. 한글을 먼저 실행해주세요.")

            if hwp is None:
                raise Exception("한글 연결 실패")

            # 열린 문서 확인
            try:
                path = hwp.Path
                if path:
                    name = os.path.basename(path)
                    self.hwp_list = [(name, path, 1)]
                    self.logger.log(f"한글 문서 감지: {name}")

                    def update_combo():
                        self.hwp_combo['values'] = [name]
                        self.hwp_combo.current(0)
                        self.selected_hwp_index.set(1)
                        self.hwp_doc_name.set(name)
                        self.status_text.set("한글 문서 감지됨")

                    self.root.after(0, update_combo)
                else:
                    # 제목 없는 문서
                    self.hwp_list = [("제목 없음", "", 1)]
                    self.logger.log("한글 제목 없는 문서 감지")

                    def update_combo():
                        self.hwp_combo['values'] = ["제목 없음"]
                        self.hwp_combo.current(0)
                        self.selected_hwp_index.set(1)
                        self.hwp_doc_name.set("제목 없음")
                        self.status_text.set("한글 문서 감지됨")

                    self.root.after(0, update_combo)
            except Exception as e:
                self.logger.log(f"한글 문서 정보 가져오기 실패: {str(e)}")
                self.hwp_list = []

                def clear_combo():
                    self.hwp_combo.set("")
                    self.hwp_combo['values'] = []
                    self.hwp_doc_name.set("열린 한글 없음")
                    self.status_text.set("한글을 먼저 열어주세요")

                self.root.after(0, clear_combo)

        except Exception as e:
            self.logger.error("한글 감지 실패", e)
            self.hwp_list = []
            err_msg = str(e)[:30]

            def show_error():
                self.hwp_combo.set("")
                self.hwp_doc_name.set("열린 한글 없음")
                self.status_text.set(f"한글 감지 실패: {err_msg}")

            self.root.after(0, show_error)

        pythoncom.CoUninitialize()

    def on_hwp_selected(self, event):
        """한글 콤보박스 선택 이벤트"""
        selected_idx = self.hwp_combo.current()
        if selected_idx >= 0 and selected_idx < len(self.hwp_list):
            name, path, hwp_index = self.hwp_list[selected_idx]
            self.selected_hwp_index.set(hwp_index)
            self.hwp_doc_name.set(name)
            self.logger.log(f"한글 선택: {name} (인덱스 {hwp_index})")

    def start_hwp_extraction(self):
        """한글 추출 시작"""
        self.logger.log("한글 추출 시작 버튼 클릭")

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
        self.logger.log("=== 한글 추출 프로세스 시작 ===")
        pythoncom.CoInitialize()

        try:
            save_path = self.hwp_save_path.get()
            save_format = self.hwp_save_format.get()
            self.root.after(0, lambda: self.status_text.set("원본 한글 연결 중..."))

            # 한글 연결
            hwp = None
            try:
                hwp = win32com.client.GetObject(Class="HWPFrame.HwpObject")
            except Exception:
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                except Exception:
                    raise Exception("한글에 연결할 수 없습니다.")

            self.logger.log(f"원본 한글 문서 연결 성공")
            self.root.after(0, lambda: self.status_text.set("새 문서로 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(30))

            # 방법 1: SaveAs 시도
            try:
                self.logger.log(f"SaveAs 시도: {save_path}")
                if save_format == "hwpx":
                    hwp.SaveAs(save_path, "HWPX")
                else:
                    hwp.SaveAs(save_path, "HWP")

                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, lambda: self.status_text.set("한글 추출 완료!"))
                self.root.after(0, lambda: messagebox.showinfo("완료",
                    f"한글 추출 완료!\n{save_path}"))
                self.logger.log(f"저장 완료: {save_path}")

            except Exception as e:
                self.logger.log(f"SaveAs 실패: {str(e)}")

                # 방법 2: 클립보드를 통한 복사 시도
                self.root.after(0, lambda: self.status_text.set("클립보드 복사 시도 중..."))
                self.root.after(0, lambda: self.progress_var.set(50))

                try:
                    # 전체 선택
                    hwp.HAction.Run("SelectAll")
                    # 복사
                    hwp.HAction.Run("Copy")

                    # 새 문서 생성
                    hwp.HAction.Run("FileNew")
                    # 붙여넣기
                    hwp.HAction.Run("Paste")

                    # 저장
                    if save_format == "hwpx":
                        hwp.SaveAs(save_path, "HWPX")
                    else:
                        hwp.SaveAs(save_path, "HWP")

                    self.root.after(0, lambda: self.progress_var.set(100))
                    self.root.after(0, lambda: self.status_text.set("한글 추출 완료!"))
                    self.root.after(0, lambda: messagebox.showinfo("완료",
                        f"한글 추출 완료 (클립보드 방식)!\n{save_path}"))
                    self.logger.log(f"클립보드 방식으로 저장 완료: {save_path}")

                except Exception as e2:
                    raise Exception(f"한글 저장 실패:\n{str(e2)}")

        except Exception as e:
            self.logger.error("한글 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.hwp_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

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
