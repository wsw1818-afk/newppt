"""
문서 추출 도구 v3
- PPT, Excel, Word, 메모장 지원
- COM으로 데이터 읽기
- python-pptx, openpyxl, python-docx로 직접 파일 생성
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
import zipfile
import ctypes
from ctypes import wintypes

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
    from openpyxl.drawing.image import Image as OpenpyxlImage
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# python-docx 관련 (Word)
try:
    from docx import Document as DocxDocument
    from docx.shared import Pt as DocxPt, RGBColor as DocxRGBColor, Emu as DocxEmu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

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
        self._lock = threading.Lock()
        self._line_count = 0
        self._closed = False
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = os.path.join(desktop, f"DocExtractor_Log_{timestamp}.txt")
        self.log_file = open(self.log_path, "w", encoding="utf-8")
        self.log(f"=== 문서 추출기 v3 로그 시작 ===")
        self.log(f"로그 파일: {self.log_path}")
        self.log(f"시작 시간: {datetime.datetime.now()}")
        self.log(f"지원 문서: PPT, Excel, 한글, Word, 메모장")
        self.log(f"한글 저장: FileSaveAs_S 액션 우선")
        self.log("")

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{timestamp}] {message}"
        with self._lock:
            if self._closed:
                return
            try:
                if sys.stdout:
                    print(line)
            except Exception:
                pass
            try:
                self.log_file.write(line + "\n")
                self._line_count += 1
                self.log_file.flush()
            except Exception:
                pass

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
        with self._lock:
            try:
                self.log_file.flush()
                self.log_file.close()
            finally:
                self._closed = True


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
    # 화살표/기타 기본값
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

    EXCEL_VALUE_CELL_LIMIT = 500_000
    EXCEL_FORMAT_CELL_LIMIT = 50_000
    EXCEL_ROW_HEIGHT_COPY_LIMIT = 5_000
    EXCEL_COLUMN_WIDTH_COPY_LIMIT = 1_024
    EXCEL_OBJECT_RETRY_COUNT = 2
    EXCEL_OBJECT_RETRY_DELAY = 0.08
    PPT_CLIPBOARD_RETRY_COUNT = 2
    PPT_CLIPBOARD_RETRY_DELAY = 0.08

    def __init__(self):
        self.logger = Logger()
        self.logger.log("DocumentExtractor v3 초기화 시작")

        self.root = tk.Tk()
        self.root.title("문서 추출 도구 v3")
        self.root.geometry("650x620")
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

        # Word 상태 변수
        self.word_doc_name = tk.StringVar(value="감지 중...")
        self.word_page_count = tk.StringVar(value="-")
        self.word_save_path = tk.StringVar(value="")
        self.word_list = []
        self.selected_word_index = tk.IntVar(value=0)

        # 메모장 상태 변수
        self.notepad_doc_name = tk.StringVar(value="감지 중...")
        self.notepad_save_path = tk.StringVar(value="")
        self.notepad_list = []

        # 탭 변경 추적 (중복 감지 방지)
        self.last_tab_index = -1
        self.tab_detected = [False, False, False, False, False]  # PPT, Excel, 한글, Word, 메모장
        self._hwp_detecting = False

        self.setup_ui()

        self.logger.log("DocumentExtractor v3 초기화 완료")

    def setup_ui(self):
        """UI 구성"""
        self.logger.log("UI 구성 시작")

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(main_frame, text="문서 추출 도구 v3",
                                font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 5))

        # 설명
        desc_label = ttk.Label(main_frame,
                               text="권한 있는 열린 문서의 내용을 새 파일로 내보냅니다",
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

        # Word 탭
        self.word_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.word_tab, text="  Word  ")
        self._setup_word_tab()

        # 메모장 탭
        self.notepad_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.notepad_tab, text="  메모장  ")
        self._setup_notepad_tab()

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

    def _connect_com_app(self, prog_id, display_name, allow_dispatch=True, use_get_active=True):
        """Office/HWP COM 애플리케이션 연결을 공통 처리한다."""
        try:
            app = win32com.client.GetObject(Class=prog_id)
            self.logger.log(f"{display_name} GetObject 연결 성공")
            return app, False
        except Exception as e1:
            self.logger.log(f"{display_name} GetObject 실패: {str(e1)[:50]}")

        if use_get_active:
            try:
                app = win32com.client.GetActiveObject(prog_id)
                self.logger.log(f"{display_name} GetActiveObject 연결 성공")
                return app, False
            except Exception as e2:
                self.logger.log(f"{display_name} GetActiveObject 실패: {str(e2)[:50]}")

        if allow_dispatch:
            try:
                app = win32com.client.Dispatch(prog_id)
                self.logger.log(f"{display_name} Dispatch 연결 성공")
                return app, True
            except Exception as e3:
                self.logger.log(f"{display_name} Dispatch 실패: {str(e3)[:50]}")

        raise Exception(f"{display_name}에 연결할 수 없습니다. {display_name}를 먼저 실행해주세요.")

    def _get_ppt_app(self, allow_dispatch=True):
        return self._connect_com_app("PowerPoint.Application", "PowerPoint", allow_dispatch=allow_dispatch)

    def _get_excel_app(self, allow_dispatch=True):
        return self._connect_com_app("Excel.Application", "Excel", allow_dispatch=allow_dispatch)

    def _get_hwp_app(self, allow_dispatch=True):
        return self._connect_com_app(
            "HWPFrame.HwpObject",
            "한글",
            allow_dispatch=allow_dispatch,
            use_get_active=False,
        )

    def _get_hwp_app_for_extraction(self):
        """추출용 HWP 연결. 새 빈 문서가 만들어지는 연결은 차단한다."""
        before_titles = self._list_hwp_window_titles()
        if not before_titles:
            raise Exception("열린 한글 문서가 없습니다. 한글에서 문서를 먼저 열어주세요.")

        try:
            return self._get_hwp_app(allow_dispatch=False)
        except Exception as active_error:
            self.logger.log(f"한글 활성 COM 연결 실패, Dispatch 연결 검증 시도: {str(active_error)[:80]}")

        hwp, created = self._get_hwp_app(allow_dispatch=True)
        after_titles = self._list_hwp_window_titles()
        if created and len(after_titles) > len(before_titles):
            raise Exception(
                "한글 COM 연결 중 새 빈 문서가 생성되어 중단했습니다.\n"
                "기존 한글 문서를 닫았다가 다시 열거나, 프로그램과 한글을 같은 권한으로 실행해 주세요."
            )

        return hwp, created

    def _get_visible_window_titles(self):
        """현재 보이는 최상위 창 제목 목록을 가져온다."""
        titles = []
        user32 = ctypes.windll.user32
        enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def enum_proc(hwnd, lparam):
            try:
                if not user32.IsWindowVisible(hwnd):
                    return True
                length = user32.GetWindowTextLengthW(hwnd)
                if length <= 0:
                    return True
                buffer = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buffer, length + 1)
                title = buffer.value.strip()
                if title:
                    titles.append(title)
            except Exception:
                pass
            return True

        user32.EnumWindows(enum_proc_type(enum_proc), 0)
        return titles

    def _list_hwp_window_titles(self):
        """한글 COM이 ROT에 없을 때 새 인스턴스를 만들지 않고 창 제목만 감지한다."""
        hwp_titles = []
        for title in self._get_visible_window_titles():
            normalized = title.strip()
            if normalized.endswith(" - 한글") or " - 한글 " in normalized:
                hwp_titles.append(normalized)
        return hwp_titles

    def _log_elapsed(self, label, start_time):
        elapsed = time.perf_counter() - start_time
        self.logger.log(f"{label}: {elapsed:.2f}초")
        return elapsed

    def _read_header_hex(self, path, size=16):
        try:
            with open(path, "rb") as f:
                return " ".join(f"{b:02X}" for b in f.read(size))
        except Exception:
            return "읽기 실패"

    def _add_filename_suffix(self, path, suffix):
        directory = os.path.dirname(path)
        stem, ext = os.path.splitext(os.path.basename(path))
        return os.path.join(directory, f"{stem}{suffix}{ext}")

    def _make_local_temp_path(self, suffix):
        """OneDrive/DRM 감시 폴더를 피해서 로컬 임시 파일 경로를 만든다."""
        fd, temp_path = tempfile.mkstemp(prefix="docextract_", suffix=suffix)
        os.close(fd)
        try:
            os.remove(temp_path)
        except Exception:
            pass
        return temp_path

    def _publish_verified_file(self, temp_path, save_path, label):
        """검증된 로컬 임시 파일을 최종 위치로 복사하고 최종본도 다시 검증한다."""
        self._validate_office_openxml(temp_path, label)
        target_dir = os.path.dirname(os.path.abspath(save_path)) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)

        ext = os.path.splitext(save_path)[1] or ".tmp"
        fd, stage_path = tempfile.mkstemp(prefix=".docextract_", suffix=ext, dir=target_dir)
        os.close(fd)
        try:
            shutil.copy2(temp_path, stage_path)
            self._validate_office_openxml(stage_path, f"{label} 최종 복사")
            os.replace(stage_path, save_path)
            self._validate_office_openxml(save_path, f"{label} 최종 파일")
        except Exception:
            if os.path.exists(stage_path):
                try:
                    os.remove(stage_path)
                except Exception:
                    pass
            raise

    def _validate_office_openxml(self, path, label):
        """확장자가 OpenXML이면 실제 ZIP 패키지인지 확인한다."""
        ext = os.path.splitext(path)[1].lower()
        required_members = {
            ".pptx": ["[Content_Types].xml", "ppt/presentation.xml"],
            ".pptm": ["[Content_Types].xml", "ppt/presentation.xml"],
            ".ppsx": ["[Content_Types].xml", "ppt/presentation.xml"],
            ".potx": ["[Content_Types].xml", "ppt/presentation.xml"],
            ".xlsx": ["[Content_Types].xml", "xl/workbook.xml"],
            ".xlsm": ["[Content_Types].xml", "xl/workbook.xml"],
            ".xlsb": ["[Content_Types].xml", "xl/workbook.bin"],
            ".docx": ["[Content_Types].xml", "word/document.xml"],
            ".docm": ["[Content_Types].xml", "word/document.xml"],
        }.get(ext)
        if not required_members:
            return

        if not os.path.exists(path) or os.path.getsize(path) == 0:
            raise Exception(f"{label} 저장 결과 파일이 없거나 비어 있습니다.")

        header = self._read_header_hex(path)
        if not zipfile.is_zipfile(path):
            hint = ""
            if header.startswith("53 43 44 53"):
                hint = " DRM/보안 컨테이너(SCDS)로 저장된 것으로 보입니다."
            raise Exception(
                f"{label} 저장 결과가 정상 {ext} 파일이 아닙니다. header={header}.{hint}"
            )

        with zipfile.ZipFile(path) as archive:
            bad_member = archive.testzip()
            if bad_member:
                raise Exception(f"{label} 저장 결과 ZIP 항목 손상: {bad_member}")
            names = set(archive.namelist())
            missing = [member for member in required_members if member not in names]
            if missing:
                raise Exception(f"{label} 저장 결과에 필수 항목이 없습니다: {missing}")

    def _run_with_heartbeat(self, label, func, interval=10, warn_after=30):
        """긴 COM 호출 동안 로그 파일이 멈춘 것처럼 보이지 않게 주기 로그를 남긴다."""
        stop_event = threading.Event()
        start_time = time.perf_counter()

        def heartbeat():
            warned = False
            while not stop_event.wait(interval):
                elapsed = time.perf_counter() - start_time
                self.logger.log(f"{label} 진행 중... {elapsed:.0f}초 경과")
                if not warned and elapsed >= warn_after:
                    self.logger.log(
                        f"{label} 지연 중: Office 저장 대화상자나 OneDrive 동기화 상태를 확인해 주세요."
                    )
                    warned = True

        thread = threading.Thread(target=heartbeat, daemon=True)
        thread.start()
        try:
            return func()
        finally:
            stop_event.set()

    def _save_native_copy(self, source_doc, save_path, label):
        """Office의 원본 복사 기능으로 서식/개체를 가장 정확하게 보존한다."""
        target_dir = os.path.dirname(os.path.abspath(save_path)) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)

        ext = os.path.splitext(save_path)[1] or ".tmp"
        temp_path = self._make_local_temp_path(ext)
        try:
            self.logger.log(f"{label} 원본 복사 로컬 임시 저장: {temp_path}")
            self._run_with_heartbeat(
                f"{label} SaveCopyAs",
                lambda: source_doc.SaveCopyAs(temp_path),
            )
            self._publish_verified_file(temp_path, save_path, label)
            self.logger.log(f"{label} 원본 복사 저장 완료: {save_path}")
        except Exception:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            raise

    def _get_powerpoint_clipboard_package(self):
        """PowerPoint slide copy clipboard package bytes."""
        try:
            import win32clipboard
        except Exception as import_error:
            raise Exception(f"win32clipboard import failed: {import_error}")

        package_format = win32clipboard.RegisterClipboardFormat("PowerPoint 14.0 Slides Package")
        last_error = None
        for _ in range(20):
            try:
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(package_format):
                        data = win32clipboard.GetClipboardData(package_format)
                        if isinstance(data, bytes) and len(data) > 0:
                            return data
                        last_error = Exception("PowerPoint clipboard package is empty")
                    else:
                        last_error = Exception("PowerPoint clipboard package format is not available")
                finally:
                    win32clipboard.CloseClipboard()
            except Exception as error:
                last_error = error
            time.sleep(0.1)
        raise Exception(str(last_error))

    def _write_ppt_clipboard_package_as_pptx(self, package_data, temp_path, target_ext=".pptx"):
        """Convert the PowerPoint clipboard OPC package into a regular PPTX package."""
        if not zipfile.is_zipfile(io.BytesIO(package_data)):
            header = " ".join(f"{byte:02X}" for byte in package_data[:16])
            raise Exception(f"PowerPoint clipboard package is not ZIP. header={header}")

        macro_content_type = b"application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
        pptx_content_type = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        with zipfile.ZipFile(io.BytesIO(package_data), "r") as source_archive:
            names = set(source_archive.namelist())
            if "clipboard/presentation.xml" not in names:
                raise Exception("PowerPoint clipboard package has no presentation.xml")

            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target_archive:
                for item in source_archive.infolist():
                    if item.is_dir():
                        continue
                    name = item.filename
                    data = source_archive.read(name)
                    if name.startswith("clipboard/"):
                        name = "ppt/" + name[len("clipboard/"):]

                    data = data.replace(b"/clipboard/", b"/ppt/")
                    data = data.replace(b'Target="clipboard/', b'Target="ppt/')
                    data = data.replace(b"Target='clipboard/", b"Target='ppt/")
                    if target_ext.lower() != ".pptm":
                        data = data.replace(macro_content_type, pptx_content_type)
                    target_archive.writestr(name, data)

    def _save_ppt_clipboard_package_copy(self, source_pres, save_path):
        """Copy slides to clipboard and rebuild the PowerPoint slide package without SaveAs."""
        temp_path = self._make_local_temp_path(".pptx")
        try:
            total_slides = source_pres.Slides.Count
            slide_indices = tuple(range(1, total_slides + 1))
            self.logger.log("PPT 클립보드 슬라이드 패키지 복원 시도")

            last_error = None
            package_data = None
            for retry in range(1, self.PPT_CLIPBOARD_RETRY_COUNT + 2):
                try:
                    source_pres.Slides.Range(slide_indices).Copy()
                    time.sleep(max(self.PPT_CLIPBOARD_RETRY_DELAY, 0.2))
                    package_data = self._get_powerpoint_clipboard_package()
                    break
                except Exception as error:
                    last_error = error
                    if retry <= self.PPT_CLIPBOARD_RETRY_COUNT:
                        time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)

            if not package_data:
                raise Exception(f"PPT 클립보드 슬라이드 패키지를 가져오지 못했습니다: {last_error}")

            target_ext = os.path.splitext(save_path)[1].lower() or ".pptx"
            self._write_ppt_clipboard_package_as_pptx(package_data, temp_path, target_ext)
            self._publish_verified_file(temp_path, save_path, "PPT 클립보드 슬라이드 패키지")
            self.logger.log(f"PPT 클립보드 슬라이드 패키지 복원 완료: {save_path}")
        except Exception:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            raise

    def _save_ppt_slide_clone(self, source_pres, save_path):
        """PowerPoint 내부 복사/붙여넣기로 슬라이드를 최대한 원본에 가깝게 복제한다."""
        target_dir = os.path.dirname(os.path.abspath(save_path)) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)

        temp_path = self._make_local_temp_path(".pptx")

        app = source_pres.Application
        target_pres = None
        try:
            self.logger.log(f"PPT 슬라이드 복제 로컬 임시 저장: {temp_path}")
            target_pres = app.Presentations.Add()

            try:
                target_pres.PageSetup.SlideWidth = source_pres.PageSetup.SlideWidth
                target_pres.PageSetup.SlideHeight = source_pres.PageSetup.SlideHeight
            except Exception as size_error:
                self.logger.log(f"PPT 슬라이드 크기 복사 실패: {str(size_error)[:60]}")

            while target_pres.Slides.Count > 0:
                target_pres.Slides(1).Delete()

            total_slides = source_pres.Slides.Count
            slide_indices = tuple(range(1, total_slides + 1))

            def copy_slide_range(indices, insert_index):
                last_error = None
                for retry in range(1, self.PPT_CLIPBOARD_RETRY_COUNT + 2):
                    try:
                        source_pres.Slides.Range(indices).Copy()
                        time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
                        target_pres.Slides.Paste(insert_index)
                        return
                    except Exception as paste_error:
                        last_error = paste_error
                        if retry >= self.PPT_CLIPBOARD_RETRY_COUNT + 1:
                            raise
                        time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
                if last_error:
                    raise last_error

            try:
                copy_slide_range(slide_indices, 1)
                self.logger.log(f"  PPT 슬라이드 일괄 복제: {total_slides}/{total_slides}")
            except Exception as bulk_error:
                self.logger.log(f"PPT 슬라이드 일괄 복제 실패, 개별 복제 전환: {str(bulk_error)[:80]}")
                while target_pres.Slides.Count > 0:
                    target_pres.Slides(1).Delete()
                for slide_idx in range(1, total_slides + 1):
                    try:
                        copy_slide_range((slide_idx,), target_pres.Slides.Count + 1)
                    except Exception as paste_error:
                        raise Exception(f"슬라이드 {slide_idx} 복제 실패: {str(paste_error)[:80]}")

                    if slide_idx == 1 or slide_idx == total_slides or slide_idx % 5 == 0:
                        self.logger.log(f"  PPT 슬라이드 복제: {slide_idx}/{total_slides}")

            self._run_with_heartbeat(
                "PPT 슬라이드 복제 SaveAs",
                lambda: target_pres.SaveAs(temp_path, 24),
            )
            self._publish_verified_file(temp_path, save_path, "PPT 슬라이드 복제")
            self.logger.log(f"PPT 슬라이드 복제 저장 완료: {save_path}")
        except Exception:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            raise
        finally:
            if target_pres is not None:
                try:
                    target_pres.Close()
                except Exception:
                    pass

    def _capture_ppt_slide_image(self, source_slide, slide_idx, temp_dir, export_width, export_height):
        """PowerPoint Export가 막힌 문서도 슬라이드를 이미지로 확보한다."""
        img_path = os.path.join(temp_dir, f"visual_slide_{slide_idx}.png")
        export_error = None
        try:
            source_slide.Export(img_path, "PNG", export_width, export_height)
            if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                return img_path
            export_error = Exception("PowerPoint Export created an empty image")
        except Exception as error:
            export_error = error

        self.logger.log(
            f"  슬라이드 {slide_idx} Export 실패, 클립보드 이미지 복사 시도: {str(export_error)[:80]}"
        )
        clipboard_error = None
        for retry in range(1, self.PPT_CLIPBOARD_RETRY_COUNT + 2):
            try:
                source_slide.Copy()
                for _ in range(12):
                    time.sleep(max(self.PPT_CLIPBOARD_RETRY_DELAY, 0.1))
                    clipboard_path = self._get_image_from_clipboard(temp_dir)
                    if clipboard_path and os.path.exists(clipboard_path) and os.path.getsize(clipboard_path) > 0:
                        return clipboard_path
                clipboard_error = Exception("클립보드에서 사용 가능한 이미지를 찾지 못했습니다.")
            except Exception as error:
                clipboard_error = error
            if retry <= self.PPT_CLIPBOARD_RETRY_COUNT:
                time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)

        raise Exception(
            f"슬라이드 {slide_idx} 이미지 캡처 실패. "
            f"Export={str(export_error)[:80]}, Clipboard={str(clipboard_error)[:80]}"
        )

    def _save_ppt_visual_copy(self, source_pres, save_path):
        """각 슬라이드를 전체 이미지로 저장해 화면 배치/서식을 가장 정확하게 보존한다."""
        if not HAS_PPTX:
            raise Exception("python-pptx 패키지가 필요합니다. pip install python-pptx")

        target_dir = os.path.dirname(os.path.abspath(save_path)) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)
        temp_dir = tempfile.mkdtemp(prefix="docextract_visual_")
        temp_path = self._make_local_temp_path(".pptx")

        try:
            visual_pres = Presentation()
            slide_width = source_pres.PageSetup.SlideWidth
            slide_height = source_pres.PageSetup.SlideHeight
            visual_pres.slide_width = Emu(int(slide_width * 12700))
            visual_pres.slide_height = Emu(int(slide_height * 12700))
            blank_layout = visual_pres.slide_layouts[6]
            export_width = max(1, int(slide_width * 2))
            export_height = max(1, int(slide_height * 2))
            total_slides = source_pres.Slides.Count

            self.logger.log(f"PPT 화면 그대로 이미지 저장 시작: {save_path}")
            for slide_idx in range(1, total_slides + 1):
                img_path = self._capture_ppt_slide_image(
                    source_pres.Slides(slide_idx),
                    slide_idx,
                    temp_dir,
                    export_width,
                    export_height,
                )
                slide = visual_pres.slides.add_slide(blank_layout)
                slide.shapes.add_picture(
                    img_path,
                    Emu(0),
                    Emu(0),
                    visual_pres.slide_width,
                    visual_pres.slide_height,
                )
                if slide_idx == 1 or slide_idx == total_slides or slide_idx % 5 == 0:
                    self.logger.log(f"  PPT 화면 그대로 이미지: {slide_idx}/{total_slides}")

            visual_pres.save(temp_path)
            self._publish_verified_file(temp_path, save_path, "PPT 화면 그대로")
            self.logger.log(f"PPT 화면 그대로 이미지 저장 완료: {save_path}")
        except Exception:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            raise
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _try_save_ppt_visual_companion(self, source_pres, save_path):
        visual_path = self._add_filename_suffix(save_path, "_화면그대로")
        try:
            self._save_ppt_visual_copy(source_pres, visual_path)
            self.logger.log(f"PPT 화면 그대로 추가본 생성: {visual_path}")
            return visual_path
        except Exception as visual_error:
            self.logger.log(f"PPT 화면 그대로 추가본 생성 실패: {str(visual_error)[:100]}")
            return None

    def _save_hwp_document(self, hwp, save_path, save_format):
        """한글 버전별 SaveAs 인자 차이를 흡수한다."""
        hwp_format = "HWPX" if save_format == "hwpx" else "HWP"
        target_dir = os.path.dirname(os.path.abspath(save_path)) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)

        try:
            hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
            file_open_save = hwp.HParameterSet.HFileOpenSave
            file_open_save.filename = save_path
            try:
                file_open_save.FileName = save_path
            except Exception:
                pass
            file_open_save.Format = hwp_format
            result = hwp.HAction.Execute("FileSaveAs_S", file_open_save.HSet)
            if result is False:
                raise Exception("FileSaveAs_S 액션이 실패를 반환했습니다.")
        except Exception as action_error:
            self.logger.log(f"한글 FileSaveAs_S 실패, SaveAs 재시도: {str(action_error)[:80]}")
            try:
                hwp.SaveAs(save_path, hwp_format, "")
            except Exception as saveas_error:
                self.logger.log(f"한글 SaveAs 3인자 실패, 2인자 재시도: {str(saveas_error)[:80]}")
                hwp.SaveAs(save_path, hwp_format)

        for _ in range(20):
            if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                return
            time.sleep(0.1)
        raise Exception("한글 저장 후 결과 파일이 생성되지 않았거나 비어 있습니다.")

    def _copy_word_document_file(self, source_doc, save_path):
        """저장된 Word 원본 파일을 원본 상태 변경 없이 복사한다."""
        try:
            source_path = source_doc.FullName
        except Exception:
            source_path = ""

        if not source_path or not os.path.exists(source_path):
            raise Exception("Word 원본 파일 경로를 확인할 수 없습니다. 문서를 먼저 저장해 주세요.")

        source_ext = os.path.splitext(source_path)[1].lower()
        target_ext = os.path.splitext(save_path)[1].lower()
        if source_ext and target_ext and source_ext != target_ext:
            raise Exception(f"원본 확장자({source_ext})와 저장 확장자({target_ext})가 달라 파일 복사를 생략합니다.")

        try:
            if not source_doc.Saved:
                raise Exception("Word 문서에 저장되지 않은 변경사항이 있습니다. 저장 후 다시 시도해 주세요.")
        except Exception as saved_error:
            if "저장되지 않은 변경사항" in str(saved_error):
                raise

        target_dir = os.path.dirname(os.path.abspath(save_path)) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)
        if os.path.abspath(source_path).lower() == os.path.abspath(save_path).lower():
            raise Exception("원본과 같은 경로로는 복사할 수 없습니다.")

        shutil.copy2(source_path, save_path)
        if not os.path.exists(save_path) or os.path.getsize(save_path) <= 0:
            raise Exception("Word 파일 복사 후 결과 파일이 생성되지 않았거나 비어 있습니다.")
        self._validate_office_openxml(save_path, "Word")

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

        self.ppt_extract_mode = tk.StringVar(value="native_copy")

        ttk.Radiobutton(mode_frame, text="원본 그대로 복사 (서식/넓이/높이/도형 보존)",
                        variable=self.ppt_extract_mode, value="native_copy").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="하이브리드 (편집 가능, 느림: 도형 속성 재생성)",
                        variable=self.ppt_extract_mode, value="hybrid").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="슬라이드 이미지만 (권장: 빠르고 안정적)",
                        variable=self.ppt_extract_mode, value="image_only").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="텍스트 중심 + 객체 보존 (도형/이미지는 그림으로 포함)",
                        variable=self.ppt_extract_mode, value="text_only").pack(anchor=tk.W)

        # 추출 버튼
        self.ppt_extract_button = ttk.Button(tab, text="새 PPT로 내보내기",
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

        self.excel_include_format = tk.BooleanVar(value=False)
        self.excel_include_formulas = tk.BooleanVar(value=False)
        self.excel_native_copy = tk.BooleanVar(value=True)

        ttk.Checkbutton(option_frame, text="원본 그대로 복사 우선 (서식/넓이/높이/도형 보존)",
                        variable=self.excel_native_copy).pack(anchor=tk.W)
        ttk.Checkbutton(option_frame, text="서식 포함 (느림: 글꼴, 색상, 행/열 크기)",
                        variable=self.excel_include_format).pack(anchor=tk.W)
        ttk.Checkbutton(option_frame, text="수식 대신 값만 저장",
                        variable=self.excel_include_formulas).pack(anchor=tk.W)

        # 추출 버튼
        self.excel_extract_button = ttk.Button(tab, text="새 Excel로 내보내기",
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
        self.hwp_extract_button = ttk.Button(tab, text="새 한글 문서로 내보내기",
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

        # debounce가 중복 이벤트를 정리하므로 탭 이동 시마다 최신 상태를 다시 확인한다.
        self.tab_detected[current_tab] = True

        # 해당 탭 감지 실행
        if current_tab == 0:  # PPT
            self.detect_open_ppt()
        elif current_tab == 1:  # Excel
            self.detect_open_excel()
        elif current_tab == 2:  # 한글
            self.detect_open_hwp()
        elif current_tab == 3:  # Word
            self.detect_open_word()
        elif current_tab == 4:  # 메모장
            self.detect_open_notepad()

    # ========== PPT 관련 메서드 ==========

    def browse_ppt_save_path(self):
        """PPT 저장 경로 선택"""
        self.logger.log("PPT 저장 경로 선택 대화상자 열기")

        doc_name = self.ppt_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 PPT 없음":
            src_ext = os.path.splitext(doc_name)[1] or ".pptx"
            default_ext = src_ext if src_ext.lower() in [".pptx", ".ppt", ".pptm"] else ".pptx"
            default_name = os.path.splitext(doc_name)[0] + "_복사본" + default_ext
        else:
            default_ext = ".pptx"
            default_name = "새문서.pptx"

        path = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=[("PowerPoint 파일", "*.pptx *.pptm *.ppt"), ("모든 파일", "*.*")],
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
            ppt, _ = self._get_ppt_app(allow_dispatch=False)
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

        save_path = self.ppt_save_path.get()
        mode = self.ppt_extract_mode.get()
        ppt_index = self.selected_ppt_index.get()
        self.logger.log(f"PPT 추출 설정: mode={mode}, index={ppt_index}, save_path={save_path}")

        if not save_path:
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.ppt_doc_name.get() == "열린 PPT 없음" or not self.ppt_list:
            messagebox.showwarning("경고", "열린 PPT가 없습니다.")
            return

        self.ppt_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_ppt, args=(save_path, mode, ppt_index))
        thread.daemon = True
        thread.start()

    def _extract_ppt(self, save_path, mode, ppt_index):
        """PPT 추출 (백그라운드)"""
        self.logger.log("=== PPT 추출 프로세스 시작 ===")
        extract_start = time.perf_counter()
        pythoncom.CoInitialize()

        temp_dir = None
        ppt_app = None
        original_alerts = None

        try:
            self.root.after(0, lambda: self.status_text.set("원본 PPT 연결 중..."))

            self.logger.log("PowerPoint COM 연결 시도")
            ppt_app, _ = self._get_ppt_app()
            try:
                original_alerts = ppt_app.DisplayAlerts
                ppt_app.DisplayAlerts = 1  # ppAlertsNone
                self.logger.log("PowerPoint 경고창 표시 비활성화")
            except Exception as alerts_error:
                self.logger.log(f"PowerPoint 경고창 설정 실패: {str(alerts_error)[:60]}")
            ppt_count = ppt_app.Presentations.Count
            if ppt_count == 0:
                raise Exception("열린 PowerPoint 문서가 없습니다. PowerPoint에서 문서를 먼저 열어주세요.")

            if ppt_index > 0 and ppt_index <= ppt_count:
                source_pres = ppt_app.Presentations(ppt_index)
            else:
                source_pres = ppt_app.ActivePresentation

            self.logger.log(f"원본 프레젠테이션: {source_pres.Name}")

            if mode == "native_copy":
                try:
                    self.root.after(0, lambda: self.status_text.set("원본 PPT 그대로 복사 중..."))
                    self.root.after(0, lambda: self.progress_var.set(20))
                    total_slides = source_pres.Slides.Count
                    self._save_native_copy(source_pres, save_path, "PPT")
                    self._log_elapsed("PPT 원본 복사 시간", extract_start)
                    self.root.after(0, lambda: self.progress_var.set(100))
                    self.root.after(0, lambda: self.status_text.set("PPT 원본 복사 완료!"))
                    self.root.after(0, lambda: messagebox.showinfo("완료",
                        f"PPT 원본 복사 완료!\n{save_path}\n\n총 {total_slides}장"))
                    return
                except Exception as copy_error:
                    self.logger.log(
                        f"PPT 원본 복사 결과 검증 실패, 슬라이드 복제로 전환: {str(copy_error)[:120]}"
                    )
                    try:
                        self.root.after(0, lambda: self.status_text.set("원본 복사 실패, 클립보드 슬라이드 패키지 복원 중..."))
                        self.root.after(0, lambda: self.progress_var.set(25))
                        self._save_ppt_clipboard_package_copy(source_pres, save_path)
                        self._log_elapsed("PPT 클립보드 슬라이드 패키지 복원 시간", extract_start)
                        self.root.after(0, lambda: self.progress_var.set(100))
                        self.root.after(0, lambda: self.status_text.set("PPT 원본 구조 복원 완료!"))
                        self.root.after(0, lambda: messagebox.showinfo("완료",
                            f"PPT 원본 구조 복원 완료!\n{save_path}\n\n총 {total_slides}장"))
                        return
                    except Exception as package_error:
                        self.logger.log(
                            f"PPT 클립보드 슬라이드 패키지 복원 실패, 슬라이드 복제로 전환: {str(package_error)[:120]}"
                        )

                    try:
                        self.root.after(0, lambda: self.status_text.set("원본 복사 실패, PowerPoint 슬라이드 복제 중..."))
                        self.root.after(0, lambda: self.progress_var.set(30))
                        self._save_ppt_slide_clone(source_pres, save_path)
                        visual_path = self._try_save_ppt_visual_companion(source_pres, save_path)
                        self._log_elapsed("PPT 슬라이드 복제 시간", extract_start)
                        self.root.after(0, lambda: self.progress_var.set(100))
                        self.root.after(0, lambda: self.status_text.set("PPT 슬라이드 복제 완료!"))
                        self.root.after(0, lambda: messagebox.showinfo("완료",
                            f"PPT 슬라이드 복제 완료!\n{save_path}\n\n"
                            f"화면 그대로 추가본:\n{visual_path or '생성 실패'}\n\n"
                            f"총 {total_slides}장"))
                        return
                    except Exception as clone_error:
                        self.logger.log(
                            f"PPT 슬라이드 복제 실패, 화면 그대로 저장 시도: {str(clone_error)[:120]}"
                        )
                        try:
                            self.root.after(0, lambda: self.status_text.set("슬라이드 복제 실패, 화면 그대로 저장 중..."))
                            self._save_ppt_visual_copy(source_pres, save_path)
                            self._log_elapsed("PPT 화면 그대로 저장 시간", extract_start)
                            self.root.after(0, lambda: self.progress_var.set(100))
                            self.root.after(0, lambda: self.status_text.set("PPT 화면 그대로 저장 완료!"))
                            self.root.after(0, lambda: messagebox.showinfo("완료",
                                f"PPT 화면 그대로 저장 완료!\n{save_path}\n\n총 {total_slides}장"))
                            return
                        except Exception as visual_error:
                            self.logger.log(
                                f"PPT 화면 그대로 저장 실패, 하이브리드 재구성으로 전환: {str(visual_error)[:120]}"
                            )
                            self.root.after(0, lambda: self.status_text.set("화면 그대로 저장 실패, PPT 재구성 중..."))
                            mode = "hybrid"

            if not HAS_PPTX:
                raise Exception("python-pptx 패키지가 필요합니다. pip install python-pptx")

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
                slide_start = time.perf_counter()
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
                    self._extract_text_with_object_images(source_slide, new_slide, temp_dir)
                else:
                    self._extract_hybrid(source_slide, new_slide, temp_dir, i, new_pres)
                self._log_elapsed(f"슬라이드 {i} 처리 시간", slide_start)

            self.root.after(0, lambda: self.status_text.set("파일 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(95))

            temp_ppt_path = self._make_local_temp_path(".pptx")
            try:
                new_pres.save(temp_ppt_path)
                self._publish_verified_file(temp_ppt_path, save_path, "PPT 재구성")
            finally:
                if os.path.exists(temp_ppt_path):
                    try:
                        os.remove(temp_ppt_path)
                    except Exception:
                        pass
            self.logger.log(f"저장 완료: {save_path}")
            self._log_elapsed("PPT 전체 추출 시간", extract_start)

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
            if ppt_app is not None and original_alerts is not None:
                try:
                    ppt_app.DisplayAlerts = original_alerts
                except Exception:
                    pass
            self.root.after(0, lambda: self.ppt_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    # ========== Excel 관련 메서드 ==========

    def browse_excel_save_path(self):
        """Excel 저장 경로 선택"""
        self.logger.log("Excel 저장 경로 선택")

        doc_name = self.excel_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 Excel 없음":
            src_ext = os.path.splitext(doc_name)[1] or ".xlsx"
            default_ext = src_ext if src_ext.lower() in [".xlsx", ".xlsm", ".xls", ".xlsb"] else ".xlsx"
            default_name = os.path.splitext(doc_name)[0] + "_복사본" + default_ext
        else:
            default_ext = ".xlsx"
            default_name = "새문서.xlsx"

        path = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls *.xlsb"), ("모든 파일", "*.*")],
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
            excel, _ = self._get_excel_app(allow_dispatch=False)
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

        native_copy = self.excel_native_copy.get()

        if not native_copy and not HAS_OPENPYXL:
            messagebox.showerror("오류", "openpyxl 패키지가 필요합니다.\npip install openpyxl")
            return

        save_path = self.excel_save_path.get()
        include_format = self.excel_include_format.get()
        values_only = self.excel_include_formulas.get()
        excel_index = self.selected_excel_index.get()

        if not save_path:
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.excel_doc_name.get() == "열린 Excel 없음" or not self.excel_list:
            messagebox.showwarning("경고", "열린 Excel이 없습니다.")
            return

        self.excel_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(
            target=self._extract_excel,
            args=(save_path, include_format, values_only, excel_index, native_copy),
        )
        thread.daemon = True
        thread.start()

    def _extract_excel(self, save_path, include_format, values_only, excel_index, native_copy=False):
        """Excel 추출 (백그라운드)"""
        self.logger.log("=== Excel 추출 시작 ===")
        extract_start = time.perf_counter()
        pythoncom.CoInitialize()

        temp_dir = None

        try:
            self.root.after(0, lambda: self.status_text.set("Excel 연결 중..."))

            excel_app, _ = self._get_excel_app()
            wb_count = excel_app.Workbooks.Count
            if wb_count == 0:
                raise Exception("열린 Excel 문서가 없습니다. Excel에서 문서를 먼저 열어주세요.")

            if excel_index > 0 and excel_index <= wb_count:
                source_wb = excel_app.Workbooks(excel_index)
            else:
                source_wb = excel_app.ActiveWorkbook

            self.logger.log(f"원본 통합문서: {source_wb.Name}")

            if native_copy:
                try:
                    self.root.after(0, lambda: self.status_text.set("원본 Excel 그대로 복사 중..."))
                    self.root.after(0, lambda: self.progress_var.set(20))
                    total_sheets = source_wb.Sheets.Count
                    self._save_native_copy(source_wb, save_path, "Excel")
                    self._log_elapsed("Excel 원본 복사 시간", extract_start)
                    self.root.after(0, lambda: self.progress_var.set(100))
                    self.root.after(0, lambda: self.status_text.set("Excel 원본 복사 완료!"))
                    self.root.after(0, lambda: messagebox.showinfo("완료",
                        f"Excel 원본 복사 완료!\n{save_path}\n\n총 {total_sheets}시트"))
                    return
                except Exception as copy_error:
                    self.logger.log(
                        f"Excel 원본 복사 결과 검증 실패, 재구성으로 전환: {str(copy_error)[:120]}"
                    )
                    self.root.after(0, lambda: self.status_text.set("원본 복사 실패, Excel 재구성 중..."))

            if not HAS_OPENPYXL:
                raise Exception("openpyxl 패키지가 필요합니다. pip install openpyxl")

            # openpyxl로 새 통합문서 생성
            new_wb = Workbook()
            # 기본 시트 제거 (나중에 추가할 것이므로)
            default_sheet = new_wb.active

            total_sheets = source_wb.Sheets.Count
            temp_dir = tempfile.mkdtemp()
            self.root.after(0, lambda: self.progress_var.set(5))

            for sheet_idx in range(1, total_sheets + 1):
                sheet_start = time.perf_counter()
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

                self._copy_excel_sheet_objects(source_sheet, new_sheet, temp_dir, sheet_name)

                # 사용 범위 가져오기
                try:
                    used_range = source_sheet.UsedRange
                    if used_range is None:
                        continue

                    row_count = used_range.Rows.Count
                    col_count = used_range.Columns.Count
                    start_row = used_range.Row
                    start_col = used_range.Column
                    cell_count = row_count * col_count

                    self.logger.log(f"    범위: {row_count}행 x {col_count}열 ({cell_count:,}셀, 시작: {start_row},{start_col})")

                    if cell_count > self.EXCEL_VALUE_CELL_LIMIT:
                        raise Exception(
                            f"사용 범위가 너무 큽니다 ({cell_count:,}셀). "
                            "Excel에서 불필요한 빈 행/열 서식을 지운 뒤 다시 시도해주세요."
                        )

                    sheet_include_format = include_format
                    if sheet_include_format and cell_count > self.EXCEL_FORMAT_CELL_LIMIT:
                        sheet_include_format = False
                        self.logger.log(
                            f"    대용량 시트 서식 복사 자동 생략: {cell_count:,}셀 "
                            f"(한도 {self.EXCEL_FORMAT_CELL_LIMIT:,}셀)"
                        )
                        self.root.after(
                            0,
                            lambda n=sheet_name: self.status_text.set(
                                f"시트 '{n}' 값 복사 중... (대용량 서식 생략)"
                            ),
                        )

                    # 값/수식은 COM Range에서 한 번에 읽어 셀 단위 왕복을 줄인다.
                    try:
                        range_data = used_range.Value if values_only else used_range.Formula
                        data_rows = self._excel_range_to_rows(range_data, row_count, col_count)
                        for r, row_values in enumerate(data_rows):
                            for c, value in enumerate(row_values[:col_count]):
                                if value is not None:
                                    new_sheet.cell(row=start_row + r, column=start_col + c).value = value
                    except Exception as data_err:
                        self.logger.log(f"    범위 데이터 일괄 읽기 실패, 셀 단위로 폴백: {str(data_err)[:50]}")
                        data_rows = None

                    if data_rows is None or sheet_include_format:
                        for r in range(row_count):
                            if r % 100 == 0:
                                progress = 5 + ((sheet_idx - 1) / total_sheets) * 85
                                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                                self.root.after(
                                    0,
                                    lambda n=sheet_name, rr=r + 1, total=row_count:
                                        self.status_text.set(f"시트 '{n}' 행 {rr}/{total} 처리 중...")
                                )
                            for c in range(col_count):
                                try:
                                    cell = source_sheet.Cells(start_row + r, start_col + c)
                                    new_cell = new_sheet.cell(row=start_row + r, column=start_col + c)

                                    # 일괄 읽기에 실패한 경우에만 값/수식을 셀 단위로 읽는다.
                                    if data_rows is None:
                                        if values_only:
                                            new_cell.value = cell.Value
                                        else:
                                            try:
                                                formula = cell.Formula
                                                if formula and str(formula).startswith('='):
                                                    new_cell.value = formula
                                                else:
                                                    new_cell.value = cell.Value
                                            except Exception:
                                                new_cell.value = cell.Value

                                    # 서식 복사
                                    if sheet_include_format:
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
                                        except Exception as font_err:
                                            self.logger.log(f"    셀 서식(폰트) 복사 실패: {str(font_err)[:40]}")

                                        try:
                                            # 배경색
                                            interior_color = cell.Interior.Color
                                            if interior_color and interior_color != 16777215:  # 흰색이 아니면
                                                hex_color = self._excel_color_to_hex(interior_color)
                                                if hex_color:
                                                    new_cell.fill = PatternFill(start_color=hex_color,
                                                                                end_color=hex_color,
                                                                                fill_type='solid')
                                        except Exception as fill_err:
                                            self.logger.log(f"    셀 서식(배경색) 복사 실패: {str(fill_err)[:40]}")

                                        try:
                                            number_format = cell.NumberFormat
                                            if number_format:
                                                new_cell.number_format = number_format
                                        except Exception:
                                            pass

                                except Exception as cell_err:
                                    self.logger.log(f"    셀 처리 실패: {str(cell_err)[:40]}")

                    if row_count <= self.EXCEL_ROW_HEIGHT_COPY_LIMIT:
                        try:
                            for r in range(1, row_count + 1):
                                height = source_sheet.Rows(start_row + r - 1).RowHeight
                                if height:
                                    new_sheet.row_dimensions[start_row + r - 1].height = height
                        except Exception:
                            pass
                    else:
                        self.logger.log(f"    행 높이 복사 생략: {row_count:,}행")

                    if col_count <= self.EXCEL_COLUMN_WIDTH_COPY_LIMIT:
                        try:
                            for c in range(1, col_count + 1):
                                width = source_sheet.Columns(start_col + c - 1).ColumnWidth
                                new_sheet.column_dimensions[get_column_letter(start_col + c - 1)].width = width
                        except:
                            pass
                    else:
                        self.logger.log(f"    열 너비 복사 생략: {col_count:,}열")

                    # 병합 셀 복사
                    try:
                        merged_areas = source_sheet.UsedRange.MergeAreas
                        for area_idx in range(1, merged_areas.Count + 1):
                            area = merged_areas(area_idx)
                            if area.Cells.Count > 1:
                                first_row = area.Row
                                first_col = area.Column
                                last_row = first_row + area.Rows.Count - 1
                                last_col = first_col + area.Columns.Count - 1
                                new_sheet.merge_cells(
                                    start_row=first_row,
                                    start_column=first_col,
                                    end_row=last_row,
                                    end_column=last_col,
                                )
                    except Exception:
                        pass

                except Exception as sheet_err:
                    self.logger.error(f"시트 '{sheet_name}' 처리 오류", sheet_err)
                finally:
                    self._log_elapsed(f"시트 '{sheet_name}' 처리 시간", sheet_start)

            # 저장
            self.root.after(0, lambda: self.status_text.set("파일 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(95))

            new_wb.save(save_path)
            self._validate_office_openxml(save_path, "Excel 재구성")
            self.logger.log(f"저장 완료: {save_path}")
            self._log_elapsed("Excel 전체 추출 시간", extract_start)

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("Excel 추출 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"Excel 추출 완료!\n{save_path}\n\n총 {total_sheets}시트"))

        except Exception as e:
            self.logger.error("Excel 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.root.after(0, lambda: self.excel_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    def _excel_range_to_rows(self, value, row_count, col_count):
        """Excel COM Range 값을 항상 행 튜플 형태로 정규화한다."""
        if row_count <= 0 or col_count <= 0:
            return []

        if row_count == 1 and col_count == 1:
            return [(value,)]

        if row_count == 1:
            if isinstance(value, (tuple, list)):
                if len(value) == 1 and isinstance(value[0], (tuple, list)):
                    return [tuple(value[0])]
                return [tuple(value)]
            return [(value,)]

        rows = value if isinstance(value, (tuple, list)) else ((value,),)
        normalized = []
        for row in rows:
            if isinstance(row, (tuple, list)):
                normalized.append(tuple(row))
            else:
                normalized.append((row,))

        return normalized

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

    def _points_to_pixels(self, points):
        """Office point 단위를 96DPI 기준 픽셀로 근사 변환한다."""
        try:
            return max(1, int(round(float(points) * 96 / 72)))
        except Exception:
            return 1

    def _copy_excel_sheet_objects(self, source_sheet, new_sheet, temp_dir, sheet_name):
        """Excel 시트의 삽입 그림/도형/차트 객체를 이미지로 복사한다."""
        try:
            shapes = source_sheet.Shapes
            shape_count = shapes.Count
        except Exception as e:
            self.logger.log(f"    시트 객체 목록 읽기 실패: {str(e)[:50]}")
            return

        if shape_count <= 0:
            return

        self.logger.log(f"    삽입 객체 복사 시작: {shape_count}개")
        copied_count = 0

        for shape_idx in range(1, shape_count + 1):
            try:
                shape = shapes(shape_idx)
                try:
                    if not bool(shape.Visible):
                        continue
                except Exception:
                    pass

                img_path = self._excel_shape_to_image(shape, temp_dir, sheet_name, shape_idx)
                if not img_path:
                    self.logger.log(f"    객체 {shape_idx} 이미지 변환 실패")
                    continue

                if os.path.splitext(img_path)[1].lower() not in [".png", ".jpg", ".jpeg", ".bmp", ".gif"]:
                    self.logger.log(f"    객체 {shape_idx} 지원하지 않는 이미지 형식 생략: {img_path}")
                    continue

                image = OpenpyxlImage(img_path)
                try:
                    image.width = self._points_to_pixels(shape.Width)
                    image.height = self._points_to_pixels(shape.Height)
                except Exception:
                    pass

                try:
                    top_left = shape.TopLeftCell
                    anchor = f"{get_column_letter(top_left.Column)}{top_left.Row}"
                except Exception:
                    anchor = "A1"

                image.anchor = anchor
                new_sheet.add_image(image)
                copied_count += 1

            except Exception as e:
                self.logger.log(f"    객체 {shape_idx} 복사 실패: {str(e)[:60]}")

        self.logger.log(f"    삽입 객체 복사 완료: {copied_count}/{shape_count}개")

    def _excel_shape_to_image(self, source_shape, temp_dir, sheet_name, shape_idx):
        """Excel Shape를 클립보드 경유 이미지 파일로 저장한다."""
        safe_sheet = "".join(ch if ch.isalnum() else "_" for ch in str(sheet_name))[:30] or "sheet"

        for retry in range(self.EXCEL_OBJECT_RETRY_COUNT):
            try:
                try:
                    source_shape.CopyPicture(Appearance=1, Format=2)  # 1=screen, 2=bitmap
                except Exception:
                    try:
                        source_shape.CopyPicture(1, 2)
                    except Exception:
                        source_shape.Copy()

                time.sleep(self.EXCEL_OBJECT_RETRY_DELAY)
                img_path = self._get_image_from_clipboard(temp_dir)
                if img_path:
                    ext = os.path.splitext(img_path)[1] or ".png"
                    final_path = os.path.join(
                        temp_dir,
                        f"excel_{safe_sheet}_{shape_idx}_{retry}{ext}"
                    )
                    try:
                        if os.path.abspath(img_path) != os.path.abspath(final_path):
                            shutil.copyfile(img_path, final_path)
                            try:
                                os.remove(img_path)
                            except Exception:
                                pass
                        return final_path
                    except Exception:
                        return img_path
            except Exception as e:
                self.logger.log(
                    f"    객체 {shape_idx} 클립보드 복사 실패 "
                    f"(시도 {retry+1}/{self.EXCEL_OBJECT_RETRY_COUNT}): {str(e)[:50]}"
                )
                time.sleep(self.EXCEL_OBJECT_RETRY_DELAY)

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
                if not self._recreate_shape(shape, target_slide, temp_dir):
                    self._add_ppt_shape_snapshot(shape, target_slide, temp_dir)
            except Exception as e:
                self.logger.log(f"도형 처리 실패, 이미지 보존 시도: {str(e)[:50]}")
                self._add_ppt_shape_snapshot(shape, target_slide, temp_dir)

    def _extract_text_with_object_images(self, source_slide, target_slide, temp_dir):
        """텍스트 중심 모드에서도 도형/이미지는 그림으로 보존한다."""
        self._extract_hybrid(source_slide, target_slide, temp_dir, 0, None)

    def _add_ppt_shape_snapshot(self, source_shape, target_slide, temp_dir, left=None, top=None, width=None, height=None):
        """재생성할 수 없는 PPT 도형을 이미지 스냅샷으로 보존한다."""
        img_path = None
        try:
            if left is None:
                left = Emu(int(source_shape.Left * 12700))
            if top is None:
                top = Emu(int(source_shape.Top * 12700))
            if width is None:
                width = Emu(int(source_shape.Width * 12700))
            if height is None:
                height = Emu(int(source_shape.Height * 12700))

            for retry in range(self.PPT_CLIPBOARD_RETRY_COUNT):
                try:
                    try:
                        source_shape.CopyPicture()
                    except Exception:
                        source_shape.Copy()
                    time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
                    img_path = self._get_image_from_clipboard(temp_dir)
                    if img_path:
                        break
                except Exception as e:
                    if retry == self.PPT_CLIPBOARD_RETRY_COUNT - 1:
                        self.logger.log(f"도형 클립보드 보존 실패: {str(e)[:60]}")
                    time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)

            if not img_path or not os.path.exists(img_path):
                try:
                    img_path = os.path.join(temp_dir, f"shape_{id(source_shape)}.png")
                    source_shape.Export(img_path, 2)
                except Exception:
                    img_path = None

            if img_path and os.path.exists(img_path):
                target_slide.shapes.add_picture(img_path, left, top, width, height)
                return True

            return False
        except Exception as e:
            self.logger.log(f"도형 이미지 스냅샷 추가 실패: {str(e)[:60]}")
            return False

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
                return self._add_ppt_shape_snapshot(source_shape, target_slide, temp_dir, left, top, width, height)

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
        export_error = None

        # 클립보드 우선: 일부 보안 PPT는 Shape.Export 호출 때 PowerPoint 경고창을 띄운다.
        for retry in range(self.PPT_CLIPBOARD_RETRY_COUNT):
            clipboard_img = None
            try:
                try:
                    source_shape.CopyPicture()
                except Exception:
                    source_shape.Copy()
                time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
                clipboard_img = self._get_image_from_clipboard(temp_dir)
                if clipboard_img:
                    target_slide.shapes.add_picture(clipboard_img, left, top, width, height)
                    return True
            except Exception as e:
                self.logger.log(
                    f"클립보드 이미지 추출 실패 "
                    f"(시도 {retry+1}/{self.PPT_CLIPBOARD_RETRY_COUNT}): {str(e)[:50]}"
                )
                time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
            finally:
                # 클립보드에서 생성된 임시 파일 정리
                if clipboard_img and os.path.exists(clipboard_img):
                    try:
                        os.remove(clipboard_img)
                    except Exception:
                        pass

        # Export 폴백
        try:
            img_path = os.path.join(temp_dir, f"img_{id(source_shape)}.png")
            source_shape.Export(img_path, 2)
            target_slide.shapes.add_picture(img_path, left, top, width, height)
            return True
        except Exception as e:
            export_error = e
            if img_path and os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass

        # Placeholder
        try:
            if export_error:
                self.logger.log(f"이미지 보존 폴백 실패, Placeholder 생성: {str(export_error)[:60]}")
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
        for retry in range(self.PPT_CLIPBOARD_RETRY_COUNT):
            try:
                source_shape.Copy()
                time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
                img_path = self._get_image_from_clipboard(temp_dir)
                if img_path:
                    target_slide.shapes.add_picture(img_path, left, top, width, height)
                    return True
            except Exception as e:
                self.logger.log(
                    f"Freeform 클립보드 추출 실패 "
                    f"(시도 {retry+1}/{self.PPT_CLIPBOARD_RETRY_COUNT}): {str(e)[:50]}"
                )
                time.sleep(self.PPT_CLIPBOARD_RETRY_DELAY)
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

                # PNG
                png_format = win32clipboard.RegisterClipboardFormat("PNG")
                if win32clipboard.IsClipboardFormatAvailable(png_format):
                    data = win32clipboard.GetClipboardData(png_format)
                    img_path = os.path.join(temp_dir, f"clipboard_{id(data)}.png")
                    with open(img_path, 'wb') as f:
                        f.write(data)
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
                orientation = -1 if height > 0 else 1
                return Image.frombytes(
                    'RGB',
                    (width, abs(height)),
                    pixel_data,
                    'raw',
                    'BGR',
                    row_size,
                    orientation,
                )
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
        if self._hwp_detecting:
            self.logger.log("한글 감지 요청 생략: 이전 감지가 아직 진행 중")
            return

        self.logger.log("한글 감지 시작")
        self.status_text.set("한글 감지 중...")
        self.hwp_doc_name.set("감지 중...")
        self._hwp_detecting = True

        thread = threading.Thread(target=self._detect_hwp)
        thread.daemon = True
        thread.start()

    def _detect_hwp(self):
        """한글 감지 (백그라운드)"""
        self.logger.log("백그라운드 한글 감지 스레드 시작")
        pythoncom.CoInitialize()

        try:
            try:
                hwp, _ = self._get_hwp_app(allow_dispatch=False)
            except Exception as connect_err:
                self.logger.log(
                    f"한글 활성 COM 연결 실패, 창 제목 감지로 대체: {str(connect_err)[:80]}"
                )
                hwp_titles = self._list_hwp_window_titles()
                if hwp_titles:
                    self.hwp_list = [(title, "", idx + 1) for idx, title in enumerate(hwp_titles)]
                    self.logger.log(f"한글 창 감지: {len(hwp_titles)}개")

                    def update_window_combo():
                        self.hwp_combo['values'] = hwp_titles
                        self.hwp_combo.current(0)
                        self.selected_hwp_index.set(1)
                        self.hwp_doc_name.set(hwp_titles[0])
                        self.status_text.set("한글 창 감지됨")

                    self.root.after(0, update_window_combo)
                else:
                    self.hwp_list = []

                    def clear_window_combo():
                        self.hwp_combo.set("")
                        self.hwp_combo['values'] = []
                        self.hwp_doc_name.set("열린 한글 없음")
                        self.status_text.set("한글을 먼저 열어주세요")

                    self.root.after(0, clear_window_combo)
                return

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

        finally:
            self._hwp_detecting = False
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

        save_path = self.hwp_save_path.get()
        save_format = self.hwp_save_format.get()

        if not save_path:
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.hwp_doc_name.get() == "열린 한글 없음" or not self.hwp_list:
            messagebox.showwarning("경고", "열린 한글 문서가 없습니다.")
            return

        self.hwp_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_hwp, args=(save_path, save_format))
        thread.daemon = True
        thread.start()

    def _extract_hwp(self, save_path, save_format):
        """한글 추출 (백그라운드)"""
        self.logger.log("=== 한글 추출 프로세스 시작 ===")
        extract_start = time.perf_counter()
        pythoncom.CoInitialize()

        try:
            self.root.after(0, lambda: self.status_text.set("원본 한글 연결 중..."))

            hwp, _ = self._get_hwp_app_for_extraction()
            try:
                current_path = hwp.Path
            except Exception:
                current_path = ""

            self.logger.log(f"원본 한글 문서 연결 성공")
            self.root.after(0, lambda: self.status_text.set("새 문서로 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(30))

            # 방법 1: SaveAs 시도
            try:
                self.logger.log(f"SaveAs 시도: {save_path}")
                self._save_hwp_document(hwp, save_path, save_format)

                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, lambda: self.status_text.set("한글 추출 완료!"))
                self.root.after(0, lambda: messagebox.showinfo("완료",
                    f"한글 추출 완료!\n{save_path}"))
                self.logger.log(f"저장 완료: {save_path}")
                self._log_elapsed("한글 전체 추출 시간", extract_start)

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
                    self._save_hwp_document(hwp, save_path, save_format)

                    self.root.after(0, lambda: self.progress_var.set(100))
                    self.root.after(0, lambda: self.status_text.set("한글 추출 완료!"))
                    self.root.after(0, lambda: messagebox.showinfo("완료",
                        f"한글 추출 완료 (클립보드 방식)!\n{save_path}"))
                    self.logger.log(f"클립보드 방식으로 저장 완료: {save_path}")
                    self._log_elapsed("한글 전체 추출 시간", extract_start)

                except Exception as e2:
                    raise Exception(f"한글 저장 실패:\n{str(e2)}")

        except Exception as e:
            self.logger.error("한글 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.hwp_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    # ========== Word 관련 메서드 ==========

    def _collect_word_runs(self, source_range, full_text):
        """Word 단락 Range에서 서식 동질 런(run) 목록을 추출.

        어절(Words) 단위로 Font 속성을 한 번씩만 읽어 서식 변경점을 감지.
        문자 단위 순회 대비 COM 왕복 횟수를 평균 5~10배 줄인다.

        반환: [(text, font_name, font_size, bold, italic, underline, color), ...]
              어절 접근 실패 시 빈 리스트.
        """
        runs_data = []
        try:
            words_col = source_range.Words
            words_count = words_col.Count
        except Exception:
            return runs_data

        if words_count <= 0:
            return runs_data

        current_text = ""
        prev_key = None
        prev_attrs = None

        for w_idx in range(1, words_count + 1):
            try:
                word_range = words_col(w_idx)
                word_text = word_range.Text
                if not word_text:
                    continue

                font = word_range.Font
                try:
                    font_name = font.Name
                except Exception:
                    font_name = None
                try:
                    font_size = font.Size
                except Exception:
                    font_size = None
                try:
                    bold = font.Bold
                except Exception:
                    bold = None
                try:
                    italic = font.Italic
                except Exception:
                    italic = None
                try:
                    underline = font.Underline
                except Exception:
                    underline = None
                try:
                    color = font.Color
                except Exception:
                    color = 0

                key = (font_name, font_size, bold, italic, underline, color)

                if prev_key is None:
                    current_text = word_text
                    prev_key = key
                    prev_attrs = (font_name, font_size, bold, italic, underline, color)
                elif key == prev_key:
                    current_text += word_text
                else:
                    runs_data.append((current_text,) + prev_attrs)
                    current_text = word_text
                    prev_key = key
                    prev_attrs = (font_name, font_size, bold, italic, underline, color)
            except Exception:
                continue

        if current_text and prev_attrs is not None:
            runs_data.append((current_text,) + prev_attrs)

        # 런 추출이 아예 실패했는데 full_text는 있는 경우의 안전망
        if not runs_data and full_text:
            runs_data.append((full_text, None, None, None, None, None, 0))

        return runs_data

    def _get_word_app(self, allow_dispatch=True):
        """Word 애플리케이션 연결 (3단계 폴백 공통화).

        반환: (word_app, created_new) — created_new는 Dispatch로 새 인스턴스를 만든 경우 True.
        실패 시 예외.
        """
        return self._connect_com_app("Word.Application", "Word", allow_dispatch=allow_dispatch)

    def _setup_word_tab(self):
        """Word 탭 설정"""
        tab = self.word_tab

        # 문서 정보 프레임
        info_frame = ttk.LabelFrame(tab, text="열린 Word 선택", padding="10")
        info_frame.pack(fill=tk.X, pady=5, padx=5)

        # Word 선택 콤보박스
        select_frame = ttk.Frame(info_frame)
        select_frame.pack(fill=tk.X, pady=2)
        ttk.Label(select_frame, text="Word 선택:", width=12).pack(side=tk.LEFT)
        self.word_combo = ttk.Combobox(select_frame, state="readonly", width=40)
        self.word_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.word_combo.bind("<<ComboboxSelected>>", self.on_word_selected)

        # 페이지 수
        page_frame = ttk.Frame(info_frame)
        page_frame.pack(fill=tk.X, pady=2)
        ttk.Label(page_frame, text="페이지 수:", width=12).pack(side=tk.LEFT)
        ttk.Label(page_frame, textvariable=self.word_page_count,
                  font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_word).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(tab, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.word_save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_word_save_path).pack(side=tk.LEFT)

        # 추출 옵션 프레임
        option_frame = ttk.LabelFrame(tab, text="추출 옵션", padding="10")
        option_frame.pack(fill=tk.X, pady=5, padx=5)

        self.word_include_format = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="서식 포함 (글꼴, 색상, 정렬, 들여쓰기)",
                        variable=self.word_include_format).pack(anchor=tk.W)
        self.word_use_saveas = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="원본 파일 복사 우선 (원본 상태 변경 없음)",
                        variable=self.word_use_saveas).pack(anchor=tk.W)

        # 추출 버튼
        self.word_extract_button = ttk.Button(tab, text="새 Word로 내보내기",
                                               command=self.start_word_extraction,
                                               style="Accent.TButton")
        self.word_extract_button.pack(pady=10)

    def browse_word_save_path(self):
        """Word 저장 경로 선택"""
        self.logger.log("Word 저장 경로 선택")

        doc_name = self.word_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 Word 없음":
            src_ext = os.path.splitext(doc_name)[1] or ".docx"
            default_ext = src_ext if src_ext.lower() in [".docx", ".docm", ".doc", ".rtf"] else ".docx"
            default_name = os.path.splitext(doc_name)[0] + "_복사본" + default_ext
        else:
            default_ext = ".docx"
            default_name = "새문서.docx"

        path = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=[("Word 파일", "*.docx *.docm *.doc *.rtf"), ("모든 파일", "*.*")],
            initialfile=default_name,
            title="저장할 위치 선택"
        )
        if path:
            self.word_save_path.set(path)
            self.logger.log(f"Word 저장 경로: {path}")

    def detect_open_word(self):
        """열려있는 Word 감지"""
        self.logger.log("Word 감지 시작")
        self.status_text.set("Word 감지 중...")
        self.word_doc_name.set("감지 중...")
        self.word_page_count.set("-")

        thread = threading.Thread(target=self._detect_word)
        thread.daemon = True
        thread.start()

    def _detect_word(self):
        """Word 감지 (백그라운드)"""
        self.logger.log("백그라운드 Word 감지 스레드 시작")
        pythoncom.CoInitialize()

        try:
            word, _ = self._get_word_app(allow_dispatch=False)

            doc_count = word.Documents.Count
            self.logger.log(f"Word 연결 성공, 열린 문서 수: {doc_count}")

            if doc_count > 0:
                word_names = []
                word_info = []

                for i in range(1, doc_count + 1):
                    try:
                        doc = word.Documents(i)
                        name = doc.Name
                        try:
                            page_count = doc.ComputeStatistics(2)  # wdStatisticPages = 2
                        except Exception:
                            page_count = 0
                        word_names.append(f"{name} ({page_count}페이지)")
                        word_info.append((name, page_count, i))
                        self.logger.log(f"  Word {i}: {name}, {page_count}페이지")
                    except Exception as e:
                        self.logger.log(f"  Word {i} 정보 가져오기 실패: {str(e)}")

                self.word_list = word_info

                def update_combo():
                    self.word_combo['values'] = word_names
                    if word_names:
                        self.word_combo.current(0)
                        self.selected_word_index.set(1)
                        self.word_doc_name.set(word_info[0][0])
                        self.word_page_count.set(f"{word_info[0][1]}페이지")
                    self.status_text.set(f"Word {len(word_names)}개 감지됨")

                self.root.after(0, update_combo)
            else:
                self.word_list = []
                def clear_combo():
                    self.word_combo.set("")
                    self.word_combo['values'] = []
                    self.word_doc_name.set("열린 Word 없음")
                    self.word_page_count.set("-")
                    self.status_text.set("Word를 먼저 열어주세요")
                self.root.after(0, clear_combo)

        except Exception as e:
            self.logger.error("Word 감지 실패", e)
            self.word_list = []
            err_msg = str(e)[:30]
            def show_error():
                self.word_combo.set("")
                self.word_doc_name.set("열린 Word 없음")
                self.word_page_count.set("-")
                self.status_text.set(f"Word 감지 실패: {err_msg}")
            self.root.after(0, show_error)

        pythoncom.CoUninitialize()

    def on_word_selected(self, event):
        """Word 콤보박스 선택 이벤트"""
        selected_idx = self.word_combo.current()
        if selected_idx >= 0 and selected_idx < len(self.word_list):
            name, page_count, word_index = self.word_list[selected_idx]
            self.selected_word_index.set(word_index)
            self.word_doc_name.set(name)
            self.word_page_count.set(f"{page_count}페이지")
            self.logger.log(f"Word 선택: {name}")

    def start_word_extraction(self):
        """Word 추출 시작"""
        self.logger.log("Word 추출 시작")

        save_path = self.word_save_path.get()
        include_format = self.word_include_format.get()
        use_saveas = self.word_use_saveas.get()
        word_index = self.selected_word_index.get()

        if not use_saveas and not HAS_DOCX:
            messagebox.showerror("오류", "python-docx 패키지가 필요합니다.\npip install python-docx")
            return

        if not save_path:
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.word_doc_name.get() == "열린 Word 없음" or not self.word_list:
            messagebox.showwarning("경고", "열린 Word가 없습니다.")
            return

        self.word_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(
            target=self._extract_word,
            args=(save_path, include_format, use_saveas, word_index),
        )
        thread.daemon = True
        thread.start()

    def _extract_word(self, save_path, include_format, use_saveas, word_index):
        """Word 추출 (백그라운드)"""
        self.logger.log("=== Word 추출 시작 ===")
        extract_start = time.perf_counter()
        pythoncom.CoInitialize()

        try:
            self.root.after(0, lambda: self.status_text.set("Word 연결 중..."))

            # Word 연결 (공통 헬퍼)
            word_app, _ = self._get_word_app()

            doc_count = word_app.Documents.Count
            if doc_count == 0:
                raise Exception("열린 Word 문서가 없습니다. Word에서 문서를 먼저 열어주세요.")

            if word_index > 0 and word_index <= doc_count:
                source_doc = word_app.Documents(word_index)
            else:
                source_doc = word_app.ActiveDocument

            # 대상 문서 유효성 검증 (빈 문서/이름 없는 문서에 SaveAs2 방지)
            try:
                doc_char_count = source_doc.Characters.Count
            except Exception:
                doc_char_count = 0
            if not source_doc.Name or doc_char_count <= 1:
                raise Exception(f"선택된 문서가 비어 있거나 유효하지 않습니다: '{source_doc.Name}'")

            self.logger.log(f"원본 문서: {source_doc.Name} ({doc_char_count}자)")

            if use_saveas:
                try:
                    self.root.after(0, lambda: self.status_text.set("원본 Word 파일 복사 중..."))
                    self.root.after(0, lambda: self.progress_var.set(10))
                    self._copy_word_document_file(source_doc, save_path)
                    self.root.after(0, lambda: self.progress_var.set(100))
                    self.root.after(0, lambda: self.status_text.set("Word 추출 완료! (원본 파일 복사)"))
                    self.root.after(0, lambda: messagebox.showinfo("완료",
                        f"Word 추출 완료 (원본 파일 복사)\n\n{save_path}"))
                    self.logger.log(f"Word 원본 파일 복사 성공: {save_path}")
                    self._log_elapsed("Word 전체 추출 시간", extract_start)
                    return
                except Exception as e:
                    self.logger.log(f"Word 원본 파일 복사 실패, 텍스트 재구성으로 폴백: {str(e)[:80]}")

            if not HAS_DOCX:
                raise Exception("python-docx 패키지가 필요합니다. pip install python-docx")

            # 방법 2: python-docx로 직접 재구성
            self.root.after(0, lambda: self.status_text.set("텍스트 추출 방식으로 진행 중..."))
            self.root.after(0, lambda: self.progress_var.set(20))

            new_doc = DocxDocument()

            total_paragraphs = source_doc.Paragraphs.Count
            self.logger.log(f"총 단락 수: {total_paragraphs}")

            for p_idx in range(1, total_paragraphs + 1):
                progress = 20 + (p_idx / total_paragraphs) * 70
                if p_idx % 50 == 0 or p_idx == total_paragraphs:
                    self.root.after(0, lambda p=progress: self.progress_var.set(p))
                    self.root.after(0, lambda n=p_idx, t=total_paragraphs:
                        self.status_text.set(f"단락 {n}/{t} 처리 중..."))

                try:
                    source_para = source_doc.Paragraphs(p_idx)
                    source_range = source_para.Range

                    # 새 단락 추가
                    new_para = new_doc.add_paragraph()

                    # 단락 서식 복사
                    if include_format:
                        try:
                            # 정렬
                            alignment = source_para.Alignment
                            # 0=Left, 1=Center, 2=Right, 3=Justify
                            align_map = {0: WD_ALIGN_PARAGRAPH.LEFT, 1: WD_ALIGN_PARAGRAPH.CENTER,
                                         2: WD_ALIGN_PARAGRAPH.RIGHT, 3: WD_ALIGN_PARAGRAPH.JUSTIFY}
                            new_para.alignment = align_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)
                        except Exception as align_err:
                            self.logger.log(f"  단락 {p_idx} 정렬 복사 실패: {str(align_err)[:40]}")

                        try:
                            # 들여쓰기 (포인트 → EMU 변환)
                            left_indent = source_para.Format.LeftIndent
                            if left_indent and left_indent > 0:
                                new_para.paragraph_format.left_indent = DocxPt(left_indent)
                            first_indent = source_para.Format.FirstLineIndent
                            if first_indent and first_indent > 0:
                                new_para.paragraph_format.first_line_indent = DocxPt(first_indent)
                        except Exception as indent_err:
                            self.logger.log(f"  단락 {p_idx} 들여쓰기 복사 실패: {str(indent_err)[:40]}")

                        try:
                            # 단락 앞/뒤 간격
                            space_before = source_para.Format.SpaceBefore
                            if space_before and space_before > 0:
                                new_para.paragraph_format.space_before = DocxPt(space_before)
                            space_after = source_para.Format.SpaceAfter
                            if space_after and space_after > 0:
                                new_para.paragraph_format.space_after = DocxPt(space_after)
                        except Exception as space_err:
                            self.logger.log(f"  단락 {p_idx} 간격 복사 실패: {str(space_err)[:40]}")

                        try:
                            # 줄 간격
                            line_spacing = source_para.Format.LineSpacing
                            if line_spacing and line_spacing > 0:
                                new_para.paragraph_format.line_spacing = DocxPt(line_spacing)
                        except Exception as ls_err:
                            self.logger.log(f"  단락 {p_idx} 줄간격 복사 실패: {str(ls_err)[:40]}")

                    # 런 단위로 텍스트 복사 — Words(어절) 단위 배치 처리로 COM 왕복 최소화
                    # (기존 문자 단위 순회는 O(n^2) COM 호출로 대용량 문서에서 UI 동결 발생)
                    try:
                        full_text = source_range.Text
                        if full_text and full_text.strip():
                            if not include_format:
                                text = full_text.rstrip('\r\n\x07\x0d')
                                if text:
                                    new_para.add_run(text)
                            else:
                                runs_data = self._collect_word_runs(source_range, full_text)
                                if runs_data:
                                    for run_text, fn, fs, b, it, ul, clr in runs_data:
                                        run_text = run_text.rstrip('\r\n\x07\x0d')
                                        if not run_text:
                                            continue
                                        run = new_para.add_run(run_text)
                                        try:
                                            if fn:
                                                run.font.name = fn
                                            if fs and fs > 0 and fs < 1000:
                                                run.font.size = DocxPt(fs)
                                            if b and b != 9999999:
                                                run.font.bold = True
                                            if it and it != 9999999:
                                                run.font.italic = True
                                            if ul and ul != 0 and ul != 9999999:
                                                run.font.underline = True
                                            if clr and clr != 0 and clr != 9999999:
                                                r_val = clr & 0xFF
                                                g_val = (clr >> 8) & 0xFF
                                                b_val = (clr >> 16) & 0xFF
                                                run.font.color.rgb = DocxRGBColor(r_val, g_val, b_val)
                                        except Exception as fmt_err:
                                            self.logger.log(f"  런 서식 적용 실패: {str(fmt_err)[:50]}")
                                else:
                                    # 런 분석 실패 시 전체 텍스트로 폴백 (서식 없음)
                                    text = full_text.rstrip('\r\n\x07\x0d')
                                    if text:
                                        new_para.add_run(text)
                    except Exception as range_err:
                        self.logger.log(f"  단락 {p_idx} Range 접근 실패: {str(range_err)[:50]}")
                        try:
                            text = source_range.Text.rstrip('\r\n\x07\x0d')
                            if text:
                                new_para.add_run(text)
                        except Exception:
                            pass

                except Exception as para_err:
                    self.logger.log(f"  단락 {p_idx} 처리 실패: {str(para_err)[:50]}")

            # 저장
            self.root.after(0, lambda: self.status_text.set("파일 저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(95))

            new_doc.save(save_path)
            self._validate_office_openxml(save_path, "Word 재구성")
            self.logger.log(f"저장 완료: {save_path}")
            self._log_elapsed("Word 전체 추출 시간", extract_start)

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("Word 추출 완료! (텍스트 재구성)"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"Word 추출 완료 (텍스트 재구성 방식)\n\n"
                f"{save_path}\n\n"
                f"총 {total_paragraphs}단락\n\n"
                f"⚠️ 이 방식은 서식이 일부 달라질 수 있습니다.\n"
                f"결과를 열어 원본과 비교해 주세요."))

        except Exception as e:
            self.logger.error("Word 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.word_extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()

    # ========== 메모장 관련 메서드 ==========

    def _setup_notepad_tab(self):
        """메모장 탭 설정"""
        tab = self.notepad_tab

        # 안내 프레임
        info_frame = ttk.LabelFrame(tab, text="메모장 텍스트 추출", padding="10")
        info_frame.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(info_frame,
                  text="현재 열려있는 메모장 창의 텍스트를 추출합니다.\n"
                       "메모장을 먼저 열고 '감지' 버튼을 누르세요.",
                  font=("맑은 고딕", 9), justify=tk.LEFT).pack(anchor=tk.W, pady=2)

        # 메모장 선택 콤보박스
        select_frame = ttk.Frame(info_frame)
        select_frame.pack(fill=tk.X, pady=5)
        ttk.Label(select_frame, text="메모장 선택:", width=12).pack(side=tk.LEFT)
        self.notepad_combo = ttk.Combobox(select_frame, state="readonly", width=40)
        self.notepad_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.notepad_combo.bind("<<ComboboxSelected>>", self.on_notepad_selected)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_notepad).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(tab, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.notepad_save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_notepad_save_path).pack(side=tk.LEFT)

        # 저장 형식 프레임
        format_frame = ttk.LabelFrame(tab, text="저장 형식", padding="10")
        format_frame.pack(fill=tk.X, pady=5, padx=5)

        self.notepad_save_format = tk.StringVar(value="txt")
        ttk.Radiobutton(format_frame, text="TXT (텍스트 파일)",
                        variable=self.notepad_save_format, value="txt").pack(anchor=tk.W)
        ttk.Radiobutton(format_frame, text="DOCX (Word 파일로 변환)",
                        variable=self.notepad_save_format, value="docx").pack(anchor=tk.W)

        # 추출 버튼
        self.notepad_extract_button = ttk.Button(tab, text="메모장 텍스트 추출",
                                                  command=self.start_notepad_extraction,
                                                  style="Accent.TButton")
        self.notepad_extract_button.pack(pady=10)

    def browse_notepad_save_path(self):
        """메모장 저장 경로 선택"""
        self.logger.log("메모장 저장 경로 선택")
        save_format = self.notepad_save_format.get()

        doc_name = self.notepad_doc_name.get()
        if doc_name and doc_name != "감지 중..." and doc_name != "열린 메모장 없음":
            base_name = os.path.splitext(doc_name)[0] + "_복사본"
        else:
            base_name = "새문서"

        if save_format == "docx":
            ext = ".docx"
            filetypes = [("Word 파일", "*.docx")]
        else:
            ext = ".txt"
            filetypes = [("텍스트 파일", "*.txt")]

        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=filetypes,
            initialfile=base_name + ext,
            title="저장할 위치 선택"
        )
        if path:
            self.notepad_save_path.set(path)
            self.logger.log(f"메모장 저장 경로: {path}")

    def detect_open_notepad(self):
        """열려있는 메모장 감지"""
        self.logger.log("메모장 감지 시작")
        self.status_text.set("메모장 감지 중...")
        self.notepad_doc_name.set("감지 중...")

        thread = threading.Thread(target=self._detect_notepad)
        thread.daemon = True
        thread.start()

    def _detect_notepad(self):
        """메모장 감지 (백그라운드) - Win32 API로 메모장 창 찾기.

        순수 Win32 ctypes만 사용하므로 COM 초기화 불필요.
        """
        self.logger.log("백그라운드 메모장 감지 스레드 시작")

        try:
            user32 = ctypes.windll.user32

            notepad_windows = []

            # EnumWindows 콜백
            def enum_callback(hwnd, lparam):
                if user32.IsWindowVisible(hwnd):
                    # 클래스명 확인
                    class_name = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, class_name, 256)

                    if class_name.value == "Notepad":
                        # 창 제목 가져오기
                        title_length = user32.GetWindowTextLengthW(hwnd)
                        title = ctypes.create_unicode_buffer(title_length + 1)
                        user32.GetWindowTextW(hwnd, title, title_length + 1)
                        notepad_windows.append((hwnd, title.value))
                        self.logger.log(f"  메모장 발견: hwnd={hwnd}, 제목='{title.value}'")
                return True

            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
            user32.EnumWindows(WNDENUMPROC(enum_callback), 0)

            if notepad_windows:
                self.notepad_list = notepad_windows
                names = [title if title else "제목 없음" for _, title in notepad_windows]

                def update_combo():
                    self.notepad_combo['values'] = names
                    self.notepad_combo.current(0)
                    self.notepad_doc_name.set(names[0])
                    self.status_text.set(f"메모장 {len(names)}개 감지됨")

                self.root.after(0, update_combo)
            else:
                self.notepad_list = []
                def clear_combo():
                    self.notepad_combo.set("")
                    self.notepad_combo['values'] = []
                    self.notepad_doc_name.set("열린 메모장 없음")
                    self.status_text.set("메모장을 먼저 열어주세요")
                self.root.after(0, clear_combo)

        except Exception as e:
            self.logger.error("메모장 감지 실패", e)
            self.notepad_list = []
            err_msg = str(e)[:30]
            def show_error():
                self.notepad_combo.set("")
                self.notepad_doc_name.set("열린 메모장 없음")
                self.status_text.set(f"메모장 감지 실패: {err_msg}")
            self.root.after(0, show_error)

    def on_notepad_selected(self, event):
        """메모장 콤보박스 선택 이벤트"""
        selected_idx = self.notepad_combo.current()
        if selected_idx >= 0 and selected_idx < len(self.notepad_list):
            hwnd, title = self.notepad_list[selected_idx]
            self.notepad_doc_name.set(title if title else "제목 없음")
            self.logger.log(f"메모장 선택: {title} (hwnd={hwnd})")

    def start_notepad_extraction(self):
        """메모장 추출 시작"""
        self.logger.log("메모장 추출 시작")

        save_path = self.notepad_save_path.get()
        save_format = self.notepad_save_format.get()
        selected_idx = self.notepad_combo.current()

        if not save_path:
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.notepad_doc_name.get() == "열린 메모장 없음" or not self.notepad_list:
            messagebox.showwarning("경고", "열린 메모장이 없습니다.")
            return

        if selected_idx < 0 or selected_idx >= len(self.notepad_list):
            messagebox.showwarning("경고", "메모장을 선택해주세요.")
            return

        hwnd, title = self.notepad_list[selected_idx]

        self.notepad_extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_notepad, args=(save_path, save_format, hwnd, title))
        thread.daemon = True
        thread.start()

    def _extract_notepad(self, save_path, save_format, hwnd, title):
        """메모장 추출 (백그라운드) - Win32 API로 텍스트 읽기.

        순수 Win32 ctypes만 사용하므로 COM 초기화 불필요.
        """
        self.logger.log("=== 메모장 추출 시작 ===")
        extract_start = time.perf_counter()

        try:
            user32 = ctypes.windll.user32

            self.logger.log(f"대상 메모장: {title} (hwnd={hwnd})")
            if not user32.IsWindow(hwnd):
                raise Exception("선택한 메모장 창이 닫혔습니다. 다시 감지 후 선택해주세요.")

            self.root.after(0, lambda: self.status_text.set("메모장 텍스트 읽는 중..."))
            self.root.after(0, lambda: self.progress_var.set(20))

            # 메모장 내부 Edit 컨트롤 찾기
            edit_hwnd = user32.FindWindowExW(hwnd, 0, "Edit", None)
            if not edit_hwnd:
                # Windows 11 구 스타일(레거시) 메모장에서 보일 수 있는 대체 클래스
                edit_hwnd = user32.FindWindowExW(hwnd, 0, "RichEditD2DPT", None)
            if not edit_hwnd:
                # Windows 11 새 메모장(UWP/WinUI)은 이 방식으로 접근할 수 없음 — 사용자 친화 안내
                raise Exception(
                    "Windows 11 기본 메모장(새 버전)은 자동 추출이 지원되지 않습니다.\n\n"
                    "해결 방법:\n"
                    "1. 메모장에서 Ctrl+A 로 전체 선택\n"
                    "2. Ctrl+C 로 복사\n"
                    "3. 새 메모장(또는 Word)에 Ctrl+V 로 붙여넣고 저장\n\n"
                    "또는 메모장에서 '파일 → 다른 이름으로 저장'을 직접 사용하세요."
                )

            self.logger.log(f"Edit 컨트롤 hwnd: {edit_hwnd}")

            # WM_GETTEXTLENGTH, WM_GETTEXT 로 텍스트 가져오기
            WM_GETTEXTLENGTH = 0x000E
            WM_GETTEXT = 0x000D

            text_length = user32.SendMessageW(edit_hwnd, WM_GETTEXTLENGTH, 0, 0)
            self.logger.log(f"텍스트 길이: {text_length}")

            if text_length <= 0:
                raise Exception("메모장에 텍스트가 없습니다.")

            self.root.after(0, lambda: self.progress_var.set(50))

            # 사용자가 읽는 사이 타이핑해 텍스트가 늘어날 수 있으므로 여유 버퍼(+16) 할당
            buffer_size = text_length + 16
            buffer = ctypes.create_unicode_buffer(buffer_size)
            copied = user32.SendMessageW(edit_hwnd, WM_GETTEXT, buffer_size, buffer)
            self.logger.log(f"WM_GETTEXT 반환값(실제 복사 문자 수): {copied}")

            # SendMessageW 반환값 검증 — 창이 응답 거부/행잉이면 0 또는 음수
            if copied <= 0:
                raise Exception(
                    "메모장 창에서 텍스트를 읽어올 수 없습니다.\n"
                    "창이 응답하지 않거나 닫혔을 수 있습니다. 다시 시도해 주세요."
                )

            text = buffer.value
            # 버퍼 안 읽힘 경계 검증 (비정상 절단 감지)
            if len(text) < copied:
                self.logger.log(f"  경고: 버퍼 길이({len(text)}) < 반환값({copied}) — 데이터 절단 가능")

            self.logger.log(f"텍스트 추출 성공: {len(text)}글자")
            self.root.after(0, lambda: self.progress_var.set(70))

            # 저장
            if save_format == "docx" and HAS_DOCX:
                self.root.after(0, lambda: self.status_text.set("DOCX 파일로 저장 중..."))
                new_doc = DocxDocument()
                for line in text.split('\n'):
                    line = line.rstrip('\r')
                    new_doc.add_paragraph(line)
                new_doc.save(save_path)
            else:
                self.root.after(0, lambda: self.status_text.set("TXT 파일로 저장 중..."))
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(text)

            self.logger.log(f"저장 완료: {save_path}")
            self._log_elapsed("메모장 전체 추출 시간", extract_start)
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("메모장 추출 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"메모장 추출 완료!\n{save_path}\n\n{len(text)}글자"))

        except Exception as e:
            self.logger.error("메모장 추출 오류", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"추출 중 오류:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.notepad_extract_button.config(state=tk.NORMAL))

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

    if not HAS_DOCX:
        errors.append("python-docx 패키지 필요 (pip install python-docx)")

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
