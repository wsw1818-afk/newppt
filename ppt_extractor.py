"""
PPT 보안 해제 도구
- 암호화된 PPT에서 COM으로 직접 데이터 읽기
- 새 PPT에 데이터 쓰기 (클립보드 우회)
- DRM 우회 가능
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
        self.log_path = os.path.join(desktop, f"PPTExtractor_Log_{timestamp}.txt")
        self.log_file = open(self.log_path, "w", encoding="utf-8")
        self.log(f"=== PPT Extractor 로그 시작 ===")
        self.log(f"로그 파일: {self.log_path}")
        self.log(f"시작 시간: {datetime.datetime.now()}")
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


class PPTExtractor:
    """PPT 추출기"""

    def __init__(self):
        self.logger = Logger()
        self.logger.log("PPTExtractor 초기화 시작")

        self.root = tk.Tk()
        self.root.title("PPT 보안 해제 도구")
        self.root.geometry("550x500")
        self.root.resizable(False, False)

        # 상태 변수
        self.current_doc_name = tk.StringVar(value="감지 중...")
        self.slide_count = tk.StringVar(value="-")
        self.save_path = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="프로그램 시작됨")
        self.progress_var = tk.DoubleVar(value=0)

        self.setup_ui()
        self.detect_open_ppt()

        self.logger.log("PPTExtractor 초기화 완료")

    def setup_ui(self):
        """UI 구성"""
        self.logger.log("UI 구성 시작")

        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(main_frame, text="PPT 보안 해제 도구",
                                font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # 설명
        desc_label = ttk.Label(main_frame,
                               text="암호화된 PPT를 COM으로 직접 읽어서 새 파일로 저장\n(클립보드를 사용하지 않아 DRM 우회 가능)",
                               font=("맑은 고딕", 9), justify=tk.CENTER)
        desc_label.pack(pady=(0, 15))

        # 문서 정보 프레임
        info_frame = ttk.LabelFrame(main_frame, text="현재 열린 PPT", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))

        # 문서명
        doc_frame = ttk.Frame(info_frame)
        doc_frame.pack(fill=tk.X, pady=2)
        ttk.Label(doc_frame, text="문서명:", width=12).pack(side=tk.LEFT)
        ttk.Label(doc_frame, textvariable=self.current_doc_name,
                  font=("맑은 고딕", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 슬라이드 수
        slide_frame = ttk.Frame(info_frame)
        slide_frame.pack(fill=tk.X, pady=2)
        ttk.Label(slide_frame, text="슬라이드 수:", width=12).pack(side=tk.LEFT)
        ttk.Label(slide_frame, textvariable=self.slide_count,
                  font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        ttk.Button(info_frame, text="다시 감지", command=self.detect_open_ppt).pack(pady=(10, 0))

        # 저장 경로 프레임
        path_frame = ttk.LabelFrame(main_frame, text="새 파일 저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=(0, 15))

        path_inner = ttk.Frame(path_frame)
        path_inner.pack(fill=tk.X)
        ttk.Entry(path_inner, textvariable=self.save_path, width=45).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_inner, text="찾아보기", command=self.browse_save_path).pack(side=tk.LEFT)

        # 옵션 프레임
        option_frame = ttk.LabelFrame(main_frame, text="추출 옵션", padding="10")
        option_frame.pack(fill=tk.X, pady=(0, 15))

        self.copy_shapes = tk.BooleanVar(value=True)
        self.copy_texts = tk.BooleanVar(value=True)
        self.copy_images = tk.BooleanVar(value=True)

        ttk.Checkbutton(option_frame, text="도형 복사", variable=self.copy_shapes).pack(anchor=tk.W)
        ttk.Checkbutton(option_frame, text="텍스트 복사", variable=self.copy_texts).pack(anchor=tk.W)
        ttk.Checkbutton(option_frame, text="이미지 복사 (임시 파일로 추출)", variable=self.copy_images).pack(anchor=tk.W)

        # 추출 버튼
        self.extract_button = ttk.Button(main_frame, text="새 PPT로 추출",
                                          command=self.start_extraction,
                                          style="Accent.TButton")
        self.extract_button.pack(pady=10)

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

        self.logger.log("UI 구성 완료")

    def browse_save_path(self):
        """저장 경로 선택"""
        self.logger.log("저장 경로 선택 대화상자 열기")

        doc_name = self.current_doc_name.get()
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
            self.save_path.set(path)
            self.logger.log(f"저장 경로 선택됨: {path}")
        else:
            self.logger.log("저장 경로 선택 취소됨")

    def detect_open_ppt(self):
        """열려있는 PPT 감지"""
        self.logger.log("PPT 감지 시작")
        self.status_text.set("PPT 감지 중...")
        self.current_doc_name.set("감지 중...")
        self.slide_count.set("-")

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
            self.logger.log(f"PowerPoint 연결 성공, 열린 프레젠테이션 수: {ppt.Presentations.Count}")

            if ppt.Presentations.Count > 0:
                presentation = ppt.ActivePresentation
                name = presentation.Name
                count = presentation.Slides.Count

                self.logger.log(f"활성 프레젠테이션: {name}")
                self.logger.log(f"슬라이드 수: {count}")
                self.logger.log(f"파일 경로: {presentation.FullName}")

                self.root.after(0, lambda: self.current_doc_name.set(name))
                self.root.after(0, lambda: self.slide_count.set(f"{count}장"))
                self.root.after(0, lambda: self.status_text.set("PPT 감지 완료"))
            else:
                self.logger.log("열린 프레젠테이션 없음")
                self.root.after(0, lambda: self.current_doc_name.set("열린 PPT 없음"))
                self.root.after(0, lambda: self.slide_count.set("-"))
                self.root.after(0, lambda: self.status_text.set("PPT를 먼저 열어주세요"))

        except Exception as e:
            self.logger.error("PPT 감지 실패", e)
            self.root.after(0, lambda: self.current_doc_name.set("열린 PPT 없음"))
            self.root.after(0, lambda: self.slide_count.set("-"))
            self.root.after(0, lambda: self.status_text.set(f"감지 실패: {str(e)[:30]}"))

        pythoncom.CoUninitialize()
        self.logger.log("백그라운드 PPT 감지 스레드 종료")

    def start_extraction(self):
        """추출 시작"""
        self.logger.log("추출 시작 버튼 클릭")

        if not self.save_path.get():
            self.logger.log("저장 경로 미선택 - 경고 표시")
            messagebox.showwarning("경고", "저장 경로를 선택해주세요.")
            return

        if self.current_doc_name.get() == "열린 PPT 없음":
            self.logger.log("열린 PPT 없음 - 경고 표시")
            messagebox.showwarning("경고", "열린 PPT가 없습니다.")
            return

        self.logger.log(f"추출 시작: 저장 경로 = {self.save_path.get()}")
        self.extract_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._extract_ppt)
        thread.daemon = True
        thread.start()

    def _extract_ppt(self):
        """PPT 추출 (백그라운드)"""
        self.logger.log("=== PPT 추출 프로세스 시작 ===")
        pythoncom.CoInitialize()

        temp_dir = None
        new_pres = None

        try:
            save_path = self.save_path.get()
            self.logger.log(f"저장 경로: {save_path}")

            # 원본 PPT 가져오기
            self.logger.log("원본 PPT 연결 시도...")
            self.root.after(0, lambda: self.status_text.set("원본 PPT 연결 중..."))

            ppt_app = win32com.client.GetObject(Class="PowerPoint.Application")
            self.logger.log("PowerPoint.Application 연결 성공")

            source_pres = ppt_app.ActivePresentation
            self.logger.log(f"원본 프레젠테이션: {source_pres.Name}")
            self.logger.log(f"원본 경로: {source_pres.FullName}")

            # 새 PPT 생성
            self.logger.log("새 PPT 생성 시도...")
            self.root.after(0, lambda: self.status_text.set("새 PPT 생성 중..."))
            self.root.after(0, lambda: self.progress_var.set(5))

            new_pres = ppt_app.Presentations.Add(WithWindow=False)
            self.logger.log("새 프레젠테이션 생성 성공")

            # 슬라이드 크기 복사
            self.logger.log(f"원본 슬라이드 크기: {source_pres.PageSetup.SlideWidth} x {source_pres.PageSetup.SlideHeight}")
            new_pres.PageSetup.SlideWidth = source_pres.PageSetup.SlideWidth
            new_pres.PageSetup.SlideHeight = source_pres.PageSetup.SlideHeight
            self.logger.log("슬라이드 크기 복사 완료")

            total_slides = source_pres.Slides.Count
            self.logger.log(f"총 슬라이드 수: {total_slides}")

            # 임시 디렉토리 생성 (이미지 추출용)
            temp_dir = tempfile.mkdtemp()
            self.logger.log(f"임시 디렉토리 생성: {temp_dir}")

            # 각 슬라이드 복사
            for i in range(1, total_slides + 1):
                self.logger.log(f"--- 슬라이드 {i}/{total_slides} 처리 시작 ---")

                progress = 5 + (i / total_slides) * 90
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                self.root.after(0, lambda n=i, t=total_slides: self.status_text.set(f"슬라이드 {n}/{t} 처리 중..."))

                source_slide = source_pres.Slides(i)
                self.logger.log(f"원본 슬라이드 {i} 가져옴")

                # 새 슬라이드 추가 (빈 레이아웃)
                self.logger.log(f"새 슬라이드 {i} 추가 시도 (ppLayoutBlank = 12)")
                new_slide = new_pres.Slides.Add(i, 12)
                self.logger.log(f"새 슬라이드 {i} 추가 성공")

                # 배경 복사 시도
                try:
                    self.logger.log("배경 복사 시도...")
                    source_slide.Background.Fill.Solid()
                    bg_color = source_slide.Background.Fill.ForeColor.RGB
                    new_slide.Background.Fill.Solid()
                    new_slide.Background.Fill.ForeColor.RGB = bg_color
                    self.logger.log(f"배경 복사 성공 (RGB: {bg_color})")
                except Exception as e:
                    self.logger.log(f"배경 복사 실패 (무시): {str(e)}")

                # 도형들 복사
                shape_count = source_slide.Shapes.Count
                self.logger.log(f"슬라이드 {i}의 도형 수: {shape_count}")

                for j, shape in enumerate(source_slide.Shapes, 1):
                    try:
                        self.logger.log(f"  도형 {j}/{shape_count} 복사 시도: Type={shape.Type}, Name={shape.Name}")
                        self._copy_shape(shape, new_slide, temp_dir)
                        self.logger.log(f"  도형 {j}/{shape_count} 복사 성공")
                    except Exception as e:
                        self.logger.error(f"  도형 {j}/{shape_count} 복사 실패", e)
                        continue

                self.logger.log(f"--- 슬라이드 {i}/{total_slides} 처리 완료 ---")

            # 저장
            self.logger.log("=== 저장 시작 ===")
            self.root.after(0, lambda: self.status_text.set("저장 중..."))
            self.root.after(0, lambda: self.progress_var.set(95))

            self.logger.log(f"SaveAs 호출: {save_path}, 형식=24 (ppSaveAsOpenXMLPresentation)")
            new_pres.SaveAs(save_path, 24)
            self.logger.log("SaveAs 완료")

            self.logger.log("프레젠테이션 닫기...")
            new_pres.Close()
            new_pres = None
            self.logger.log("프레젠테이션 닫기 완료")

            # 파일 존재 확인
            if os.path.exists(save_path):
                file_size = os.path.getsize(save_path)
                self.logger.log(f"저장된 파일 확인: {save_path} (크기: {file_size} bytes)")
            else:
                self.logger.log(f"[경고] 저장된 파일을 찾을 수 없음: {save_path}")

            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set("추출 완료!"))
            self.root.after(0, lambda: messagebox.showinfo("완료",
                f"PPT 추출 완료!\n{save_path}\n\n"
                f"총 {total_slides}장의 슬라이드가 복사되었습니다.\n\n"
                f"로그 파일: {self.logger.log_path}"))

            self.logger.log("=== PPT 추출 프로세스 성공적으로 완료 ===")

        except Exception as e:
            self.logger.error("PPT 추출 중 오류 발생", e)
            self.root.after(0, lambda: self.status_text.set(f"오류: {str(e)[:50]}"))
            self.root.after(0, lambda: messagebox.showerror("오류",
                f"추출 중 오류 발생:\n{str(e)}\n\n"
                f"상세 로그: {self.logger.log_path}"))

            # 실패 시 새 프레젠테이션 닫기 시도
            if new_pres:
                try:
                    self.logger.log("오류 발생으로 새 프레젠테이션 닫기 시도...")
                    new_pres.Close()
                    self.logger.log("새 프레젠테이션 닫기 성공")
                except Exception as close_e:
                    self.logger.error("새 프레젠테이션 닫기 실패", close_e)

        finally:
            # 임시 파일 정리
            if temp_dir and os.path.exists(temp_dir):
                try:
                    self.logger.log(f"임시 디렉토리 정리: {temp_dir}")
                    shutil.rmtree(temp_dir)
                    self.logger.log("임시 디렉토리 정리 완료")
                except Exception as e:
                    self.logger.error("임시 디렉토리 정리 실패", e)

            self.root.after(0, lambda: self.extract_button.config(state=tk.NORMAL))
            pythoncom.CoUninitialize()
            self.logger.log("COM 해제 완료")

    def _copy_shape(self, source_shape, target_slide, temp_dir):
        """개별 도형 복사"""
        shape_type = source_shape.Type
        shape_name = source_shape.Name

        self.logger.log(f"    _copy_shape: Type={shape_type}, Name={shape_name}")

        # msoTextBox = 17, msoAutoShape = 1, msoFreeform = 5
        # msoPicture = 13, msoLinkedPicture = 11
        # msoPlaceholder = 14, msoTextEffect = 15
        # msoGroup = 6, msoTable = 19, msoChart = 3

        try:
            # 위치 및 크기 정보
            left = source_shape.Left
            top = source_shape.Top
            width = source_shape.Width
            height = source_shape.Height
            self.logger.log(f"    위치: Left={left}, Top={top}, Width={width}, Height={height}")

            # 그림인 경우
            if shape_type in [13, 11]:  # msoPicture, msoLinkedPicture
                self.logger.log(f"    이미지 타입 감지 (Type={shape_type})")
                if self.copy_images.get():
                    img_path = os.path.join(temp_dir, f"img_{id(source_shape)}.png")
                    self.logger.log(f"    이미지 Export 시도: {img_path}")
                    source_shape.Export(img_path, 2)  # ppShapeFormatPNG = 2
                    self.logger.log(f"    이미지 Export 성공")

                    self.logger.log(f"    AddPicture 시도")
                    target_slide.Shapes.AddPicture(
                        img_path,
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=left,
                        Top=top,
                        Width=width,
                        Height=height
                    )
                    self.logger.log(f"    AddPicture 성공")
                return

            # 텍스트 박스/도형인 경우
            if shape_type in [1, 5, 14, 17]:  # AutoShape, Freeform, Placeholder, TextBox
                self.logger.log(f"    텍스트/도형 타입 감지 (Type={shape_type})")
                if self.copy_shapes.get():
                    if shape_type == 17:  # TextBox
                        self.logger.log(f"    AddTextbox 시도")
                        new_shape = target_slide.Shapes.AddTextbox(
                            1,  # msoTextOrientationHorizontal
                            left, top, width, height
                        )
                        self.logger.log(f"    AddTextbox 성공")
                    else:
                        self.logger.log(f"    AddShape (Rectangle) 시도")
                        new_shape = target_slide.Shapes.AddShape(
                            1,  # msoShapeRectangle
                            left, top, width, height
                        )
                        self.logger.log(f"    AddShape 성공")

                    # 텍스트 복사
                    if self.copy_texts.get():
                        try:
                            if source_shape.HasTextFrame:
                                if source_shape.TextFrame.HasText:
                                    text = source_shape.TextFrame.TextRange.Text
                                    self.logger.log(f"    텍스트 복사: '{text[:50]}...' (길이: {len(text)})")
                                    new_shape.TextFrame.TextRange.Text = text

                                    # 폰트 복사 시도
                                    try:
                                        src_font = source_shape.TextFrame.TextRange.Font
                                        dst_font = new_shape.TextFrame.TextRange.Font
                                        dst_font.Name = src_font.Name
                                        dst_font.Size = src_font.Size
                                        dst_font.Bold = src_font.Bold
                                        dst_font.Italic = src_font.Italic
                                        dst_font.Color.RGB = src_font.Color.RGB
                                        self.logger.log(f"    폰트 복사 성공: {src_font.Name}, {src_font.Size}pt")
                                    except Exception as e:
                                        self.logger.log(f"    폰트 복사 실패 (무시): {str(e)}")
                        except Exception as e:
                            self.logger.log(f"    텍스트 복사 실패 (무시): {str(e)}")

                    # 도형 스타일 복사 시도
                    try:
                        new_shape.Fill.ForeColor.RGB = source_shape.Fill.ForeColor.RGB
                        self.logger.log(f"    Fill 색상 복사 성공")
                    except Exception as e:
                        self.logger.log(f"    Fill 색상 복사 실패 (무시): {str(e)}")

                    try:
                        new_shape.Line.ForeColor.RGB = source_shape.Line.ForeColor.RGB
                        new_shape.Line.Weight = source_shape.Line.Weight
                        self.logger.log(f"    Line 스타일 복사 성공")
                    except Exception as e:
                        self.logger.log(f"    Line 스타일 복사 실패 (무시): {str(e)}")

                return

            # 테이블인 경우
            if shape_type == 19:  # msoTable
                self.logger.log(f"    테이블 타입 감지")
                if self.copy_shapes.get():
                    rows = source_shape.Table.Rows.Count
                    cols = source_shape.Table.Columns.Count
                    self.logger.log(f"    테이블 크기: {rows}행 x {cols}열")

                    self.logger.log(f"    AddTable 시도")
                    new_table = target_slide.Shapes.AddTable(rows, cols, left, top, width, height)
                    self.logger.log(f"    AddTable 성공")

                    # 셀 데이터 복사
                    for r in range(1, rows + 1):
                        for c in range(1, cols + 1):
                            try:
                                src_cell = source_shape.Table.Cell(r, c)
                                dst_cell = new_table.Table.Cell(r, c)

                                if src_cell.Shape.HasTextFrame:
                                    if src_cell.Shape.TextFrame.HasText:
                                        text = src_cell.Shape.TextFrame.TextRange.Text
                                        dst_cell.Shape.TextFrame.TextRange.Text = text
                            except Exception as e:
                                self.logger.log(f"    셀({r},{c}) 복사 실패 (무시): {str(e)}")

                    self.logger.log(f"    테이블 데이터 복사 완료")
                return

            # Connector (선/연결선) - Export 불가, 직접 그리기
            if shape_type == 9:  # msoConnector
                self.logger.log(f"    Connector(선) 타입 감지 - 직접 선 그리기 시도")
                if self.copy_shapes.get():
                    try:
                        # 선의 시작점과 끝점 가져오기
                        begin_x = source_shape.BeginConnectedShape if hasattr(source_shape, 'BeginConnectedShape') else left
                        begin_y = top
                        end_x = left + width
                        end_y = top + height

                        # 직선으로 추가
                        self.logger.log(f"    AddLine 시도: ({left}, {top}) -> ({left+width}, {top+height})")
                        new_line = target_slide.Shapes.AddLine(left, top, left + width, top + height)
                        self.logger.log(f"    AddLine 성공")

                        # 선 스타일 복사
                        try:
                            new_line.Line.ForeColor.RGB = source_shape.Line.ForeColor.RGB
                            new_line.Line.Weight = source_shape.Line.Weight
                            self.logger.log(f"    선 스타일 복사 성공")
                        except Exception as e:
                            self.logger.log(f"    선 스타일 복사 실패 (무시): {str(e)}")
                    except Exception as e:
                        self.logger.log(f"    Connector 그리기 실패 (건너뜀): {str(e)}")
                return

            # 그룹인 경우
            if shape_type == 6:  # msoGroup
                self.logger.log(f"    그룹 타입 감지, 내부 도형 수: {source_shape.GroupItems.Count}")
                group_count = source_shape.GroupItems.Count
                for idx, item in enumerate(source_shape.GroupItems, 1):
                    try:
                        self.logger.log(f"      그룹 내 도형 {idx}/{group_count} 처리")
                        self._copy_shape(item, target_slide, temp_dir)
                    except Exception as e:
                        self.logger.log(f"      그룹 내 도형 {idx}/{group_count} 복사 실패 (무시): {str(e)}")
                return

            # 기타 도형: 이미지로 내보내기 후 삽입
            self.logger.log(f"    기타 도형 타입 (Type={shape_type}) - 이미지로 변환 시도")
            if self.copy_shapes.get():
                try:
                    img_path = os.path.join(temp_dir, f"shape_{id(source_shape)}.png")
                    self.logger.log(f"    Export 시도: {img_path}")
                    source_shape.Export(img_path, 2)
                    self.logger.log(f"    Export 성공")

                    self.logger.log(f"    AddPicture 시도")
                    target_slide.Shapes.AddPicture(
                        img_path,
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=left,
                        Top=top,
                        Width=width,
                        Height=height
                    )
                    self.logger.log(f"    AddPicture 성공")
                except Exception as e:
                    self.logger.log(f"    기타 도형 이미지 변환 실패 (건너뜀): {str(e)}")

        except Exception as e:
            self.logger.error(f"    도형 복사 메인 로직 실패, fallback 시도", e)
            # Connector 타입(9)은 Export 불가능하므로 건너뛰기
            if shape_type == 9:
                self.logger.log(f"    Connector(선) 타입은 fallback도 불가 - 건너뜀")
                return
            # 실패 시 이미지로 추출 시도
            try:
                img_path = os.path.join(temp_dir, f"fallback_{id(source_shape)}.png")
                self.logger.log(f"    Fallback Export 시도: {img_path}")
                source_shape.Export(img_path, 2)

                target_slide.Shapes.AddPicture(
                    img_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=source_shape.Left,
                    Top=source_shape.Top,
                    Width=source_shape.Width,
                    Height=source_shape.Height
                )
                self.logger.log(f"    Fallback 성공")
            except Exception as e2:
                self.logger.log(f"    Fallback도 실패 (건너뜀): {str(e2)}")

    def run(self):
        """프로그램 실행"""
        self.logger.log("메인 루프 시작")
        self.root.mainloop()
        self.logger.log("메인 루프 종료")
        self.logger.close()


def check_dependencies():
    """의존성 확인"""
    if not HAS_WIN32COM:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("의존성 오류",
            "pywin32 패키지가 필요합니다.\n\n"
            "설치 명령어:\npip install pywin32")
        sys.exit(1)


if __name__ == "__main__":
    check_dependencies()
    app = PPTExtractor()
    app.run()
