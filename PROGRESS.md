# PROGRESS.md (현재 진행: 얇게 유지)

## Dashboard
- Progress: 100%
- Risk: 낮음
- Last updated: 2026-05-15

## Latest Build
- 2026-05-15 17:10 KST: rebuilt onefile and folder EXE after HWP dialog path normalization.
- Onefile SHA256: `B1931432F693DFF2D1667877813B2C02955D05DA65E244F8AF81D189C532515C`
- Folder SHA256: `39A9855F9204EEC2C6BD5FCB2478FAC1AC81CF3024433A78E6F1360889856214`
- 2026-05-15 16:21 KST: rebuilt onefile and folder EXE after HWP Save As dialog control-submit changes.
- Onefile SHA256: `5B4C5495A3E5547C488B5F051C3CB3279AE5C5CF8079EBE686E1AD52EB60E22D`
- Folder SHA256: `DAD0AF50E1B5D05F2955262FC33C60348E81AC2A1E932F13DFB92264E295BB1D`
- 2026-05-15 15:44 KST: rebuilt onefile and folder EXE after expected Office-not-open log downgrade.
- Onefile SHA256: `F41381776A88FC2CA8092CDAADE3FEA5D2C1DF1CC93939649C3734C42CEF2783`
- Folder SHA256: `C476AB17632712F7CAE029CD924B1C0DC0A5C5C11D1677271CF8C5FE6A613C6E`
- 2026-05-15 15:26 KST: rebuilt onefile and folder EXE after HWP COM fallback changes.
- Onefile SHA256: `58EEFED4142BD4C39F997E23649A4369308EB2AD45A96809F8D345EF1C1D71BE`
- Folder SHA256: `2F3A14C4E443817944EA4A8E510E3A0AECB7D79B94AA119D2C0FF5C9228C702E`
- 2026-05-15 14:16 KST: rebuilt onefile and folder EXE with PyInstaller.
- Onefile EXE: `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`
- Folder EXE: `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3_folder\DocumentExtractor_v3.exe`
- Onefile SHA256: `D118893EA7CAC244B105EE90411E7B4397838A083D5D38FDD0F47682391133E0`
- Folder SHA256: `2F852FBABDED65D7FB06D3BE8D58FDA288B240B0EAAB18AA862AE3DD11F5A569`

## Today Goal
- v3 추출기의 PPT/Excel/Word/HWP/메모장 지원 상태 점검
- 도형, 이미지, 서식, 원본 크기 보존 중심으로 변환 안정성 개선
- 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt`에 배포

## Current State
- PPT/Excel은 기본값을 "원본 그대로 복사"로 두어 도형, 이미지, 차트, 슬라이드/시트 크기, 서식 보존을 우선한다.
- Word 빠른 복사는 `SaveAs2` 대신 파일 시스템 복사를 사용해 열린 원본 문서 상태를 바꾸지 않는다.
- HWP/HWPX 저장은 `FileSaveAs_S` 액션을 우선 사용하고, 실제 파일 생성/크기 검증을 추가했다.
- HWP 감지는 `Dispatch`로 빈 한글 문서를 만들지 않고, 열린 창 제목 폴백과 중복 실행 가드로 처리한다.
- PPT 하이브리드 도형 매핑의 중복 키를 제거했다.
- `scripts\goal_verify_v3.py`와 `run_goal_verify_v3.bat`로 목표 검증을 자동 실행할 수 있다.
- 로그 파일은 매 줄 즉시 flush하며, PPT/Excel 원본 복사 `SaveCopyAs`가 길어지면 10초마다 진행 중 로그를 남긴다.
- PPT/Excel/Word 원본 복사 결과가 OpenXML ZIP인지 검증한다. DRM 컨테이너(`SCDS`)처럼 확장자만 Office 파일인 결과는 완료 처리하지 않는다.
- PPT/Excel 원본 복사 검증 실패 시 자동으로 재구성 경로로 전환한다.
- PPT/Excel/Word 재구성 저장 결과도 OpenXML ZIP인지 검증한다.
- PPT 이미지 도형은 1차 `Export` 실패 후 클립보드 폴백이 성공하면 실패 로그를 남기지 않는다.
- PPT 변환 중 PowerPoint 경고창을 `DisplayAlerts=ppAlertsNone`으로 비활성화하고, 이미지 보존은 `Export`보다 클립보드 방식을 우선 사용한다.
- PPT `SaveCopyAs`가 DRM 컨테이너로 실패하면 `python-pptx` 재구성 전에 PowerPoint 내부 슬라이드 복제 방식으로 먼저 저장한다.
- PPT `SaveCopyAs` 실패 시 PowerPoint `SaveAs`를 거치지 않는 클립보드 슬라이드 패키지(`PowerPoint 14.0 Slides Package`) 직접 복원 경로를 우선 시도한다.
- PPT 원본 복사가 DRM으로 실패하면 요청 파일 외에 `_화면그대로.pptx` 추가본을 생성한다. 이 파일은 편집성보다 서식/크기/배치 보존을 우선한다.
- PPT 슬라이드 복제/화면 그대로/재구성 저장은 OneDrive 대상 경로에 직접 저장하지 않고 로컬 `%TEMP%`에 정상 PPTX를 만든 뒤 최종 경로로 복사한다.
- PPT `Slide.Export`까지 DRM/보안 정책으로 막히면 `slide.Copy()` 클립보드 PNG/DIB 이미지를 받아 화면 그대로 PPTX를 생성한다.
- PPT 기본 원본 복사 모드는 이미지/하이브리드 자동 전환을 하지 않는다. 편집 가능한 구조 복원 경로가 모두 실패하면 실패 처리한다.
- PPT 하이브리드/텍스트 중심 재구성에서도 텍스트가 포함된 도형/그룹은 이미지 스냅샷으로 변환하지 않는다.
- 최신 `DocumentExtractor_v3.exe`는 결과물 폴더에 복사되었고, 빌드본과 SHA256 해시가 같다.
- 회사 보안 PC 대응으로 단일 EXE 빌드의 UPX 압축을 비활성화했다.
- 회사 보안 솔루션이 단일 EXE 자가해제 방식을 차단할 때를 대비해 폴더형 배포본(`DocumentExtractor_v3_folder`)을 추가했다.
- 회사 보안 솔루션이 PyInstaller EXE 자체를 차단할 때를 대비해 Python 소스 실행 배포본(`DocumentExtractor_v3_source`)을 추가했다.
- GUI 생성 전 예외가 발생하면 `DocExtractor_Startup_Error_YYYYMMDD_HHMMSS.txt`를 바탕화면/문서/임시폴더 순서로 남긴다.
- Word 기본 원본 복사 모드는 이미지/표/머리글/도형 보존을 위해 원본 파일 복사 실패 시 텍스트 재구성으로 자동 하락하지 않는다.
- Word 원본 파일 복사가 DRM/보안 컨테이너(`SCDS`)로 감기면 Word 내부 `WordOpenXML`을 받아 표/이미지/바닥글이 포함된 정상 `.docx` 패키지 복원을 시도한다.
- Word 텍스트 재구성 모드는 `.docx` 저장만 허용한다. `.doc`, `.rtf`, `.docm`은 원본 복사 경로에서만 처리한다.
- Word/메모장 텍스트 재구성 DOCX 저장 시 XML 비호환 제어문자를 제거한다.
- 메모장 감지는 `Notepad` 클래스 외에 `- Notepad`/`- 메모장` 제목 창도 잡고, 중첩된 `Edit`/`RichEditD2DPT`/`RICHEDIT50W` 컨트롤을 탐색한다.
- 메모장 DOCX/TXT 저장 결과 검증을 추가했다.

## Verification
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 HWP UI Save As now writes normalized Windows paths (`D:\...`) to avoid "invalid file name" errors from slash paths.
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 HWP UI fallback now logs the Save As dialog tree and tries direct `WM_SETTEXT`/button click before keyboard fallback.
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 PPT/Excel/Word expected "app not open" detection is now logged as info instead of `[ERROR]`.
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 HWP COM fallback log case addressed: if COM creates a blank document, the blank instance is closed and the selected HWP window is tried through a guarded UI Save As path.
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 0, FAIL 0
- 2026-05-15 HWP checks -> PASS (`hwp_getobject_no_spawn`, `hwp_action_save`)
- 2026-05-15 Notepad check -> PASS (`notepad_legacy_read`, chars=26)
- `py -m py_compile ppt_extractor_v3.py document_extractor.py pdf_printer.py ppt_extractor.py ppt_extractor_v2.py` -> 성공
- `git diff --check` -> 성공
- `AUTOSHAPE_MAPPING` AST 중복 키 검사 -> `duplicate_keys: []`
- HWP 무생성 감지 검증 -> 성공 (`before=0`, `after=0`, 새 `Hwp.exe` 없음)
- HWP 안전 추출 검증 -> 성공 (`build\goal_verify\hwp_safe_extract_output.hwp`, 13KB)
- Word 안전 복사 검증 -> 성공 (원본 `FullName` 유지, 복사본 13KB)
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- `dist\DocumentExtractor_v3.exe` -> `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe` 복사 성공
- 배포본 SHA256 -> `C665F94300DA25625AA8C297A42BDAA29E3C11516B62805168F23A90473862B4`
- `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 사용자 PPT 로그 분석: `=== PPT 추출 프로세스 시작 ===` 이후 로그가 멈춘 것처럼 보이는 원인은 로그 flush 지연 가능성이 큼.
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `D5CC841B8719596A9C5B5B62E2D8E5AB54E73B20BFD87FA8CEE67899B3E3EF2E`
- 2026-05-12 `D:\OneDrive\프로젝트\새문서.pptx` 진단 -> PPTX ZIP 아님, 헤더 `SCDSA004`, PowerPoint에서 열 수 없는 DRM/보안 컨테이너 형태
- 2026-05-12 원본 복사 결과 OpenXML 검증 및 PPT/Excel 자동 폴백 추가
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `B87C8BA87E5DD00ABEA8D2AF08C7B6E4605C56AC20AAEC3D5DB3995D913EC9CC`
- 2026-05-12 `D:\OneDrive\프로젝트\외교부 유선지역B 사업수행계획서250117 V2 (착수보고서)_d복사본.pptx` 진단 -> 정상 PPTX, PowerPoint 열기 성공, 40장/1696도형/818그림/664미디어
- 2026-05-12 PPT 이미지 Export 1차 실패 로그를 최종 실패 시에만 남기도록 정리
- 2026-05-12 PPT/Excel/Word 재구성 저장 결과 OpenXML 검증 추가
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `8DBF0EE03DBFBA2FF7E9CA4DF58F72AAFFCA63EB987670DD1AC599EA85E95B1F`
- 2026-05-12 `D:\OneDrive\프로젝트\1111111본.pptx` 진단 -> 정상 PPTX, PowerPoint 열기 성공, 40장/1696도형/664미디어
- 2026-05-12 변환 중 "지원하지 않는 확장자로는 저장할 수 없습니다" PowerPoint 경고창 대응: PPT 추출 중 `DisplayAlerts` 비활성화
- 2026-05-12 이미지/도형 보존 순서를 `Export` 우선에서 클립보드 우선으로 변경
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `B80128F0FE9BE904619032F08603B4904CC92D7DC7F7EDC9C115A724953F3CD4`
- 2026-05-12 PPT 변환 품질 개선: PowerPoint 내부 `Slides.Range(...).Copy()` + 새 프레젠테이션 `Paste()` + `SaveAs(..., 24)` 경로 추가
- 2026-05-12 슬라이드 복제 COM 테스트 -> 성공 (`slide_clone_py_test.pptx`, 40장/1696도형)
- 2026-05-12 `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `744AA4C59BF15B47E7C599C0CD754FED2D05060DCF5CCC5EF1BA58E85259ABC9`
- 2026-05-12 산출물 크기/도형 수 확인: `1111111본.pptx`, `_d복사본.pptx`, `slide_clone_py_test.pptx` 모두 40장/538.625x808pt/1696도형
- 2026-05-12 슬라이드 전체 이미지 Export 테스트 -> 성공 (`EXPORT_OK=40`)
- 2026-05-12 원본 복사 실패 시 `_화면그대로.pptx` 추가 생성 경로 추가
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `0B86B4D46A8868FB796ADE48EEC75B136565CB302F3D5EBA6637CE9E4FF9C761`
- 2026-05-12 `121212.pptx` 로그 분석: 슬라이드 복제 자체는 성공했지만 OneDrive 대상 경로에 `SaveAs`하면서 결과가 다시 `SCDS`로 감겨 하이브리드 재구성으로 하락
- 2026-05-12 로컬 `%TEMP%` 슬라이드 복제 저장 후 OneDrive 복사 테스트 -> 정상 PPTX (`PK`, ZIP OK, 40장)
- 2026-05-12 PPT 로컬 임시 저장 후 검증된 파일만 최종 경로로 복사하는 publish 경로 추가
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `556624A0C9AA1D169FD291945211CE07CD65103F4B961B50C95F24501B53C862`
- 2026-05-12 `121212121212.pptx` 로그 분석 -> `SaveCopyAs`, 슬라이드 복제 `SaveAs`, `Slide.Export`가 모두 DRM/보안 정책에 막혀 하이브리드 재구성으로 하락했고, 그래서 양식이 크게 깨졌음
- 2026-05-12 PPT 화면 그대로 저장에 `slide.Copy()` 클립보드 PNG/DIB 폴백 추가
- 2026-05-12 강제 Export 실패 테스트 -> 클립보드 폴백으로 40장 PPTX 생성 성공 (`visual_clipboard_fallback_test.pptx`, ZIP OK, 40 slides)
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 7, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `FF1CC984F625DABD89269A96897B6FBC4304286266A56383DAF3E80CBD0027AE`
- 2026-05-12 `D:\OneDrive\1212121212.pptx` 진단 -> 정상 PPTX지만 40장 모두 그림 1개짜리 슬라이드, 편집 가능한 원본 구조는 남아 있지 않음
- 2026-05-12 PowerPoint 클립보드 `PowerPoint 14.0 Slides Package`가 ZIP/OPC 슬라이드 패키지임을 확인하고 `clipboard/` -> `ppt/` 변환으로 편집 가능한 PPTX 복원 테스트 성공
- 2026-05-12 PPT 원본 복사 실패 후 클립보드 슬라이드 패키지 복원 경로 추가
- 2026-05-12 샘플 검증 -> 텍스트/도형이 살아 있는 PPTX 생성 및 PowerPoint 열기 성공
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 8, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `0FD49DB86FE7B277C969B3E56090F17B7063B5011F3D468F09C0BF216401F795`
- 2026-05-12 사용자 요구 반영: 기본 PPT 원본 복사 모드에서 이미지/하이브리드 자동 폴백 금지
- 2026-05-12 PPT 원본 구조 복원 실패 시 이미지 산출물 생성 대신 명확한 오류 메시지를 표시하도록 변경
- 2026-05-12 UI의 `슬라이드 이미지만` 문구를 `화면 캡처용, 편집 불가`로 변경
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 8, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `972F44269DA26241D33B3BEDDE9A99F66083809588FF94CAF88913AB06AA69E6`
- 2026-05-12 사용자 요구 반영: 원본 이미지 이동은 허용하되, 텍스트가 포함된 도형/그룹의 이미지 변환 금지
- 2026-05-12 `슬라이드 이미지만` PPT 모드 UI 제거 및 기존 값 방어 로직 추가
- 2026-05-12 하이브리드 재구성의 실패 폴백을 `이미지 스냅샷`에서 `텍스트 포함 여부 검사 후 편집 가능 텍스트 복원`으로 변경
- 2026-05-12 그룹 내부 텍스트 샘플 검증 -> 결과 PPTX에 그림 0개, 편집 가능 텍스트 유지
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 8, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사 -> 성공
- 2026-05-12 배포본 SHA256 -> `C3AF3B5F9DEE46C1151D4F3B583DECBDF26043F33C35555C27CDF6E7ED2D714D`
- 2026-05-12 회사 보안 PC 대응: `DocumentExtractor_v3.spec`의 `upx=False` 변경
- 2026-05-12 회사 보안 PC 대응: `DocumentExtractor_v3_folder.spec`, `build_security_pc.bat` 추가
- 2026-05-12 `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- 2026-05-12 `git diff --check` -> 성공
- 2026-05-12 `py scripts\goal_verify_v3.py --clean` -> PASS 8, SKIP 1, FAIL 0
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-12 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3_folder.spec` -> 성공
- 2026-05-12 단일 EXE 배포본 SHA256 -> `A5DA1358FE1364185435D0DA90DDC5DE91E3AAAD4A982F7B51B809C7509EC2F3`
- 2026-05-12 폴더형 배포본 EXE SHA256 -> `F93E4D00443C54EB43CEC1300E7DD2DEB55032228E8BBBD7867B48604A056FC0`
- 2026-05-15 Word/HWP/메모장 재검토: Word 저품질 자동 폴백 차단, 메모장 감지/저장 검증 보강
- 2026-05-15 `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- 2026-05-15 `git diff --check` -> 성공
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 8, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3_folder.spec` -> 성공
- 2026-05-15 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt`로 복사 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `E3F3171B85234E0D52B80C270FAAAE3B992572498F3C92340DE68049EDB117C0`
- 2026-05-15 폴더형 배포본 EXE SHA256 -> `4D8E732CC9EBD66D017F46AE803C30BA0EC40998AC0AA2B33243549CB366B821`
- 2026-05-15 사용자 Word 로그 분석: `저장되지 않은 변경사항` 때문에 원본 복사가 실패했고, 이전 EXE가 텍스트 재구성으로 자동 하락하면서 XML 비호환 제어문자 오류가 반복됨
- 2026-05-15 Word/메모장 DOCX 텍스트 저장 전 XML 비호환 제어문자 제거 추가
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 9, SKIP 1, FAIL 0
- 2026-05-15 `word_xml_text_sanitizer` 검증 추가 -> PASS
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3_folder.spec` -> 성공
- 2026-05-15 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt`로 복사 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `343D8569F4352E22AD3EFA0A7A79BC8A9232E249A78B6E1B242073DB1E89D302`
- 2026-05-15 폴더형 배포본 EXE SHA256 -> `FEA2572FE8A905B164032B26E8ECFF291182B7ECCE50FB6A3758DBEAAFC87AB4`
- 2026-05-15 사용자 Word 로그 분석: 저장된 문서여도 파일 시스템 복사 결과가 `SCDSA004` DRM 컨테이너로 나와 정상 `.docx` 검증 실패
- 2026-05-15 Word `WordOpenXML` Flat OPC -> 일반 DOCX ZIP 패키지 복원 경로 추가
- 2026-05-15 `word_openxml_copy` 검증 추가 -> 표 1개, 이미지 1개, 바닥글 유지 PASS
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 10, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3_folder.spec` -> 성공
- 2026-05-15 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt`로 복사 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `04EC1199B63A5C151172A5F8B22197F9451723E697A9D71AF0A6B873BCFE43E3`
- 2026-05-15 폴더형 배포본 EXE SHA256 -> `0C8EF8EEBFAC6F275B1F90AF7A02EBF98245AFB37A29FB4D1AD7484217C62772`
- 2026-05-15 회사 보안 PC 대응: EXE 차단 우회용 `DocumentExtractor_v3_source` 소스 실행 배포본 추가
- 2026-05-15 `build_source_distribution.bat/.ps1`, `run_from_source.bat`, `install_requirements.bat`, `requirements_runtime.txt`, `SECURITY_PC_README.txt` 추가
- 2026-05-15 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3_source` 복사 -> 성공

- 2026-05-15 사용자 HWP 로그 분석: UI 저장 fallback은 성공했지만 `F12` 단축키가 한글 사전/부가창을 열 수 있어 저장 단축키 후보에서 제거.
- 2026-05-15 한글 UI 저장 대화상자 파일명 입력칸 탐색을 visible 필터 없이 수행하도록 보강하고, 후보/입력값 확인 로그 추가.
- 2026-05-15 `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `333569218BACFC026766735D21BB0A8A9FD41DA186E9FD0C6281FEE3FEF1EE76`

- 2026-05-15 사용자 HWP 결과물 `D:\OneDrive\프로젝트\작업중중\11한글_복사본.hwp` 확인 -> 파일 생성됨, 255468 bytes
- 2026-05-15 HWP 결과물 헤더 `SCDSA004` 확인 -> 회사 보안/DRM 컨테이너로 저장된 상태
- 2026-05-15 한글 UI 저장의 `파일명 입력값 확인 실패` 로그를 실제 실패로 보이지 않게 완화하고 SCDS 저장 결과 로그 추가
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `C0773370504CCF670FC97A387AAD4B6BEA877F1A6CA7F870941E8E9F0D644F0C`

- 2026-05-15 앱 차단 진단: 결과물 EXE `DocumentExtractor_v3.exe`는 Authenticode `NotSigned`, Zone.Identifier 없음, SHA256 `C0773370504CCF670FC97A387AAD4B6BEA877F1A6CA7F870941E8E9F0D644F0C`
- 2026-05-15 로컬 Windows Defender 상태는 비활성화되어 있고 Defender/AppLocker/CodeIntegrity 이벤트에서 `DocumentExtractor` 차단 흔적 없음
- 2026-05-15 로컬 `Start-Process` 실행 확인 -> `DocumentExtractor_v3` 프로세스 시작 성공 후 수동 종료
- 2026-05-15 결론: 현재 작업 PC에서는 차단 재현 안 됨. 다른 회사 PC에서 차단된다면 서명되지 않은 PyInstaller 단일 EXE 정책 차단 가능성이 높음

- 2026-05-15 사용자 피드백 반영: 한글 결과가 같은 회사 보안 프로그램에서만 열리면 변환 목적을 달성하지 못하므로 SCDS 결과를 성공으로 처리하지 않도록 변경
- 2026-05-15 HWP COM 저장 및 UI 저장 fallback에서 결과 헤더가 `SCDS`이면 결과 파일을 제거하고 명시적 실패 메시지 표시
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `57D52337D818A82B081D4E6068587C57194D29E7746A662728376999CF77DA8C`

- 2026-05-15 외부 자료 조사: `SCDSA004`는 SoftCamp Document Security/SDF 보안문서 헤더로 확인됨
- 2026-05-15 가능한 합법 경로는 SoftCamp SDF 공식 복호화 API(`/api/decryption`) 또는 SDF SDK(`CreateDecryptFile`) 연동이며, 둘 다 회사 IT/보안의 licenseKey/key file/권한 제공이 필요함
- 2026-05-15 pyhwp/OpenHWP/Hancom Hwp SDK는 일반 HWP/HWPX 처리에는 유효하지만 SCDS 컨테이너 자체를 권한 없이 일반 HWP로 복원하는 대안은 아님

- 2026-05-15 HWP는 SoftCamp 공식 복호화 권한 없이는 목적 달성이 불가능하므로 앱 UI에서 한글 탭 제거 결정
- 2026-05-15 앱 지원 문서 로그를 `PPT, Excel, Word, 메모장`으로 변경하고 Notebook 탭도 PPT/Excel/Word/메모장만 노출
- 2026-05-15 탭 변경 감지 인덱스를 한글 제거 후 순서(PPT=0, Excel=1, Word=2, 메모장=3)에 맞게 수정
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `7756EA5D0E4F34075F6B15ED9B0AD8166A34146A87EF9EB13BEE091AD8E74489`

- 2026-05-15 다른 PC 앱 차단/DRM 진단을 위해 `diagnose_app_block.bat`와 `diagnose_app_block.ps1` 추가
- 2026-05-15 진단 스크립트는 EXE 헤더(`MZ` vs `SCDSA004`), Authenticode 서명, Zone.Identifier/ADS, Defender/AppLocker/CodeIntegrity 이벤트, 선택적 실행 테스트를 Desktop 보고서로 저장
- 2026-05-15 로컬 진단 결과: 결과물 EXE 헤더는 정상 `MZ`, 서명은 `NotSigned`, 차단 이벤트 없음
- 2026-05-15 진단 파일을 `D:\OneDrive\코드작업\결과물\newppt`에 복사 완료

## Open Issues
- 사용자 실제 문서 기준의 HWP/Word/메모장 수동 검증은 아직 필요하다.
- Word 원본 파일 복사는 저장된 문서와 동일 확장자일 때만 안전 경로를 탄다. 기본값에서는 저장 안 된 문서나 확장자 변환을 텍스트 재구성으로 자동 하락시키지 않는다.
- Excel 재구성 모드의 도형/이미지 위치는 TopLeftCell 기준 보존이며, 셀 내부 픽셀 오프셋은 근사치다.
- PPT `_화면그대로.pptx` 추가본은 시각 보존용이라 각 슬라이드가 이미지로 들어간다. 현재 기본 UI에서는 텍스트 이미지화를 막기 위해 직접 선택 경로를 제거했다.
- PPT 원본 복사 실패 후 클립보드 슬라이드 패키지가 제공되면 편집 가능한 원본 구조 복원을 우선한다. 이 패키지가 DRM/보안 정책으로 제공되지 않으면 슬라이드 복제까지만 시도하고, 실패 시 이미지 산출물 없이 오류 처리한다.
- PowerPoint 경고창이 계속 뜨면 기존 실행 중인 EXE/PowerPoint를 닫고 최신 EXE로 재실행해야 한다.
- `SCDSA004`처럼 이미 DRM 컨테이너로 저장된 파일은 그 파일만으로 정상 PPTX 복구가 불가능하다. 원본 문서를 PowerPoint에 연 상태에서 새 EXE로 다시 추출해야 한다.
- Windows 11 새 메모장은 자동 검증에서 SKIP될 수 있다. 현재 Win32 Edit/RichEdit 계열 컨트롤 탐색 방식으로 제한적이라 UI Automation 검토 여지가 있다.
- `ppt_extractor_v3.py`에는 진단 없이 삼키는 `except:`/`except: pass`가 일부 남아 있다.
- 회사 보안 PC에서 시작 오류 로그도 생성되지 않으면 Python 코드 진입 전 보안 솔루션이 실행 파일을 차단한 것으로 보고 격리/차단 로그 또는 허용 정책 확인이 필요하다.
- 회사 보안 PC에서는 먼저 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3_folder\DocumentExtractor_v3.exe`를 실행한다. 폴더 내부 파일을 분리하면 실행이 깨질 수 있다.
- 폴더형 EXE도 차단되면 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3_source\run_from_source.bat`를 사용한다. 이 방식은 대상 PC에 Python 3와 런타임 패키지가 필요하다.
- 회사망에서 `pip install`이 차단되면 소스 실행 배포본도 IT 보안팀의 Python 패키지 설치/내부 저장소 설정이 필요하다.
- HWP는 실제 사용자 문서 기준 다중 창/다중 문서 선택 검증이 아직 필요하다. COM 연결이 특정 창을 정확히 지정하지 못하는 환경에서는 사용자가 대상 문서를 활성화한 뒤 추출해야 한다.
- Windows 11 새 메모장이 WinUI 전용 텍스트 컨트롤만 노출하면 Win32 방식으로 자동 추출할 수 없다. 현재 검증 환경에서도 Notepad 자동 검증은 SKIP 상태다.
- WordOpenXML 복원은 `.docx` 대상으로 검증했다. `.docm` 매크로, OLE 임베딩, 특수 보안 플러그인 개체는 추가 실문서 검증이 필요하다.

- 2026-05-15 UI/UX 개편: Notebook 탭을 좌측 문서 선택 사이드바 + 우측 작업 패널 구조로 변경
- 2026-05-15 지원 메뉴는 PPT, Excel, Word, 메모장 4개로 정리하고 HWP/한글은 DRM 제약 안내 문구만 사이드바 하단에 표시
- 2026-05-15 `_do_detect` 감지 흐름을 새 `current_doc_index` 기반으로 정리하고 기존 PPT/Excel/Word/메모장 감지 함수는 그대로 재사용
- 2026-05-15 병렬 검토 결과 반영: 기존 탭별 `_setup_*_tab` 구성은 유지하고 프레임 raise 방식으로 화면 전환
- 2026-05-15 `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `F9BAFED94EEC3CF697BEDB040B512F50E9C015283DAA1CA1ECBE0546427F2F9E`

- 2026-05-15 사용자 피드백 반영: 검은 사이드바와 기본 Windows 회색 폼 느낌을 줄이고 밝은 내비게이션 + 흰색 카드형 섹션으로 UI 재정리
- 2026-05-15 선택 메뉴는 배지형 문서 타입(PPT/XLS/DOC/TXT) + 설명 텍스트로 정리하고, 섹션/입력창/버튼 색상을 같은 톤으로 맞춤
- 2026-05-15 `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `58E225FF0CB5AFDCE0F536027962883DE0FF29177780D373DFB8571733122FE1`

- 2026-05-15 일괄 변환 기능 추가: 좌측 메뉴에 `일괄 변환` 화면을 추가하고 파일/폴더 추가, 선택 제거, 목록 비우기, 출력 폴더 선택, 일괄 변환 시작 UI 구현
- 2026-05-15 일괄 변환 지원 확장자: PPT/PPTX/PPTM/PPSX/POTX, XLS/XLSX/XLSM/XLSB, DOC/DOCX/DOCM, TXT. HWP는 기존 결정대로 제외
- 2026-05-15 일괄 변환 처리: Office 파일은 COM으로 열어 원본 복사/구조 복원 경로를 재사용하고 TXT는 파일 복사로 처리
- 2026-05-15 TXT 일괄 변환 smoke test -> `sample_복사본.txt` 생성 성공
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `C872C62A71A00992771275DBFA5795D625A88DBA39060ECDCE6A07FED217FBDD`

- 2026-05-15 개별 PPT/Excel/Word/메모장 화면에도 `파일 선택` 직접 변환 경로 추가. 문서를 미리 열지 않아도 선택 파일을 내부적으로 열어 변환 가능
- 2026-05-15 파일 선택 시 저장 경로는 원본 폴더의 `_복사본` 이름으로 자동 채움. 기존 열린 문서 감지/추출 경로도 계속 유지
- 2026-05-15 TXT 직접 변환 smoke test -> TXT 복사 및 DOCX 변환 성공
- 2026-05-15 화면 높이를 `900x660`으로 조정해 직접 파일 선택 UI와 추출 버튼이 잘리지 않게 수정
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `3CAA4102C71E8E07F97C8BFF058FCFD3E11B4BCA407F3970AB412FB2FE9C920A`

- 2026-05-15 사용자 로그 반영: PPT 일괄 변환 중 `SaveCopyAs` 실패 후 PowerPoint `다른 이름으로 저장` 대화상자가 남는 문제 확인
- 2026-05-15 PPT 일괄/직접 변환 개선: 정상 PPTX/PPTM/PPSX/POTX는 PowerPoint를 열지 않고 OpenXML 패키지 검증 후 파일 직접 복사
- 2026-05-15 PPT 내부 복원이 필요한 DRM/보안 케이스에서는 `SaveCopyAs`를 건너뛰고 클립보드 패키지 -> 슬라이드 복제 순서만 시도하도록 변경
- 2026-05-15 PowerPoint 프로세스의 남은 modal dialog(`#32770`)를 자동 닫는 정리 루틴 추가
- 2026-05-15 정상 PPTX no-open smoke test -> PowerPoint 연결 없이 직접 복사 성공
- 2026-05-15 `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- 2026-05-15 `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 2026-05-15 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 2026-05-15 단일 EXE 배포본 SHA256 -> `5C246F6F0AA2CC1B58FABF105B240CC75301DE49FE901D410F4CDE6E11321CAE`

## Next
1. 실제 사용자 문서로 Word/메모장 탭 수동 확인
2. Windows 11 새 메모장 지원 강화를 위해 UI Automation 경로 검토
3. 회사 보안 PC에서 폴더형 배포본 실행 여부와 `DocExtractor_Startup_Error_*.txt` 생성 여부 확인

---
## Archive Rule
- 완료 항목이 20개를 넘거나 파일이 5KB를 넘으면 완료된 내용을 `ARCHIVE_YYYY_MM.md`로 옮기고, `PROGRESS.md`는 현재 이슈만 남긴다.
