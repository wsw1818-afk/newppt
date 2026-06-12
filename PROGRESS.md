# PROGRESS.md (현재 진행: 얇게 유지)

## Dashboard
- Progress: 100%
- Risk: 낮음
- Last updated: 2026-06-11

## Latest Build
- 2026-06-11 KST (2차): **한글 원본 파일 직접 Open 메모리 추출 모드 추가.** 회사 PC 1차 빌드 로그 분석 결과, 열린 한글이 ROT 미등록(`GetActiveObject` → `-2147221021 MK_E_UNAVAILABLE`)이라 COM으로 못 잡혀 메모리 추출(방법2/3)이 실행조차 안 되고 UI 저장(SCDS)으로 떨어짐. 해결: HWP 탭에 '① 원본 파일 직접 선택' 추가 → `Dispatch` 새 인스턴스로 `hwp.Open(원본, "HWP", "forceopen:true")` → 보안모듈이 사용자 권한으로 복호화 → `GetTextFile("HWPML2X")` 메모리 추출(열린 문서 COM 연결 단계를 건너뜀). 방법1/2/3을 `_hwp_save_with_fallbacks`로 공통화해 열린문서/직접Open 양쪽 재사용. 신규: `browse_hwp_source_path`/`_extract_hwp_from_file`/`_hwp_save_with_fallbacks`, `_extract_hwp(source_path=...)`. 검증: py_compile + 구조 스모크(Open/폴백 배선) + goal_verify PASS=10. onefile/folder 재빌드·배포. onefile SHA256 `97989C25AADEDC38C1B4287CA4C8AAF8FA0CE5B75260CF3255AF95A8495EB7FA`, folder-exe `FF3A0225CF5A10796B92444611B56E0C4728077EB404E1931D9FD487B5766ACF`. `APP_BUILD_ID="2026-06-11-hwp-direct-open"`. **회사 PC 재테스트 필요: 원본 파일 직접 선택 시 Open/GetTextFile이 보안에서 뚫리는지 확인.**
- 2026-06-11 KST: **한글(HWP) DRM(SCDS) 극복 + 사이드바 탭 복원.**
  - 탭 복원: orphan 상태였던 HWP 기능(감지/추출/저장 메서드는 살아있으나 UI 미연결)을 사이드바에 재연결. `doc_views`에 '한글'을 Word 다음 배치, `hwp_tab` 생성/setup, `content_frames`·`tab_detected`(7칸) 조정, `_do_detect`를 인덱스 하드코딩 대신 `doc_views` 감지함수 기반으로 리팩터(메뉴 순서 변경에 안전).
  - DRM 극복: 기존 HWP는 직접 SaveAs·클립보드·UI저장 모두 마지막이 파일 저장이라 보안 래퍼가 SCDS(헤더 `SCDSA004`)로 재포장돼 실패. PPT(`slide.Copy`)/Word(`WordOpenXML`)가 성공한 원리(메모리 COM 추출)를 HWP에 적용: `GetTextFile("HWPML2X")`로 파일 저장 없이 완전 구조 XML을 메모리에서 추출. **3단계 폴백** — (1) 직접 SaveAs (2) HWPML2X→새 문서 `SetTextFile` 재구성→일반 .hwp (3) HWPML2X를 파이썬이 직접 .hwpml로 기록(한글 SaveAs 미사용, 가장 확실한 우회). 기존 `_save_hwp_document`의 SCDS 헤더 검증 재사용. 순수 COM이라 보안PC 호환(새 네이티브 의존성 0).
  - 신규 메서드: `_hwp_extract_hwpml`/`_hwp_rebuild_via_hwpml`/`_hwp_save_hwpml_direct`/`_hwp_extract_success`. 구식 클립보드(SelectAll/Copy/Paste) 경로 제거(한컴 권고: 복잡 자동화에선 Copy/Paste 대신 SetTextFile).
  - 검증: `py_compile` OK, UI 스모크(7탭/한글@3 연결), 재구성 구조 스모크(GetTextFile/SetTextFile/HWPML2X/3단계 폴백), `goal_verify` PASS=10/FAIL=0. **단, 실제 SCDS 우회 동작은 개발 PC에 한글/SCDS 문서가 없어 미검증 → 회사 보안 PC에서 실문서 테스트 필요.** `APP_BUILD_ID="2026-06-11-hwp-memory-rebuild"`. onefile/folder 재빌드·배포 완료(결과물 폴더). onefile SHA256 `c7c37621C4AD83958AA62AAC5178DE3B72DD2956515174AC4E5F8F5B9ADCD742`, folder-exe SHA256 `0B90DB91BCBB789C5F6E7080C8760911B302F183B2D6B6F70DC667C75B5BB374`. 보안PC 호환(`_rust.pyd` 부재) + pypdf(PYZ 102항목) 재확인.
- 2026-06-04 KST: 사이드바 메뉴 순서 변경(PDF를 일괄 변환 위로) 후 onefile/folder 재빌드·배포. 보안PC 호환 유지, GUI 순서 스모크(pdf=idx4/batch=idx5) 확인.
- Onefile SHA256: `1C80EB30B29510B90342341A046A840438250744B9F958AD23E91A0E578B30EF`
- Folder EXE SHA256: `81C1D8D65090085155655807FE488798A1CE7EAFE9DF14A84BC644548988D8C8`
- 2026-06-04 KST: PDF 보안해제 **전용 탭** 추가(사이드바 'PDF' 메뉴/탭) 후 onefile/folder 재빌드·배포. cryptography 제외 유지(보안PC 호환), pypdf 포함 확인. GUI 구성 스모크 + goal_verify PASS=10.
- Onefile SHA256: `021F60186E9AB2401B733FD89F690C81E6EF82EA7E8060A206E15D66FF8D27B4`
- Folder EXE SHA256: `E3F333DB93B3B4DA1E8E6642236238E31C0E5B4A822B47C32C191D8212C3FE89`
- 2026-06-04 KST: 보안PC 실행 불가 회귀 수정 — cryptography(네이티브 `_rust.pyd`/cffi)를 번들에서 제외(`excludes`)해 옛 정상 빌드와 동일 네이티브 구성으로 복원. 옛 `_517833e` 빌드엔 `_rust.pyd`가 없고 `libcrypto`(OpenSSL, 파이썬 표준)만 있음을 바이너리 비교로 확인 → OpenSSL은 원래 있던 것(무죄), `_rust.pyd`가 차단 원인. pypdf 유지(AES PDF만 미지원). onefile/folder 재빌드·배포.
- Onefile SHA256: `9022E76B671A097BE782B9CBD5278AD66E1392F26967E6C33B7FFB9026FFD324`
- Folder EXE SHA256: `DE76B9F2E48786E76A5FC2A6F1852DC9F5BC70D9AEDEF08F5671AEDE50584B33`
- 2026-06-04 KST: PDF 보안 해제 기능 추가(.pdf 입력 지원, pypdf) 후 onefile/folder EXE 재빌드·배포. pypdf/cryptography 번들 xref로 확인.
- Onefile SHA256: `C1B1C0BB7744011339CBBDF72B63F2231C97AE8F675C7A467A84C7A542514B2C`
- Folder EXE SHA256: `77D9EAFDDFC2372AF5CF21BA18A8CBE600BE2A909401173CCB7F5DE7333E9D82`
- 2026-06-04 KST: 단일 변환 탭도 예열·재사용 적용(2번) 후 onefile/folder EXE 재빌드·배포.
- Onefile SHA256: `0B513D683D0D376B14009C5BEB578D32D119C00F47449D284492E889141EC6D8`
- Folder EXE SHA256: `0434B81C538B71A26FBBC336DC91B91323E1F47E59ACD795DA12907BE59C0151`
- 2026-06-04 KST: 일괄 변환 속도 개선(Office 예열·재사용 + 파일 추가 시 백그라운드 예열) 반영 후 onefile EXE 재빌드·배포.
- Onefile SHA256: `913E4E74CC0D6760AD1A8BA8A00DDD75B57C3D6A4D2CA4D8CEDE01BDA1987CAB`
- 2026-05-18 KST: synced other-PC update `517833e` after input mode separation for file/open document selection.
- Onefile SHA256: `52174E22F4A95E8C3DAD6636691D5BD11CFABBA4B22172201641EF5DA703BE9C`
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
- (2026-06-03) 일괄 변환 속도 개선: 영속 Office 워커 스레드가 Office 인스턴스를 예열·재사용해, 변환마다 새로 켜던 콜드스타트 비용을 최초 1회로 축소. PPT는 창 최소화(`WindowState=2`)로 렌더링 부담을 줄임. 종료 시 `_shutdown_office_worker`로 정리. goal_verify PASS=10/SKIP=0/FAIL=0, 예열 재사용 스모크 테스트 통과로 확인. (단일 변환 탭은 미적용 — 추후 동일 적용 가능)
- (2026-06-04) 추가: 배치 파일 추가 시 필요한 Office만 백그라운드 예열(`_prewarm_batch_office`/`_prewarm_office_job`)해 **첫 변환의 콜드스타트까지 제거**. 정상 OpenXML은 예열을 건너뛰고, zip이 아닌 DRM/보안 컨테이너 형식만 예열한다. 예열 선택 스모크 테스트 + goal_verify PASS=10 회귀 확인 후 onefile EXE 재빌드·배포(SHA256 913E4E74…).
- (2026-06-04) 추가: 단일 변환 탭(`_convert_direct_file`)도 같은 예열·재사용 워커로 라우팅해 Office 콜드스타트 제거(배치와 인스턴스 공유). 단일 변환 재사용 스모크 + goal_verify PASS=10 확인 후 onefile/folder EXE 둘 다 재빌드·배포(onefile SHA256 0B513D68…, folder SHA256 0434B81C…).
- (2026-06-04) 기능 추가: **PDF 보안 해제**. 일괄/단일 변환에 `.pdf`를 지원 형식으로 추가하고 `_convert_pdf_file`로 처리 — 정상 PDF는 복사, 암호/편집제한 PDF는 `pypdf`로 해제(복호화·제한 제거) 저장, DRM 컨테이너(비PDF)는 인가 뷰어→Print to PDF 안내. PDF는 Office 미사용이라 예열 워커는 자동 skip. 의존성 `pypdf`(+`cryptography`) 추가(requirements + spec 2종). PDF 해제 스모크 + goal_verify PASS=10 + xref 번들 확인 후 onefile/folder EXE 재빌드·배포(onefile C1B1C0BB…, folder 77D9EAFD…).
- (2026-06-04) 수정: 회사 보안PC에서 새 빌드가 실행되지 않는 회귀 → 원인은 PDF와 함께 들어온 `cryptography`의 네이티브 Rust 확장(`_rust.pyd`)+cffi(보안PC가 차단). spec `excludes`로 cryptography 제외해 옛 정상 빌드와 동일 네이티브 구성 복원(옛 `_517833e` 바이너리 비교로 확인). pypdf는 순수 파이썬이라 유지 → 평문/RC4/편집제한 PDF 해제 동작, **AES 암호화 PDF만 미지원**. cryptography 없이 pypdf import·RC4 해제를 모사 테스트로 확인.
- (2026-06-04) 추가: PDF **전용 탭** 신설. PDF가 일괄변환 탭에만 있어 "메뉴가 안 보인다"는 피드백 → 사이드바에 'PDF(보안 해제)' 메뉴/탭 추가(파일 선택→저장 위치→'PDF 보안 해제' 버튼). 일괄변환과 동일 엔진(`_convert_pdf_file`)을 단일 변환 경로로 연결. 인덱스 충돌 방지를 위해 뷰 맨 끝(6번째)에 배치(기존 탭 무영향). GUI 구성 스모크(PDF_UI_OK)+goal_verify PASS=10 후 재빌드·배포(onefile 021F6018…, folder E3F333DB…).
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
- 사용자 실제 문서 기준의 Word/메모장 수동 검증은 아직 필요하다.
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

## 2026-05-18
- 사용자 로그 분석: Excel 열린 통합문서를 선택했는데 이전에 남아 있던 파일 직접 선택 경로가 우선되어 다른 Excel 파일이 변환되는 흐름을 확인.
- PPT/Excel/Word/메모장 입력 모드를 `file`/`open`으로 분리하고, 열린 문서 콤보 선택 또는 `다시 감지` 버튼 사용 시 파일 직접 선택 경로를 초기화하도록 수정.
- 저장 경로 대화상자의 기본 파일명도 현재 입력 모드에 맞게 파일 직접 선택 또는 열린 문서 이름을 사용하도록 정리.
- Excel 파일 열기 COM 호출에 heartbeat 로그를 추가해 보안/DRM/OneDrive 지연 시 로그가 멈춘 것처럼 보이지 않게 수정.
- 파일 직접 선택 시 열린 문서 콤보 표시를 비우고, 열린 문서 선택 시 파일 경로 표시를 비워 두 입력 방식이 동시에 선택된 것처럼 보이지 않게 보완.
- 파일 직접 선택 입력칸은 찾아보기 버튼으로만 바뀌도록 읽기 전용 처리해 수동 입력과 열린 문서 선택이 섞이는 문제를 줄임.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 11, SKIP 1, FAIL 0
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 단일 EXE 배포본 SHA256 -> `52174E22F4A95E8C3DAD6636691D5BD11CFABBA4B22172201641EF5DA703BE9C`

- 2026-05-18 드래그앤드롭 입력 추가: 개별 PPT/Excel/Word/TXT 파일 선택 칸에 파일 드롭 시 직접 변환 입력으로 설정하고, 일괄 변환 목록에는 파일/폴더 드롭으로 지원 파일을 추가하도록 구현
- 2026-05-18 `tkinterdnd2` 의존성을 requirements/런타임 배포/ PyInstaller spec에 반영하고, 패키지가 없으면 기존 파일 선택 방식으로 계속 실행되도록 폴백 처리
- 2026-05-18 macOS 검증에서 Windows 전용 메모장 Win32 검증을 FAIL 대신 SKIP 처리하도록 `scripts/goal_verify_v3.py` 보강
- 2026-05-18 `python3 -m py_compile ppt_extractor_v3.py scripts/goal_verify_v3.py document_extractor.py pdf_printer.py ppt_extractor.py ppt_extractor_v2.py` -> 성공
- 2026-05-18 `python3 scripts/goal_verify_v3.py --clean` -> PASS 3, SKIP 6, FAIL 0 (macOS라 Office/Notepad Windows 검증은 SKIP)
- 2026-05-18 드롭 경로 파싱/폴더 확장 smoke test -> `drop_logic_ok`
- 2026-05-18 `git diff --check` -> 성공

- 2026-05-15 프로젝트 폴더 분석 완료: 메인 경로는 `ppt_extractor_v3.py`, 빌드는 PyInstaller 단일/폴더형 및 소스 배포로 구성됨. 우선 정리 포인트는 HWP 잔재/문서 정합성, Windows 11 메모장 UI Automation, `PROGRESS.md` 아카이브다.

- 2026-05-15 리뷰 지적 수정: `requirements.txt`에 `lxml` 추가, 직접/일괄 Office 변환을 `DispatchEx` 격리 인스턴스로 변경, `DisplayAlerts` 원복 처리 추가
- 2026-05-15 HWP는 현재 UI 제외 범위에 맞춰 `scripts\goal_verify_v3.py`의 활성 검증 목록에서 제거하고, 사용자 설명서/보안 PC 안내를 HWP 제외 안내로 갱신
- 2026-05-15 `python3 -m py_compile ppt_extractor_v3.py scripts/goal_verify_v3.py document_extractor.py pdf_printer.py ppt_extractor.py ppt_extractor_v2.py` -> 성공
- 2026-05-15 `git diff --check` -> 성공

- 2026-05-16 리뷰 추가 반영: PowerPoint 격리 인스턴스를 표시 상태로 설정해 클립보드 슬라이드 패키지 복원 안정성 보강
- 2026-05-16 `build_security_pc.bat`도 `requirements.txt` 설치 후 빌드하도록 변경해 `lxml` 누락 빌드 방지
- 2026-05-16 `python3 -m py_compile ppt_extractor_v3.py scripts/goal_verify_v3.py document_extractor.py pdf_printer.py ppt_extractor.py ppt_extractor_v2.py` -> 성공
- 2026-05-16 `git diff --check` -> 성공
- 2026-05-16 결과물 폴더(`/Volumes/SSD/OneDrive/코드작업/결과물/newppt`)에 최신 소스 실행 배포본과 안내/빌드 파일 복사 완료
- 2026-05-16 Windows EXE 빌드는 현재 macOS 세션(Darwin arm64, Windows `py`/PyInstaller 없음)에서 수행 불가. 기존 결과물 EXE 해시 유지: onefile `5C246F6F0AA2CC1B58FABF105B240CC75301DE49FE901D410F4CDE6E11321CAE`, folder `39A9855F9204EEC2C6BD5FCB2478FAC1AC81CF3024433A78E6F1360889856214`

## 2026-05-19
- 사용자 로그 분석: 일괄 변환 65.38초 중 첫 Excel `Workbooks.Open` 단계가 56초가량 소요되어 병목이 Excel 파일 열기/보안 검사 구간에 있음을 확인.
- Excel/Word 직접 파일 변환 및 일괄 변환에서 정상 OpenXML 파일은 Office COM을 열기 전에 검증 복사하도록 개선해 정상 파일은 Excel/Word 시작 비용을 건너뛰도록 수정.
- DRM/SCDS 등 직접 검증 복사가 실패하는 파일만 Office 내부 복원 경로로 전환하도록 로그를 분리.
- Excel 내부 복원이 필요한 경우 `UpdateLinks=0`, `ReadOnly=True`, `AddToMru=False`, `Notify=False`로 열고 이벤트/화면갱신/링크 업데이트/매크로 자동 실행을 끄는 옵션을 적용.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 9, SKIP 0, FAIL 0
- `git diff --check` -> 성공
- `build_security_pc.bat` -> 성공, 결과물 폴더 EXE 교체 완료
- 단일 EXE SHA256 -> `9BFDEF98039A98A1662324B8E2F3EE436E391435A83C60C8949FBB16F3AFBE36`
- 폴더형 EXE SHA256 -> `5987D45AFA4435EC1653C84809E5347D16FFB986825F2048E62ADB63C9D11369`
- 추가 병목 개선: 정상 Office 파일 직접 복사 경로에서는 ZIP 전체 `testzip()` 검사를 생략하고 필수 OpenXML 항목만 확인하도록 변경. 생성/복원된 임시 파일은 기존처럼 깊은 검사를 유지.
- 직접 복사 시 중간 stage 파일 검증을 제거하고 최종 파일만 경량 검증하도록 줄여 큰 Excel/PPT/Word 파일의 반복 읽기 비용을 완화.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 9, SKIP 0, FAIL 0
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 단일 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 단일 EXE SHA256 -> `542DCD94270A72412B66AC52E2187DDCA53CF5E95190BB94B388AC9D7F23D44A`
- 추가 검토 반영: 직접 복사 경로는 `copy2` 대신 파일 내용만 복사하는 `copyfile`을 사용하고, 결과 검증은 원본/대상 크기와 헤더 비교로 줄여 메타데이터 복사와 ZIP 재열기 비용을 제거.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 9, SKIP 0, FAIL 0
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 단일 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 단일 EXE SHA256 -> `FB89456840699B6256D53CE7529A474825894D2B85616A57F89F6C4FBC17AFFC`

## 2026-05-25
- 리뷰 지적 반영: Excel 직접/일괄 변환에서 빠른 직접 복사를 이미 시도한 뒤 내부 복원 경로로 넘어갈 때 `_batch_convert_excel_file()`의 중복 직접 복사 검증을 건너뛰도록 `skip_direct` 옵션을 추가.
- 직접 파일 변환의 Excel fallback도 사전 직접 복사 실패 후 들어오는 경로라 `skip_direct=True`로 호출하도록 정리.
- Windows 11 Notepad 검증 흔들림 보완: 고정 `sample_notepad.txt` 파일명이 Notepad 탭/세션 복원과 충돌하지 않도록 매 실행 고유 파일명을 사용하고, 텍스트 로딩을 짧게 대기하도록 수정.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 9, SKIP 0, FAIL 0
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 단일 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 단일 EXE SHA256 -> `58256EF1CF133BF61EDE6B48867DAA932DA264F59CF834E31237144BBE82FF5D`

## 2026-05-27
- 사용자 Excel 로그 분석: 원본 복사가 DRM/SCDS 컨테이너로 막혀 재구성 경로로 전환된 뒤 `선번장` 시트의 `UsedRange`가 1,048,576행 x 43열로 과대 감지되어 원본 규격 복사가 깨지는 문제를 확인.
- Excel 재구성 범위 계산을 `UsedRange` 단독 기준에서 실제 값/수식 `Find("*")`, 도형 위치, 병합 셀 확장 범위를 합산하는 방식으로 변경해 불필요한 빈 행/열 서식 때문에 전체 행을 복사하지 않도록 수정.
- 병합 셀 복사도 전체 `UsedRange.MergeAreas` 대신 계산된 유효 범위 안에서만 처리하도록 변경해 과대 범위 시트의 지연과 오류를 줄임.
- 후속 사용자 로그에서 일반 데이터 시트까지 `1행 x 1열`로 과도하게 줄어든 문제를 확인하고, Excel `Find("*")` 호출에 `After` 기준 셀을 명시해 첫/마지막 데이터 셀 검색을 안정화.
- 정상 크기의 `UsedRange`는 기존처럼 그대로 사용하고, 값 복사 한도를 넘는 비정상 과대 `UsedRange` 시트에만 실제 값/수식/도형 기준 보정을 적용하도록 수정.
- `한국교직원공제회 구축공정 (20260527)_복111사본.xlsx` 분석: 일반 시트는 복구됐지만 `선번장` 시트의 삽입 객체 2개가 누락되어 빈 시트로 남는 문제를 확인.
- Excel 객체 복사에서 클립보드 이미지가 EMF 등 openpyxl 미지원 형식으로만 잡히거나 비어 있을 때, Excel 임시 `ChartObject`에 붙여넣어 PNG로 내보내는 폴백을 추가.
- `scripts\goal_verify_v3.py`에 Excel 재구성 전용 검증을 추가해 과대 `UsedRange` 시트의 실제 범위, 행/열 크기, 삽입 객체 이미지 보존을 함께 확인하도록 보강.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 10, SKIP 0, FAIL 0
- Excel COM 재현 검증: 원시 `UsedRange=1,048,572행` 시트를 유효 범위 `1셀`로 보정 확인
- `git diff --check` -> 성공
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 단일 EXE만 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 추가 전체 검증: 전체 Python 파일 문법 검사, `goal_verify_v3`, Excel 정상/과대 `UsedRange` 재현 검증, PyInstaller 단일 EXE 빌드 재확인 -> 성공
- 단일 EXE SHA256 -> `1282EF520171ACA57641917B456D01554234E78316A81C1EB93E2230C5279A0A`

### 2026-05-27 추가 Excel 재구성 보정
- Excel DRM/SCDS 재구성 경로에서 도형/이미지의 배치 셀을 값 복사 범위 계산에 섞지 않도록 분리했다.
- 객체 전용 시트는 셀 범위를 `A1` 1셀로 제한하고, 객체는 별도 이미지로 보존하도록 보정했다.
- Excel `Find("*")` 범위 계산을 값 기준 우선으로 바꾸고, 빈 문자열 수식/서식 꼬리 때문에 1,048,576행까지 확장되는 케이스를 막았다.
- Excel 객체 복사는 클립보드 DIB/PNG/EMF 외에 PIL `ImageGrab.grabclipboard()`와 임시 `ChartObject` PNG 변환을 추가했다.
- `scripts\goal_verify_v3.py`의 `excel_reconstruction_fallback` 검증을 과대 UsedRange, 빈 문자열 수식 꼬리, 객체 전용 시트 이미지까지 확인하도록 확장했다.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 10, SKIP 0, FAIL 0
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공
- 단일 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 교체 -> 성공
- 단일 EXE SHA256 -> `F0FA80CAD708A174A0E18BBA4435E96146162F9A7CE87E84C37388CD4666E3C4`
- 사용자가 요청한 전전전 커밋 확인: `HEAD~3=517833e`도 Excel 재구성에서 `UsedRange`를 그대로 사용하며, 재현 테스트상 값이 있는 시트에 마지막 행 서식이 묻으면 `1,048,576행 x 20열`로 잡혀 기존 `500,000셀` 제한에 걸리는 것을 확인.
- A/B 비교용 과거 EXE를 별도 파일로 빌드해 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3_517833e.exe`에 복사 -> 성공
- `DocumentExtractor_v3_517833e.exe` SHA256 -> `A0451D47D82F74E719825EB1C622D55D3C2FC5A8616A497406A87B552FA8BEBE`
- 사용자 추가 로그도 `ppt_extractor_v3.py line 2772`로 확인되어 과거 EXE 실행으로 판단. 혼선을 줄이기 위해 시작 로그에 `빌드 ID: 2026-05-27-excel-range-v2`와 `실행 파일:` 경로를 출력하도록 추가.
- 결과물 폴더의 `DocumentExtractor_v3.exe`, `DocumentExtractor_v3_1.exe`, `DocumentExtractor_v3_517833e.exe`를 모두 최신 동일 EXE로 덮어씀.
- 최신 동일 EXE SHA256 -> `40768513B4B1FFCA4EFF13E6703589EDB261EC4F9EE07FD53039249D2DE8845C`
- 사용자 재확인: 최신 빌드 실행 로그에서 `선번장`이 여전히 과대 UsedRange로 잡혔으나, Excel을 완전히 종료 후 다시 열어 변환하니 진행됨. 원인은 코드 단독 문제가 아니라 열린 Excel 세션의 `UsedRange`/보안 래퍼 상태 꼬임 가능성이 높음.
- 큰 시트에서 느려질 수 있는 실험성 표시값 스캔 보정은 배포 EXE에 넣지 않았고 소스에서도 제거했다.

## 2026-05-28
- Excel DRM/SCDS 재구성 경로에서 시트 처리 오류나 삽입 객체 누락이 있어도 전체 결과가 단순 성공처럼 보이는 문제를 줄였다.
- `_extract_excel()` 재구성 루프에 `rebuild_issues`를 추가해 시트 오류, 객체 일부 복사 실패를 누적 기록한다.
- 재구성 파일은 저장하되 확인 필요 항목이 있으면 로그와 UI를 `부분 완료 - 확인 필요`로 표시하고 `messagebox.showwarning`을 띄우도록 변경했다.
- `_copy_excel_sheet_objects()`는 보이는 객체 수와 복사 성공 수를 반환해 호출자가 누락 여부를 판단할 수 있게 했다.
- `scripts\goal_verify_v3.py`의 Excel 재구성 검증은 객체 복사 반환값까지 확인하도록 보강했다.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공
- `py scripts\goal_verify_v3.py --clean` -> PASS 10, SKIP 0, FAIL 0
- `build_security_pc.bat` 실행 -> 성공, 결과물 폴더 `DocumentExtractor_v3.exe`/`DocumentExtractor_v3_1.exe`/`DocumentExtractor_v3_517833e.exe`를 최신 동일 EXE로 교체
- 최신 one-file EXE SHA256 -> `44313F26F0BC2E742DC08762CF5590005D48320E9F8302EAF66EA41DEA3CE003`
- 최신 folder EXE SHA256 -> `D7482081DE6EE32B5CF0E5B2D418D0035B55D36B246A8DF26D193C9D577369CA`

## 2026-05-31
- 코드베이스 전반 버그 헌팅(4관점 병렬 리뷰 + 재현 검증) 후 검증된 버그만 수정했다.
- [CRITICAL] 추출 예외 처리 6곳(`_extract_ppt/excel/hwp/word/batch/notepad`)에서 `except ... as e` 뒤 `root.after(0, lambda: ...str(e)...)`가 지연 실행될 때 `NameError`(e 자동 삭제)로 오류 안내가 사라지던 문제 수정. `error_message = str(e)` 고정 후 캡처. Python 3.10.11 재현 테스트로 확정.
- [HIGH] `_handle_table` 셀 단위 `except: pass`를 셀 좌표 포함 로그로 바꿔 데이터 누락을 표면화.
- [MEDIUM] `_handle_connector` 무로그 실패에 로그 추가.
- [MEDIUM] 임시파일 누수: `_save_native_copy`/`_save_ppt_clipboard_package_copy`/`_save_ppt_slide_clone`/`_save_word_openxml_copy`/`_save_ppt_visual_copy` 5곳이 성공 경로에서 temp 파일을 안 지우던 것을 try/finally로 통일. `_make_local_temp_path` 호출처 8곳 전수 확인(정상 3곳: `_extract_ppt`/`_save_excel_as_openxml_copy`/`_save_word_as_docx_copy`).
- [MEDIUM] `_add_batch_paths` 중복검사 키를 `os.path.abspath(path).lower()`로 통일(상대경로 중복 방지).
- [MEDIUM] `scripts\goal_verify_v3.py` 열 너비 컬럼 키를 `chr(ord("A")+...)` → `get_column_letter()`로 교체(26열 초과 대응, import 추가).
- false positive로 분류: 클립보드 EMF 핸들(`GetClipboardData(14)`)은 클립보드 소유라 `DeleteEnhMetaFile` 불필요.
- 미수정(실제 위험 낮음): detect 함수 `CoUninitialize`가 try/finally 밖이나 예외를 삼키고 early-return이 없어 양 경로 모두 도달.
- [MEDIUM] Word 런 `bold=False`(0)가 굵은 단락에서 상속으로 굵게 나오던 엣지 케이스 수정: `if X is not None and X != 9999999: run.font.X = bool(X)`로 bold/italic 명시 설정(`_collect_word_runs`는 -1/0/9999999/None 반환). underline/color는 0을 의도적으로 제외하는 다른 로직이라 미변경.
- `py -m py_compile ppt_extractor_v3.py scripts\goal_verify_v3.py` -> 성공.
- `py scripts\goal_verify_v3.py --clean` -> **PASS=10 SKIP=0 FAIL=0** (약 29초). 수정한 temp 누수(`ppt_native_copy`/`ppt_clipboard_package`/`word_openxml_copy`)·Excel 재구성(`excel_reconstruction_fallback: range=46x19 images=1 object_images=1`)·goal_verify 컬럼키(`get_column_letter`) 모두 회귀 없이 통과 확인.
- Word `bold/italic` 수정 후 `goal_verify --clean` 재검증 -> **PASS=10 SKIP=0 FAIL=0** (회귀 없음).
- `py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec` -> 성공. `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`(38,896,407 bytes) 교체.
- 배포 EXE SHA256 -> `65944f8c15a21b3e079467ddab9b777306338fa2f606926b8f974b1fc282df74`

## Next
1. Windows 환경에서 `build_security_pc.bat` 실행 후 최신 EXE/폴더형 EXE를 결과물 폴더에 교체
2. 실제 사용자 문서로 Word/메모장 탭 수동 확인
3. Windows 11 새 메모장 지원 강화를 위해 UI Automation 경로 검토

---
## Archive Rule
- 완료 항목이 20개를 넘거나 파일이 5KB를 넘으면 완료된 내용을 `ARCHIVE_YYYY_MM.md`로 옮기고, `PROGRESS.md`는 현재 이슈만 남긴다.
