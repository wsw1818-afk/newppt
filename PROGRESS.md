# PROGRESS.md (현재 진행: 얇게 유지)

## Dashboard
- Progress: 100%
- Risk: 낮음
- Last updated: 2026-05-12

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
- PPT 원본 복사가 DRM으로 실패하면 요청 파일 외에 `_화면그대로.pptx` 추가본을 생성한다. 이 파일은 편집성보다 서식/크기/배치 보존을 우선한다.
- PPT 슬라이드 복제/화면 그대로/재구성 저장은 OneDrive 대상 경로에 직접 저장하지 않고 로컬 `%TEMP%`에 정상 PPTX를 만든 뒤 최종 경로로 복사한다.
- PPT `Slide.Export`까지 DRM/보안 정책으로 막히면 `slide.Copy()` 클립보드 PNG/DIB 이미지를 받아 화면 그대로 PPTX를 생성한다.
- 최신 `DocumentExtractor_v3.exe`는 결과물 폴더에 복사되었고, 빌드본과 SHA256 해시가 같다.

## Verification
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

## Open Issues
- 사용자 실제 문서 기준의 HWP/Word/메모장 수동 검증은 아직 필요하다.
- Word 원본 파일 복사는 저장된 문서와 동일 확장자일 때만 안전 경로를 탄다. 저장 안 된 문서나 확장자 변환은 기존 재구성 경로를 사용한다.
- Excel 재구성 모드의 도형/이미지 위치는 TopLeftCell 기준 보존이며, 셀 내부 픽셀 오프셋은 근사치다.
- PPT `_화면그대로.pptx` 추가본은 시각 보존용이라 각 슬라이드가 이미지로 들어간다. 편집이 필요하면 기본 산출물, 화면 일치가 필요하면 `_화면그대로` 산출물을 사용한다.
- PPT 원본 복사와 PowerPoint 슬라이드 복제 저장이 모두 DRM으로 실패하면 화면 그대로 이미지 PPTX를 우선 생성한다. 단, 이 결과물은 시각 보존용이라 슬라이드 안의 개별 텍스트/도형 편집성은 낮다.
- PowerPoint 경고창이 계속 뜨면 기존 실행 중인 EXE/PowerPoint를 닫고 최신 EXE로 재실행해야 한다.
- `SCDSA004`처럼 이미 DRM 컨테이너로 저장된 파일은 그 파일만으로 정상 PPTX 복구가 불가능하다. 원본 문서를 PowerPoint에 연 상태에서 새 EXE로 다시 추출해야 한다.
- Windows 11 새 메모장은 자동 검증에서 SKIP됨. 현재 Win32 Edit/RichEdit 방식으로 제한적이라 UI Automation 검토 여지가 있다.
- `ppt_extractor_v3.py`에는 진단 없이 삼키는 `except:`/`except: pass`가 일부 남아 있다.

## Next
1. Windows 11 새 메모장 지원 강화를 위해 UI Automation 경로 검토
2. 실제 사용자 문서로 HWP/Word/메모장 탭 수동 확인
3. 멈춤 재현 시 바탕화면 `DocExtractor_Log_*.txt`의 `SaveCopyAs 진행 중` 로그와 처리 시간 로그로 병목 구간 확인

---
## Archive Rule
- 완료 항목이 20개를 넘거나 파일이 5KB를 넘으면 완료된 내용을 `ARCHIVE_YYYY_MM.md`로 옮기고, `PROGRESS.md`는 현재 이슈만 남긴다.
