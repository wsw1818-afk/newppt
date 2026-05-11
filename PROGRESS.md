# PROGRESS.md (현재 진행: 얇게 유지)

## Dashboard
- Progress: 100%
- Risk: 낮음
- Last updated: 2026-05-11

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

## Open Issues
- 사용자 실제 문서 기준의 HWP/Word/메모장 수동 검증은 아직 필요하다.
- Word 원본 파일 복사는 저장된 문서와 동일 확장자일 때만 안전 경로를 탄다. 저장 안 된 문서나 확장자 변환은 기존 재구성 경로를 사용한다.
- Excel 재구성 모드의 도형/이미지 위치는 TopLeftCell 기준 보존이며, 셀 내부 픽셀 오프셋은 근사치다.
- Windows 11 새 메모장은 현재 Win32 Edit/RichEdit 방식으로 제한적일 수 있어 UI Automation 검토 여지가 있다.
- `ppt_extractor_v3.py`에는 진단 없이 삼키는 `except:`/`except: pass`가 일부 남아 있다.

## Next
1. 실제 사용자 문서로 HWP/Word/메모장 탭 수동 확인
2. 멈춤 재현 시 바탕화면 `DocExtractor_Log_*.txt`의 처리 시간 로그로 병목 구간 확인
3. Windows 11 새 메모장 지원 강화를 위해 UI Automation 경로 검토

---
## Archive Rule
- 완료 항목이 20개를 넘거나 파일이 5KB를 넘으면 완료된 내용을 `ARCHIVE_YYYY_MM.md`로 옮기고, `PROGRESS.md`는 현재 이슈만 남긴다.
