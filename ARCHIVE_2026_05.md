# ARCHIVE_2026_05.md

## 2026-05-11 문서 추출기 v3 개선 기록

### 완료된 변경
- 프로젝트 상태와 `AGENTS.md`/`MEMORY.md`/`PROGRESS.md` 기준 작업 흐름을 확인했다.
- PPT/Excel 원본 그대로 복사 모드를 기본값으로 설정해 도형, 이미지, 차트, 서식, 원본 크기 보존을 우선하도록 했다.
- Excel 재구성 경로에 값/수식 범위 일괄 읽기, 행 높이, 열 너비, 병합 셀, 숫자 서식, 삽입 이미지/도형/차트 복사를 추가했다.
- PPT 하이브리드/텍스트 중심 모드에서 미지원 도형을 이미지 스냅샷으로 보존하도록 했다.
- Word/메모장 탭을 추가하고 Word 텍스트 재구성 경로와 빠른 복사 옵션을 분리했다.
- Word 빠른 복사를 `SaveAs2`에서 파일 시스템 복사로 바꿔 원본 문서의 저장 위치와 상태가 바뀌지 않도록 했다.
- HWP/HWPX 저장을 `FileSaveAs_S` 액션 우선으로 바꾸고, 저장 결과와 실제 파일 크기를 검증하도록 했다.
- HWP 감지 단계에서 `Dispatch`로 빈 한글 문서가 생성되는 병목을 제거했다.
- HWP 추출 연결 시 새 빈 문서가 만들어지는 경우를 차단하는 안전장치를 추가했다.
- `AUTOSHAPE_MAPPING` 중복 키를 제거했다.
- 시작 로그 문구를 현재 지원 범위에 맞춰 `PPT, Excel, 한글, Word, 메모장`으로 정리했다.
- `DocumentExtractor_v3.spec`, `build_ppt_v3.bat`, `.gitignore`를 v3 빌드에 맞게 정리했다.

### 주요 검증
- Python 컴파일 검증 성공.
- PyInstaller 6.17.0 기준 `DocumentExtractor_v3.spec` 빌드 성공.
- PPT COM 샘플 변환 검증 성공 (`shapes=2`, `pictures=1`, `autoshapes=1`, `media=1`).
- Excel COM 샘플 변환 검증 성공 (`images=2`, `media=2`, `drawings=1`).
- PPT `SaveCopyAs` 원본 복사 검증 성공.
- Excel `SaveCopyAs` 원본 복사 검증 성공.
- HWP COM 연결과 `FileSaveAs_S` 저장 검증 성공.
- HWP/HWPX 추출 경로 검증 성공.
- HWP 무생성 감지 검증 성공 (`before=0`, `after=0`, 새 `Hwp.exe` 없음).
- HWP 안전 추출 검증 성공 (`hwp_safe_extract_output.hwp`, 13KB).
- Word 안전 복사 검증 성공.
- `AUTOSHAPE_MAPPING` AST 중복 키 검사 성공 (`duplicate_keys: []`).
- 최신 EXE를 `D:\OneDrive\코드작업\결과물\newppt\DocumentExtractor_v3.exe`로 복사했고, 빌드본과 SHA256 해시가 일치했다.

### 배포 해시
- SHA256: `C665F94300DA25625AA8C297A42BDAA29E3C11516B62805168F23A90473862B4`
