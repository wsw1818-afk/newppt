DocumentExtractor v3 - 회사 보안 PC 실행 안내

1. EXE가 회사 보안에 차단되면 이 폴더의 run_from_source.bat를 실행하세요.
2. Python 패키지 오류가 나오면 install_requirements.bat를 먼저 실행하세요.
3. Python이 없다는 메시지가 나오면 Python 3 설치가 필요합니다.
4. 회사망에서 pip 설치가 막히면 IT 보안팀에 아래 패키지 설치 또는 내부 저장소 사용을 요청해야 합니다.
   - pywin32
   - python-pptx
   - openpyxl
   - python-docx
   - Pillow
   - lxml
   - tkinterdnd2

주의:
- 이 방식은 PyInstaller EXE를 사용하지 않으므로 EXE 차단을 우회하는 데 도움이 됩니다.
- 그래도 Python 실행 자체가 차단되면 IT 보안 예외/허용 정책이 필요합니다.
- Word/PPT/Excel 문서는 먼저 해당 프로그램에서 열어 둔 상태로 사용할 수 있고, 파일 선택 칸 또는 일괄 변환 목록에 파일을 드래그앤드롭할 수 있습니다. HWP는 회사 DRM 환경에서 일반 파일 변환이 불가해 제외되었습니다.
