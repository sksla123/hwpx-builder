document.assets를 추가했습니다. assets와 resources는 동일한 리소스 형식을 사용하고, 실행 시 교체 범위만 구분합니다.

- assets.yaml에는 고정 표지·공통 이미지, resources.yaml에는 보고서별 변수·본문 표를 둡니다.
- --resources는 resources만 전체 교체하고 assets는 유지합니다.
- 두 목록의 같은 종류·이름이 겹치면 DUPLICATE_RESOURCE로 실패합니다.
- 각 file_path 파일의 상대 데이터 경로는 해당 파일 디렉터리 기준이며 상위/symlink 이탈을 금지합니다.
- 기존 assets 없는 선언도 지원합니다. 스키마는 1.0을 유지합니다.
- README의 컴포넌트·config·템플릿 실행 예제를 업데이트했습니다.

macOS Apple Silicon과 Windows x64용 실행 파일, Python 패키지를 제공합니다. 실행 파일 ZIP에는 바이너리와 라이선스만 포함합니다.
