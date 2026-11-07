document.assets와 리소스 네임스페이스를 추가했습니다. assets/resources는 동일한 파일 형식을 사용하지만 참조 범위는 격리됩니다.

- assets.yaml에는 고정 표지·공통 이미지, resources.yaml에는 보고서별 변수·본문 표를 둡니다.
- 콘텐츠는 assets.이름 또는 resources.이름으로 참조합니다. 양쪽에 같은 이름을 사용할 수 있습니다.
- 기존 접두사 없는 리소스 참조는 resources만 조회하며 assets로 자동 탐색하지 않습니다.
- --resources는 resources만 전체 교체하고 assets는 유지합니다.
- 각 file_path 파일의 상대 데이터 경로는 해당 파일 디렉터리 기준이며 상위/symlink 이탈을 금지합니다.
- README의 컴포넌트·config·템플릿 실행 예제와 로컬 개발 명세를 업데이트했습니다. 스키마는 1.0입니다.

macOS Apple Silicon과 Windows x64용 실행 파일, Python 패키지를 제공합니다. 실행 파일 ZIP에는 바이너리와 라이선스만 포함합니다.
