# 가상 공장 운영보고서 예제

“농Shim 제 1공장”은 공개 예제용 가상 설정이다. 모든 문구·수치·로고는 새로 만든 가상 자료다.

표지 2개 표는 template/data/cover.xlsx의 절대 경로를 공유한다. 본문 2개 표와 변수 6개는 각 data/ 폴더에서 읽는다. 섹션 3개와 재사용 스타일은 별도 YAML로 관리한다. 새 체크아웃에서는 저장소 루트에서 `uv run python -m scripts.split_cover`를 실행한다.

사용법과 전체 config 예시는 루트 README.md를 따른다.
