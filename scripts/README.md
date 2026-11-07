# 저장소 준비 도구

scripts는 공개 빌더 패키지에 포함하지 않는다. 원본 자료나 실제 회사 데이터는 공개 저장소에 두지 않는다.

- extract_tables.py / extract_images.py / reference_styles.py: 일반 HWPX 자료에서 표·이미지·서식 메모를 추출하는 도구.
- extract_reference.py: 위 추출 CLI. 상대 경로는 실행 디렉터리 기준이다.
- repair_xlsx.py: 기본 네임스페이스와 공백 전용 rich text의 xml:space 보존을 복구한다.
- split_cover.py: 가상 공개 예제의 표지 분리 및 공유 표지 절대 경로 갱신. 이미 분리된 Excel은 다시 작성하지 않는다.

```sh
uv run python -m scripts.split_cover
uv run python -m scripts.extract_reference --help
```

실제 자료는 reports/ 등 Git에서 제외된 로컬 폴더에 두고 별도 백업한다. 문서 빌드 시 추출/복구/이관 도구를 자동 실행하지 않는다.

- build_standalone.py: 해당 OS에서 독립 실행 파일을 빌드하고, 실행 파일·라이선스만 묶는다. 검증용 예제는 ZIP 외부의 임시 폴더에 만들고 버전/검증/문서 생성을 검사한다. `uv run --group release python -m scripts.build_standalone`으로 실행한다.
