# HWPX Builder

## Description

YAML로 문서 구조를 선언하고, 완성된 Excel 표·문서 변수·이미지를 조립해 HWPX를 만드는 템플릿 기반 빌더입니다.

표의 값·병합·치수·글꼴·서식은 XLSX에서 가져옵니다. 문서 변수는 별도의 2열 XLSX에서 읽습니다. 계산식 평가, 집계, 업무 규칙, 표 생성은 입력 자료를 준비하는 쪽에서 처리하며 빌더는 선언된 내용을 조립합니다.

저장소의 `template/`과 `examples/` 공개 자료는 **“농Shim 제 1공장”이라는 가상 설정으로 새로 만든 예제**입니다. 실제 기업의 보고서·실적·로고를 포함하지 않습니다. 스키마 버전은 `1.0`, 패키지 버전은 `pyproject.toml`에서 관리합니다.

## 설치

### macOS / Windows 실행 파일

[최신 릴리즈](https://github.com/sksla123/hwpx-builder/releases/latest)에서 OS에 맞는 ZIP을 받으세요. **Python이나 uv를 따로 설치할 필요가 없습니다.** 터미널에서 실행하는 CLI이며, ZIP에는 실행 파일과 라이선스만 들어 있습니다. 예제·템플릿은 [GitHub 저장소](https://github.com/sksla123/hwpx-builder/tree/main/template)에서 별도로 가져옵니다.

| OS / CPU | 다운로드 |
|---|---|
| macOS Apple Silicon (arm64) | [hwpx-builder-0.1.3-macos-arm64.zip](https://github.com/sksla123/hwpx-builder/releases/download/v0.1.3/hwpx-builder-0.1.3-macos-arm64.zip) |
| Windows x64 | [hwpx-builder-0.1.3-windows-x64.zip](https://github.com/sksla123/hwpx-builder/releases/download/v0.1.3/hwpx-builder-0.1.3-windows-x64.zip) |

macOS에서는 다운로드 폴더에서 다음을 실행합니다.

```sh
unzip hwpx-builder-0.1.3-macos-arm64.zip -d hwpx-builder
cd hwpx-builder
./hwpx-builder --version
./hwpx-builder --help
```

Windows PowerShell에서는 다음을 실행합니다.

```powershell
Expand-Archive hwpx-builder-0.1.3-windows-x64.zip -DestinationPath hwpx-builder
cd hwpx-builder
.\hwpx-builder.exe --version
.\hwpx-builder.exe --help
```

문서를 만들 때는 `build /path/to/my-report --output output/report.hwpx`처럼 document.yaml이 있는 폴더를 지정합니다. 저장소 템플릿을 복사하면 assets.yaml이 표지·로고를 함께 참조합니다. 실행 파일의 폴더를 PATH에 등록하면 다른 위치에서도 hwpx-builder 명령을 사용할 수 있습니다. 현재 macOS 배포본은 arm64 전용이며 Apple 개발자 서명/공증을 포함하지 않습니다. Windows 배포본도 코드 서명을 포함하지 않습니다.

### Python 패키지 / 소스 설치

Python **3.11 이상**이 필요합니다. 개발 및 예제 실행에는 [uv](https://docs.astral.sh/uv/getting-started/installation/)를 사용합니다. HWPX 생성에는 Excel이나 한컴 GUI 설치가 필요하지 않습니다. 결과를 화면에서 확인할 때는 호환 뷰어와 문서에 지정한 글꼴이 필요합니다.

```sh
git clone https://github.com/sksla123/hwpx-builder.git
cd hwpx-builder
uv sync --locked
```

템플릿의 고정 입력은 assets.yaml, 보고서별 입력은 resources.yaml에 분리되어 있습니다. 각 파일 내부 상대 경로는 해당 YAML 폴더 기준이므로 체크아웃 위치를 바꿔도 그대로 사용할 수 있습니다.

CLI만 설치하려면 저장소 루트에서 다음을 실행할 수 있습니다. 템플릿과 예제는 저장소에서 별도로 가져옵니다.

```sh
python -m pip install .
hwpx-builder --help
hwpx-builder --version
```

## 컴포넌트

| 컴포넌트 | 역할 / 현재 지원 |
|---|---|
| Document | `document.yaml` 진입점, 기본값·리소스·변수·섹션 순서 |
| Assets | 템플릿에 고정되는 표지 XLSX·공통 이미지. resources와 같은 형식이며 실행 인자로 교체하지 않음 |
| Resources | 보고서별 XLSX 표·변수·이미지 등록. `--resources`로 이 부분만 전체 교체 |
| Section | 논리 구역과 글자/들여쓰기 설정. 파일로 분리 가능 |
| PageSetup | 용지·방향·여백. 흐름 중 용지 변경은 미지원 |
| Layer | 섹션의 콘텐츠 흐름. 현재 `single_column`만 지원 |
| Paragraph / Run | 문단과 부분 서식 구간 |
| Table | 완성 XLSX 표 가져오기. YAML 행·셀 생성/서식 덮어쓰기 없음 |
| Note | `headernote`/`footernote`로 표·흐름 Container 앞뒤 설명 |
| Container | 콘텐츠 묶음과 Image의 직접 부모. 제한된 고정 그룹 지원 |
| Image | PNG/JPEG/BMP. Container 안에서 크기·정렬 선언 |
| PageBreak | 명시적 쪽 나눔 |
| Include | 별도 `type: blocks` 파일을 현재 흐름에 삽입 |
| Reference | 2열 변수 표의 평면 키를 `<ref:키>`로 치환 |
| Styles | 인라인 또는 `{src: ...}` 공통 스타일 파일 |
| SectionStyle | mm 단위 들여쓰기 단계표와 문단의 `indent_depth` |

TextStyle·ParagraphStyle·SectionStyle·ImageStyle·ContainerStyle을 사용할 수 있습니다. Row/Cell/TableStyle은 독립 YAML 입력으로 제공하지 않으며 Excel의 서식을 사용합니다. 페이지별 Header/Footer/PageNumber, Field, Repeat/Condition, 일반 자유 배치는 현재 미지원입니다.

## Config 작성 예시

다음은 최소 문서 구조입니다. 외부 선언 파일에도 `schema_version: '1.0'`과 해당 `type`을 씁니다.

```text
my-report/
  document.yaml
  assets.yaml
  resources.yaml
  data/
    variables.xlsx
    tables.xlsx
    logo.png
  styles/
    body.yaml
  sections/
    overview.yaml
```

```yaml
# my-report/document.yaml
schema_version: '1.0'
type: document
defaults:
  page:
    size: A4
    margin: {left: 20, right: 20, top: 20, bottom: 15}
  text_style: {src: styles/body.yaml}
  paragraph_style: {line_spacing: '160%', space_before: 2, space_after: 1}
  section_style: {indent_by_depth: [0, 5, 10, 15]}
assets: {file_path: assets.yaml}
resources: {file_path: resources.yaml}
variables: {resource: document_variables}
sections:
  - src: sections/overview.yaml
```

```yaml
# my-report/styles/body.yaml
schema_version: '1.0'
type: text_style
font_family: 휴먼명조
font_size: 11
```

```yaml
# my-report/resources.yaml
schema_version: '1.0'
type: resources
tables:
  document_variables:
    path: data/variables.xlsx
    sheet: variables
    header: true
  production:
    path: data/tables.xlsx
    sheet: production
```

```yaml
# my-report/assets.yaml — resources.yaml과 동일한 파일 형식
schema_version: '1.0'
type: resources
images:
  logo: {path: data/logo.png}
```

assets와 resources 모두 tables/images를 등록할 수 있습니다. 두 목록을 합쳐 같은 이름으로 참조하므로 콘텐츠에서는 어느 파일에 있는지 구별하지 않습니다. 같은 종류의 이름이 겹치면 `DUPLICATE_RESOURCE` 오류이며, 자동 덮어쓰기는 없습니다.

`variables.xlsx`의 `variables` 시트는 아래처럼 정확히 두 열로 작성합니다. 완성 표와 **다른 XLSX 파일**이어야 합니다. `header: true`이면 첫 행을 제목으로 제외합니다. 값은 빈칸도 허용하며 변수명은 비어 있거나 중복되면 안 됩니다.

| name | value |
|---|---|
| factory_name | 농Shim 제 1공장 |
| reporting_period | 2030.04.01.~2030.04.30. |

```yaml
# my-report/sections/overview.yaml
schema_version: '1.0'
type: section
layer:
  type: single_column
  contents:
    - type: paragraph
      text: '<bold><ref:factory_name></bold> 운영보고서'
    - type: paragraph
      indent_depth: 1
      text: '기간: <ref:reporting_period>'
    - type: container
      contents:
        - type: image
          resource: logo
          image_style: {width: 20, fit: contain}
    - type: table
      resource: production
      headernote:
        text: 생산 현황
        paragraph_style: {line_spacing: '100%', space_before: 0, space_after: 0}
      footernote:
        text: '※ 가상 예제 자료입니다.'
        paragraph_style: {line_spacing: '100%', space_before: 0, space_after: 0}
```

**YAML 안의 파일 경로는 기본적으로 `document.yaml`이 있는 루트 폴더를 기준으로 해석합니다.** 섹션 파일 안의 스타일 src도 섹션 폴더가 아닌 문서 루트 기준입니다. 예외로 `assets: {file_path: ...}`, `resources: {file_path: ...}` 또는 `--resources`로 리소스 YAML을 불러오면, **그 YAML 안의 데이터 path만 해당 리소스 YAML이 있는 폴더 기준**입니다. 기존 `{src: ...}` 방식의 resources는 문서 루트 기준을 유지합니다.

상대 데이터 path는 기준 폴더보다 상위로 벗어날 수 없습니다. `..` 또는 심볼릭 링크 이탈도 거부합니다. 폴더 밖 공유 파일은 절대 경로로 지정합니다. Section·스타일·Include src는 항상 문서 루트 내부로 제한됩니다. **CLI의 `PROJECT`, `--output`, `--resources`는 모두 현재 터미널 pwd 기준**입니다. 기본 출력도 pwd의 `output.hwpx`입니다. YAML의 `assets.file_path`와 `resources.file_path`는 계속 document.yaml 기준이며, 선택한 리소스 YAML 내부의 데이터 path는 그 파일 기준입니다.

길이·간격·좌표는 **mm 숫자**, 글자 크기는 **pt 숫자**, 줄간격은 **`'160%'` 문자열**입니다. `indent_depth: 2`는 단계표의 세 번째 값인 10mm를 선택합니다. 문단의 부모나 탭 개수를 해석하지 않습니다. 표 설명의 줄간격도 선언에서 변경할 수 있습니다. 표의 headernote/footernote는 실제 표의 좌우 경계 안으로 제한합니다. 표의 정렬·위치·바깥 여백을 반영하며, 지정한 문단 들여쓰기는 경계 안에 있는 경우 유지합니다. 들여쓰기와 여백 때문에 설명 폭이 사라지면 `TABLE_NOTE_NO_SPACE` 오류를 반환합니다.

## 텍스트 태그

Paragraph.text와 Run.text, 표 설명 문단에서 사용합니다. 태그는 대소문자를 구분하며 XML/HTML 문법 전체를 지원하지 않습니다.

| 태그 | 예시 | 효과 |
|---|---|---|
| ref | `<ref:factory_name>` | 변수 표의 정확한 평면 키 치환 |
| bold | `<bold>내용</bold>` | 굵게 |
| italic | `<italic>내용</italic>` | 기울임 |
| underline | `<underline>내용</underline>` | 밑줄 |
| strike | `<strike>내용</strike>` | 취소선 |
| highlight | `<highlight:yellow>내용</highlight>` | 글자 배경색 |
| color | `<color:#FF0000>내용</color>` | 글자색 |
| size | `<size:14>내용</size>` | 14pt 글자 크기 |
| linebreak | `첫 줄<linebreak>다음 줄` | 같은 문단 안의 명시적 줄바꿈 |
| nat | `<nat><ref:그대로><\nat>` | 내부를 태그/엔터티 해석 없이 출력 |

서식 태그는 중첩할 수 있고 정확한 순서로 닫아야 합니다. `ref`와 `linebreak`에는 종료 태그가 없습니다. `nat`의 종료는 `</nat>`가 아닌 **`<\nat>`**입니다. nat 밖에서는 `&lt;`, `&gt;`, `&amp;`만 기호로 변환합니다. 색은 black/white/red/green/blue/yellow/gray 또는 `#RRGGBB`를 사용합니다.

변수 치환값과 Excel 셀 문자열은 다시 태그로 해석하지 않습니다. 일반 문자열의 LF 대신 `<linebreak>`로 명시적 줄바꿈을 선언합니다.

## 실행 인자

```text
hwpx-builder build PROJECT [--output PATH] [--debug] [--warnings-as-errors] [--resources FILE_PATH]
hwpx-builder validate PROJECT [--resources FILE_PATH]
hwpx-builder --version
```

| 인자 | 명령 | 설명 |
|---|---|---|
| `PROJECT` | build / validate | document.yaml이 있는 디렉터리. 현재 작업 디렉터리 기준 |
| `--output PATH` | build | 출력 경로. 기본 `output.hwpx`, 상대 경로는 **현재 터미널 pwd 기준** |
| `--debug` | build | 모든 ref 치환 구간을 노란색으로 강조. 고정 문구·표 내부 제외 |
| `--warnings-as-errors` | build | 리소스 경고/진단이 있으면 출력 전에 실패 |
| `--resources FILE_PATH` | build / validate | Document.resources만 전체 교체, assets 유지. 파일 선택 경로는 **현재 터미널 pwd 기준** |
| `--version` | 최상위 명령 | 설치된 패키지의 현재 버전 출력. 예: `hwpx-builder 0.1.3` |
| `-h`, `--help` | 전체 / 각 명령 | 해당 도움말 출력 |

validate는 선언·리소스·일부 배치를 검사하며 파일 생성/재열기/화면 렌더는 하지 않습니다. build는 HWPX 생성, ZIP/XML/관계/격자 검사, 재열기를 마친 뒤 출력 파일을 원자적으로 교체합니다. 성공은 종료 코드 0과 stdout JSON, 실패는 종료 코드 1과 stderr 진단 JSON으로 반환합니다. YAML 오류는 실제 파일·줄·열·속성 경로와 참조 경로를 제공합니다.

## 실행 예제

저장소 루트에서 실행합니다.

```sh
uv run hwpx-builder validate template
uv run hwpx-builder build template --output output/factory-report.hwpx
uv run hwpx-builder build template --output output/factory-report-debug.hwpx --debug
uv run hwpx-builder build examples/variables --output output/variables.hwpx --debug
uv run pytest -n auto -q
```

테스트는 `pytest-xdist`의 `-n auto`를 기본 설정으로 사용해 모든 CPU에서 병렬 실행합니다. 로컬과 CI 모두 같은 설정을 사용합니다.

Python에서도 같은 빌더를 사용할 수 있습니다. **Python API의 상대 output/resources 인자는 document.yaml 디렉터리 기준**입니다. CLI는 터미널 기준 경로를 절대 경로로 변환해 API에 전달합니다.

```python
from hwpx_builder import build, load_project

model = load_project('template')
result = build('template', '../output/factory-report.hwpx', debug=True)
print(result.output, result.tables, result.images)
```

일반 XLSX 열 너비는 근사 변환이므로 경고가 발생할 수 있습니다. 캐시가 없는 수식, 지원하지 않는 Excel 서식, 본문보다 넓은 표는 오류를 반환합니다. 표는 자동 축소하지 않습니다. 물리 쪽 수와 전체 텍스트 넘침은 뷰어에서 별도로 확인합니다.

## 고정 assets와 보고서별 resources.yaml

템플릿의 섹션·스타일·표지·공통 로고는 공유하고, 보고서별 변수와 본문 표만 바꿉니다. `assets`와 `resources`는 **동일한 리소스 형식**이며 교체 범위만 다릅니다. 실제 업무 자료를 넣는 reports/는 Git에서 제외됩니다.

```text
template/
  document.yaml
  assets.yaml                 # 고정 표지·공통 로고
  resources.yaml              # 기본 보고서 입력
  sections/
  styles/
  data/cover.xlsx
  data/demo-logo.png
reports/2030-05/
  resources.yaml
  data/variables.xlsx
  data/tables.xlsx
```

```yaml
# template/document.yaml의 리소스 설정
assets: {file_path: assets.yaml}
resources: {file_path: resources.yaml}
```

```yaml
# template/assets.yaml
schema_version: '1.0'
type: resources
tables:
  table_01: {path: data/cover.xlsx, sheet: table_01}
  table_02: {path: data/cover.xlsx, sheet: table_02}
images:
  demo_logo: {path: data/demo-logo.png}
```

```yaml
# reports/2030-05/resources.yaml
schema_version: '1.0'
type: resources
tables:
  document_variables: {path: data/variables.xlsx, sheet: variables, header: true}
  table_03: {path: data/tables.xlsx, sheet: table_03}
  table_04: {path: data/tables.xlsx, sheet: table_04}
```

```sh
uv run hwpx-builder validate template --resources reports/2030-05/resources.yaml
uv run hwpx-builder build template --resources reports/2030-05/resources.yaml --output output/2030-05.hwpx
```

`--resources`는 **Document.resources 전체를 교체하며 Document.assets는 유지**합니다. 고정 표지나 로고를 보고서마다 다시 등록하지 않습니다. 교체 파일에는 필요한 보고서 입력 이름을 모두 등록해야 합니다. assets와 같은 종류·이름을 등록하면 오류입니다.

이 예의 assets 안 `data/...`는 template 기준, 보고서 resources 안 `data/...`는 reports/2030-05 기준입니다. 상대 데이터는 각 기준 폴더 밖으로 벗어날 수 없으며 외부 공유 파일은 절대 경로로 지정합니다. CLI의 --resources/--output은 터미널 pwd 기준, YAML의 섹션·스타일·file_path는 document.yaml 기준입니다. Python API의 상대 resources/output은 기존 문서 루트 기준을 유지합니다.

바로 실행 가능한 동일한 패턴의 공개 예제도 제공합니다.

```sh
uv run hwpx-builder build template --resources examples/factory_report/resources.yaml --output output/override-example.hwpx
```

```python
from hwpx_builder import build

build(
    'template',
    '../output/override-example.hwpx',
    resources='../examples/factory_report/resources.yaml',
    debug=True,
)
```

## 문제 제보

오류나 지원이 필요한 기능은 [GitHub Issues](https://github.com/sksla123/hwpx-builder/issues)에 남겨 주세요. 실행 명령, 패키지/Python/OS 버전, 진단 JSON, 재현 가능한 **가상 데이터**를 함께 주시면 확인하기 쉽습니다. 실제 회사 자료, 개인정보, 비공개 Excel·문서·로고는 올리지 마세요.

## License

[MIT License](LICENSE). 의존 라이브러리에는 각 라이브러리의 라이선스가 적용됩니다.
