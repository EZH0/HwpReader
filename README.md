# hwpreader

같은 양식의 HWP 문서 안 표를 읽어서 Excel로 누적 저장하는 도구입니다.

## 현재 구현 범위

- `data` 폴더의 `YYYYMMDD.hwp` 파일을 순서대로 읽습니다.
- 설정된 표 번호(`resource/config.json`의 `table_index`)의 내용을 가져옵니다. 기본값은 `all`이라 문서 안의 모든 표를 읽습니다.
- HWP 표의 열을 하나의 공사 레코드로 보고, 엑셀의 행 구조로 변환합니다.
- 줄바꿈은 설정값(`newline_replacement`)으로 정리합니다.
- 전체 데이터는 `result/full data`, 주소가 지정 지역명으로 시작하는 결과는 `result/sorted data`에 저장합니다.

## 준비

Windows에서 한글이 설치되어 있어야 실제 `.hwp` 자동 추출이 가능합니다.

```powershell
py -m uv sync
```

기본값(`--backend auto`)은 한글 COM 자동화(`pywin32`)를 먼저 사용하고, 실패하면 `pyhwpx`를 시도합니다.

## 실행

```powershell
py -m uv run python -m hwpreader
```

옵션:

```powershell
py -m uv run python -m hwpreader --config resource/config.json
py -m uv run python -m hwpreader --dry-run
py -m uv run python -m hwpreader --backend pyhwpx
py -m uv run python -m hwpreader --backend com
```

HWP 표 추출이 실패하거나 표가 아닌 값이 추출되면 엑셀 파일을 만들지 않고 종료합니다.

## 설정

`resource/config.json`

- `table_index`: HWP 문서 안에서 가져올 표 번호입니다. 0부터 시작하며, `all`이면 모든 표를 순서대로 읽습니다.
- `input_orientation`: 현재는 `columns` 기본값입니다. 표의 각 열이 공사 1건이라는 뜻입니다.
- `fields`: 엑셀 행으로 저장할 필드 순서입니다.
- `address_field`: 주소 필드명입니다.
- `address_prefixes`: `sorted data`로 따로 저장할 주소 시작 문자열입니다.
