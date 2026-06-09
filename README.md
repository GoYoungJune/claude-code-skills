# Claude Code Skills

Claude Code에서 `/skill-name` 으로 바로 호출할 수 있는 커스텀 스킬 모음입니다.

## 스킬 목록

### `/excel-report` — CSV → Excel 변환

CSV 파일을 서식이 적용된 Excel 파일로 변환합니다.

- 헤더: 파란 배경 + 흰 굵은 글씨
- 숫자 열: 천 단위 콤마 포맷
- 합계 행 자동 추가
- 바 차트 자동 삽입

```bash
python skills/excel-report/scripts/excel_report.py data.csv
# → data.xlsx 생성
```

의존성: `pip install openpyxl`

---

### `/markdown-converter` — 파일 → Markdown 변환

PDF, Word, PowerPoint, Excel, HTML, CSV, 이미지 등 다양한 형식을 Markdown으로 변환합니다.  
`uvx markitdown` 을 사용하므로 별도 설치 불필요.

지원 형식: PDF, `.docx`, `.pptx`, `.xlsx`, HTML, CSV, JSON, XML, 이미지(EXIF/OCR), 오디오(전사), ZIP, YouTube URL, ePub

```bash
uvx markitdown input.pdf              # stdout 출력
uvx markitdown input.pdf -o out.md    # 파일로 저장
```

---

### `/gog` — Google Workspace CLI

`gog` CLI를 통해 Gmail, Calendar, Drive, Contacts, Sheets, Docs를 터미널에서 제어합니다.

```bash
gog gmail search 'newer_than:7d' --max 10
gog calendar events <calendarId> --from <iso> --to <iso>
gog drive search "query" --max 10
gog sheets get <sheetId> "Sheet1!A1:D10" --json
```

사전 설정: `gog auth credentials /path/to/client_secret.json`

## 설치 방법

```bash
# Claude Code 프로젝트 루트에서
cp -r skills/ ~/.claude/skills/
```

또는 이 저장소를 `~/.claude/skills/` 경로에 클론합니다.
