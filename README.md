# revision-comment-mcp

도면 리비전(개정판) 변경사항 비교 + 발주처 Comment Sheet 답변 초안 자동화를 하나의
흐름으로 연결한 MCP 서버입니다.

- 도면이 개정되면 이전 리비전과 새 리비전 PDF를 비교해 어디가 바뀌었는지 자동으로 찾아주고
- 발주처가 과거에 준 코멘트와 비슷한 코멘트가 다시 왔을 때, 과거 회신 이력에서 유사 사례를
  찾아준 뒤, 최신 개정 내용을 반영한 답변 초안(Word)을 만듭니다.

## 설계 원칙

이 코드(서버)는 "정확한 계산·검색·파일 처리"까지만 담당합니다.

- 어디가 실질적으로 바뀐 것인지(단순 렌더링 오차 vs 의미 있는 변경) 최종 판단
- 어떤 과거 사례가 지금 코멘트와 진짜 비슷한지 최종 판단
- 사내 어투에 맞춰 완결된 답변 문장으로 다듬는 것

은 이 tool을 호출하는 AI(Claude 등)가 담당하고, 최종 방향 지시와 검수는 담당자가 합니다.

## 제공 tool

1. `diff_drawing_revisions` — 두 PDF를 페이지 단위로 비교해 격자 단위로 변경 의심 영역을
   빨간 박스로 표시한 이미지를 반환 (의미 해석은 AI가 이미지를 보고 판단)
2. `parse_comment_sheet` — 발주처 Comment Sheet(.docx)의 표를 구조화된 행으로 파싱
3. `search_similar_comments` — `history/` 폴더의 과거 Comment Sheet들에서 유사 코멘트·과거
   회신을 문자열 유사도로 검색 (표면적 유사도이므로 최종 판단은 AI가 다시 함)
4. `create_comment_sheet_draft` — 답변 초안을 담은 새 Comment Sheet(.docx) 생성

## 설치 및 실행

```bash
uv venv .venv --python 3.12
uv pip install --python .venv/bin/python mcp pymupdf pillow numpy python-docx
.venv/bin/python revision_comment_mcp_server.py
```

Claude Desktop 설정(`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "revision-comment-tools": {
      "command": "/절대/경로/.venv/bin/python",
      "args": ["/절대/경로/revision_comment_mcp_server.py"]
    }
  }
}
```

## 주의 / 한계

- `COMMENT_SHEET_COLUMN_KEYWORDS`, 표 구조 가정은 예시입니다. 실제 발주처 Comment Sheet
  양식(컬럼 구성)에 맞춰 확인/수정이 필요합니다.
- 유사도 검색은 `difflib` 기반 문자열 비교(MVP 수준)라, 의미상 유사한지는 검색 결과를
  AI가 다시 읽고 재판단하는 걸 전제로 합니다.
- `history/`, `output/` 폴더는 실제 회사 문서(도면·코멘트시트)가 들어갈 자리라 `.gitignore`
  처리했습니다 — 폴더 구조만 저장소에 남아있고 내용물은 로컬에서 각자 채워야 합니다.
- 리비전 비교 threshold(기본 0.012)는 "벡터 기반 PDF끼리, 래스터화 잡음이 거의 없는" 조건을
  가정한 값입니다. 스캔본처럼 페이지마다 미세하게 흔들리는 PDF는 threshold를 0.03~0.05로
  올려야 오탐이 줄어듭니다.
