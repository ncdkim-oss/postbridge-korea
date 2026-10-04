# AI 보안 전문 에이전트 — 인수인계 문서 (2026-10-04)

클라우드 세션에서 작성한 산출물을 집 PC 터미널로 옮기기 위한 문서다.
집 PC에서 Claude Code를 열고 이 폴더에서 작업을 이어간다.

## 1. 현재 상태

| 단계 | 상태 |
|---|---|
| 기획서 v0.1 (초안) 작성 | 완료 — `docs/20261004_기획서_AI보안전문에이전트구축.hwp` |
| 기획서 v1.0 확정 | 대기 — 아래 「미결 사항」 6개 결정 후 개정 |
| 1단계 개발 (기반 구축, 2주) | 미착수 |

## 2. 파일 구성

```
ai-security-agent/
├── CLAUDE.md                  # 이 프로젝트의 작업 규칙 (Claude Code가 자동으로 읽음)
├── README.md                  # 이 문서
├── docs/
│   ├── 20261004_기획서_AI보안전문에이전트구축.hwp    # 기본 산출물 (한글)
│   ├── 20261004_기획서_AI보안전문에이전트구축.hwpx   # 같은 내용의 OWPML 형식 (프로그램 편집용)
│   └── 20261004_기획서_AI보안전문에이전트구축.docx   # 같은 내용의 Word 형식 (참고용)
└── tools/
    ├── build_hwp.py           # 기획서 .hwp/.hwpx 생성 스크립트 (python-hwpx)
    └── build_docx.js          # 기획서 .docx 생성 스크립트 (docx npm)
```

## 3. 집 PC 환경 준비

```powershell
# 저장소 받기 (이미 있으면 pull)
git fetch origin claude/ai-security-agent-handoff
git checkout claude/ai-security-agent-handoff
cd ai-security-agent

# .hwp 생성 환경 (Python 3.11 이상)
pip install python-hwpx

# .docx 생성 환경 (선택, Node 18 이상)
npm install docx
```

## 4. 문서 재생성

```powershell
cd ai-security-agent/tools
python build_hwp.py ../docs/20261004_기획서_AI보안전문에이전트구축     # .hwp 와 .hwpx 를 함께 만든다
node build_docx.js ../docs/20261004_기획서_AI보안전문에이전트구축.docx  # 선택
```

내용을 고칠 때는 스크립트 안의 본문 텍스트(표·문단 데이터)를 수정한 뒤 다시 실행한다.
기획서를 개정하면 파일명 날짜를 바꾸고 부록 B 개정 이력표에 한 줄을 추가한다.

## 5. 미결 사항 (기획서 10.1) — 집 PC에서 가장 먼저 결정

| 번호 | 결정 사항 | 권고 |
|---|---|---|
| 1 | 운영체제 확인 | Windows 11 전제. 다르면 센서 모듈 조정 |
| 2 | 회사 PC 적용 범위 | 정보보안 담당부서 협의 후 「읽기 전용 점검」부터 |
| 3 | AI 운영 방식 | 집 PC는 클라우드 API, 회사 PC는 로컬 모델 또는 규칙 모드 |
| 4 | 알림 채널 | 모바일 메신저 봇 + 이메일 병행 |
| 5 | 자동 대응 초기 등급 | 2단계까지 Tier 0, 3단계부터 Tier 1 |
| 6 | 백업 체계 | 외장 디스크 + 클라우드 병행 (3-2-1 원칙) |

## 6. 다음 작업 순서

1. 미결 사항 6개 결정 → 기획서 v1.0 개정 (`build_hwp.py` 수정 후 재생성)
2. 에이전트 코드용 **별도 비공개 저장소** 생성 (이 저장소는 웹사이트용이라 분리 권장)
3. 1단계 착수: 보안 위생 점검 모듈(`hygiene/`), 일일 리포트(`report/`), 코드 자가 업데이트 뼈대(`updater/`)
4. 집 PC 실측 1개월 후 기획서 v1.1 (비용·오탐·리포트 유용성 반영)

## 7. 집 PC에서 Claude Code 시작 프롬프트 (예시)

```
ai-security-agent 폴더의 README.md와 CLAUDE.md, docs의 기획서를 읽고 현재 상태를 요약해 줘.
그 다음 10.1 미결 사항을 하나씩 물어보고, 결정이 끝나면 기획서를 v1.0으로 개정해서 .hwp로 만들어 줘.
```

## 8. 참고: 클라우드 세션에서 확인된 기술 메모

- `python-hwpx` 6.7.0: `HwpxDocument.new()` → `add_paragraph` / `add_table` → `save_to_path("x.hwp")` 로 HWP 5.0 바이너리를 직접 쓴다. 저장 후 `conversion_report` 가 `None` 이면 손실 없음.
- 표 테두리 면 이름은 대문자(`TOP`, `BOTTOM`, `LEFT`, `RIGHT`). 셀 세로 정렬 기본값은 `CENTER`.
- 글자 크기는 pt 단위로 넘기면 내부적으로 100배(10pt → 1000)로 저장된다.
- `apply_paragraph_format` 은 앞 문단의 문단 모양을 상속하므로 쪽 나눔(`page_break_before`)·들여쓰기는 모든 문단에 명시적으로 지정한다 (`build_hwp.py` 의 `FMT_DEFAULTS` 참고).
- 한글 프로그램이 없는 환경에서는 `hwpx.tools.layout_preview.render_layout_preview` 로 HTML 미리보기를 만들어 확인할 수 있다 (쪽 나눔은 표시되지 않는 근사치).
