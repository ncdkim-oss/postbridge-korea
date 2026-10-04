const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
        HeadingLevel, PageNumber, TabStopType, LevelFormat, VerticalAlign, PageBreak } = require('docx');
const fs = require('fs');

const W = 10206;
const FONT = "한컴산뜻돋음";
const DATE = "2026. 10. 4.";
const TITLE = "AI 보안 전문 에이전트 구축 기획서";

const thin = { style: BorderStyle.SINGLE, size: 1, color: "B0BEC5" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const noShade = { type: ShadingType.CLEAR, fill: "FFFFFF" };

function cellBorders(isFirst, isLast) {
  return { top: thin, bottom: thin, left: isFirst ? none : thin, right: isLast ? none : thin };
}
function hCell(text, width, isFirst, isLast) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(isFirst, isLast),
    shading: { fill: "525252", type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, border: noBorders, shading: noShade,
      children: [new TextRun({ text, bold: true, size: 18, color: "FFFFFF", font: FONT })] })]
  });
}
function dCell(text, width, align, isFirst, isLast, opts = {}) {
  const lines = Array.isArray(text) ? text : [text];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(isFirst, isLast),
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: lines.map(l => new Paragraph({ alignment: align, border: noBorders, shading: noShade,
      children: [new TextRun({ text: l, size: 18, color: "2C3E50", font: FONT, bold: !!opts.bold })] }))
  });
}
// headers: [], rows: [[...]], widths: [], aligns: []
function table(headers, rows, widths, aligns, opts = {}) {
  const n = headers.length;
  const al = aligns || headers.map((_, i) => i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT);
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => hCell(h, widths[i], i === 0, i === n - 1)) }),
      ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) =>
        dCell(c, widths[i], al[i], i === 0, i === n - 1, { bold: opts.boldFirst && i === 0 })) }))
    ]
  });
}
function H1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, border: noBorders, shading: noShade, children: [new TextRun({ text: t, font: FONT })] }); }
function H2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, border: noBorders, shading: noShade, children: [new TextRun({ text: t, font: FONT })] }); }
function H3(t) { return new Paragraph({ heading: HeadingLevel.HEADING_3, border: noBorders, shading: noShade, children: [new TextRun({ text: t, font: FONT })] }); }
function P(t, o = {}) {
  const runs = Array.isArray(t) ? t : [{ text: t }];
  return new Paragraph({ border: noBorders, shading: noShade, alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { after: o.after ?? 120, line: 300 },
    children: runs.map(r => new TextRun({ font: FONT, size: 20, color: "2C3E50", ...r })) });
}
function B(t, level = 0) {
  const runs = Array.isArray(t) ? t : [{ text: t }];
  return new Paragraph({ border: noBorders, shading: noShade, numbering: { reference: "bul", level },
    spacing: { after: 60, line: 300 },
    children: runs.map(r => new TextRun({ font: FONT, size: 20, color: "2C3E50", ...r })) });
}
function N(t, level = 0) {
  return new Paragraph({ border: noBorders, shading: noShade, numbering: { reference: "num", level },
    spacing: { after: 60, line: 300 },
    children: [new TextRun({ font: FONT, size: 20, color: "2C3E50", text: t })] });
}
function Note(t) {
  return new Paragraph({ border: noBorders, shading: noShade, spacing: { before: 60, after: 160 },
    children: [new TextRun({ font: FONT, size: 18, color: "6B7B8C", italics: true, text: t })] });
}
function Gap(n = 120) { return new Paragraph({ border: noBorders, shading: noShade, spacing: { after: n }, children: [] }); }
function Box(title, lines) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [
      new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders: cellBorders(true, true),
        shading: { fill: "525252", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({ border: noBorders, shading: noShade, children: [new TextRun({ text: title, bold: true, size: 20, color: "FFFFFF", font: FONT })] })] })] }),
      new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders: cellBorders(true, true),
        shading: { fill: "F4F6F8", type: ShadingType.CLEAR }, margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children: lines.map(l => new Paragraph({ border: noBorders, shading: noShade, spacing: { after: 80, line: 300 },
          children: [new TextRun({ text: l, size: 19, color: "2C3E50", font: FONT })] })) })] })
    ]
  });
}
const PB = () => new Paragraph({ children: [new PageBreak()] });
const C = AlignmentType.CENTER, L = AlignmentType.LEFT;

// ---------------- 표지 ----------------
const cover = [
  Gap(2400),
  new Paragraph({ alignment: C, border: noBorders, shading: noShade, spacing: { after: 200 },
    children: [new TextRun({ text: "기 획 서", size: 28, color: "6B7B8C", font: FONT, bold: true })] }),
  new Paragraph({ alignment: C, border: noBorders, shading: noShade, spacing: { after: 160 },
    children: [new TextRun({ text: "AI 보안 전문 에이전트", size: 56, bold: true, color: "2C3E50", font: FONT })] }),
  new Paragraph({ alignment: C, border: noBorders, shading: noShade, spacing: { after: 160 },
    children: [new TextRun({ text: "구축 기획서", size: 56, bold: true, color: "2C3E50", font: FONT })] }),
  Gap(200),
  new Paragraph({ alignment: C, border: noBorders, shading: noShade, spacing: { after: 120 },
    children: [new TextRun({ text: "지속 업데이트를 기본으로 하는 집·회사 PC 실시간 방어 체계", size: 26, color: "34495E", font: FONT })] }),
  new Paragraph({ alignment: C, border: noBorders, shading: noShade,
    children: [new TextRun({ text: "“창은 매일 날카로워진다. 방패도 매일 새로워져야 한다.”", size: 22, color: "6B7B8C", font: FONT, italics: true })] }),
  Gap(3200),
  table(["구분", "내용"], [
    ["문서명", TITLE],
    ["문서 버전", "v0.1 (초안) — 검토 후 v1.0 확정"],
    ["작성일", DATE],
    ["적용 대상", "① 집 PC (개인 소유)  ② 회사 PC (대학 정보보안 부서 승인 후)"],
    ["문서 성격", "구축 전 기획 단계 문서. 본 기획서 확정 후 단계별 개발에 착수"],
    ["개정 주기", "분기 1회 정기 개정 + 중대 위협 변화 시 수시 개정 (부록 B 개정 이력 관리)"],
  ], [2200, 8006]),
  PB(),
];

// ---------------- 목차 ----------------
const toc = [
  H1("목  차"),
  ...[
    "요약 (Executive Summary)",
    "1. 추진 배경 및 목적",
    "2. 보호 대상 및 위협 분석",
    "3. 에이전트 개념 설계",
    "4. 지속 업데이트 체계 (핵심)",
    "5. 기술 구성",
    "6. 추진 단계 및 일정",
    "7. 소요 비용",
    "8. 성과지표 (KPI)",
    "9. 리스크 및 대응",
    "10. 결정 필요 사항 및 다음 단계",
    "부록 A. 용어 설명",
    "부록 B. 기획서 개정 이력",
  ].map(t => P(t, { align: L, after: 80 })),
  PB(),
];

// ---------------- 요약 ----------------
const summary = [
  H1("요약 (Executive Summary)"),
  Box("한 눈에 보는 기획 요지", [
    "■ 무엇을: 집 PC와 회사 PC를 해킹·악성코드·피싱으로부터 실시간으로 지키는 「AI 보안 전문 에이전트」를 구축한다.",
    "■ 어떻게: 기존 보안 도구(Windows Defender, 방화벽, OS 업데이트)를 대체하지 않고, 그 위에서 ① 상태 감시 ② 위협 탐지 ③ AI 판단 ④ 자동 대응 ⑤ 지속 업데이트의 5개 층을 쌓는다.",
    "■ 핵심 원칙: 「지속 업데이트」를 선택 기능이 아닌 기본 설계로 둔다. 위협정보는 매시간, 탐지규칙은 매일, 에이전트 코드는 매주, AI 판단 로직은 매월 갱신하며 업데이트가 멈추면 에이전트 스스로 경보를 낸다.",
    "■ AI의 역할: 수집된 이벤트를 종합해 위험도를 판단하고, 사용자가 이해할 수 있는 말로 설명하며, 되돌릴 수 있는 범위에서만 자동 대응한다. 파괴적 조치는 반드시 사용자 승인을 거친다.",
    "■ 일정: 12주 4단계. 1단계(2주)에서 점검·리포트 에이전트를 먼저 가동해 효과를 확인하고 단계적으로 확장한다.",
    "■ 비용: 오픈소스와 무료 위협정보 피드를 기본으로 하며, 실비는 AI API 사용료 수준이다.",
    "■ 전제: 회사 PC 적용은 대학 정보보안 담당부서와 사전 협의·승인이 필수이며, 승인 전에는 집 PC에서만 운영한다.",
  ]),
  Gap(),
  P("본 기획서는 「해킹과 보안은 창과 방패」라는 전제에서 출발한다. 공격자는 이미 AI를 활용해 취약점 탐색, 피싱 메일 작성, 침투 자동화를 수행하고 있으며, 공격 도구는 매일 갱신된다. 따라서 방어 체계가 한 번 구축하고 끝나는 「완성품」이라면 수개월 내에 무력화된다. 이 기획서는 에이전트를 「계속 자라는 방패」로 정의하고, 업데이트·검증·자가진단의 순환 구조를 설계의 중심에 둔다."),
  PB(),
];

// ---------------- 1. 배경 ----------------
const s1 = [
  H1("1. 추진 배경 및 목적"),
  H2("1.1 추진 배경: AI가 바꾼 공격 환경"),
  P("최근 금융기관, 공공기관, 교육기관의 웹사이트와 내부 시스템을 겨냥한 침해사고가 잇따르고 있으며, 그 수법은 AI의 등장으로 질적으로 달라졌다. 과거에는 숙련된 공격자가 수 주에 걸쳐 수행하던 정찰·취약점 분석·침투·정보 탈취 과정이 AI 에이전트에 의해 자동화되어 소수 인력으로도 대규모·동시다발 공격이 가능해졌다. 개인 PC 사용자 입장에서 체감되는 변화는 다음과 같다."),
  B([{ text: "피싱의 정교화: ", bold: true }, { text: "AI가 작성한 한국어 피싱 메일·문자는 문법 오류가 없고 수신자의 소속·업무 맥락을 반영해 기존 「어색한 문장」 기준의 판별이 통하지 않는다." }]),
  B([{ text: "취약점 악용 속도: ", bold: true }, { text: "보안 패치가 공개된 뒤 공격 코드가 만들어지기까지의 시간이 며칠에서 수 시간으로 단축되어, 「나중에 업데이트」하는 습관이 곧 침해로 이어진다." }]),
  B([{ text: "악성코드의 변형: ", bold: true }, { text: "AI가 악성코드를 자동 변형해 서명(시그니처) 기반 백신을 우회하므로, 「무엇인가」보다 「무엇을 하는가」(행위)를 보는 탐지가 필요하다." }]),
  B([{ text: "개인 PC가 관문: ", bold: true }, { text: "대학 교직원의 집 PC와 회사 PC는 학사·행정 시스템, 이메일, 클라우드 문서로 연결되는 입구이며, 한 대의 감염이 조직 전체의 정보 유출로 이어질 수 있다." }]),
  H2("1.2 목적"),
  P("본 사업의 목적은 사용자가 보안 전문가가 아니어도, 전문가 수준의 감시·판단·대응을 24시간 받을 수 있는 「개인 전담 보안 에이전트」를 구축하는 것이다. 구체적 목표는 다음 네 가지다."),
  N("실시간 방어: 의심 프로세스·네트워크 연결·피싱 URL·악성 파일을 탐지 즉시 알리고, 되돌릴 수 있는 조치는 자동으로 수행한다."),
  N("정보 보호: 개인정보·업무문서·계정정보가 외부로 유출되는 경로(원격 접속, 비정상 전송, 정보 탈취형 악성코드)를 차단한다."),
  N("지속 업데이트: 위협정보·탐지규칙·에이전트 코드·AI 판단 로직을 정해진 주기로 자동 갱신하고, 갱신이 멈추면 스스로 경보한다."),
  N("설명 가능성: 모든 탐지와 조치를 사용자가 이해할 수 있는 한국어 보고로 남겨, 「왜 막았는지」를 알 수 있게 한다."),
  H2("1.3 기본 철학: 창과 방패, 그리고 지속 업데이트 5대 원칙"),
  P("공격(창)은 날마다 새로워지므로 방어(방패)도 날마다 새로워져야 한다. 이 철학을 설계 원칙으로 옮기면 다음 다섯 가지가 된다. 이후 모든 장의 설계는 이 원칙에 따라 검증한다."),
  table(["원칙", "내용", "설계 반영"], [
    ["① 업데이트가 기본값", "업데이트는 선택 기능이 아니라 에이전트의 생존 조건이다. 업데이트가 멈춘 에이전트는 「고장」으로 취급한다.", "4.1 업데이트 주기표, 4.4 노후화 지표"],
    ["② 다층 방어", "하나의 층이 뚫려도 다음 층이 막는다. AI는 기존 보안 도구를 대체하지 않고 그 위에서 조율한다.", "3.2 5층 아키텍처"],
    ["③ 행위 중심 탐지", "「알려진 악성코드인가」보다 「정상에서 벗어난 행동인가」를 본다. 기준선(baseline)을 학습해 이탈을 감지한다.", "3.3 핵심 기능, L2 탐지 엔진"],
    ["④ 되돌릴 수 있는 자동화", "자동 대응은 복구 가능한 범위로 제한하고, 파괴적 조치는 사람이 승인한다. AI의 오판이 사고가 되지 않게 한다.", "3.4 대응 자동화 단계"],
    ["⑤ 방패도 창으로 시험", "월 1회 안전한 모의 공격으로 방패의 구멍을 찾고, 결과를 다음 업데이트에 반영한다.", "4.5 자가 진단"],
  ], [2000, 5206, 3000], [C, L, L], { boldFirst: true }),
  PB(),
];

// ---------------- 2. 현황 ----------------
const s2 = [
  H1("2. 보호 대상 및 위협 분석"),
  H2("2.1 보호 대상"),
  P("보호 대상은 집 PC와 회사 PC 두 대이며, 소유권과 관리 권한이 달라 적용 방식도 달라야 한다. 특히 회사 PC는 대학 정보보안 정책(내PC지키미, 백신, 매체제어, 망분리 등)이 이미 적용되어 있을 가능성이 높으므로 반드시 정보보안 담당부서와 협의 후 적용한다."),
  table(["구분", "집 PC", "회사 PC"], [
    ["소유·관리 권한", "본인 (관리자 권한 보유)", "대학 (정보보안 정책 적용 중, 관리자 권한 제한 가능)"],
    ["주요 자산", "개인정보, 금융정보, 가족 사진·문서, 각종 계정", "학사·행정 문서, 평가 자료, 교직원·학생 개인정보, 내부 시스템 접속 권한"],
    ["주요 위협", "피싱, 랜섬웨어, 정보탈취형 악성코드, 원격제어, 가정용 공유기 취약점", "스피어피싱(업무 사칭), 문서형 악성코드(HWP·DOCX), 내부망 확산, USB 매체"],
    ["적용 범위", "전체 기능 (감시·탐지·AI 판단·자동 대응·업데이트)", "1차: 읽기 전용 감시·점검·리포트 → 승인 범위에 따라 확대"],
    ["AI 데이터 전송", "허용 (단, 해시·메타데이터 위주로 최소화)", "원칙적 금지. 로컬 모델 사용 또는 전송 없는 규칙 기반 운영"],
    ["적용 시점", "1단계부터 즉시", "4단계 (정보보안 담당부서 승인 후)"],
  ], [2000, 4103, 4103], [C, L, L], { boldFirst: true }),
  Note("※ 전제: 두 PC 모두 Windows 11 기준으로 설계한다. macOS 또는 Linux인 경우 센서 층(L1)의 수집 모듈만 교체하고 상위 설계는 그대로 적용한다."),
  H2("2.2 주요 위협 시나리오"),
  P("개인 PC에서 실제로 발생 빈도가 높은 위협을 공격 경로와 함께 정리하고, 각 위협을 어느 층에서 막을지 대응시킨다. 이 표는 3장의 기능 설계와 4장의 업데이트 대상을 정하는 근거가 된다."),
  table(["위협", "공격 경로", "피해", "주 대응 층"], [
    ["피싱·스미싱", "업무 사칭 메일, 택배·공공기관 사칭 문자, 가짜 로그인 페이지", "계정 탈취, 2차 침투", "L2 URL 평판 조회 + L3 AI 메일·URL 분석"],
    ["문서형 악성코드", "HWP·DOCX·PDF 첨부 파일의 매크로·취약점", "원격제어, 정보 유출", "L1 프로세스 트리 감시 + L2 행위 규칙(오피스→파워셸 실행 등)"],
    ["정보탈취형 악성코드", "불법 소프트웨어, 가짜 업데이트, 브라우저 확장", "저장된 비밀번호·쿠키·지갑 탈취", "L1 브라우저 데이터 접근 감시 + L2 해시 조회"],
    ["랜섬웨어", "RDP 노출, 피싱, 취약점", "파일 암호화, 업무 마비", "L1 대량 파일 변경 감시 + L4 즉시 격리 + 백업"],
    ["원격제어·백도어", "원격 지원 도구 악용, 취약한 공유기", "장기 잠복, 지속 유출", "L1 신규 자동실행·계정·서비스 감시 + L2 외부 연결 평판"],
    ["취약점 미패치", "OS·브라우저·오피스·한글 미업데이트", "모든 공격의 기반", "L0 패치 상태 점검 + L5 CVE 피드 대조"],
    ["AI 에이전트 자체 공격", "에이전트가 읽는 로그·메일에 숨긴 명령(프롬프트 인젝션), API 키 탈취", "방패가 창으로 전환", "3.6 에이전트 자체 보안"],
  ], [2000, 3000, 2200, 3006], [C, L, L, L], { boldFirst: true }),
  H2("2.3 기존 보안 수단의 한계"),
  P("Windows Defender와 상용 백신은 알려진 악성코드 차단에는 효과적이지만, 다음 영역에서 공백이 있다. 에이전트는 바로 이 공백을 메우는 것을 목표로 한다."),
  B("여러 신호를 종합하는 「판단」이 없다. 백신은 파일 하나를 보지만, 실제 침해는 「메일 열람 → 문서 실행 → 파워셸 → 외부 접속」처럼 연쇄로 나타난다."),
  B("사용자에게 「왜 위험한지」 설명하지 않는다. 경고창은 무시되거나 「허용」으로 눌러진다."),
  B("업데이트 상태, 백업 상태, 계정 보안(MFA), 공유기 설정 같은 「위생 상태」를 통합 점검하지 않는다."),
  B("새로운 위협 동향(KISA 보안공지, 신규 CVE)을 사용자의 환경에 대조해 「당신의 PC에 해당된다」고 알려주지 않는다."),
  PB(),
];

// ---------------- 3. 설계 ----------------
const s3 = [
  H1("3. 에이전트 개념 설계"),
  H2("3.1 역할 정의: 에이전트가 하는 일과 하지 않는 일"),
  table(["에이전트가 하는 일", "에이전트가 하지 않는 일"], [
    [["• 기존 보안 도구의 상태를 감시하고 꺼져 있으면 되살린다",
      "• 프로세스·네트워크·자동실행·로그를 수집해 기준선과 비교한다",
      "• 위협정보와 규칙으로 1차 탐지하고, AI가 종합 판단한다",
      "• 되돌릴 수 있는 범위에서 자동 차단·격리한다",
      "• 모든 탐지·조치를 한국어로 설명하고 일·주 보고서를 만든다",
      "• 스스로 업데이트하고, 업데이트가 멈추면 경보한다"],
     ["• Windows Defender·백신을 대체하지 않는다 (대체가 아니라 조율)",
      "• 파일 삭제, 계정 삭제, 디스크 포맷 등 파괴적 조치를 자동으로 하지 않는다",
      "• 회사 보안 정책을 우회하거나 무력화하지 않는다",
      "• 사용자의 문서·메일 전문을 외부 AI로 보내지 않는다 (해시·메타데이터 위주)",
      "• 「완전 방어」를 약속하지 않는다. 제로데이는 백업·복구 체계로 보완한다"]],
  ], [5103, 5103], [L, L]),
  H2("3.2 아키텍처: 5층 방어 구조"),
  P("에이전트는 아래 다섯 층으로 구성된다. 아래 층일수록 「검증된 기존 도구」, 위 층일수록 「AI와 자동화」이며, 가장 바깥의 L5 지속 업데이트 층이 나머지 네 층 전체를 매일 새롭게 한다."),
  table(["층", "명칭", "역할", "주요 구성요소"], [
    ["L5", "지속 업데이트", "위협정보·규칙·코드·AI 로직을 주기적으로 갱신하고 검증·롤백한다. 모든 층을 감싸는 바깥 층", "업데이트 스케줄러, 서명 검증, 카나리 배포, 노후화 지표"],
    ["L4", "대응", "탐지 결과에 따라 알림·차단·격리·복구 안내를 수행한다", "방화벽 규칙 적용, 프로세스 종료, 파일 격리함, 모바일 알림"],
    ["L3", "AI 판단", "여러 신호를 종합해 위험도를 산정하고, 근거를 설명하며, 대응 등급을 결정한다", "Claude API (클라우드) 또는 로컬 LLM, 판단 프롬프트, 사례 기억"],
    ["L2", "탐지 엔진", "규칙·위협정보·행위 기준선으로 이벤트를 1차 선별한다", "Sigma 규칙, YARA, 해시·URL·IP 평판 조회, 기준선 이상탐지"],
    ["L1", "센서·수집", "PC에서 일어나는 일을 구조화된 이벤트로 수집한다", "Sysmon, Windows 이벤트 로그, 프로세스·네트워크·자동실행 스냅샷, 파일 변경 감시"],
    ["L0", "기반 방어", "OS가 제공하는 기본 보안을 항상 켜진 상태로 유지한다", "Defender 실시간 보호, 방화벽, 자동 업데이트, BitLocker, 계정 MFA"],
  ], [800, 1800, 4000, 3606], [C, C, L, L], { boldFirst: true }),
  Gap(),
  P([{ text: "이벤트 흐름: ", bold: true }, { text: "L1 수집 → L2 규칙 선별(대부분의 정상 이벤트는 여기서 걸러짐) → L3 AI 판단(의심 이벤트만 전달, 1일 수십 건 수준) → L4 대응 등급 결정·실행 → 사용자 보고. L5는 이 흐름과 독립적으로 매시간·매일 돌며 L0~L4의 재료를 교체한다." }]),
  H2("3.3 핵심 기능"),
  table(["기능", "설명", "우선순위", "단계"], [
    ["보안 위생 점검", "Defender·방화벽·OS 패치·브라우저·오피스·한글 업데이트 상태, BitLocker, 계정 MFA, 백업 최신 여부를 매일 점검하고 미비 항목을 안내·자동 복구", "필수", "1"],
    ["일일·주간 보안 리포트", "점검 결과, 탐지 건수, 조치 내역, 주요 위협 뉴스 요약을 한국어로 작성해 아침에 전달", "필수", "1"],
    ["프로세스·네트워크 감시", "신규 프로세스, 부모-자식 관계(오피스→파워셸 등), 외부 연결 목적지를 실시간 수집해 기준선과 비교", "필수", "2"],
    ["자동실행·지속성 감시", "레지스트리 Run 키, 예약 작업, 서비스, 신규 계정, 브라우저 확장의 변경을 감지 (공격자의 「잠복」 수법)", "필수", "2"],
    ["위협정보 대조", "실행 파일 해시, 접속 URL·IP를 무료 위협정보 피드와 대조해 알려진 악성 여부 판정", "필수", "2"],
    ["피싱 분석 도우미", "사용자가 메일·문자·URL을 붙여넣으면 발신자·링크·요구 행동을 분석해 위험도와 근거를 제시", "필수", "3"],
    ["AI 종합 판단", "L2가 올린 의심 이벤트를 전후 맥락과 함께 분석해 위험도(정보/주의/경고/위급)와 대응 등급을 결정", "필수", "3"],
    ["자동 대응", "등급에 따라 알림, 외부 연결 차단, 프로세스 중지, 파일 격리를 수행하고 되돌리기 기능 제공", "필수", "3"],
    ["랜섬웨어 조기 감지", "짧은 시간 대량 파일 변경·확장자 변경을 감지하면 즉시 프로세스 중지 및 알림", "필수", "3"],
    ["자가 업데이트", "4장의 업데이트 파이프라인 전체 (피드·규칙·코드·AI 로직)", "필수", "1~4"],
    ["자가 진단(모의 공격)", "월 1회 안전한 테스트(EICAR, 피싱 샘플, 정상 범위 내 행위 테스트)로 탐지 성능 검증", "권장", "4"],
    ["보안 상담", "「이 프로그램 설치해도 되나요?」 같은 질문에 설치 전 평판·권한을 조사해 답변", "권장", "4"],
  ], [2000, 5406, 1200, 1600], [C, L, C, C], { boldFirst: true }),
  H2("3.4 대응 자동화 단계 (되돌릴 수 있는 자동화)"),
  P("AI의 판단이 틀릴 수 있다는 전제에서, 자동 대응의 범위를 「되돌릴 수 있는가」로 나눈다. 등급은 사용자가 설정에서 조정할 수 있으며, 회사 PC는 승인 전까지 Tier 0으로 고정한다."),
  table(["등급", "자동 조치 범위", "예시", "되돌리기"], [
    ["Tier 0 알림", "조치 없음. 알림과 권고만", "「새 브라우저 확장이 설치되었습니다. 설치한 적이 없다면 삭제를 권합니다」", "해당 없음"],
    ["Tier 1 자동 차단", "복구 가능한 차단·격리 (즉시 실행 후 사후 보고)", "악성 평판 IP로의 연결 차단, 의심 프로세스 중지, 의심 파일 격리함 이동", "원클릭 복원 (규칙 삭제, 파일 복원)"],
    ["Tier 2 승인 후 조치", "사용자 승인 후 실행 (알림에서 승인/거부)", "자동실행 항목 삭제, 예약 작업 제거, 계정 비활성화, 네트워크 전체 차단", "조치 전 스냅샷 보관"],
    ["금지", "에이전트가 절대 수행하지 않음", "파일 영구 삭제, 디스크 포맷, 백신·방화벽 해제, 회사 정책 변경", "—"],
  ], [1800, 2600, 3806, 2000], [C, L, L, L], { boldFirst: true }),
  H2("3.5 AI 활용 설계 및 데이터 보호"),
  P("AI(LLM)는 L3 판단과 피싱 분석, 보고서 작성, 보안 상담에 사용한다. 보안 에이전트가 오히려 정보 유출 통로가 되지 않도록 다음 원칙으로 데이터를 다룬다."),
  B([{ text: "최소 전송: ", bold: true }, { text: "AI에 보내는 것은 프로세스명·해시·접속 주소·이벤트 요약 등 메타데이터이며, 문서 내용·메일 본문·비밀번호는 보내지 않는다. 피싱 분석은 사용자가 직접 붙여넣은 내용에 한한다." }]),
  B([{ text: "PC별 모드 분리: ", bold: true }, { text: "집 PC는 클라우드 모델(Claude API) 사용, 회사 PC는 로컬 모델 또는 AI 없는 규칙 모드로 운영한다. 모드는 설정 파일에서 PC별로 고정한다." }]),
  B([{ text: "모델 선택: ", bold: true }, { text: "일상 판단은 비용 효율 모델, 위급 등급 재검토와 월간 심층 분석은 상위 모델을 사용하는 2단 구성으로 비용과 정확도를 함께 잡는다." }]),
  B([{ text: "판단 기록: ", bold: true }, { text: "AI의 모든 판단은 입력 요약·근거·결론을 로컬 DB에 남겨 사후 검토와 오탐 교정에 쓴다. 이 기록은 4장의 「AI 로직 업데이트」의 재료가 된다." }]),
  H2("3.6 에이전트 자체 보안 (방패가 창이 되지 않게)"),
  P("보안 에이전트는 관리자 권한으로 PC 전체를 보고 외부와 통신하므로, 공격자에게는 가장 매력적인 목표가 된다. 다음을 설계 요건으로 둔다."),
  table(["위험", "대책"], [
    ["프롬프트 인젝션 (로그·메일·파일명에 숨긴 명령으로 AI를 조종)", "에이전트가 읽는 모든 외부 데이터는 「데이터」로만 취급하고 명령으로 해석하지 않도록 프롬프트 구조를 분리한다. AI 출력은 허용된 조치 목록(화이트리스트) 안에서만 실행한다."],
    ["API 키·설정 탈취", "키는 Windows 자격 증명 관리자(DPAPI)에 보관하고 평문 설정 파일에 두지 않는다. 키별 사용 한도와 월 비용 상한을 설정한다."],
    ["업데이트 경로 위·변조 (가짜 업데이트 주입)", "업데이트 파일은 서명과 해시를 검증한 뒤에만 적용한다. 검증 실패 시 적용하지 않고 경보한다. (4.3 참조)"],
    ["에이전트 권한 남용", "수집은 관리자 권한, AI 판단은 일반 권한으로 분리한다. 자동 조치는 3.4의 Tier 범위를 코드 수준에서 강제한다."],
    ["에이전트 중단·무력화", "에이전트 프로세스가 종료되면 감시 서비스가 재시작하고, 10분 이상 중단 시 모바일로 알림을 보낸다."],
  ], [3600, 6606], [L, L]),
  PB(),
];

// ---------------- 4. 지속 업데이트 ----------------
const s4 = [
  H1("4. 지속 업데이트 체계 (핵심)"),
  P("이 장은 본 기획서의 중심이다. 에이전트의 모든 구성요소를 「업데이트 대상」으로 정의하고, 각각의 주기·출처·검증·실패 시 대응을 정한다. 설계 기준은 단순하다. 「업데이트되지 않은 방패는 방패가 아니다.」"),
  H2("4.1 업데이트 대상과 주기"),
  table(["계층", "업데이트 대상", "주기", "방식", "출처"], [
    ["L0 기반 방어", "Windows·Defender 정의 파일, 브라우저, 오피스, 한글, 주요 앱", "매일 점검, 발견 즉시 적용", "OS 자동 업데이트 + 에이전트가 상태 확인·독촉", "Microsoft, 각 제조사"],
    ["L2 위협정보", "악성 해시·URL·IP·도메인 목록", "매시간", "자동 (피드 다운로드·검증·적용)", "abuse.ch(URLhaus·MalwareBazaar·Feodo), KISA 보호나라, VirusTotal 조회"],
    ["L2 탐지 규칙", "Sigma 행위 규칙, YARA 서명, 자체 규칙", "매일", "자동 적용 + 신규 규칙은 24시간 「관찰 모드」 후 활성화", "SigmaHQ, YARA-Rules, 자체 작성"],
    ["L1·L4 에이전트 코드", "센서·대응·스케줄러 프로그램 본체", "매주 (보안 수정은 수시)", "자동 (서명 검증 → 집 PC 선적용 → 24시간 후 회사 PC)", "자체 저장소(GitHub Releases)"],
    ["L3 AI 판단 로직", "판단 프롬프트, 위험도 기준, 오탐 교정 사례, 사용 모델 버전", "매월 + 오탐 누적 시 수시", "반자동 (제안 생성 → 사용자 검토 → 적용)", "자체 판단 기록 DB, 모델 제공사 릴리스"],
    ["지식베이스", "신규 CVE, KISA·교육부 보안공지, 보안 뉴스 요약", "매일", "자동 수집 → AI가 「내 PC 해당 여부」 판정 → 리포트 반영", "NVD, KISA 보호나라, 교육부, 주요 보안 매체"],
    ["기획서·운영 규정", "본 기획서, 대응 등급 설정, 운영 루틴", "분기", "수동 (자가진단 결과·위협 변화 반영)", "운영 기록, 분기 검토 회의"],
  ], [1600, 2600, 1500, 2506, 2000], [C, L, C, L, L], { boldFirst: true }),
  H2("4.2 외부 정보원"),
  P("무료·공개 정보원을 기본으로 하고, 유료 서비스는 필요 시 추가한다. 모든 정보원은 「수집 → 형식 검증 → 적용 → 적용 결과 기록」의 공통 절차를 거치며, 정보원이 48시간 이상 응답하지 않으면 노후화 경보(4.4)에 반영된다."),
  table(["정보원", "제공 내용", "활용", "비용"], [
    ["KISA 보호나라(boho.or.kr)", "국내 보안공지, 취약점 경보, 악성코드 동향", "지식베이스, 국내 특화 위협 반영", "무료"],
    ["abuse.ch (URLhaus, MalwareBazaar, Feodo Tracker)", "악성 URL, 악성코드 해시, 봇넷 C2 IP", "L2 평판 대조 (매시간 갱신)", "무료"],
    ["VirusTotal API", "파일 해시·URL 다중 백신 판정", "의심 파일·URL 2차 확인 (일 한도 내)", "무료 등급 (한도 있음)"],
    ["NVD / CVE", "공식 취약점 데이터베이스", "설치 소프트웨어 버전과 대조해 미패치 알림", "무료"],
    ["SigmaHQ / YARA-Rules", "공개 탐지 규칙 저장소", "L2 행위·서명 규칙 갱신", "무료 (오픈소스)"],
    ["Microsoft 보안 업데이트 가이드", "월간 패치 내역, 긴급 패치", "L0 패치 점검 기준", "무료"],
    ["교육부·대학 정보보안 공지", "교육기관 대상 보안 지침·사고 사례", "회사 PC 운영 기준 반영", "무료"],
    ["모델 제공사 릴리스 노트", "AI 모델 버전·기능 변경", "L3 모델 교체 판단", "API 사용료에 포함"],
  ], [2800, 2800, 3006, 1600], [L, L, L, C]),
  H2("4.3 자가 업데이트 파이프라인"),
  P("에이전트 코드와 규칙의 업데이트는 다음 7단계를 자동으로 거친다. 어느 단계에서든 실패하면 이전 버전을 유지하고 경보한다. 이 파이프라인 자체가 공격 통로가 되지 않도록 서명 검증을 가장 앞에 둔다."),
  N("확인: 정해진 주기에 업데이트 저장소에서 최신 버전 정보를 가져온다."),
  N("검증: 다운로드 파일의 서명과 해시를 확인한다. 불일치 시 즉시 중단·경보."),
  N("격리 테스트: 새 규칙·코드를 「관찰 모드」로 실행해 기존 정상 이벤트에 오탐이 급증하지 않는지 확인한다 (규칙은 24시간, 코드는 자가 테스트 통과)."),
  N("카나리 적용: 집 PC에 먼저 적용하고 24시간 운영한다."),
  N("확대 적용: 이상이 없으면 회사 PC(승인 범위)에 적용한다."),
  N("건강 점검: 적용 후 센서 수집률, 탐지 지연, 오류 로그를 확인한다. 기준 미달 시 자동 롤백."),
  N("기록: 버전, 적용 시각, 결과를 업데이트 이력 DB에 남기고 주간 리포트에 포함한다."),
  H2("4.4 방패 노후화 지표와 경보"),
  P("「업데이트가 멈춘 상태」를 눈에 보이게 만드는 것이 지속 업데이트의 핵심이다. 다음 지표를 매일 계산해 대시보드와 리포트의 맨 위에 표시하고, 임계값을 넘으면 경보한다."),
  table(["지표", "정의", "정상", "경보 기준"], [
    ["위협정보 신선도", "마지막 피드 갱신 후 경과 시간", "6시간 이내", "24시간 초과"],
    ["규칙 신선도", "마지막 규칙 갱신 후 경과 일수", "3일 이내", "7일 초과"],
    ["코드 버전 차이", "최신 릴리스 대비 뒤처진 버전 수", "0", "2개 이상 또는 보안 수정 미적용"],
    ["패치 미적용 수", "설치 소프트웨어 중 알려진 취약점이 있는 미패치 항목 수", "0", "1개 이상 (긴급은 즉시)"],
    ["AI 로직 검토 경과", "마지막 판단 로직 검토 후 경과 일수", "30일 이내", "45일 초과"],
    ["업데이트 실패 횟수", "최근 7일 파이프라인 실패 건수", "0", "2회 이상 연속"],
    ["자가진단 통과율", "최근 월간 모의 공격 탐지율", "90% 이상", "80% 미만"],
  ], [2200, 4206, 1800, 2000], [C, L, C, C], { boldFirst: true }),
  H2("4.5 창으로 방패 시험하기: 월간 자가 진단"),
  P("방패가 실제로 막는지는 창으로 찔러 봐야 안다. 월 1회, 실제 피해가 없는 안전한 방법으로 탐지·대응을 시험하고, 놓친 항목은 다음 달 업데이트의 1순위로 반영한다. 모든 테스트는 집 PC에서만 수행하며, 회사 PC에서는 정보보안 담당부서가 허용한 항목만 수행한다."),
  table(["시험 항목", "방법 (안전)", "확인 사항"], [
    ["백신 연동", "EICAR 표준 테스트 파일 생성", "Defender 탐지 → 에이전트 인지 → 알림까지의 시간"],
    ["피싱 탐지", "알려진 피싱 샘플 URL(보고된 사례)로 분석 요청", "위험도 「경고」 이상 판정 여부, 근거 설명의 적절성"],
    ["행위 탐지", "오피스 문서에서 파워셸 실행 등 공격 패턴을 흉내 낸 무해한 스크립트 실행", "L2 규칙 발동, L3 판단, Tier 1 조치 여부"],
    ["지속성 탐지", "테스트용 자동실행 항목·예약 작업 추가", "변경 감지 및 알림 시간"],
    ["랜섬웨어 모사", "테스트 폴더 안에서 다수 파일을 빠르게 변경", "대량 변경 감지 → 프로세스 중지 시간"],
    ["업데이트 무결성", "해시가 틀린 가짜 업데이트 파일 투입", "검증 실패 → 적용 거부 → 경보 여부"],
    ["중단 복구", "에이전트 프로세스 강제 종료", "자동 재시작 시간, 알림 발송 여부"],
  ], [2000, 4206, 4000], [C, L, L], { boldFirst: true }),
  H2("4.6 운영 루틴"),
  table(["주기", "자동 (에이전트)", "수동 (사용자)", "소요 시간"], [
    ["매시간", "위협정보 피드 갱신, 센서 상태 점검", "—", "0분"],
    ["매일", "보안 위생 점검, 규칙 갱신, 지식베이스 수집, 일일 리포트 발송", "아침 리포트 확인 (경보 시 승인/거부 응답)", "3분"],
    ["매주", "에이전트 코드 업데이트, 주간 리포트(탐지 통계·오탐·업데이트 이력)", "주간 리포트 검토, 오탐 표시", "10분"],
    ["매월", "AI 판단 로직 개선안 생성, 자가 진단 실행", "개선안 승인, 자가 진단 결과 검토", "30분"],
    ["분기", "운영 통계 종합", "기획서·대응 등급·정보원 재검토, 위협 환경 변화 반영", "2시간"],
  ], [1200, 3806, 3600, 1600], [C, L, L, C], { boldFirst: true }),
  PB(),
];

// ---------------- 5. 기술 구성 ----------------
const s5 = [
  H1("5. 기술 구성"),
  H2("5.1 기술 스택"),
  P("개인이 유지보수할 수 있는 범위에서 검증된 오픈소스를 우선 선택한다. 모든 구성요소는 교체 가능하도록 모듈로 분리한다."),
  table(["구분", "선택", "선정 이유"], [
    ["운영체제", "Windows 11 (집·회사 공통 가정)", "대학 업무 환경 표준. 센서 모듈만 교체하면 타 OS 확장 가능"],
    ["개발 언어", "Python 3.12", "보안 라이브러리와 Windows API 연동이 풍부하고 유지보수가 쉬움"],
    ["센서", "Sysmon + Windows 이벤트 로그, psutil, WMI", "Microsoft 공식 도구로 프로세스·네트워크·파일 이벤트를 상세 수집"],
    ["탐지 규칙", "Sigma(pySigma), YARA", "공개 규칙 생태계가 가장 크고 매일 갱신됨"],
    ["AI", "Claude API (Claude Agent SDK) / 회사 PC는 로컬 모델 또는 규칙 모드", "도구 호출·장기 작업에 적합. 모델은 설정으로 교체 가능"],
    ["저장소", "SQLite (이벤트·판단·업데이트 이력)", "설치 불필요, 단일 파일, 백업 용이"],
    ["실행", "Windows 서비스 + 작업 스케줄러", "부팅 시 자동 시작, 중단 시 자동 재시작"],
    ["알림", "모바일 메신저 봇(텔레그램 등) + 이메일 + Windows 알림", "PC를 보고 있지 않을 때도 위급 알림 수신"],
    ["대시보드", "로컬 웹(FastAPI + 단일 HTML)", "외부 노출 없이 localhost에서만 열람"],
    ["업데이트 저장소", "비공개 GitHub 저장소 + Releases", "버전·서명·이력 관리를 표준 도구로 처리"],
    ["비밀 보관", "Windows 자격 증명 관리자(DPAPI)", "API 키를 평문 파일에 두지 않음"],
  ], [1800, 3600, 4806], [C, L, L], { boldFirst: true }),
  H2("5.2 모듈 구성"),
  table(["모듈", "책임", "대응 층"], [
    ["sensors/", "프로세스·네트워크·자동실행·파일·이벤트 로그 수집기", "L1"],
    ["baseline/", "정상 상태 기준선 학습·저장·비교", "L2"],
    ["rules/", "Sigma·YARA 규칙 로딩, 평가, 관찰 모드 관리", "L2"],
    ["intel/", "위협정보 피드 수집·검증·조회 (해시·URL·IP)", "L2, L5"],
    ["brain/", "AI 판단: 이벤트 종합, 위험도 산정, 근거 생성, 등급 결정", "L3"],
    ["responder/", "Tier별 조치 실행, 되돌리기, 조치 이력", "L4"],
    ["hygiene/", "보안 위생 점검 (패치·Defender·방화벽·백업·MFA)", "L0"],
    ["updater/", "4.3 파이프라인: 확인·검증·관찰·카나리·롤백·기록", "L5"],
    ["selftest/", "월간 자가 진단 시나리오 실행·채점", "L5"],
    ["report/", "일·주·월 리포트 생성, 노후화 지표 계산", "공통"],
    ["notify/", "메신저·이메일·Windows 알림 발송, 승인 응답 수신", "L4"],
    ["dashboard/", "로컬 웹 대시보드", "공통"],
  ], [2200, 6406, 1600], [L, L, C]),
  PB(),
];

// ---------------- 6. 일정 ----------------
const s6 = [
  H1("6. 추진 단계 및 일정"),
  P("총 12주, 4단계로 추진한다. 1단계에서 「점검·리포트」 에이전트를 먼저 가동해 매일 가치를 체감하면서, 이후 단계에서 탐지·판단·대응·업데이트를 쌓아 올린다. 각 단계는 완료 기준을 충족해야 다음 단계로 넘어간다."),
  table(["단계", "기간", "주요 과업", "산출물", "완료 기준"], [
    ["1단계 기반 구축", "1~2주", "• 개발 환경·저장소 구성  • hygiene 점검 모듈  • 일일 리포트  • updater 뼈대(코드 자가 업데이트)", "점검 에이전트 v0.1, 일일 리포트", "집 PC에서 매일 아침 리포트 자동 수신, 코드 자가 업데이트 1회 성공"],
    ["2단계 감시·탐지", "3~6주", "• Sysmon 설치·센서  • 기준선 학습(2주)  • Sigma·YARA 규칙 적용  • 위협정보 피드 연동  • 모바일 알림", "탐지 에이전트 v0.2", "자가 진단 항목 중 행위·지속성·백신 연동 탐지 성공, 오탐 일 5건 이하"],
    ["3단계 AI 판단·대응", "7~10주", "• brain 모듈(AI 판단)  • 피싱 분석 도우미  • responder Tier 0~2  • 랜섬웨어 조기 감지  • 되돌리기", "방어 에이전트 v0.3", "자가 진단 7개 항목 중 6개 이상 통과, Tier 1 조치 되돌리기 성공"],
    ["4단계 지속 업데이트 완성", "11~12주", "• 규칙 관찰 모드·카나리·롤백  • 노후화 지표·경보  • 월간 자가 진단 자동화  • 회사 PC 적용 협의·승인·Tier 0 적용", "에이전트 v1.0, 운영 매뉴얼", "4.4 지표 전 항목 「정상」, 가짜 업데이트 거부 테스트 통과, 회사 PC 승인 범위 적용"],
    ["운영", "13주~", "4.6 운영 루틴에 따른 상시 운영, 분기 기획서 개정", "주·월·분기 리포트", "노후화 경보 없이 분기 운영"],
  ], [1700, 1000, 3606, 1800, 2100], [C, C, L, L, L], { boldFirst: true }),
  Note("※ 회사 PC 적용(4단계)은 대학 정보보안 담당부서 승인 일정에 따라 변동될 수 있으며, 승인이 늦어져도 집 PC 운영은 계획대로 진행한다."),
  PB(),
];

// ---------------- 7. 비용 ----------------
const s7 = [
  H1("7. 소요 비용"),
  P("오픈소스와 무료 피드를 기본으로 하여 고정비는 거의 없고, 실비는 AI API 사용료가 대부분이다. 아래는 월 기준 추정이며, 실제 사용량에 따라 변동한다. 월 상한을 설정해 예기치 않은 과금을 막는다."),
  table(["항목", "내용", "월 추정 비용", "비고"], [
    ["AI API 사용료", "일상 판단(의심 이벤트 일 30~50건) + 일일 리포트 + 월간 심층 분석", "약 1~3만 원", "월 상한 설정. 회사 PC는 전송 없음"],
    ["위협정보 피드", "abuse.ch, KISA, NVD, SigmaHQ", "0원", "공개·무료"],
    ["VirusTotal", "무료 등급 (분당·일일 조회 한도)", "0원", "한도 초과 시 유료 전환 검토"],
    ["알림 채널", "메신저 봇, 이메일", "0원", "무료 범위"],
    ["저장소", "비공개 GitHub 저장소", "0원", "개인 계정 무료 범위"],
    ["백업(권장)", "외장 디스크 또는 클라우드 백업 (3-2-1 원칙)", "0~1만 원", "랜섬웨어 최종 방어선. 별도 결정"],
    ["합계", "", "약 1~4만 원/월", "개발 인건비는 본인 수행으로 미산정"],
  ], [2000, 4206, 1800, 2200], [C, L, C, L], { boldFirst: true }),
  Note("※ 비용 수치는 2026년 10월 기준 공개 요금을 바탕으로 한 추정치이며, 모델 요금과 사용량에 따라 달라진다. 1단계 운영 1개월 후 실측값으로 갱신한다."),
  H1("8. 성과지표 (KPI)"),
  P("「잘 막고 있는가」와 「계속 새로워지고 있는가」를 함께 측정한다. 지표는 주간 리포트에 자동 집계되며, 분기 검토 시 목표를 조정한다."),
  table(["영역", "지표", "목표", "측정 방법"], [
    ["탐지", "자가 진단 탐지율", "90% 이상", "월간 7개 항목 통과 수"],
    ["탐지", "탐지 지연 시간(이벤트 발생→알림)", "5분 이내", "Sysmon 시각 대비 알림 시각"],
    ["대응", "Tier 1 자동 조치 소요 시간", "1분 이내", "탐지→조치 로그"],
    ["정확도", "오탐율 (전체 경보 중 사용자가 오탐 표시한 비율)", "10% 이하", "사용자 피드백 집계"],
    ["업데이트", "위협정보 적시율 (24시간 이내 반영 비율)", "99% 이상", "피드 갱신 로그"],
    ["업데이트", "업데이트 파이프라인 성공률", "95% 이상", "updater 이력"],
    ["위생", "패치 적용률 (알려진 취약점 기준)", "100%", "hygiene 점검 결과"],
    ["위생", "백업 성공률", "100%", "백업 로그"],
    ["가용성", "에이전트 가동률", "99% 이상", "서비스 중단 시간"],
    ["사용자", "리포트 확인률 / 평균 대응 시간", "매일 / 1시간 이내", "알림 응답 로그"],
  ], [1400, 4006, 1800, 3000], [C, L, C, L], { boldFirst: true }),
  PB(),
];

// ---------------- 9. 리스크 ----------------
const s9 = [
  H1("9. 리스크 및 대응"),
  table(["리스크", "영향", "가능성", "대응"], [
    ["오탐으로 인한 업무 방해", "정상 프로그램 차단, 경보 피로로 알림 무시", "높음", "기준선 학습 2주 확보, 신규 규칙 24시간 관찰 모드, Tier 1 되돌리기, 오탐 피드백을 AI 로직 업데이트에 반영"],
    ["AI 오판 (위험을 놓치거나 과잉 대응)", "침해 미탐지 또는 불필요한 차단", "중간", "AI 판단 전 L2 규칙 선별, 자동 조치는 복구 가능 범위로 제한, 위급 등급은 상위 모델 재검토"],
    ["에이전트 자체 취약점", "에이전트가 공격 통로가 됨", "중간", "3.6 자체 보안 요건, 업데이트 서명 검증, 최소 권한, 월간 가짜 업데이트 테스트"],
    ["회사 보안정책과 충돌", "정책 위반, 징계 또는 설치 불가", "높음", "승인 전 설치 금지, 1차 읽기 전용, 정보보안 담당부서와 범위 서면 합의"],
    ["성능 부하", "PC 느려짐, 업무 지장", "중간", "Sysmon 필터 최적화, 수집 주기 조정, CPU 5% 상한 모니터링"],
    ["API 비용 초과", "예산 초과", "낮음", "월 상한 설정, L2 선별로 AI 호출 최소화, 사용량 주간 리포트"],
    ["제로데이·완전 방어 불가", "탐지 없이 침해", "중간", "백업 3-2-1 원칙 병행, 피해 최소화(격리·복구) 중심 설계, 「완전 방어」 약속 금지"],
    ["개인 유지보수 한계", "개발·운영 중단 시 방패 노후화", "중간", "노후화 지표로 가시화, 자동 업데이트로 수동 개입 최소화, 분기 검토로 지속 여부 판단"],
  ], [2200, 2400, 1000, 4606], [C, L, C, L], { boldFirst: true }),
  H1("10. 결정 필요 사항 및 다음 단계"),
  H2("10.1 착수 전 결정 사항"),
  table(["번호", "결정 사항", "선택지", "권고"], [
    ["1", "운영체제 확인", "Windows 11 / Windows 10 / macOS", "Windows 11 전제로 설계. 다를 경우 센서 모듈 조정"],
    ["2", "회사 PC 적용 범위", "미적용 / 읽기 전용 점검 / 전체 기능", "정보보안 담당부서 협의 후 「읽기 전용 점검」부터"],
    ["3", "AI 운영 방식", "클라우드 API / 로컬 모델 / 혼합", "집 PC는 클라우드 API, 회사 PC는 로컬 또는 규칙 모드"],
    ["4", "알림 채널", "텔레그램 / 이메일 / 카카오톡 / Windows 알림", "모바일 메신저 봇 + 이메일 병행"],
    ["5", "자동 대응 초기 등급", "Tier 0 / Tier 1", "집 PC는 2단계까지 Tier 0, 3단계부터 Tier 1"],
    ["6", "백업 체계", "외장 디스크 / 클라우드 / 병행", "병행 (3-2-1 원칙)"],
  ], [800, 2200, 3600, 3606], [C, L, L, L]),
  H2("10.2 다음 단계"),
  N("본 기획서 검토·확정 (v0.1 → v1.0): 10.1의 결정 사항 확정"),
  N("1단계 착수: 저장소 생성, 개발 환경 구성, 점검 모듈과 일일 리포트 개발 (2주)"),
  N("회사 PC 적용을 위한 정보보안 담당부서 협의 개시 (1단계와 병행, 서면 합의 목표)"),
  N("1단계 완료 후 1개월 운영 실측(비용·오탐·리포트 유용성)으로 기획서 1차 개정"),
  PB(),
];

// ---------------- 부록 ----------------
const appendix = [
  H1("부록 A. 용어 설명"),
  table(["용어", "설명"], [
    ["EDR", "Endpoint Detection and Response. PC(엔드포인트)에서 행위를 수집·분석해 위협을 탐지·대응하는 체계. 본 에이전트는 개인용 경량 EDR에 AI 판단을 더한 것"],
    ["Sysmon", "Microsoft가 제공하는 시스템 감시 도구. 프로세스 생성, 네트워크 연결, 파일 생성 등을 상세 로그로 남김"],
    ["Sigma 규칙", "로그 기반 탐지 규칙의 공개 표준 형식. 「오피스 프로그램이 파워셸을 실행」 같은 공격 패턴을 기술"],
    ["YARA", "파일 내용의 패턴으로 악성코드를 식별하는 규칙 언어"],
    ["IOC", "Indicator of Compromise. 침해 흔적 지표(악성 해시, URL, IP 등)"],
    ["기준선(Baseline)", "정상 상태에서의 프로세스·연결·자동실행 목록. 이탈을 감지하는 비교 기준"],
    ["C2", "Command and Control. 감염된 PC를 원격 조종하는 공격자 서버"],
    ["프롬프트 인젝션", "AI가 읽는 데이터 안에 명령을 숨겨 AI의 행동을 조종하는 공격"],
    ["카나리 배포", "변경을 일부 대상에 먼저 적용해 이상 여부를 확인한 뒤 전체로 확대하는 방식"],
    ["3-2-1 백업", "3개 사본, 2종류 매체, 1개는 외부(오프라인·클라우드) 보관 원칙"],
    ["MFA", "Multi-Factor Authentication. 비밀번호 외 추가 인증(앱·문자·키) 요구"],
    ["EICAR", "백신 동작 확인용 표준 무해 테스트 파일"],
  ], [2200, 8006], [C, L], { boldFirst: true }),
  H1("부록 B. 기획서 개정 이력"),
  P("본 기획서는 「지속 업데이트」 원칙에 따라 분기 1회 정기 개정하며, 중대한 위협 환경 변화나 자가 진단 결과에 따라 수시 개정한다. 개정 시 아래 표에 기록한다."),
  table(["버전", "일자", "개정 내용", "작성자"], [
    ["v0.1", DATE, "최초 작성 (초안)", ""],
    ["v1.0", "(예정)", "10.1 결정 사항 반영 후 확정", ""],
    ["v1.1", "(예정)", "1단계 1개월 운영 실측 반영 (비용·오탐·리포트)", ""],
    ["", "", "", ""],
  ], [1200, 1800, 5606, 1600], [C, C, L, C]),
];

const doc = new Document({
  creator: "기획조정처 경영전략팀",
  title: TITLE,
  styles: {
    default: { document: { run: { font: FONT, size: 20, color: "2C3E50" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: "2C3E50" },
        paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0, border: noBorders, shading: noShade } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: "34495E" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1, border: noBorders, shading: noShade } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: FONT, color: "34495E" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2, border: noBorders, shading: noShade } },
    ]
  },
  numbering: {
    config: [
      { reference: "bul", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 220 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 880, hanging: 220 } } } },
      ] },
      { reference: "num", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 300 } } } },
      ] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 567, bottom: 567, left: 850, right: 850, header: 567, footer: 567 } } },
    headers: { default: new Header({ children: [ new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "CFD8DC" } },
      tabStops: [{ type: TabStopType.RIGHT, position: W }],
      children: [ new TextRun({ text: TITLE, size: 20, color: "9E9E9E", font: FONT }),
                  new TextRun({ text: "\t" + DATE, size: 20, color: "9E9E9E", font: FONT }) ] }) ] }) },
    footers: { default: new Footer({ children: [ new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 1, color: "CFD8DC" } },
      tabStops: [{ type: TabStopType.RIGHT, position: W }],
      children: [ new TextRun({ text: "AI 보안 전문 에이전트 구축 기획서 v0.1 (초안)", size: 20, color: "9E9E9E", font: FONT }),
                  new TextRun({ text: "\t", size: 20 }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 20, color: "9E9E9E", font: FONT }) ] }) ] }) },
    children: [...cover, ...toc, ...summary, ...s1, ...s2, ...s3, ...s4, ...s5, ...s6, ...s7, ...s9, ...appendix]
  }]
});

const OUT = process.argv[2] || "20261004_기획서_AI보안전문에이전트구축.docx";
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log("written", OUT, buf.length); });
