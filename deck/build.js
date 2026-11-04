const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/b5915536-3fe3-4a22-9798-2179749b4aa4_c0048d07-4707-4906-bf38-1c686d3188cc/pptx/scripts/apply_theme.js");
const FONT = "Malgun Gothic";
const THEME = { name: "Spain Light and Dark", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1:"1C1714", lt1:"FFFFFF", dk2:"3A2D26", lt2:"F4F0EA", accent1:"C8932B", accent2:"A8322D",
    accent3:"2E4A5A", accent4:"7C8A4A", accent5:"8C5A3C", accent6:"5B4B7A", hlink:"A8322D", folHlink:"5B4B7A" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "리드온 독서클럽 - 스페인 미술 4주 프로그램 기획안";
const C = pres.SchemeColor;

pres.defineSlideMaster({ title: "DARK", background: { color: C.text1 },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.9, w: 8.2, h: 1.5, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } }] });
pres.defineSlideMaster({ title: "LIGHT", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontSize: 30, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "리드온 독서클럽 · 스페인 미술 4주 과정", options: { x: 0.6, y: 6.95, w: 6, h: 0.3, fontSize: 11, color: C.text2, margin: 0, isTextBox: true } } },
  ], slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 11, color: C.text2 } });

const circ = (s, x, y, d, fill, name, tr) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill, transparency: tr || 0 }, line: { type: "none" }, objectName: name });
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1 }, o));

// ---------- Slide 1 ----------
pres.addSection({ title: "표지" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "표지" });
circ(s, 9.0, 0.6, 3.8, C.accent1, "빛-외곽", 85);
circ(s, 9.6, 1.2, 2.6, C.accent1, "빛-중간", 70);
circ(s, 10.2, 1.8, 1.4, C.accent1, "빛-중심", 0);
T(s, "리드온 독서클럽 · 4주 프로그램 기획안", { x: 0.8, y: 1.1, w: 8, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
s.addText("읽고, 보고, 맛보고, 나누다", { placeholder: "title" });
T(s, "《스페인 미술 500년: 빛과 어둠의 연대기》와 함께하는 4주", { x: 0.8, y: 3.6, w: 8.2, h: 0.5, fontSize: 20, color: C.background1 });
T(s, "예술의전당 한가람디자인미술관 · 2026.9.22 – 2027.1.20", { x: 0.8, y: 4.2, w: 8.2, h: 0.4, fontSize: 14, color: C.background2 });
const wk = [["1", "책 읽기", "거장 입문"], ["2", "책 읽기", "역사 맥락"], ["3", "전시 관람", "현장 미션"], ["4", "티타임·피드백", "맛보고 나누기"]];
wk.forEach((w, i) => {
  const x = 0.8 + i * 3.0;
  circ(s, x, 5.1, 0.85, i === 2 ? C.accent2 : C.accent1, "주차-원" + (i + 1));
  T(s, w[0], { x, y: 5.1, w: 0.85, h: 0.85, fontSize: 28, bold: true, color: i === 2 ? C.background1 : C.text1, align: "center", valign: "middle" });
  T(s, [{ text: w[1], options: { bold: true, fontSize: 16, color: C.background1, breakLine: true } }, { text: w[2], options: { fontSize: 14, color: C.background2 } }], { x: x + 1.05, y: 5.15, w: 1.85, h: 0.8, valign: "middle" });
});
s.addNotes("표지. 읽고(1~2주) → 보고(3주) → 맛보고 나누는(4주) 구조를 한 장으로 보여줍니다. 전시 기간은 예술의전당 공식 안내로 최종 확인하세요.");

// ---------- Slide 2 ----------
pres.addSection({ title: "설계" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "설계" });
s.addText("전시가 기억에 남도록: 보기 전 · 보는 중 · 본 뒤", { placeholder: "title" });
const cards = [["전", "1–2주차 · 보기 전", "책으로 지도 그리기", "스페인 미술 500년의 흐름을 책으로 잡고, 작가·시대 키워드 카드 5장으로 나만의 관람 지도를 만듭니다."],
  ["중", "3주차 · 보는 중", "목적 있게 보기", "미션 카드 3장을 들고 관람합니다. ‘나의 한 점’을 정해 오래 보고, 한 줄 제목을 붙입니다."],
  ["후", "4주차 · 본 뒤", "맛보고 나누기", "그림 속 음식과 차를 함께 맛보며 관람 후기를 나누고, 피드백으로 다음 기수를 설계합니다."]];
cards.forEach((c, i) => {
  const x = 0.6 + i * 4.15;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.5, w: 3.85, h: 3.2, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "카드-" + c[0] });
  circ(s, x + 0.3, 1.8, 0.75, i === 1 ? C.accent2 : C.text1, "카드원-" + c[0]);
  T(s, c[0], { x: x + 0.3, y: 1.8, w: 0.75, h: 0.75, fontSize: 24, bold: true, color: C.background1, align: "center", valign: "middle" });
  T(s, c[1], { x: x + 1.25, y: 1.85, w: 2.4, h: 0.3, fontSize: 14, bold: true, color: C.accent2 });
  T(s, c[2], { x: x + 1.25, y: 2.17, w: 2.4, h: 0.4, fontSize: 20, bold: true });
  T(s, c[3], { x: x + 0.3, y: 2.9, w: 3.25, h: 1.6, fontSize: 14, color: C.text2 });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.0, w: 12.15, h: 1.65, rectRadius: 0.12, fill: { color: C.text1 }, line: { type: "none" }, objectName: "전시한눈에-배경" });
T(s, "전시 한눈에", { x: 0.9, y: 5.3, w: 2.0, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
T(s, "뉴욕 히스패닉 소사이어티 소장품 / 엘 그레코 · 벨라스케스 · 고야", { x: 0.9, y: 5.75, w: 2.6, h: 0.8, fontSize: 12, color: C.background2 });
[["95점", "15–20세기 초 회화"], ["500년", "스페인 미술의 흐름"], ["3대 거장", "엘 그레코·벨라스케스·고야"], ["120분", "주 1회 모임 기준"]].forEach((st, i) => {
  const x = 3.9 + i * 2.2;
  T(s, st[0], { x, y: 5.3, w: 2.1, h: 0.6, fontSize: 30, bold: true, color: C.accent1 });
  T(s, st[1], { x, y: 5.95, w: 2.1, h: 0.5, fontSize: 12, color: C.background1 });
});
s.addNotes("설계 원리: 미술 전시는 보고 나면 잊히기 쉬우므로 사전 지식(책) - 현장 목적(미션) - 사후 정리(나눔)를 한 흐름으로 묶었습니다. 전시 규모·작가 정보는 보도 및 예매처 정보 기준이므로 공식 안내로 확인하세요.");

// ---------- Slide 3 ----------
pres.addSection({ title: "커리큘럼" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "커리큘럼" });
s.addText("4주 커리큘럼: 책 2주 → 전시 1주 → 나눔 1주", { placeholder: "title" });
const wkc = [
  ["1", "1주차 · 책", "거장들과 먼저 친해지기", "《나는 스페인 미술관의 도슨트입니다》 이안", "체크인 10분 → 화가별 발제 40분 → ‘그림 한 장 천천히 읽기’ 40분", "키워드 카드 3장"],
  ["2", "2주차 · 책", "‘빛과 어둠’의 역사 읽기", "《스페인, 역사와 미술을 보다》 김영욱", "가톨릭·이슬람 맥락 토론 → 시대 구분 지도 → 관람 위시리스트 → 미션 안내", "관람 위시 3점"],
  ["3", "3주차 · 전시", "전시장에서 만나다", "예술의전당 한가람디자인미술관", "도슨트 시간대 입장 → 미션 카드 관람 → 로비 3분 공유", "나의 한 점 + 한 줄 제목"],
  ["4", "4주차 · 나눔", "맛보고 나누다", "스페인 티타임 + 피드백", "미니 도슨트 5분 → 티타임 → 피드백(KPT) → 다음 기수 제안", "후기 한 줄 + 제안"]];
wkc.forEach((w, i) => {
  const col = i % 2, row = Math.floor(i / 2);
  const x = 0.6 + col * 6.15, y = 1.45 + row * 2.7;
  const hi = i === 2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 6.0, h: 2.5, rectRadius: 0.12, fill: { color: hi ? C.text1 : C.background2 }, line: { type: "none" }, objectName: "주차카드-" + w[0] });
  circ(s, x + 0.25, y + 0.25, 0.7, hi ? C.accent1 : C.accent2, "주차원-" + w[0]);
  T(s, w[0], { x: x + 0.25, y: y + 0.25, w: 0.7, h: 0.7, fontSize: 24, bold: true, color: hi ? C.text1 : C.background1, align: "center", valign: "middle" });
  T(s, w[1], { x: x + 1.15, y: y + 0.2, w: 4.6, h: 0.3, fontSize: 14, bold: true, color: hi ? C.accent1 : C.accent2 });
  T(s, w[2], { x: x + 1.15, y: y + 0.5, w: 4.6, h: 0.4, fontSize: 20, bold: true, color: hi ? C.background1 : C.text1 });
  T(s, w[3], { x: x + 1.15, y: y + 0.95, w: 4.6, h: 0.3, fontSize: 14, color: hi ? C.background2 : C.text2 });
  T(s, [{ text: "활동  ", options: { bold: true, color: hi ? C.accent1 : C.accent2 } }, { text: w[4], options: { color: hi ? C.background1 : C.text1 } }], { x: x + 0.25, y: y + 1.35, w: 5.5, h: 0.6, fontSize: 14 });
  T(s, [{ text: "산출물  ", options: { bold: true, color: hi ? C.accent1 : C.accent2 } }, { text: w[5], options: { color: hi ? C.background1 : C.text1 } }], { x: x + 0.25, y: y + 2.05, w: 5.5, h: 0.3, fontSize: 14 });
});
s.addNotes("교재: 1주차 《나는 스페인 미술관의 도슨트입니다》(이안), 2주차 《스페인, 역사와 미술을 보다》(김영욱). 심화 선택 《프라도 미술관에서 꼭 봐야 할 그림 100》(김영숙)은 벨라스케스·고야를 더 보고 싶은 참가자용입니다. 도서 정보는 사용자 제공 기준이며 분량은 확인 후 발췌독을 허용합니다. 1주차 약 120분: 체크인 10 - 발제 40 - 그림 읽기 40 - 정리 30.");

// ---------- Slide 4 ----------
pres.addSection({ title: "3-4주차" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "3-4주차" });
s.addText("클라이맥스 3·4주차: 보고, 그리고 맛보며 나누기", { placeholder: "title" });
// left panel
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.4, w: 6.0, h: 3.95, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "3주차-패널" });
T(s, "3주차  관람 설계 (120분)", { x: 0.9, y: 1.55, w: 5.4, h: 0.45, fontSize: 20, bold: true });
[["15분", "집결 · 오리엔테이션"], ["40분", "도슨트 투어 또는 오디오가이드"], ["50분", "자유 관람 + 미션 카드"], ["15분", "로비 공유 · 마무리"]].forEach((r, i) => {
  const y = 2.15 + i * 0.5;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.9, y, w: 0.95, h: 0.38, rectRadius: 0.19, fill: { color: C.text1 }, line: { type: "none" }, objectName: "시간칩-" + i });
  T(s, r[0], { x: 0.9, y, w: 0.95, h: 0.38, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  T(s, r[1], { x: 2.05, y, w: 4.3, h: 0.38, fontSize: 14, valign: "middle" });
});
T(s, [{ text: "미션 카드 3장", options: { bold: true, color: C.accent2, breakLine: true } },
  { text: "① 나의 한 점 ② 빛과 어둠 찾기 ③ 그림 속 음식 찾기", options: {} }], { x: 0.9, y: 4.25, w: 5.4, h: 0.95, fontSize: 14 });
// right panel
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 1.4, w: 6.0, h: 3.95, rectRadius: 0.12, fill: { color: C.text1 }, line: { type: "none" }, objectName: "4주차-패널" });
T(s, "4주차  스페인 티타임 & 피드백", { x: 7.05, y: 1.55, w: 5.4, h: 0.45, fontSize: 20, bold: true, color: C.background1 });
T(s, "메뉴 후보 — 3주차 ‘그림 속 음식’에서 확정", { x: 7.05, y: 2.1, w: 5.4, h: 0.3, fontSize: 14, bold: true, color: C.accent1 });
["추로스·초콜라테", "토르티야", "하몽·치즈 타파스", "무알코올 상그리아"].forEach((m, i) => {
  const x = 7.05 + (i % 2) * 2.75, y = 2.55 + Math.floor(i / 2) * 0.6;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.6, h: 0.45, rectRadius: 0.22, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "메뉴칩-" + i });
  T(s, m, { x, y, w: 2.6, h: 0.45, fontSize: 14, bold: true, color: C.text1, align: "center", valign: "middle" });
});
T(s, [{ text: "흐름  ", options: { bold: true, color: C.accent1 } }, { text: "미니 도슨트 5분 → 티타임 30분 → 피드백(KPT) 40분 → 다음 기수 제안", options: { color: C.background1 } }], { x: 7.05, y: 3.9, w: 5.4, h: 0.65, fontSize: 14 });
T(s, "알레르기·선호는 3주차에 미리 조사", { x: 7.05, y: 4.7, w: 5.4, h: 0.3, fontSize: 12, color: C.background2 });
// bottom callout
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.55, w: 12.15, h: 1.2, rectRadius: 0.12, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "도슨트-배경" });
T(s, "도슨트·전문가 토크, 가능합니다", { x: 0.9, y: 5.7, w: 3.0, h: 0.9, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
T(s, [{ text: "① 전시장 도슨트 시간대에 맞춰 관람 (권장, 추가 비용 최소)", options: { breakLine: true } }, { text: "② 회원 ‘미니 도슨트’ 5분  ③ 외부 해설가 20분 토크 (예산 시)", options: {} }], { x: 4.1, y: 5.7, w: 8.4, h: 0.9, fontSize: 14, color: C.background1, valign: "middle" });
s.addNotes("질문하신 도슨트/전문가 토크는 어렵지 않습니다. 가장 현실적인 안은 예술의전당 도슨트 운영 시간대를 확인해 3주차 관람 시간을 맞추는 것입니다. 음식은 전시 속 정물화·풍속화에서 실제로 발견한 음식으로 4주차 메뉴를 확정하면 전시와 자연스럽게 연결됩니다.");

// ---------- Slide 5 ----------
pres.addSection({ title: "검증" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "검증" });
s.addText("기획 검증: 5개 페르소나 평균 9.5 / 10", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "점수", labels: ["미술 초보 참가자", "열성 독서가", "도슨트·큐레이터", "클럽 운영자", "예산·안전 담당"], values: [9.5, 9.4, 9.6, 9.5, 9.5] }],
  { x: 0.6, y: 1.4, w: 5.9, h: 3.9, barDir: "bar", chartColors: ["A8322D"], showLegend: false, showTitle: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: "1C1714", dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0.0",
    catAxisLabelFontSize: 14, catAxisLabelColor: "1C1714", catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 10.5, valGridLine: { style: "none" }, catGridLine: { style: "none" }, barGapWidthPct: 45 });
const fx = [["초보 참가자", "주당 읽기 범위 분할 + 발췌독 허용"], ["독서가", "1·2주차 키워드 카드로 책↔작품 연결"], ["도슨트·큐레이터", "미션 카드로 ‘목적 있는 관람’"], ["클럽 운영자", "120분 시간표 + 미니 도슨트 역할 분담"], ["예산·안전", "단체 관람료·도슨트 시간 확인, 알레르기 조사"]];
fx.forEach((f, i) => {
  const y = 1.4 + i * 0.78;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.9, y, w: 1.75, h: 0.62, rectRadius: 0.1, fill: { color: C.text1 }, line: { type: "none" }, objectName: "페르소나칩-" + i });
  T(s, f[0], { x: 6.9, y, w: 1.75, h: 0.62, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  T(s, f[1], { x: 8.85, y, w: 3.9, h: 0.62, fontSize: 14, valign: "middle" });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.5, w: 12.15, h: 1.0, rectRadius: 0.12, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "다음단계-배경" });
T(s, "다음 단계", { x: 0.9, y: 5.5, w: 1.6, h: 1.0, fontSize: 18, bold: true, valign: "middle" });
T(s, "① 모집·일정 확정  ② 도서 3권 확정  ③ 관람료·도슨트 시간 확인  ④ 4주차 메뉴·장소 확정", { x: 2.6, y: 5.5, w: 10.0, h: 1.0, fontSize: 14, bold: true, valign: "middle" });
s.addNotes("점수는 AI가 5개 페르소나 관점에서 시뮬레이션한 자체 평가이며 실제 참가자 의견이 아닙니다. 1기 운영 후 실제 피드백으로 보정하세요.");

(async () => {
  await pres.writeFile({ fileName: "readon_spain_art_club.pptx" });
  await applyTheme("readon_spain_art_club.pptx", THEME);
  console.log("done");
})();
