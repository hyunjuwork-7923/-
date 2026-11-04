const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/b5915536-3fe3-4a22-9798-2179749b4aa4_c0048d07-4707-4906-bf38-1c686d3188cc/pptx/scripts/apply_theme.js");
const FONT = "Pretendard";
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
s.addText("책 두 권 읽고, 전시 보러 가는 한 달", { placeholder: "title" });
T(s, "전시 《스페인 미술 500년: 빛과 어둠의 연대기》 · 4주", { x: 0.8, y: 3.6, w: 8.2, h: 0.5, fontSize: 20, color: C.background1 });
T(s, "예술의전당 한가람디자인미술관 · 2026.9.22 – 2027.1.20", { x: 0.8, y: 4.2, w: 8.2, h: 0.4, fontSize: 14, color: C.background2 });
T(s, "운영 장소  서울 서초구 남부순환로315길 85 · 정원 10명", { x: 0.8, y: 4.6, w: 8.2, h: 0.35, fontSize: 14, color: C.background2 });
const wk = [["1", "책 읽기", "화가 익히기"], ["2", "책 읽기", "역사 배경"], ["3", "전시 관람", "한가람 2시간"], ["4", "티타임·피드백", "먹으며 후기"]];
wk.forEach((w, i) => {
  const x = 0.8 + i * 3.0;
  circ(s, x, 5.5, 0.85, i === 2 ? C.accent2 : C.accent1, "주차-원" + (i + 1));
  T(s, w[0], { x, y: 5.5, w: 0.85, h: 0.85, fontSize: 28, bold: true, color: i === 2 ? C.background1 : C.text1, align: "center", valign: "middle" });
  T(s, [{ text: w[1], options: { bold: true, fontSize: 16, color: C.background1, breakLine: true } }, { text: w[2], options: { fontSize: 14, color: C.background2 } }], { x: x + 1.05, y: 5.55, w: 1.85, h: 0.8, valign: "middle" });
});
s.addNotes("표지. 1~2주는 책, 3주는 전시, 4주는 티타임입니다. 전시 기간은 예술의전당 공식 안내로 최종 확인하세요.");

// ---------- Slide 3 ----------
pres.addSection({ title: "커리큘럼" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "커리큘럼" });
s.addText("4주 커리큘럼: 책 2주 → 전시 1주 → 나눔 1주", { placeholder: "title" });
const wkc = [
  ["1", "1주차 · 책", "화가 세 명 얼굴부터 익히기", "《나는 스페인 미술관의 도슨트입니다》 이안", "체크인 10분 → 화가별 발제 40분 → ‘그림 한 장 천천히 읽기’ 40분", "키워드 카드 3장"],
  ["2", "2주차 · 책", "가톨릭과 이슬람, 그림에 남은 흔적", "《스페인, 역사와 미술을 보다》 김영욱", "가톨릭·이슬람 맥락 토론 → 시대 구분 지도 → 관람 위시리스트 → 미션 안내", "관람 위시 3점"],
  ["3", "3주차 · 전시", "한가람디자인미술관에서 2시간", "예술의전당 한가람디자인미술관", "도슨트 시간대 입장 → 미션 카드 관람 → 로비 3분 공유", "나의 한 점 + 한 줄 제목"],
  ["4", "4주차 · 나눔", "추로스 먹으면서 후기 말하기", "스페인 티타임 + 피드백", "미니 도슨트 5분 → 티타임 → 피드백(KPT) → 다음 기수 제안", "후기 한 줄 + 제안"]];
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
s.addText("3·4주차: 전시 보고 와서 추로스 먹기", { placeholder: "title" });
// left panel
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.4, w: 6.0, h: 3.95, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "3주차-패널" });
T(s, "3주차  관람 설계 (120분)", { x: 0.9, y: 1.55, w: 5.4, h: 0.45, fontSize: 20, bold: true });
[["15분", "공간 집결 · 오리엔테이션(이동 포함)"], ["40분", "도슨트 투어 또는 오디오가이드"], ["50분", "자유 관람 + 미션 카드"], ["15분", "로비 공유 · 마무리"]].forEach((r, i) => {
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
T(s, "음식 반입·취식 가능 · 알레르기는 3주차에 조사", { x: 7.05, y: 4.7, w: 5.4, h: 0.3, fontSize: 12, color: C.background2 });
// bottom callout
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.55, w: 12.15, h: 1.2, rectRadius: 0.12, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "도슨트-배경" });
T(s, "도슨트나 전문가 토크는 이렇게 넣습니다", { x: 0.9, y: 5.7, w: 3.0, h: 0.9, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
T(s, [{ text: "① 전시장 도슨트 시간대에 맞춰 관람 (권장, 추가 비용 최소)", options: { breakLine: true } }, { text: "② 회원 ‘미니 도슨트’ 5분  ③ 외부 해설가 20분 토크 (예산 시)", options: {} }], { x: 4.1, y: 5.7, w: 8.4, h: 0.9, fontSize: 14, color: C.background1, valign: "middle" });
s.addNotes("질문하신 도슨트/전문가 토크는 어렵지 않습니다. 가장 현실적인 안은 예술의전당 도슨트 운영 시간대를 확인해 3주차 관람 시간을 맞추는 것입니다. 음식은 전시 속 정물화·풍속화에서 실제로 발견한 음식으로 4주차 메뉴를 확정하면 전시와 자연스럽게 연결됩니다.");


// ---------- Slide 4 market ----------
pres.addSection({ title: "시장분석" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "시장분석" });
s.addText("시장 비교: 시즌 모임 17만~40만 원, 원데이 1.8만~8만 원", { placeholder: "title" });
const X0 = 1.2, Y0 = 5.1, PX = 0.36; // 축 원점, 1만 원당 높이
const yv = (v) => Y0 - PX * v;
[5, 10].forEach((v) => s.addShape(pres.shapes.LINE, { x: X0, y: yv(v), w: 6.8, h: 0, line: { color: C.background2, width: 1 }, objectName: "눈금선-" + v }));
s.addShape(pres.shapes.LINE, { x: X0, y: 1.5, w: 0, h: Y0 - 1.5, line: { color: C.text2, width: 1.25 }, objectName: "세로축" });
s.addShape(pres.shapes.LINE, { x: X0, y: Y0, w: 6.8, h: 0, line: { color: C.text2, width: 1.25 }, objectName: "가로축" });
[[0, "0"], [5, "5만"], [10, "10만"]].forEach((r) => T(s, r[1], { x: 0.55, y: yv(r[0]) - 0.15, w: 0.55, h: 0.3, fontSize: 12, color: C.text2, align: "right", valign: "middle" }));
T(s, "회당 가격", { x: 0.6, y: 1.2, w: 1.5, h: 0.25, fontSize: 12, bold: true, color: C.text2 });
[["1회", 2.1], ["4주", 3.5], ["3개월", 5.7], ["4개월", 7.3]].forEach((r) => T(s, r[0], { x: r[1] - 0.5, y: Y0 + 0.08, w: 1.0, h: 0.28, fontSize: 12, color: C.text2, align: "center" }));
T(s, "약정 기간 · 몰입 깊이 →", { x: 4.5, y: Y0 + 0.42, w: 3.5, h: 0.28, fontSize: 12, bold: true, color: C.text2, align: "right" });
const cap = (cx, lo, hi, name) => { const y = yv(hi), h = Math.max(yv(lo) - yv(hi), 0.3); s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx - 0.25, y, w: 0.5, h, rectRadius: 0.25, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "범위-" + name }); };
cap(1.7, 1.8, 8, "원데이"); cap(2.5, 1.4, 2.3, "전시"); cap(5.7, 4.3, 5.8, "정규"); cap(7.3, 6.3, 10, "트레바리");
const lab = (name, sub, x, y, w, al, col) => T(s, [{ text: name, options: { bold: true, fontSize: 14, color: col || C.text1, breakLine: true } }, { text: sub, options: { fontSize: 12, color: col || C.text2 } }], { x, y, w, h: 0.5, align: al });
lab("넷플연가 원데이", "1.8–8만", 1.3, 1.65, 1.8, "left");
lab("전시 단독 관람", "1.4–2.3만", 2.9, 4.45, 1.6, "left");
lab("넷플연가 정규", "4.3–5.8만", 3.9, 2.5, 1.65, "right");
lab("트레바리", "6.3–10만", 5.3, 1.65, 1.6, "right");
circ(s, 3.0, 3.36, 1.0, C.accent2, "가격대-고리", 80);
circ(s, 3.2, 3.56, 0.6, C.accent2, "우리-점");
lab("리드온 (우리)", "3.2–3.7만", 4.1, 3.62, 1.7, "left", C.accent2);
T(s, "가격은 사용자 제공 자료 기준이며 공식 안내로 재확인이 필요합니다. 문토·소셜링 독서모임과 전시 연계 도슨트 모임은 가격을 확인하지 못했습니다. 막대는 회당 가격 범위입니다.", { x: 0.6, y: 5.95, w: 7.4, h: 0.6, fontSize: 12, color: C.text2 });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.3, y: 1.45, w: 4.45, h: 5.05, rectRadius: 0.12, fill: { color: C.text1 }, line: { type: "none" }, objectName: "우리의자리-배경" });
T(s, "우리 모임", { x: 8.6, y: 1.65, w: 3.9, h: 0.45, fontSize: 20, bold: true, color: C.background1 });
[["가격", "4회에 12.9만~14.9만 원. 넷플연가 정규 하단(17만 원)보다 낮게 잡았다"], ["기간", "4개월 약정 없이 4주에 끝난다. 전시 관람료와 다과는 참가비에 포함한다"], ["대상", "미술은 처음인 직장인, 리드온 회원, 전시는 가는데 혼자 보면 남는 게 적다는 사람"]].forEach((r, i) => {
  const y = 2.3 + i * 1.35;
  T(s, r[0], { x: 8.6, y, w: 3.9, h: 0.3, fontSize: 14, bold: true, color: C.accent1 });
  T(s, r[1], { x: 8.6, y: y + 0.33, w: 3.9, h: 0.8, fontSize: 14, color: C.background1 });
});
s.addNotes("포지셔닝 맵: 가로축은 약정 기간·몰입 깊이, 세로축은 회당 가격(막대는 범위)입니다. 시장 지도: 시즌 구독형(트레바리), 호스트형 멤버십(넷플연가), 원데이, 전시 단독 관람. 우리는 4주 안에 끝나고 전시 관람과 다과가 포함된 중간 가격대입니다. 트레바리·넷플연가 가격은 사용자 제공 자료 기준입니다.");

// ---------- Slide 5 pricing ----------
pres.addSection({ title: "참가비" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "참가비" });
s.addText("참가비: 129,000원부터, 10명까지", { placeholder: "title" });
[["얼리버드", "129,000원", "회당 약 3.2만 원 · 마감 전 신청", false], ["정가", "149,000원", "회당 약 3.7만 원 · 10명 한정", true], ["별도", "책값", "도서 3권 개별 구입 · 3번째는 선택", false]].forEach((c, i) => {
  const x = 0.6 + i * 4.15;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.45, w: 3.85, h: 1.85, rectRadius: 0.12, fill: { color: c[3] ? C.accent2 : C.background2 }, line: { type: "none" }, objectName: "가격카드-" + i });
  T(s, c[0], { x: x + 0.3, y: 1.62, w: 3.3, h: 0.3, fontSize: 14, bold: true, color: c[3] ? C.background1 : C.accent2 });
  T(s, c[1], { x: x + 0.3, y: 1.98, w: 3.3, h: 0.7, fontSize: 36, bold: true, color: c[3] ? C.background1 : C.text1 });
  T(s, c[2], { x: x + 0.3, y: 2.78, w: 3.3, h: 0.35, fontSize: 14, color: c[3] ? C.background1 : C.text2 });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.55, w: 6.0, h: 3.1, rectRadius: 0.12, fill: { color: C.text1 }, line: { type: "none" }, objectName: "수지-배경" });
T(s, "10명 정가 기준 수지", { x: 0.9, y: 3.7, w: 5.4, h: 0.4, fontSize: 18, bold: true, color: C.background1 });
[["수입", "1,490,000원"], ["비용", "580,000원"], ["남는 돈", "910,000원"]].forEach((r, i) => {
  const y = 4.25 + i * 0.5;
  T(s, r[0], { x: 0.9, y, w: 1.6, h: 0.4, fontSize: 14, color: C.background2, valign: "middle" });
  T(s, r[1], { x: 2.5, y, w: 3.8, h: 0.4, fontSize: 22, bold: true, color: i === 2 ? C.accent1 : C.background1, valign: "middle" });
});
T(s, "비용: 관람료 23만 · 다과 12만 · 자료 3만 · 외부 해설 20만 (가정). 남는 돈은 진행자 보상·예비비. 최소 개설 8명", { x: 0.9, y: 5.8, w: 5.4, h: 0.75, fontSize: 12, color: C.background2 });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 3.55, w: 6.0, h: 3.1, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "운영-배경" });
T(s, "운영 조건", { x: 7.05, y: 3.7, w: 5.4, h: 0.4, fontSize: 18, bold: true });
const bl = (a, b, last) => [{ text: a + "  ", options: { bold: true, color: C.accent2, bullet: false } }, { text: b, options: { breakLine: !last } }];
T(s, [].concat(bl("장소", "서초구 남부순환로315길 85 (무료 대관)"), bl("정원", "10명 · 음식 반입·취식 가능"), bl("포함", "관람료 · 다과 · 자료"), bl("환불", "1주차 전 전액 · 관람 이후 불가", true)), { x: 7.05, y: 4.25, w: 5.4, h: 2.2, fontSize: 14, paraSpaceAfter: 8 });
s.addNotes("정원 10명이라 수입 상한이 정해져 있어 정가를 149,000원으로 두고, 얼리버드 129,000원으로 초기 모집을 유도합니다. 비용은 가정값이며 관람료 단체 할인 여부, 외부 해설 견적, 진행자 포함 여부를 확인해 갱신하세요.");
(async () => {
  await pres.writeFile({ fileName: "readon_spain_art_club.pptx" });
  await applyTheme("readon_spain_art_club.pptx", THEME);
  console.log("done");
})();
