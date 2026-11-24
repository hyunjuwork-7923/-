// v11: 13.33인치 슬라이드에 맞춰 글자를 키우고(본문 20pt 안팎), 한 장에 담던 내용을 나눠 8장으로 구성
const fs = require("fs");
const NOTES = JSON.parse(fs.readFileSync(process.env.NOTES_JSON, "utf8")).map((x) => JSON.parse(x));
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/b5915536-3fe3-4a22-9798-2179749b4aa4_c0048d07-4707-4906-bf38-1c686d3188cc/pptx/scripts/apply_theme.js");
const FONT = "Pretendard";
const THEME = { name: "Spain Light and Dark", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1:"1C1714", lt1:"FFFFFF", dk2:"3A2D26", lt2:"F4F0EA", accent1:"C8932B", accent2:"A8322D",
    accent3:"2E4A5A", accent4:"7C8A4A", accent5:"8C5A3C", accent6:"5B4B7A", hlink:"A8322D", folHlink:"5B4B7A" } };
const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE"; pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "문득클럽 스페인 미술 3회 프로그램 기획안";
const C = pres.SchemeColor;
pres.defineSlideMaster({ title: "DARK", background: { color: C.text1 },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.6, w: 9.4, h: 2.2, fontSize: 48, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } }] });
pres.defineSlideMaster({ title: "LIGHT", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 34, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "문득클럽 · 스페인 미술 3회 과정", options: { x: 0.6, y: 6.95, w: 6, h: 0.3, fontSize: 12, color: C.text2, margin: 0, isTextBox: true } } },
  ], slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 12, color: C.text2 } });
const T = (s, t, o) => s.addText(t, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1 }, o));
const circ = (s, x, y, d, fill, name, tr) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill, transparency: tr || 0 }, line: { type: "none" }, objectName: name });
const card = (s, x, y, w, h, fill, name, r) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: r || 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });
const line = (s, x, y, w, h, col, wd, name) => s.addShape(pres.shapes.LINE, { x, y, w, h, line: { color: col, width: wd }, objectName: name });

// ===== 1 표지
let s = pres.addSlide({ masterName: "DARK" });
circ(s, 9.5, 0.6, 3.5, C.accent1, "빛-외곽", 85); circ(s, 10.05, 1.15, 2.4, C.accent1, "빛-중간", 70); circ(s, 10.6, 1.7, 1.3, C.accent1, "빛-중심", 0);
T(s, "문득클럽 · 파일럿 프로그램 기획안", { x: 0.8, y: 0.9, w: 8, h: 0.45, fontSize: 20, bold: true, color: C.accent1 });
s.addText("벨라스케스 보러 가기 전, 책으로 친해지기", { placeholder: "title" });
T(s, "《스페인 미술 500년: 빛과 어둠의 연대기》", { x: 0.8, y: 3.8, w: 10, h: 0.55, fontSize: 24, color: C.background1 });
[["1", "12/16 (수)", "19:00 · 책 읽기"], ["2", "12/30 (수)", "19:00 · 책 읽기"], ["3", "1/13 (수)", "19:00 · 전시와 해설"]].forEach((w, i) => {
  const x = 0.8 + i * 4.0;
  circ(s, x, 4.95, 1.05, i === 2 ? C.accent2 : C.accent1, "회차-원" + (i + 1));
  T(s, w[0], { x, y: 4.95, w: 1.05, h: 1.05, fontSize: 34, bold: true, color: i === 2 ? C.background1 : C.text1, align: "center", valign: "middle" });
  T(s, [{ text: w[1], options: { bold: true, fontSize: 24, color: C.background1, breakLine: true } }, { text: w[2], options: { fontSize: 18, color: C.background2 } }], { x: x + 1.25, y: 4.95, w: 2.7, h: 1.05, valign: "middle" });
});
T(s, "서초 · 참가자 10명 · 파일럿 특가 76,000원", { x: 0.8, y: 6.45, w: 10, h: 0.35, fontSize: 16, color: C.background2 });
s.addNotes(NOTES[0]);

// ===== 2 커리큘럼
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("3회 커리큘럼: 책 2회, 그리고 전시와 나눔", { placeholder: "title" });
card(s, 0.6, 1.35, 12.15, 0.9, C.text1, "목표-배경");
T(s, "모임 목표", { x: 0.9, y: 1.35, w: 1.7, h: 0.9, fontSize: 18, bold: true, color: C.accent1, valign: "middle" });
T(s, "3회 과정을 지나면, 스페인 미술에 대해 설명할 수 있어요.", { x: 2.7, y: 1.35, w: 9.9, h: 0.9, fontSize: 24, bold: true, color: C.background1, valign: "middle" });
const cur = [
  ["1", "12/16 (수) · 책", "화가 세 명 얼굴부터 익히기", "《나는 스페인 미술관의 도슨트입니다》 이안", "발제 · 그림 읽기 · 구분 게임", "키워드 카드 3장", false],
  ["2", "12/30 (수) · 책", "가톨릭과 이슬람, 그림에 남은 흔적", "《스페인, 역사와 미술을 보다》 김영욱", "토론 · 시대 지도 · 위시리스트", "위시 3점 · 티켓", false],
  ["3", "1/13 (수) · 전시", "각자 보고, 저녁에 모여 3분 해설", "관람 티켓 제공 · 큐레이터 게스트(섭외 중)", "큐레이터 토크 · 간식 · 발표", "나의 한 점 해설", true]];
cur.forEach((c, i) => {
  const x = 0.6 + i * 4.125, hi = c[6];
  card(s, x, 2.5, 3.9, 4.25, hi ? C.text1 : C.background2, "회차카드-" + c[0]);
  circ(s, x + 0.3, 2.7, 0.75, hi ? C.accent1 : C.accent2, "회차원-" + c[0]);
  T(s, c[0], { x: x + 0.3, y: 2.7, w: 0.75, h: 0.75, fontSize: 28, bold: true, color: hi ? C.text1 : C.background1, align: "center", valign: "middle" });
  T(s, c[1], { x: x + 1.25, y: 2.7, w: 2.55, h: 0.75, fontSize: 18, bold: true, color: hi ? C.accent1 : C.accent2, valign: "middle" });
  T(s, c[2], { x: x + 0.3, y: 3.65, w: 3.4, h: 1.0, fontSize: 26, bold: true, color: hi ? C.background1 : C.text1 });
  T(s, c[3], { x: x + 0.3, y: 4.75, w: 3.4, h: 0.6, fontSize: 16, color: hi ? C.background2 : C.text2 });
  T(s, c[4], { x: x + 0.3, y: 5.45, w: 3.4, h: 0.7, fontSize: 20, color: hi ? C.background1 : C.text1 });
  T(s, [{ text: "산출물  ", options: { bold: true, color: hi ? C.accent1 : C.accent2 } }, { text: c[5], options: { color: hi ? C.background1 : C.text1 } }], { x: x + 0.3, y: 6.2, w: 3.45, h: 0.4, fontSize: 18 });
});
s.addNotes(NOTES[1]);

// ===== 3 1·2회차 시간표
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("1·2회차 시간표: 19:00~21:30", { placeholder: "title" });
const tt = [
  ["1회차 · 12/16 (수)", "화가 세 명 얼굴부터", [[15, "체크인 · 자기평가"], [40, "화가별 발제"], [10, "휴식"], [35, "그림 한 장 읽기"], [25, "화가 구분 게임"], [25, "카드 작성 · 정리"]]],
  ["2회차 · 12/30 (수)", "가톨릭과 이슬람", [[10, "체크인 · 카드 공유"], [35, "맥락 토론"], [10, "휴식"], [35, "시대 구분 지도"], [35, "관람 위시리스트"], [25, "티켓 · 미션 안내"]]]];
tt.forEach((c, i) => {
  const x = 0.6 + i * 6.15;
  card(s, x, 1.4, 6.0, 5.3, C.background2, "시간표-패널-" + i);
  T(s, c[0], { x: x + 0.4, y: 1.55, w: 5.2, h: 0.35, fontSize: 18, bold: true, color: C.accent2 });
  T(s, c[1], { x: x + 0.4, y: 1.95, w: 5.2, h: 0.5, fontSize: 26, bold: true });
  c[2].forEach((r, j) => {
    const y = 2.7 + j * 0.66;
    card(s, x + 0.4, y, 1.05, 0.5, C.text1, `분칩-${i}-${j}`, 0.25);
    T(s, r[0] + "분", { x: x + 0.4, y, w: 1.05, h: 0.5, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, r[1], { x: x + 1.7, y, w: 4.1, h: 0.5, fontSize: 22, valign: "middle" });
  });
});
s.addNotes(NOTES[2]);

// ===== 4 3회차 시간표 (A/B)
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("3회차 시간표: A안과 B안", { placeholder: "title" });
const ab = [
  ["A안 · 큐레이터가 오시면", "1/13 (수) 19:00~21:30", [[15, "체크인 · 간식"], [50, "큐레이터 토크 · 질문"], [10, "휴식"], [25, "3분 해설 초안"], [40, "해설 발표"], [10, "자기평가 · 설문"]], C.text1, C.accent1, C.background1],
  ["B안 · 못 모시면", "큐레이터 확정 기한 12/30", [[15, "체크인 · 간식"], [30, "미션 카드 후기"], [10, "휴식"], [30, "해설 초안 · 짝 연습"], [50, "해설 발표"], [15, "설문 · 정리"]], C.accent2, C.background1, C.background1]];
ab.forEach((c, i) => {
  const x = 0.6 + i * 6.15;
  card(s, x, 1.4, 6.0, 5.3, c[3], "안-패널-" + i);
  T(s, c[0], { x: x + 0.4, y: 1.55, w: 5.2, h: 0.5, fontSize: 26, bold: true, color: C.background1 });
  T(s, c[1], { x: x + 0.4, y: 2.1, w: 5.2, h: 0.35, fontSize: 18, color: c[4] });
  c[2].forEach((r, j) => {
    const y = 2.7 + j * 0.66;
    card(s, x + 0.4, y, 1.05, 0.5, c[4], `안-분칩-${i}-${j}`, 0.25);
    T(s, r[0] + "분", { x: x + 0.4, y, w: 1.05, h: 0.5, fontSize: 20, bold: true, color: i === 0 ? C.text1 : C.accent2, align: "center", valign: "middle" });
    T(s, r[1], { x: x + 1.7, y, w: 4.1, h: 0.5, fontSize: 22, color: C.background1, valign: "middle" });
  });
});
s.addNotes(NOTES[2]);

// ===== 5 3회차 진행
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("3회차: 관람은 각자, 저녁에는 같이", { placeholder: "title" });
[["12/30", "티켓과 미션 카드 받기"], ["1/13 낮 또는 미리", "각자 전시 보기"], ["1/13 19:00", "같은 공간에 모이기"]].forEach((f, i) => {
  const x = 0.6 + i * 4.125, hi = i === 2;
  card(s, x, 1.4, 3.9, 1.9, hi ? C.text1 : C.background2, "흐름카드-" + i);
  T(s, f[0], { x: x + 0.3, y: 1.6, w: 3.4, h: 0.45, fontSize: 20, bold: true, color: hi ? C.accent1 : C.accent2 });
  T(s, f[1], { x: x + 0.3, y: 2.15, w: 3.4, h: 0.9, fontSize: 26, bold: true, color: hi ? C.background1 : C.text1 });
  if (i < 2) T(s, "→", { x: x + 3.9, y: 1.95, w: 0.225, h: 0.7, fontSize: 24, bold: true, color: C.text2, align: "center", valign: "middle" });
});
card(s, 0.6, 3.6, 12.15, 1.5, C.accent2, "큐레이터-배경");
T(s, "큐레이터는 12/30까지 확정하고, 못 모시면 B안으로 갑니다", { x: 1.0, y: 3.6, w: 11.4, h: 0.9, fontSize: 28, bold: true, color: C.background1, valign: "middle" });
T(s, "발표 때 그림을 보여줄 화면을 준비합니다", { x: 1.0, y: 4.45, w: 11.4, h: 0.5, fontSize: 20, color: C.background1, valign: "middle" });
T(s, "간식 후보", { x: 0.6, y: 5.55, w: 2.0, h: 0.6, fontSize: 20, bold: true, color: C.accent2, valign: "middle" });
["추로스·초콜라테", "토르티야", "하몽·치즈 타파스", "무알코올 상그리아"].forEach((m, i) => {
  const x = 2.55 + i * 2.55;
  card(s, x, 5.55, 2.4, 0.6, C.accent1, "간식칩-" + i, 0.3);
  T(s, m, { x, y: 5.55, w: 2.4, h: 0.6, fontSize: 18, bold: true, color: C.text1, align: "center", valign: "middle" });
});
s.addNotes(NOTES[3]);

// ===== 6 시장 비교
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("시장 비교: 시즌 17~40만 원, 원데이 1.8~8만 원", { placeholder: "title" });
const X0 = 1.2, Y0 = 5.6, PX = 0.36, yv = (v) => Y0 - PX * v;
[5, 10].forEach((v) => line(s, X0, yv(v), 6.8, 0, C.background2, 1, "눈금선-" + v));
line(s, X0, 1.95, 0, Y0 - 1.95, C.text2, 1.25, "세로축"); line(s, X0, Y0, 6.8, 0, C.text2, 1.25, "가로축");
[[0, "0"], [5, "5만"], [10, "10만"]].forEach((r) => T(s, r[1], { x: 0.45, y: yv(r[0]) - 0.15, w: 0.65, h: 0.3, fontSize: 14, color: C.text2, align: "right", valign: "middle" }));
T(s, "회당 가격", { x: 0.6, y: 1.5, w: 2.0, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
[["1회", 2.1], ["4주", 3.5], ["3개월", 5.7], ["4개월", 7.3]].forEach((r) => T(s, r[0], { x: r[1] - 0.5, y: Y0 + 0.08, w: 1.0, h: 0.32, fontSize: 16, color: C.text2, align: "center" }));
T(s, "약정 기간 · 몰입 깊이 →", { x: 4.0, y: Y0 + 0.5, w: 4.0, h: 0.32, fontSize: 16, bold: true, color: C.text2, align: "right" });
const cap = (cx, lo, hi, n) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx - 0.25, y: yv(hi), w: 0.5, h: Math.max(yv(lo) - yv(hi), 0.3), rectRadius: 0.25, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "범위-" + n });
cap(1.7, 1.8, 8, "원데이"); cap(2.5, 1.4, 2.3, "전시"); cap(5.7, 4.3, 5.8, "정규"); cap(7.3, 6.3, 10, "트레바리");
const lab = (n, sub, x, y, w, al, col) => T(s, [{ text: n, options: { bold: true, fontSize: 18, color: col || C.text1, breakLine: true } }, { text: sub, options: { fontSize: 16, color: col || C.text2 } }], { x, y, w, h: 0.6, align: al });
lab("넷플연가 원데이", "1.8–8만", 1.3, 2.05, 2.6, "left"); lab("넷플연가 정규", "4.3–5.8만", 3.3, 2.85, 2.25, "right"); lab("트레바리", "6.3–10만", 5.0, 2.0, 1.95, "right");
T(s, [{ text: "전시 단독 관람 ", options: { bold: true, fontSize: 18 } }, { text: "1.4–2.3만", options: { fontSize: 16, color: C.text2 } }], { x: 1.95, y: 5.15, w: 4.2, h: 0.35 });
circ(s, 3.1, yv(2.5) - 0.4, 0.8, C.accent2, "가격대-고리", 80); circ(s, 3.2, yv(2.5) - 0.3, 0.6, C.accent2, "우리-점");
lab("문득클럽 (우리)", "약 2.5만", 4.1, yv(2.5) - 0.3, 2.4, "left", C.accent2);
T(s, "가격은 제공 자료 기준입니다. 공식 안내로 다시 확인하세요.", { x: 0.6, y: 6.5, w: 7.4, h: 0.3, fontSize: 14, color: C.text2 });
card(s, 8.3, 1.45, 4.45, 5.1, C.text1, "우리모임-배경");
T(s, "우리 모임", { x: 8.6, y: 1.65, w: 3.9, h: 0.5, fontSize: 26, bold: true, color: C.background1 });
[["가격", "3회 76,000원\n파일럿 특가"], ["기간", "격주 수요일, 3회"], ["대상", "서초 인근 직장인,\n미술이 처음인 사람"]].forEach((r, i) => {
  const y = 2.45 + i * 1.4;
  T(s, r[0], { x: 8.6, y, w: 3.9, h: 0.35, fontSize: 18, bold: true, color: C.accent1 });
  T(s, r[1], { x: 8.6, y: y + 0.4, w: 3.9, h: 0.9, fontSize: 22, color: C.background1 });
});
s.addNotes(NOTES[4]);

// ===== 7 참가비
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("참가비: 파일럿 특가 76,000원, 10명까지", { placeholder: "title" });
card(s, 0.6, 1.45, 6.4, 3.65, C.accent2, "가격-배경");
T(s, "파일럿 특가", { x: 1.0, y: 1.7, w: 5.6, h: 0.45, fontSize: 22, bold: true, color: C.background1 });
T(s, "76,000원", { x: 1.0, y: 2.25, w: 5.6, h: 1.3, fontSize: 72, bold: true, color: C.background1 });
T(s, "3회 · 회당 약 2.5만 원", { x: 1.0, y: 3.75, w: 5.6, h: 0.45, fontSize: 24, color: C.background1 });
T(s, "한·스페인 수교 76주년 전시라서 76,000원 [수교 연도 확인 후 사용]", { x: 1.0, y: 4.35, w: 5.6, h: 0.6, fontSize: 16, color: C.background1 });
[["포함", "티켓 · 간식 · 자료"], ["별도", "책값 (2권)"], ["정원", "참가자 10명"], ["환불", "1회차 전까지 전액"]].forEach((r, i) => {
  const y = 1.45 + i * 0.93;
  card(s, 7.3, y, 5.45, 0.8, C.background2, "조건-배경-" + i);
  T(s, r[0], { x: 7.6, y, w: 1.1, h: 0.8, fontSize: 20, bold: true, color: C.accent2, valign: "middle" });
  T(s, r[1], { x: 8.8, y, w: 3.8, h: 0.8, fontSize: 22, valign: "middle" });
});
card(s, 0.6, 5.4, 12.15, 1.3, C.text1, "수지-배경");
T(s, [{ text: "10명이면 약 18만 원이 남고, ", options: { color: C.background1 } }, { text: "6명부터 손익분기", options: { color: C.accent1, bold: true } }, { text: "입니다", options: { color: C.background1 } }], { x: 0.95, y: 5.4, w: 11.5, h: 0.8, fontSize: 28, valign: "middle" });
T(s, "비용은 가정값입니다 (티켓 23만, 간식 12만, 자료 3만, 외부 강사 20만)", { x: 0.95, y: 6.15, w: 11.5, h: 0.4, fontSize: 16, color: C.background2 });
s.addNotes(NOTES[5]);

// ===== 8 모집
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("모집 계획: 12월 9일까지 6~10명", { placeholder: "title" });
[["11/25 (수)", "모집 오픈"], ["12/2 (수)", "마감 임박 알림"], ["12/9 (수)", "마감 · 티켓 구매"], ["12/16 (수)", "19:00 1회차"]].forEach((d, i) => {
  const x = 0.6 + i * 3.075, hi = i === 3;
  card(s, x, 1.4, 2.95, 1.4, hi ? C.text1 : C.background2, "일정카드-" + i);
  T(s, d[0], { x: x + 0.3, y: 1.55, w: 2.5, h: 0.55, fontSize: 26, bold: true, color: hi ? C.accent1 : C.accent2 });
  T(s, d[1], { x: x + 0.3, y: 2.15, w: 2.5, h: 0.45, fontSize: 20, color: hi ? C.background1 : C.text1 });
});
card(s, 0.6, 3.05, 6.9, 3.65, C.background2, "채널-배경");
T(s, [{ text: "어디서 모으나", options: { bold: true, fontSize: 26 } }, { text: "   인원은 가정", options: { fontSize: 16, color: C.text2 } }], { x: 0.95, y: 3.2, w: 6.3, h: 0.5 });
[["당근", "5명"], ["미술 커뮤니티", "2명"], ["카페 · 서점 · 지인", "2명"], ["인스타 · 문토", "1명"]].forEach((r, i) => {
  const y = 3.9 + i * 0.68;
  T(s, r[0], { x: 0.95, y, w: 4.6, h: 0.55, fontSize: 24, bold: true, color: C.accent2, valign: "middle" });
  card(s, 6.1, y + 0.07, 1.1, 0.42, C.text1, "인원칩-" + i, 0.21);
  T(s, r[1], { x: 6.1, y: y + 0.07, w: 1.1, h: 0.42, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
});
card(s, 7.7, 3.05, 5.05, 3.65, C.text1, "모집글-배경");
T(s, "모집글에 넣을 것", { x: 8.0, y: 3.2, w: 4.5, h: 0.5, fontSize: 26, bold: true, color: C.background1 });
T(s, ["미술이 처음이어도 신청 가능", "격주 수요일 3회, 전시는 각자 관람", "76,000원, 10명 한정, 티켓 포함", "진행자 소개와 환불 규정 공개"].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })),
  { x: 8.0, y: 3.85, w: 4.55, h: 2.7, fontSize: 22, color: C.background1, paraSpaceAfter: 10 });
s.addNotes(NOTES[6]);

(async () => {
  await pres.writeFile({ fileName: "readon_spain_art_club_v11_raw.pptx" });
  await applyTheme("readon_spain_art_club_v11_raw.pptx", THEME);
  console.log("done");
})();
