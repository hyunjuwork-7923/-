// 가독성 개편: 슬라이드당 문장 수를 줄이고 글자 크기를 키운 버전
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
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.7, w: 9.2, h: 2.0, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } }] });
pres.defineSlideMaster({ title: "LIGHT", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "문득클럽 · 스페인 미술 3회 과정", options: { x: 0.6, y: 6.95, w: 6, h: 0.3, fontSize: 12, color: C.text2, margin: 0, isTextBox: true } } },
  ], slideNumber: { x: 12.2, y: 6.95, w: 0.5, h: 0.3, fontSize: 12, color: C.text2 } });
const T = (s, t, o) => s.addText(t, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1 }, o));
const circ = (s, x, y, d, fill, name, tr) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill, transparency: tr || 0 }, line: { type: "none" }, objectName: name });
const card = (s, x, y, w, h, fill, name, r) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: r || 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });
const line = (s, x, y, w, h, col, wd, name) => s.addShape(pres.shapes.LINE, { x, y, w, h, line: { color: col, width: wd }, objectName: name });

// ===== 1 표지
let s = pres.addSlide({ masterName: "DARK" });
circ(s, 9.3, 0.7, 3.6, C.accent1, "빛-외곽", 85); circ(s, 9.85, 1.25, 2.5, C.accent1, "빛-중간", 70); circ(s, 10.4, 1.8, 1.4, C.accent1, "빛-중심", 0);
T(s, "문득클럽 · 파일럿 프로그램 기획안", { x: 0.8, y: 1.0, w: 8, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
s.addText("벨라스케스 보러 가기 전, 책으로 친해지기", { placeholder: "title" });
T(s, "《스페인 미술 500년: 빛과 어둠의 연대기》", { x: 0.8, y: 3.85, w: 9, h: 0.5, fontSize: 22, color: C.background1 });
[["1", "12/16 (수)", "책 읽기"], ["2", "12/30 (수)", "책 읽기"], ["3", "1/13 (수) 저녁", "전시 + 해설"]].forEach((w, i) => {
  const x = 0.8 + i * 4.0;
  circ(s, x, 5.0, 0.95, i === 2 ? C.accent2 : C.accent1, "회차-원" + (i + 1));
  T(s, w[0], { x, y: 5.0, w: 0.95, h: 0.95, fontSize: 30, bold: true, color: i === 2 ? C.background1 : C.text1, align: "center", valign: "middle" });
  T(s, [{ text: w[1], options: { bold: true, fontSize: 20, color: C.background1, breakLine: true } }, { text: w[2], options: { fontSize: 16, color: C.background2 } }], { x: x + 1.15, y: 5.05, w: 2.7, h: 0.85, valign: "middle" });
});
T(s, "서초 · 참가자 10명 · 파일럿 특가 76,000원", { x: 0.8, y: 6.5, w: 9, h: 0.3, fontSize: 14, color: C.background2 });
s.addNotes("표지. 1~2회차는 책, 3회차는 각자 전시를 본 뒤 저녁에 모여 해설합니다. 장소는 서울 서초구 남부순환로315길 85(무료 대관, 음식 반입 가능). 전시 기간(2026.9.22~2027.1.20)은 예술의전당 공식 안내로 최종 확인하세요.");

// ===== 2 커리큘럼
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("3회 커리큘럼: 책 2회, 그리고 전시와 나눔", { placeholder: "title" });
card(s, 0.6, 1.35, 12.15, 0.85, C.text1, "목표-배경");
T(s, "모임 목표", { x: 0.9, y: 1.35, w: 1.5, h: 0.85, fontSize: 16, bold: true, color: C.accent1, valign: "middle" });
T(s, "3회 과정을 지나면, 스페인 미술에 대해 설명할 수 있어요.", { x: 2.5, y: 1.35, w: 10, h: 0.85, fontSize: 22, bold: true, color: C.background1, valign: "middle" });
const cur = [
  ["1", "12/16 (수) · 책", "화가 세 명 얼굴부터 익히기", "《나는 스페인 미술관의 도슨트입니다》 이안", "발제 · 그림 한 장 읽기 · 화가 구분 게임", "키워드 카드 3장", false],
  ["2", "12/30 (수) · 책", "가톨릭과 이슬람, 그림에 남은 흔적", "《스페인, 역사와 미술을 보다》 김영욱", "맥락 토론 · 시대 지도 · 관람 위시리스트", "위시 3점 · 티켓 받기", false],
  ["3", "1/13 (수) 저녁 · 전시", "각자 보고, 저녁에 모여 3분 해설", "관람 티켓 제공 · 큐레이터 게스트(섭외 중)", "큐레이터 토크 · 간식 · 해설 발표", "나의 한 점 3분 해설", true]];
cur.forEach((c, i) => {
  const x = 0.6 + i * 4.125, hi = c[6];
  card(s, x, 2.45, 3.9, 4.25, hi ? C.text1 : C.background2, "회차카드-" + c[0]);
  circ(s, x + 0.3, 2.7, 0.75, hi ? C.accent1 : C.accent2, "회차원-" + c[0]);
  T(s, c[0], { x: x + 0.3, y: 2.7, w: 0.75, h: 0.75, fontSize: 26, bold: true, color: hi ? C.text1 : C.background1, align: "center", valign: "middle" });
  T(s, c[1], { x: x + 1.25, y: 2.7, w: 2.5, h: 0.75, fontSize: 16, bold: true, color: hi ? C.accent1 : C.accent2, valign: "middle" });
  T(s, c[2], { x: x + 0.3, y: 3.65, w: 3.35, h: 0.95, fontSize: 22, bold: true, color: hi ? C.background1 : C.text1 });
  T(s, c[3], { x: x + 0.3, y: 4.65, w: 3.35, h: 0.55, fontSize: 14, color: hi ? C.background2 : C.text2 });
  T(s, c[4], { x: x + 0.3, y: 5.3, w: 3.35, h: 0.65, fontSize: 18, color: hi ? C.background1 : C.text1 });
  T(s, [{ text: "산출물  ", options: { bold: true, color: hi ? C.accent1 : C.accent2 } }, { text: c[5], options: { color: hi ? C.background1 : C.text1 } }], { x: x + 0.3, y: 6.1, w: 3.4, h: 0.4, fontSize: 16 });
});
s.addNotes("목표 설계: 3회(약 8시간)로 ‘준전문가’까지는 어렵고, ‘처음 보는 사람 앞에서 그림 한 점을 3분 해설’이 현실적인 도달점입니다. 1회차 시작과 3회차 끝에 같은 자기평가 3문항(화가 구분, 시대 순서 설명, 한 점 3분 해설)을 받아 전후를 비교하세요. 2회차에서 관람 티켓과 미션 카드를 나눠주고, 전시는 12/30 이후 1/13 낮까지 각자 봅니다. 심화 선택 도서: 《프라도 미술관에서 꼭 봐야 할 그림 100》(김영숙).");

// ===== 3 시간표
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("회차별 시간표: 1·2회차 180분, 3회차 150분", { placeholder: "title" });
const tt = [
  ["1회차 · 12/16 (수)", "화가 세 명 얼굴부터", [[15, "체크인 · 자기평가"], [50, "화가별 발제"], [10, "휴식"], [45, "그림 한 장 읽기"], [30, "화가 구분 게임"], [30, "카드 작성 · 정리"]], false],
  ["2회차 · 12/30 (수)", "가톨릭과 이슬람", [[15, "체크인 · 카드 공유"], [45, "맥락 토론"], [10, "휴식"], [40, "시대 구분 지도"], [40, "관람 위시리스트"], [30, "티켓·미션 안내"]], false],
  ["3회차 · 1/13 (수) 저녁", "큐레이터 토크와 해설", [[15, "체크인 · 간식"], [50, "큐레이터 토크·질문"], [10, "휴식"], [25, "3분 해설 초안"], [40, "해설 발표"], [10, "자기평가 · 설문"]], true]];
tt.forEach((c, i) => {
  const x = 0.6 + i * 4.125, hi = c[3];
  card(s, x, 1.4, 3.9, 5.0, hi ? C.text1 : C.background2, "시간표-패널-" + i);
  T(s, c[0], { x: x + 0.3, y: 1.55, w: 3.4, h: 0.3, fontSize: 16, bold: true, color: hi ? C.accent1 : C.accent2 });
  T(s, c[1], { x: x + 0.3, y: 1.9, w: 3.4, h: 0.45, fontSize: 22, bold: true, color: hi ? C.background1 : C.text1 });
  c[2].forEach((r, j) => {
    const y = 2.6 + j * 0.62;
    card(s, x + 0.3, y, 0.85, 0.42, hi ? C.accent1 : C.text1, `분칩-${i}-${j}`, 0.21);
    T(s, r[0] + "분", { x: x + 0.3, y, w: 0.85, h: 0.42, fontSize: 16, bold: true, color: hi ? C.text1 : C.background1, align: "center", valign: "middle" });
    T(s, r[1], { x: x + 1.35, y, w: 2.4, h: 0.42, fontSize: 16, color: hi ? C.background1 : C.text1, valign: "middle" });
  });
});
T(s, "큐레이터를 못 모시면 3회차는 다음 장의 B안으로 진행합니다.", { x: 0.6, y: 6.5, w: 12, h: 0.3, fontSize: 14, color: C.text2 });
s.addNotes("1회차: 체크인 15분과 자기평가 3문항(화가 구분 / 시대 순서 설명 / 한 점 3분 해설), 화가별 발제 50분(책 구성에 맞춰 세 명 확정), 휴식 10분, 그림 한 장 천천히 읽기 45분(보이는 것부터 말하기), 화가 구분 게임 30분(도판만 보고 맞히기), 키워드 카드 3장 작성과 정리 30분(2회차 책 범위 안내).\n2회차: 체크인과 카드 공유 15분, 맥락 토론 45분, 휴식 10분, 시대 구분 지도 40분(연표에 화가와 사건 붙이기), 관람 위시리스트 40분(전시에서 꼭 볼 3점), 관람 티켓·미션 카드 배부와 안내 정리 30분.\n3회차: 체크인·간식 15분, 큐레이터 토크 35분과 질문 15분, 휴식 10분, 3분 해설 초안 25분(장면 → 빛과 어둠 → 내 한 줄), 간식 먹으며 해설 발표 40분(1인 3분 + 질문 1분, 10명 기준), 자기평가 재측정·만족도·제안 설문 10분.");

// ===== 4 3회차
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("3회차: 관람은 각자, 저녁에는 같이", { placeholder: "title" });
[["12/30", "티켓과 미션 카드 받기"], ["~1/13", "수요일 낮이나 미리 관람"], ["1/13 저녁", "같은 공간에 모이기"]].forEach((f, i) => {
  const x = 0.6 + i * 4.125;
  card(s, x, 1.4, 3.9, 1.35, i === 2 ? C.text1 : C.background2, "흐름카드-" + i);
  T(s, f[0], { x: x + 0.3, y: 1.55, w: 3.3, h: 0.4, fontSize: 18, bold: true, color: i === 2 ? C.accent1 : C.accent2 });
  T(s, f[1], { x: x + 0.3, y: 2.0, w: 3.4, h: 0.6, fontSize: 20, bold: true, color: i === 2 ? C.background1 : C.text1 });
  if (i < 2) T(s, "→", { x: x + 3.9, y: 1.7, w: 0.225, h: 0.6, fontSize: 20, bold: true, color: C.text2, align: "center", valign: "middle" });
});
const ab = [["큐레이터가 오시면 (A안)", ["기획 이야기와 질문 50분", "3분 해설 초안 25분", "간식 먹으며 해설 발표 40분"], false],
            ["큐레이터를 못 모시면 (B안)", ["미션 카드 후기 라운드 30분", "해설 초안과 짝 연습 30분", "해설 발표 50분 (1인 3분 + 질문 2분)"], true]];
ab.forEach((a, i) => {
  const x = 0.6 + i * 6.15;
  card(s, x, 3.0, 6.0, 2.75, a[2] ? C.accent2 : C.text1, "안-패널-" + i);
  T(s, a[0], { x: x + 0.35, y: 3.2, w: 5.3, h: 0.45, fontSize: 22, bold: true, color: C.background1 });
  T(s, a[1].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < a[1].length - 1 } })), { x: x + 0.35, y: 3.85, w: 5.3, h: 1.7, fontSize: 18, color: C.background1, paraSpaceAfter: 8 });
});
T(s, "간식 후보", { x: 0.6, y: 6.05, w: 1.5, h: 0.5, fontSize: 16, bold: true, color: C.accent2, valign: "middle" });
["추로스·초콜라테", "토르티야", "하몽·치즈 타파스", "무알코올 상그리아"].forEach((m, i) => {
  const x = 2.1 + i * 2.65;
  card(s, x, 6.05, 2.5, 0.5, C.accent1, "간식칩-" + i, 0.25);
  T(s, m, { x, y: 6.05, w: 2.5, h: 0.5, fontSize: 16, bold: true, color: C.text1, align: "center", valign: "middle" });
});
s.addNotes("관람은 12/30(2회차) 이후 1/13 당일 낮까지 각자 합니다. 관람 티켓은 모집 마감 후 확정 인원만큼 구매해 2회차에 나눠줍니다. 미션 카드는 ① 나의 한 점 ② 빛과 어둠 찾기 ③ 그림 속 음식 찾기이며, 관람 후 사진을 보내게 합니다. A안: 체크인·간식 15 / 큐레이터 토크 35 + 질문 15 / 휴식 10 / 초안 25 / 발표 40 / 설문 10. B안: 체크인·간식 15 / 후기 라운드 30 / 휴식 10 / 초안·짝 연습 30 / 발표 50 / 설문·정리 15. 큐레이터는 12/30까지 확정하고, 못 모시면 B안으로 갑니다. 간식 메뉴는 관람 중 찾은 ‘그림 속 음식’에서 확정하고 알레르기는 2회차에 조사합니다. 발표 때 그림을 보여줄 화면과 전시장 사진 촬영 규정은 확인이 필요합니다.");

// ===== 5 시장 비교
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("시장 비교: 시즌 17만~40만 원, 원데이 1.8만~8만 원", { placeholder: "title" });
const X0 = 1.2, Y0 = 5.6, PX = 0.36, yv = (v) => Y0 - PX * v;
[5, 10].forEach((v) => line(s, X0, yv(v), 6.8, 0, C.background2, 1, "눈금선-" + v));
line(s, X0, 1.95, 0, Y0 - 1.95, C.text2, 1.25, "세로축"); line(s, X0, Y0, 6.8, 0, C.text2, 1.25, "가로축");
[[0, "0"], [5, "5만"], [10, "10만"]].forEach((r) => T(s, r[1], { x: 0.55, y: yv(r[0]) - 0.15, w: 0.55, h: 0.3, fontSize: 12, color: C.text2, align: "right", valign: "middle" }));
T(s, "회당 가격", { x: 0.6, y: 1.55, w: 1.6, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
[["1회", 2.1], ["4주", 3.5], ["3개월", 5.7], ["4개월", 7.3]].forEach((r) => T(s, r[0], { x: r[1] - 0.5, y: Y0 + 0.08, w: 1.0, h: 0.3, fontSize: 14, color: C.text2, align: "center" }));
T(s, "약정 기간 · 몰입 깊이 →", { x: 4.5, y: Y0 + 0.45, w: 3.5, h: 0.3, fontSize: 14, bold: true, color: C.text2, align: "right" });
const cap = (cx, lo, hi, n) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx - 0.25, y: yv(hi), w: 0.5, h: Math.max(yv(lo) - yv(hi), 0.3), rectRadius: 0.25, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "범위-" + n });
cap(1.7, 1.8, 8, "원데이"); cap(2.5, 1.4, 2.3, "전시"); cap(5.7, 4.3, 5.8, "정규"); cap(7.3, 6.3, 10, "트레바리");
const lab = (n, sub, x, y, w, al, col) => T(s, [{ text: n, options: { bold: true, fontSize: 16, color: col || C.text1, breakLine: true } }, { text: sub, options: { fontSize: 14, color: col || C.text2 } }], { x, y, w, h: 0.55, align: al });
lab("넷플연가 원데이", "1.8–8만", 1.3, 2.1, 2.0, "left"); T(s, [{ text: "전시 단독 관람 ", options: { bold: true, fontSize: 16 } }, { text: "1.4–2.3만", options: { fontSize: 14, color: C.text2 } }], { x: 1.95, y: 5.13, w: 3.0, h: 0.32 });
lab("넷플연가 정규", "4.3–5.8만", 3.7, 2.9, 1.85, "right"); lab("트레바리", "6.3–10만", 5.3, 2.05, 1.6, "right");
circ(s, 3.1, yv(2.5) - 0.4, 0.8, C.accent2, "가격대-고리", 80); circ(s, 3.2, yv(2.5) - 0.3, 0.6, C.accent2, "우리-점");
lab("문득클럽 (우리)", "약 2.5만", 4.1, yv(2.5) - 0.28, 1.9, "left", C.accent2);
T(s, "가격은 제공 자료 기준입니다. 공식 안내로 다시 확인하세요.", { x: 0.6, y: 6.4, w: 7.4, h: 0.3, fontSize: 12, color: C.text2 });
card(s, 8.3, 1.45, 4.45, 5.0, C.text1, "우리모임-배경");
T(s, "우리 모임", { x: 8.6, y: 1.65, w: 3.9, h: 0.5, fontSize: 22, bold: true, color: C.background1 });
[["가격", "3회 76,000원, 파일럿 특가"], ["기간", "격주 수요일, 3회"], ["대상", "서초 인근 직장인 중 미술이 처음이거나, 혼자 보면 남는 게 적은 사람"]].forEach((r, i) => {
  const y = 2.45 + i * 1.3;
  T(s, r[0], { x: 8.6, y, w: 3.9, h: 0.3, fontSize: 16, bold: true, color: C.accent1 });
  T(s, r[1], { x: 8.6, y: y + 0.38, w: 3.9, h: 0.85, fontSize: 18, color: C.background1 });
});
s.addNotes("포지셔닝 맵: 가로축은 약정 기간·몰입 깊이, 세로축은 회당 가격(막대는 범위)입니다. 우리는 3회 76,000원이라 회당 약 2.5만 원이고, 넷플연가 정규 하단(17만 원)보다 낮게 잡았습니다. 트레바리·넷플연가 가격은 사용자 제공 자료 기준입니다. 4주 눈금은 12/16~1/13이 4주에 걸치기 때문입니다.");

// ===== 6 참가비
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("참가비: 파일럿 특가 76,000원, 10명까지", { placeholder: "title" });
card(s, 0.6, 1.45, 6.4, 3.6, C.accent2, "가격-배경");
T(s, "파일럿 특가", { x: 1.0, y: 1.75, w: 5.6, h: 0.4, fontSize: 18, bold: true, color: C.background1 });
T(s, "76,000원", { x: 1.0, y: 2.3, w: 5.6, h: 1.2, fontSize: 64, bold: true, color: C.background1 });
T(s, "3회 · 회당 약 2.5만 원", { x: 1.0, y: 3.75, w: 5.6, h: 0.4, fontSize: 20, color: C.background1 });
T(s, "한·스페인 수교 76주년 전시라서 76,000원 [수교 연도 확인 후 사용]", { x: 1.0, y: 4.3, w: 5.6, h: 0.5, fontSize: 14, color: C.background1 });
[["포함", "관람 티켓 · 간식 · 자료"], ["별도", "책값 (2권)"], ["정원", "참가자 10명"], ["환불", "1회차 시작 전까지 전액"]].forEach((r, i) => {
  const y = 1.45 + i * 0.9;
  card(s, 7.3, y, 5.45, 0.78, C.background2, "조건-배경-" + i);
  T(s, r[0], { x: 7.6, y, w: 1.0, h: 0.78, fontSize: 18, bold: true, color: C.accent2, valign: "middle" });
  T(s, r[1], { x: 8.7, y, w: 3.9, h: 0.78, fontSize: 20, valign: "middle" });
});
card(s, 0.6, 5.35, 12.15, 1.15, C.text1, "수지-배경");
T(s, [{ text: "10명이면 약 18만 원이 남고, ", options: { color: C.background1 } }, { text: "6명부터 손익분기", options: { color: C.accent1, bold: true } }, { text: "입니다", options: { color: C.background1 } }], { x: 0.95, y: 5.35, w: 11.5, h: 0.75, fontSize: 24, valign: "middle" });
T(s, "비용은 가정값입니다 (관람 티켓 23만, 간식 12만, 자료 3만, 외부 강사 20만)", { x: 0.95, y: 6.05, w: 11.5, h: 0.35, fontSize: 14, color: C.background2 });
s.addNotes("수입 76,000원 × 10명 = 760,000원, 비용 580,000원(관람 티켓 23만, 간식 12만, 자료 3만, 외부 강사 20만, 모두 가정), 남는 돈 180,000원은 진행자 보상과 예비비입니다. 1인당 남는 돈 4.1만 원, 고정비 23만 원이라 손익분기는 6명입니다. 관람 티켓은 신청 마감 후 확정 인원만큼 구매하며, 날짜 지정권이면 노쇼 손실이 생기므로 전시장에 기간권·단체권이 있는지 확인하세요. 장소는 서울 서초구 남부순환로315길 85, 무료 대관, 음식 반입·취식 가능, 참가자 10명(진행자·큐레이터 별도).");

// ===== 7 모집
s = pres.addSlide({ masterName: "LIGHT" });
s.addText("모집 계획: 12월 9일까지 6~10명", { placeholder: "title" });
[["11/25 (수)", "모집 오픈"], ["12/2 (수)", "마감 임박 알림"], ["12/9 (수)", "마감 · 티켓 구매"], ["12/16 (수)", "1회차 시작"]].forEach((d, i) => {
  const x = 0.6 + i * 3.075, hi = i === 3;
  card(s, x, 1.4, 2.95, 1.25, hi ? C.text1 : C.background2, "일정카드-" + i);
  T(s, d[0], { x: x + 0.3, y: 1.55, w: 2.5, h: 0.5, fontSize: 24, bold: true, color: hi ? C.accent1 : C.accent2 });
  T(s, d[1], { x: x + 0.3, y: 2.1, w: 2.5, h: 0.4, fontSize: 18, color: hi ? C.background1 : C.text1 });
});
card(s, 0.6, 2.95, 7.1, 3.55, C.background2, "채널-배경");
T(s, [{ text: "어디서 모으나", options: { bold: true, fontSize: 22 } }, { text: "   인원은 가정", options: { fontSize: 14, color: C.text2 } }], { x: 0.95, y: 3.1, w: 6.4, h: 0.45 });
[["당근", "5명"], ["미술 커뮤니티", "2명"], ["인근 카페 · 서점 · 지인", "2명"], ["인스타 · 문토 · 소셜링", "1명"]].forEach((r, i) => {
  const y = 3.75 + i * 0.68;
  T(s, r[0], { x: 0.95, y, w: 4.8, h: 0.5, fontSize: 20, bold: true, color: C.accent2, valign: "middle" });
  card(s, 6.5, y + 0.05, 0.95, 0.4, C.text1, "인원칩-" + i, 0.2);
  T(s, r[1], { x: 6.5, y: y + 0.05, w: 0.95, h: 0.4, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
});
card(s, 7.85, 2.95, 4.9, 3.55, C.text1, "모집글-배경");
T(s, "모집글에 넣을 것", { x: 8.15, y: 3.1, w: 4.3, h: 0.45, fontSize: 22, bold: true, color: C.background1 });
T(s, ["미술이 처음이어도 신청할 수 있습니다", "격주 수요일 3회, 전시는 각자 관람", "76,000원, 10명 한정, 관람 티켓 포함", "진행자 소개와 환불 규정 공개"].map((t, j, a) => ({ text: t, options: { bullet: true, breakLine: j < a.length - 1 } })),
  { x: 8.15, y: 3.7, w: 4.4, h: 2.6, fontSize: 18, color: C.background1, paraSpaceAfter: 8 });
s.addNotes("기존 회원이 없고 개인 SNS도 작아서 당근이 주 채널입니다. 당근은 거주 동네 기준이라 서초에서 일해도 다른 동네에 살면 게시글을 못 볼 수 있고, 게시 범위와 광고비는 확인하지 못했습니다. 채널 설명: 당근은 동네생활·모임에 서초 인근으로 게시(11/25 오픈, 12/2 마감 알림), 미술 커뮤니티는 전시 관련 카페·오픈채팅, 인근 카페·서점에는 포스터 부탁과 지인 직접 연락, 인스타는 클럽 계정(wakeon.club) 게시와 문토·소셜링 보조(수수료 20%, 사용자 제공 자료 기준). 신청 폼 문항: 이름·연락처·알레르기·1/13 저녁 참석 가능 여부. 문의 25~30명에서 신청 10명(전환 3분의 1)을 가정하고 12/2에 채널을 조정합니다. 1회차까지 책 읽을 시간이 1주일이라 1회차는 앞부분 범위를 정해 안내하세요. 12/30은 연말이라 참석률이 떨어질 수 있습니다.");

(async () => {
  await pres.writeFile({ fileName: "readon_spain_art_club_v10_raw.pptx" });
  await applyTheme("readon_spain_art_club_v10_raw.pptx", THEME);
  console.log("done");
})();
