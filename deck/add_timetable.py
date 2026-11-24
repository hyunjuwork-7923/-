# v6 → v7: 2번 뒤에 "주차별 시간표" 슬라이드 추가 (사용자 제공 시간표 기준, 기존 슬라이드 글자는 수정하지 않음)
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.oxml.ns import qn
from lxml import etree
FONT = "Pretendard"
prs = Presentation(sys.argv[1])
layout = [l for l in prs.slide_layouts if l.name == "LIGHT"][0]
s = prs.slides.add_slide(layout)

def setfont(run, size, bold=False, color=TC.TEXT_1):
    f = run.font; f.size = Pt(size); f.bold = bold; f.name = FONT; f.color.theme_color = color
    rPr = run._r.get_or_add_rPr(); latin = rPr.find(qn("a:latin"))
    ea = etree.Element(qn("a:ea")); ea.set("typeface", FONT); latin.addnext(ea)
    cs = etree.Element(qn("a:cs")); cs.set("typeface", FONT); ea.addnext(cs)
def box(x, y, w, h, fill, name, radius=0.1):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)); sh.name = name
    sh.adjustments[0] = min(0.5, radius / min(w, h)); sh.fill.solid(); sh.fill.fore_color.theme_color = fill
    sh.line.fill.background(); sh.shadow.inherit = False; return sh
def text(x, y, w, h, runs, name, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align
    for (t, size, bold, color) in runs:
        r = p.add_run(); r.text = t; setfont(r, size, bold, color)

title = [p for p in s.placeholders if p.placeholder_format.type == 1][0]
title.text_frame.text = "주차별 시간표: 한 번에 180분"
for r in title.text_frame.paragraphs[0].runs: setfont(r, 30, True)
for sh in prs.slides[4].shapes:
    if sh.is_placeholder and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
        s.shapes._spTree.append(copy.deepcopy(sh._element))

cols = [
 ("1주차 · 11/18 (수)", "화가 세 명 얼굴부터 익히기", [
   (15, "체크인 · 자기평가 3문항"), (50, "화가별 발제 (책 구성에 맞춰 세 명 확정)"), (10, "휴식"),
   (45, "그림 한 장 천천히 읽기 (보이는 것부터 말하기)"), (30, "화가 구분 게임 (도판만 보고 맞히기)"),
   (15, "키워드 카드 3장 작성"), (15, "정리 · 2주차 책 범위 안내")], False),
 ("2주차 · 12/2 (수)", "가톨릭과 이슬람, 그림에 남은 흔적", [
   (15, "체크인 · 키워드 카드 공유"), (45, "가톨릭·이슬람 맥락 토론"), (10, "휴식"),
   (40, "시대 구분 지도 (연표에 화가·사건 붙이기)"), (40, "관람 위시리스트 (전시에서 꼭 볼 3점)"),
   (20, "3주차 미션 카드 안내 · 동선 공유"), (10, "정리")], False),
 ("3주차 · 12/19 (토)", "전시 관람 + 한 점 3분 해설", [
   (10, "전시장 집결 · 오리엔테이션"), (35, "도슨트 투어 또는 오디오가이드"), (50, "자유 관람 + 미션 카드"),
   (10, "모임 장소로 이동 · 휴식"), (25, "3분 해설 초안 (장면 → 빛과 어둠 → 내 한 줄)"),
   (40, "간식 먹으며 해설 발표 (1인 3분 + 질문 1분)"), (10, "자기평가 재측정 · 만족도·제안 설문")], True)]
for ci, (tag, ttl, rows, hi) in enumerate(cols):
    x = 0.6 + ci * 4.1
    box(x, 1.35, 3.95, 5.15, TC.TEXT_1 if hi else TC.BACKGROUND_2, f"주차패널-{ci}", 0.12)
    text(x + 0.25, 1.5, 3.5, 0.3, [(tag, 14, True, TC.ACCENT_1 if hi else TC.ACCENT_2)], f"주차태그-{ci}")
    text(x + 0.25, 1.82, 3.5, 0.6, [(ttl, 18, True, TC.BACKGROUND_1 if hi else TC.TEXT_1)], f"주차제목-{ci}")
    for ri, (mins, desc) in enumerate(rows):
        y = 2.5 + ri * 0.57
        rest = desc.startswith("휴식") or desc.startswith("모임 장소로")
        box(x + 0.25, y + 0.07, 0.72, 0.36, TC.ACCENT_1 if hi else TC.TEXT_1, f"분칩-{ci}-{ri}", 0.18)
        text(x + 0.25, y + 0.07, 0.72, 0.36, [(f"{mins}분", 14, True, TC.TEXT_1 if hi else TC.BACKGROUND_1)], f"분-{ci}-{ri}", PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        text(x + 1.12, y, 2.7, 0.5, [(desc, 14, False, TC.BACKGROUND_1 if hi else TC.TEXT_1)], f"내용-{ci}-{ri}", anchor=MSO_ANCHOR.MIDDLE)
text(0.6, 6.58, 12.1, 0.28, [("집결 시각과 도슨트는 12/19 전시장 시간표를 확인한 뒤 확정합니다. 4주차 시간표는 아직 없습니다.", 12, False, TC.TEXT_2)], "시간표-메모")

if True:
    ns_ = s.notes_slide
    if ns_.notes_text_frame is None:
        for sh in prs.slides[4].notes_slide.shapes:
            if sh.is_placeholder and "BODY" in str(sh.placeholder_format.type):
                ns_.shapes._spTree.append(copy.deepcopy(sh._element))
    ns_.notes_text_frame.text = ("1주차 15분: 체크인, 자기평가 3문항(화가 구분 / 시대 순서 설명 / 한 점 3분 해설). 50분: 화가별 발제, 책 구성에 맞춰 세 명 확정. 45분: 그림 한 장 천천히 읽기, 한 점을 두고 보이는 것부터 말하기. 30분: 화가 구분 게임. "
        "2주차 40분: 시대 구분 지도 만들기, 연표에 화가와 사건 붙이기. 40분: 관람 위시리스트, 도판을 보며 전시에서 꼭 볼 3점 고르기. 20분: 3주차 미션 카드 안내와 동선 공유. "
        "3주차 25분: 3분 해설 초안 쓰기, 구성은 장면 → 빛과 어둠 → 내 한 줄. 40분: 스페인 간식을 먹으며 해설 발표, 1인당 3분 해설과 1분 질문(10명 기준). 마지막 10분: 자기평가 재측정, 만족도 조사, 다음 기수 제안은 설문 문항 하나로 받습니다. "
        "주의: 3번 슬라이드의 ‘3주차 관람 설계 120분’은 이 시간표(180분)와 다릅니다.")

pres = prs.part._element
lst = pres.find(qn("p:sldIdLst"))
new = lst[-1]; lst.remove(new); lst.insert(2, new)
nid = new.get("id"); second = lst[1].get("id")
ns14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
for sec in pres.iter("{%s}section" % ns14):
    ids = sec.find("{%s}sldIdLst" % ns14)
    sl = [e.get("id") for e in ids]
    if second in sl:
        e = etree.Element("{%s}sldId" % ns14); e.set("id", nid)
        ids.insert(sl.index(second) + 1, e)
prs.save(sys.argv[2]); print("saved")
