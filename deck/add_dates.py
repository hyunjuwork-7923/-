# v2 → v3: 1~3번 슬라이드에 날짜 추가 (기존 글자는 수정하지 않고 새 텍스트 상자만 추가)
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.oxml.ns import qn
from lxml import etree
FONT = "Pretendard"
prs = Presentation(sys.argv[1])
def tb(slide, x, y, w, h, t, size, color, align, name):
    b = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); b.name = name
    tf = b.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = t
    f = r.font; f.size = Pt(size); f.bold = True; f.name = FONT; f.color.theme_color = color
    rPr = r._r.get_or_add_rPr(); latin = rPr.find(qn("a:latin"))
    for i, tag in enumerate(("a:ea", "a:cs")):
        e = etree.Element(qn(tag)); e.set("typeface", FONT); latin.addnext(e) if i == 0 else rPr.find(qn("a:ea")).addnext(e)
dates = ["11/18", "12/2", "12/19", "12/30"]
# 1번: 표지 4개 원 위
s1 = prs.slides[0]
for i, d in enumerate(dates):
    label = d + (" (토)" if i == 2 else " (수)")
    tb(s1, 0.8 + i * 3.0, 5.1, 2.8, 0.3, label, 14, TC.ACCENT_1, PP_ALIGN.LEFT, f"표지-날짜-{i}")
# 2번: 카드 번호 원 아래
s2 = prs.slides[1]
for i, d in enumerate(dates):
    col, row = i % 2, i // 2
    x = 0.6 + col * 6.15; y = 2.2 + row * 2.3
    tb(s2, x + 0.12, y + 0.88, 0.95, 0.28, d, 14, TC.ACCENT_1 if i == 2 else TC.ACCENT_2, PP_ALIGN.CENTER, f"커리큘럼-날짜-{i}")
# 3번: 패널 제목 오른쪽
s3 = prs.slides[2]
tb(s3, 4.0, 1.55, 2.3, 0.45, "12/19 (토)", 14, TC.ACCENT_2, PP_ALIGN.RIGHT, "3주차-날짜")
tb(s3, 11.0, 1.55, 1.45, 0.45, "12/30 (수)", 14, TC.ACCENT_1, PP_ALIGN.RIGHT, "4주차-날짜")
# 4번: 대상 문구 (요청에 따라 수정)
for sh in prs.slides[3].shapes:
    if sh.has_text_frame and "리드온 회원" in sh.text_frame.text:
        sh.text_frame.paragraphs[0].runs[0].text = "서초 인근 직장인 중 미술은 처음이거나, "
prs.save(sys.argv[2]); print("saved")
