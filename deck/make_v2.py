# 사용자가 수정한 PPT(원본)를 기준으로: 1) 서체 통일(테마 ea/Hang → Pretendard) 2) 6번 모집 슬라이드 추가
import sys, zipfile, re, copy, shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.oxml.ns import qn
from lxml import etree

SRC, TMP, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
FONT = "Pretendard"

# 1) 테마 서체: ea/cs/Hang 스크립트까지 Pretendard로
zin = zipfile.ZipFile(SRC); zout = zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    data = zin.read(it.filename)
    if re.match(r"ppt/theme/theme\d+\.xml", it.filename):
        t = data.decode("utf8")
        t = re.sub(r'<a:ea typeface="[^"]*"/>', f'<a:ea typeface="{FONT}"/>', t)
        t = re.sub(r'<a:cs typeface="[^"]*"/>', f'<a:cs typeface="{FONT}"/>', t)
        t = re.sub(r'(<a:font script="Hang" typeface=")[^"]*(")', rf'\g<1>{FONT}\2', t)
        t = re.sub(r'<a:latin typeface="[^"]*"/>', f'<a:latin typeface="{FONT}"/>', t)
        data = t.encode("utf8")
    zout.writestr(it, data)
zout.close()

prs = Presentation(TMP)
layout = [l for l in prs.slide_layouts if l.name == "LIGHT"][0]
s = prs.slides.add_slide(layout)

def setfont(run, size, bold=False, color=TC.TEXT_1):
    f = run.font; f.size = Pt(size); f.bold = bold; f.name = FONT; f.color.theme_color = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = etree.SubElement(rPr, qn(tag))
        e.set("typeface", FONT)
    # 순서 보정: latin, ea, cs 가 연속되도록
    latin = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea")); cs = rPr.find(qn("a:cs"))
    for e in (ea, cs):
        rPr.remove(e)
    idx = list(rPr).index(latin) + 1
    rPr.insert(idx, ea); rPr.insert(idx + 1, cs)

def box(x, y, w, h, fill, name, radius=0.1, line=False):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.name = name
    sh.adjustments[0] = min(0.5, radius / min(w, h))
    sh.fill.solid(); sh.fill.fore_color.theme_color = fill
    sh.line.fill.background(); sh.shadow.inherit = False
    return sh

def text(x, y, w, h, paras, name, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=0):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    first = True
    for runs in paras:
        para = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        para.alignment = align
        if space: para.space_after = Pt(space)
        for (t, size, bold, color) in runs:
            r = para.add_run(); r.text = t; setfont(r, size, bold, color)
    return tb

# 제목 (레이아웃 placeholder)
title = [p for p in s.placeholders if p.placeholder_format.type == 1][0]
title.text_frame.text = "모집 계획: 11월 13일까지 8~10명"
for r in title.text_frame.paragraphs[0].runs: setfont(r, 30, True)

# 슬라이드 번호 (5번 슬라이드의 번호 도형 복제)
for sh in prs.slides[4].shapes:
    if sh.is_placeholder and sh.placeholder_format.type is not None and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
        s.shapes._spTree.append(copy.deepcopy(sh._element))

# 상단: 모집 일정 4칸
steps = [("11/6 (금)", "모집 오픈", "얼리버드 129,000원"), ("11/10 (화)", "얼리버드 마감", "이후 149,000원"),
         ("11/13 (금)", "정가 마감", "최소 8명 확인 · 확정 안내"), ("11/18 (수)", "1회차, 이후 격주", "12/2 · 12/19(토) · 12/30")]
for i, (d, a, b) in enumerate(steps):
    x = 0.6 + i * 3.075
    hi = i == 3
    box(x, 1.4, 2.95, 1.3, TC.TEXT_1 if hi else TC.BACKGROUND_2, f"일정카드-{i}", 0.12)
    c1 = TC.ACCENT_1 if hi else TC.ACCENT_2
    c2 = TC.BACKGROUND_1 if hi else TC.TEXT_1
    text(x + 0.25, 1.52, 2.5, 0.35, [[(d, 18, True, c1)]], f"일정날짜-{i}")
    text(x + 0.25, 1.92, 2.5, 0.3, [[(a, 16, True, c2)]], f"일정제목-{i}")
    text(x + 0.25, 2.27, 2.5, 0.3, [[(b, 14, False, c2)]], f"일정설명-{i}")

# 좌: 채널별 모집
box(0.6, 2.95, 7.1, 3.55, TC.BACKGROUND_2, "채널-배경", 0.12)
text(0.9, 3.08, 6.5, 0.4, [[("어디서 모으나", 18, True, TC.TEXT_1), ("   목표 인원은 가정", 12, False, TC.TEXT_2)]], "채널-제목")
rows = [("리드온 기존 회원", "단톡·뉴스레터 공지, 11/6 첫 안내", "4명"),
        ("회원의 지인", "한 명씩 데려오기 요청", "2명"),
        ("인스타그램", "게시 2회: 11/6 오픈, 11/10 마감 알림", "2명"),
        ("모집 플랫폼", "문토·소셜링 등, 수수료 20% 별도(자료 기준)", "2명")]
for i, (a, b, n) in enumerate(rows):
    y = 3.62 + i * 0.7
    text(0.9, y, 2.1, 0.55, [[(a, 14, True, TC.ACCENT_2)]], f"채널명-{i}", anchor=MSO_ANCHOR.MIDDLE)
    text(3.0, y, 3.6, 0.55, [[(b, 14, False, TC.TEXT_1)]], f"채널설명-{i}", anchor=MSO_ANCHOR.MIDDLE)
    box(6.7, y + 0.08, 0.8, 0.4, TC.TEXT_1, f"채널인원칩-{i}", 0.2)
    text(6.7, y + 0.08, 0.8, 0.4, [[(n, 14, True, TC.BACKGROUND_1)]], f"채널인원-{i}", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 우: 모집글에 넣을 것
box(7.85, 2.95, 4.9, 3.55, TC.TEXT_1, "모집글-배경", 0.12)
text(8.15, 3.08, 4.3, 0.4, [[("모집글과 신청 폼", 18, True, TC.BACKGROUND_1)]], "모집글-제목")
items = [("첫 줄", "미술을 처음 보는 사람도 신청할 수 있습니다"),
         ("일정", "격주 수요일, 3회차만 12/19(토) 전시장"),
         ("가격", "129,000원, 11/10 이후 149,000원 (관람료·다과 포함)"),
         ("정원", "10명, 11/13 마감"),
         ("폼 문항", "이름·연락처·알레르기·12/19(토) 참석 가능 여부")]
paras = []
for a, b in items:
    paras.append([(a + "  ", 14, True, TC.ACCENT_1), (b, 14, False, TC.BACKGROUND_1)])
text(8.15, 3.6, 4.3, 2.8, paras, "모집글-목록", space=7)

# 하단 메모
text(0.6, 6.58, 12.1, 0.3, [[("문의 25~30명에서 신청 10명이 나온다고 가정(전환 3분의 1). 실제 반응을 보고 11/10에 채널을 조정합니다.", 12, False, TC.TEXT_2)]], "가정-메모")

ns_ = s.notes_slide
if ns_.notes_text_frame is None:
    for sh in prs.slides[4].notes_slide.shapes:
        if sh.is_placeholder and "BODY" in str(sh.placeholder_format.type):
            ns_.shapes._spTree.append(copy.deepcopy(sh._element))
notes = ns_.notes_text_frame
notes.text = ("모집 기간이 11/6~11/13으로 8일뿐입니다. 1회차(11/18)까지 책 읽을 시간이 5일이라, 1회차 책은 앞부분만 읽고 오도록 범위를 정해 신청 확정 때 안내하세요. "
              "일정은 11/18, 12/2, 12/19(토), 12/30이며 3회차는 직장인 참석을 위해 주말인 12/19(토)로 옮겼습니다. 집결 시각은 전시장 도슨트 시간표를 확인한 뒤 정하고(운영 10:00~19:00, 입장 마감 18:00), 주말은 관람객이 많아 단체 도슨트 예약 여부를 미리 물어보세요. "
              "채널별 목표 인원과 전환율은 가정값입니다. 문토 수수료 20%는 사용자 제공 자료 기준이라 모집 플랫폼을 쓰면 수지표에 반영해야 합니다.")

# 섹션 목록에 새 슬라이드 추가
pres = prs.part._element
new_id = pres.find(qn("p:sldIdLst"))[-1].get("id")
ns14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
lst = pres.find(".//{%s}sectionLst" % ns14)
if lst is not None:
    sec = etree.SubElement(lst, "{%s}section" % ns14); sec.set("name", "모집"); sec.set("id", "{7B3C1F20-5A4D-4E1B-9C77-2D6A8E0F4B11}")
    l2 = etree.SubElement(sec, "{%s}sldIdLst" % ns14); e = etree.SubElement(l2, "{%s}sldId" % ns14); e.set("id", new_id)
prs.save(OUT)
print("saved", OUT)
