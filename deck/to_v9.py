# v8 → v9: 1회차 12/16 시작(12/16·12/30·1/13), 관람 티켓 제공 + 각자 관람, 3회차는 저녁 모임(큐레이터 유무 2안)
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.oxml.ns import qn
from lxml import etree
FONT = "Pretendard"
prs = Presentation(sys.argv[1]); S = list(prs.slides)
def sh(slide, i):
    for x in slide.shapes:
        if x.shape_id == i: return x
    raise KeyError(i)
def byname(slide, name):
    for x in slide.shapes:
        if x.name == name: return x
    raise KeyError(name)
def kill(slide, *ids):
    for i in ids:
        e = sh(slide, i)._element; e.getparent().remove(e)
def settext(shape, *paras):
    ps = shape.text_frame.paragraphs
    assert len(paras) <= len(ps), (shape.name, len(paras), len(ps))
    for k, new in enumerate(paras):
        runs = ps[k].runs; new = [new] if isinstance(new, str) else list(new)
        for j, t in enumerate(new): runs[j].text = t
        for r in runs[len(new):]: r._r.getparent().remove(r._r)
    for p in ps[len(paras):]: p._p.getparent().remove(p._p)
def move(shape, x=None, y=None, w=None, h=None):
    if x is not None: shape.left = Inches(x)
    if y is not None: shape.top = Inches(y)
    if w is not None: shape.width = Inches(w)
    if h is not None: shape.height = Inches(h)
def notes(slide, text): slide.notes_slide.notes_text_frame.text = text
def setfont(run, size, bold, color):
    f = run.font; f.size = Pt(size); f.bold = bold; f.name = FONT; f.color.theme_color = color
    rPr = run._r.get_or_add_rPr(); latin = rPr.find(qn("a:latin"))
    ea = etree.Element(qn("a:ea")); ea.set("typeface", FONT); latin.addnext(ea)
    cs = etree.Element(qn("a:cs")); cs.set("typeface", FONT); ea.addnext(cs)
def addchip(slide, x, y, w, h, fill, name):
    c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)); c.name = name
    c.adjustments[0] = 0.5; c.fill.solid(); c.fill.fore_color.theme_color = fill; c.line.fill.background(); c.shadow.inherit = False
def addtext(slide, x, y, w, h, t, size, bold, color, name, align=PP_ALIGN.LEFT):
    b = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); b.name = name
    tf = b.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = align; r = p.add_run(); r.text = t; setfont(r, size, bold, color)

# 1 표지
s = S[0]
settext(sh(s, 22), "12/16 (수)"); settext(sh(s, 23), "12/30 (수)"); settext(sh(s, 24), "1/13 (수) 저녁")
settext(sh(s, 18), "관람 + 발표", "저녁에 모여 3분 해설")
notes(s, "표지. 1~2회차는 책, 3회차는 각자 전시를 보고 저녁에 모여 해설하는 모임입니다. 관람 티켓은 제공합니다. 전시 기간(2027.1.20까지)은 예술의전당 공식 안내로 최종 확인하세요.")

# 2 커리큘럼
s = S[1]
settext(sh(s, 20), ["활동  ", "맥락 토론 45분 → 시대 구분 지도 40분 → 관람 위시리스트 40분 → 미션 카드·티켓 배부 20분"])
settext(sh(s, 26), "각자 전시를 보고, 저녁에 모여 3분 해설하기")
settext(sh(s, 27), "관람 티켓 제공 · 같은 공간에서 저녁 모임 · 큐레이터 게스트(섭외 중)")
settext(sh(s, 28), ["활동  ", "각자 관람(미션 카드) → 큐레이터 토크 → 3분 해설 초안 → 간식 먹으며 해설 발표(1인 3분 + 질문 1분)"])
settext(sh(s, 39), "12/16"); settext(sh(s, 40), "12/30"); settext(sh(s, 41), "1/13")

# 3 시간표
s = S[2]
for x in s.shapes:
    if x.has_text_frame and x.text_frame.text.startswith("회차별 시간표"): settext(x, "회차별 시간표: 1·2회차 180분, 3회차 150분")
settext(sh(s, 29), "1회차 · 12/16 (수)"); settext(sh(s, 53), "2회차 · 12/30 (수)"); settext(sh(s, 77), "3회차 · 1/13 (수) 저녁")
settext(sh(s, 72), "미션 카드·관람 티켓 배부 · 안내")
settext(sh(s, 78), "큐레이터 토크 + 한 점 3분 해설")
rows = [(15, "체크인 · 간식 세팅"), (35, "큐레이터 토크 (기획 이야기)"), (15, "질문"), (10, "휴식"),
        (25, "3분 해설 초안 (장면 → 빛과 어둠 → 내 한 줄)"), (40, "간식 먹으며 해설 발표 (1인 3분 + 질문 1분)"), (10, "자기평가 재측정 · 만족도·제안 설문")]
for i, (m, d) in enumerate(rows):
    settext(byname(s, f"분-2-{i}"), f"{m}분")
    settext(byname(s, f"내용-2-{i}"), d)
settext(sh(s, 100), "전시는 각자 보고(티켓 제공) 1/13 저녁에 모입니다. 큐레이터를 못 모시면 4번 슬라이드 구성으로 진행하며, 12/30(2회차)까지 확정합니다.")
notes(s, "3회차는 수요일 저녁이라 전시는 각자 12/30 이후 1/13 낮까지 보거나 미리 봅니다. 관람 티켓은 2회차에 배부합니다. 큐레이터 토크 35분과 질문 15분은 섭외가 안 되면 4번 슬라이드의 후기 라운드와 해설 연습 구성으로 바뀝니다.")

# 4 3회차 상세
s = S[3]
for x in s.shapes:
    if x.has_text_frame and x.text_frame.text.startswith("3회차:"): settext(x, "3회차: 관람은 각자, 저녁에 모여 해설하기")
settext(sh(s, 4), "관람은 각자 (티켓 제공)")
for (c, t), (a, b) in zip([(6, 7), (9, 10), (12, 13), (15, 16)],
                         [("12/30", "2회차에서 미션 카드·티켓 받기"), ("~1/13", "수요일 낮이나 미리 관람"), ("관람 후", "미션 카드 사진 보내기"), ("1/13", "저녁 모임 (같은 공간)")]):
    settext(sh(s, c), a); settext(sh(s, t), b)
kill(s, 35)
settext(sh(s, 19), "큐레이터를 못 모셨을 때 (150분)")
kill(s, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30)
for i, (m, d) in enumerate([(15, "체크인 · 간식 세팅"), (30, "미션 카드 후기 라운드"), (10, "휴식"), (30, "3분 해설 초안 + 짝 연습"),
                            (50, "해설 발표 (1인 3분 + 질문 2분)"), (15, "자기평가·설문 · 정리")]):
    y = 2.15 + i * 0.5
    addchip(s, 7.05, y, 0.95, 0.38, TC.ACCENT_1, f"B안-칩-{i}")
    addtext(s, 7.05, y, 0.95, 0.38, f"{m}분", 14, True, TC.TEXT_1, f"B안-분-{i}", PP_ALIGN.CENTER)
    addtext(s, 8.2, y, 4.25, 0.38, d, 14, False, TC.BACKGROUND_1, f"B안-내용-{i}")
settext(sh(s, 32), "간식과 큐레이터 확정")
settext(sh(s, 33), "① 메뉴 후보: 추로스·초콜라테, 토르티야, 하몽·치즈 타파스, 무알코올 상그리아 (관람 중 찾은 ‘그림 속 음식’에서 확정)", "② 큐레이터는 12/30(2회차)까지 확정하고, 못 모시면 오른쪽 구성으로 진행합니다")
notes(s, "A안(큐레이터 있음)은 3번 슬라이드의 3회차 열입니다. B안은 큐레이터 토크 35분과 질문 15분(50분)을 미션 카드 후기 라운드 30분, 해설 연습 5분 추가, 발표 시간 10분 추가로 채웁니다. 어느 쪽이든 3분 해설 발표와 자기평가 재측정은 그대로 들어갑니다. 발표 때 그림을 보여줄 화면과 전시장 사진 촬영 규정은 확인이 필요합니다.")

# 6 참가비
s = S[5]
settext(sh(s, 6), "관람 티켓 제공 · 자료 포함")
ps = sh(s, 26).text_frame.paragraphs
ps[1].runs[1].text = "참가자 10명 · 음식 반입·취식 가능"
ps[2].runs[1].text = "관람 티켓 · 간식 · 자료"
notes(s, "파일럿이라 76,000원으로 낮췄습니다(한·스페인 수교 76주년과 맞춘 숫자, 수교 연도는 확인 필요). 관람 티켓은 신청 마감 후 확정 인원만큼 구매해 2회차에 배부합니다. 날짜 지정 티켓이면 노쇼 손실이 생기므로 전시장에 기간권이나 단체권이 있는지 확인하세요. 비용은 가정값입니다.")

# 7 모집
s = S[6]
settext(sh(s, 2), "모집 계획: 12월 9일까지 6~10명")
for i, (d, a, b) in zip((29, 33, 37, 41), [("11/25 (수)", "모집 오픈", "파일럿 특가 76,000원"), ("12/2 (수)", "마감 임박 알림", "인스타·당근에 다시 게시"),
                                          ("12/9 (수)", "마감", "최소 6명 확인 · 티켓 구매"), ("12/16 (수)", "1회차, 이후 격주", "12/30 · 1/13 (수)")]):
    settext(sh(s, i), d); settext(sh(s, i + 1), a); settext(sh(s, i + 2), b)
settext(sh(s, 47), "동네생활·모임에 서초 인근으로 게시, 11/25 오픈·12/2 마감 알림")
ps = sh(s, 64).text_frame.paragraphs
def body(p, t):
    rs = p.runs; rs[1].text = t
    for r in rs[2:]: r._r.getparent().remove(r._r)
body(ps[1], "격주 수요일 12/16 · 12/30 · 1/13 (전시는 각자 관람, 3회차는 저녁 모임)")
body(ps[2], "76,000원 파일럿 특가, 10명 한정 (관람 티켓·간식 포함)")
body(ps[3], "진행자 소개·환불 규정·책 두 권 공개. 외부 강사는 깜짝 게스트")
body(ps[4], "이름·연락처·알레르기·1/13 저녁 참석 가능 여부")
settext(sh(s, 65), "기존 회원이 없어 채널별 인원은 모두 가정입니다. 문의 25~30명에서 신청 10명(전환 3분의 1)을 가정하고, 12/2에 채널을 조정합니다.")
notes(s, "모집 기간은 11/25~12/9, 2주입니다. 1회차(12/16)까지 책 읽을 시간이 1주일이라 1회차 책은 앞부분 범위를 정해 신청 확정 때 안내하세요. 12/30은 연말이라 참석률이 떨어질 수 있습니다. 클럽 이름은 문득클럽(wakeon.club)입니다. 당근은 거주 동네 기준이라 서초에서 일해도 다른 동네에 살면 게시글을 못 볼 수 있습니다. 채널별 목표 인원과 전환율은 가정값입니다.")
prs.save(sys.argv[2]); print("saved")
