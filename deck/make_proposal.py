# 큐레이터 초청 제안서 (A4, 2쪽 내외)
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Pretendard"
INK = RGBColor(0x1C, 0x17, 0x14); RED = RGBColor(0xA8, 0x32, 0x2D); GRAY = RGBColor(0x5A, 0x4E, 0x46)
d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.2); sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(1.8)

def setfont(run, size=10.5, bold=False, color=INK, hl=False):
    run.font.name = FONT; run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr(); rf = rPr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"): rf.set(qn(a), FONT)
    if hl: run.font.highlight_color = WD_COLOR_INDEX.YELLOW
st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

def para(parts, size=10.5, after=4, before=0, align=None, indent=None, line=1.45):
    p = d.add_paragraph(); pf = p.paragraph_format
    pf.space_after = Pt(after); pf.space_before = Pt(before); pf.line_spacing = line
    if align: p.alignment = align
    if indent is not None: pf.left_indent = Cm(indent)
    if isinstance(parts, str): parts = [(parts, {})]
    for t, o in parts:
        r = p.add_run(t); setfont(r, o.get("size", size), o.get("bold", False), o.get("color", INK), o.get("hl", False))
    return p
def heading(n, t):
    para([(f"{n}  ", {"bold": True, "color": RED, "size": 12.5}), (t, {"bold": True, "size": 12.5})], before=10, after=4).paragraph_format.keep_with_next = True
def bullet(parts, after=2):
    if isinstance(parts, str): parts = [(parts, {})]
    p = para([("•  ", {"color": RED})] + parts, after=after, indent=0.5); p.paragraph_format.first_line_indent = Cm(-0.5); return p
def ph(t): return (t, {"hl": True})  # 빈칸 표시

para([("큐레이터 토크 초청 제안서", {"bold": True, "size": 22})], after=2)
para([("문득클럽 스페인 미술 모임  ·  3회차 게스트 제안 (2027년 1월 13일)", {"size": 12, "color": RED, "bold": True})], after=10)

t = d.add_table(rows=3, cols=2); t.alignment = WD_TABLE_ALIGNMENT.LEFT; t.autofit = False
t.columns[0].width = Cm(2.2); t.columns[1].width = Cm(14.4)
for i, (a, b) in enumerate([("수신", [("오은수 선임 큐레이터님", {})]), ("발신", [("문득클럽 ", {}), ph("[진행자 이름]"), (" (wakeon.club)", {})]), ("일자", [("2026년 11월 25일", {})])]):
    c0, c1 = t.rows[i].cells; c0.width = Cm(2.2); c1.width = Cm(14.4)
    for c, parts, bold in ((c0, [(a, {"bold": True, "color": GRAY})], True), (c1, b, False)):
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(1)
        for tx, o in parts:
            r = p.add_run(tx); setfont(r, 10.5, o.get("bold", False), o.get("color", INK), o.get("hl", False))

heading("1", "인사")
para([("안녕하세요. 문득클럽에서 모임을 진행하는 ", {}), ph("[진행자 이름]"), ("입니다. 디자인을 전공했고, 지금은 전시와 상품을 기획하는 일을 합니다. 전시를 만드는 쪽에 있다 보니 작품 하나가 어떤 맥락에서 어떻게 보이도록 놓이는지가 늘 궁금했습니다.", {})])
para("《스페인 미술 500년: 빛과 어둠의 연대기》를 중심으로 작은 모임을 준비하고 있습니다. 이 전시를 기획하신 분의 이야기를 참가자들과 함께 듣고 싶어서, 마지막 모임에 게스트로 모시고자 조심스럽게 제안드립니다.")

heading("2", "모임 소개")
para("문득클럽은 ‘문득 문화가 고플 때’ 찾아오는 모임입니다. 이번 프로그램은 스페인 미술을 주제로 한 3회 파일럿입니다. 전시는 혼자 보면 남는 게 적다고 느끼는 분, 미술이 처음이라 어디서부터 봐야 할지 막막한 분이 대상이고, 서초동 인근 직장인 10명 안팎이 참여합니다.", after=4)
para("진행자는 스페인 미술 전문가로서 강의하지 않습니다. 책을 같이 읽고 전시를 같이 보는 사람으로 참여하고, 처음 보는 사람도 그림 앞에서 할 말이 생기도록 질문과 활동을 짰습니다. 모임이 끝났을 때 참가자가 그림 한 점을 자기 말로 3분 동안 설명할 수 있는 것이 목표입니다.", after=6)

tb = d.add_table(rows=4, cols=3); tb.style = "Table Grid"; tb.autofit = False
for ci, w in enumerate([1.9, 4.4, 10.3]): tb.columns[ci].width = Cm(w)
rows = [("회차", "일시 (모두 수요일)", "내용"),
        ("1회차", "12월 16일 19:00~21:30", "책 《나는 스페인 미술관의 도슨트입니다》 (화가 구분)"),
        ("2회차", "12월 30일 19:00~21:30", "책 《스페인, 역사와 미술을 보다》 (시대와 배경), 관람 티켓과 미션 카드 배부"),
        ("3회차", "1월 13일 19:00~21:30", "각자 전시를 보고 모여 큐레이터 토크, 한 점 3분 해설 발표")]
for i, r in enumerate(rows):
    for j, tx in enumerate(r):
        c = tb.rows[i].cells[j]; c.width = Cm([1.9, 4.4, 10.3][j])
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
        run = p.add_run(tx); setfont(run, 10, i == 0 or j == 0, RED if (i == 3 and j == 0) else INK)
        if i == 0:
            tcPr = c._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "F4F0EA"); tcPr.append(shd)

heading("3", "요청 드리는 내용")
bullet([("일시: ", {"bold": True}), ("2027년 1월 13일(수) 저녁 19:00~21:30 중 약 19:15부터 50분", {})])
bullet([("구성: ", {"bold": True}), ("큐레이터 토크 약 35분, 참가자 질문 약 15분", {})])
bullet([("토크 주제(제안): ", {"bold": True}), ("가제 ‘이 전시는 어떻게 놓였나’. 전시를 어떤 이야기로 구성했는지, 작품 한두 점을 두고 왜 그 자리에 그렇게 놓았는지를 들려주시면 좋겠습니다. 주제와 방식은 편하신 대로 정해주셔도 됩니다.", {})])
bullet([("참가자 상태: ", {"bold": True}), ("12월 30일 이후 전시를 각자 보고 옵니다(관람 티켓은 저희가 드립니다). 미션 카드 세 장(나의 한 점, 빛과 어둠 찾기, 그림 속 음식 찾기)을 들고 와서, 직접 본 작품을 바탕으로 질문합니다.", {})])
bullet([("그 뒤 순서: ", {"bold": True}), ("토크가 끝나면 참가자가 간식을 먹으며 한 점씩 3분 해설을 발표합니다. 발표를 지켜보시거나 먼저 자리를 비우셔도 괜찮습니다.", {})], after=4)

heading("4", "저희가 준비하는 것")
bullet([("장소: ", {"bold": True}), ("서울 서초구 남부순환로315길 85. 최대 10명을 수용하는 작은 공간이며, 음식 반입이 가능합니다.", {})])
bullet([("진행 환경: ", {"bold": True}), ("작품 이미지를 보여줄 화면, 스페인 간식(추로스 등)", {})])
bullet([("사례: ", {"bold": True}), ph("[사례 금액]"), (" 기준으로 협의드리고 싶습니다. 선생님의 기준이 있으시면 말씀해 주세요.", {})])
bullet([("이름 공개: ", {"bold": True}), ("기본은 ‘깜짝 게스트’로, 모집글에 이름을 쓰지 않고 당일에 소개합니다. 이름을 미리 알려도 괜찮으시면 알려주세요.", {})])
bullet([("소속 기관 절차: ", {"bold": True}), ("외부 행사 참여에 기관의 사전 승인이 필요하시면 맞춰 진행하겠습니다.", {})], after=4)

heading("5", "회신 부탁드립니다")
para([("2회차(12월 30일) 전에 참가자에게 안내해야 해서, ", {}), ("12월 23일(수)까지", {"bold": True}), (" 가능 여부를 알려주시면 감사하겠습니다. 1월 13일이 어려우시면 짧은 영상 인사말이나 서면 답변으로 참여하는 방법도 말씀해 주시면 같이 찾아보겠습니다. 거절하셔도 괜찮습니다.", {})], after=6)

heading("6", "첨부와 연락처")
bullet("프로그램 기획안 (PPT, 3쪽 발췌)")
bullet([("연락처: ", {"bold": True}), ph("[이름 · 전화번호 · 이메일]"), ("  /  wakeon.club", {})])
para("바쁘신 중에 긴 글 읽어주셔서 감사합니다.", before=8, after=2)
d.save(sys.argv[1]); print("saved")
