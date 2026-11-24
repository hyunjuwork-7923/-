# v7 → v8: 3회 구성, 76,000원, 새 일정(12/9·12/23·1/6), 외부 강사 게스트, 클럽 이름 문득클럽 반영
import sys, re
from pptx import Presentation
from pptx.util import Inches
prs = Presentation(sys.argv[1])
S = list(prs.slides)
def sh(slide, i):
    for x in slide.shapes:
        if x.shape_id == i: return x
    raise KeyError(i)
def kill(slide, *ids):
    for i in ids:
        e = sh(slide, i)._element; e.getparent().remove(e)
def settext(shape, *paras):
    ps = shape.text_frame.paragraphs
    assert len(paras) <= len(ps), (shape.name, len(paras), len(ps))
    for k, new in enumerate(paras):
        p = ps[k]; runs = p.runs
        new = [new] if isinstance(new, str) else list(new)
        for j, t in enumerate(new): runs[j].text = t
        for r in runs[len(new):]: r._r.getparent().remove(r._r)
    for p in ps[len(paras):]: p._p.getparent().remove(p._p)
def move(shape, x=None, y=None, w=None, h=None):
    if x is not None: shape.left = Inches(x)
    if y is not None: shape.top = Inches(y)
    if w is not None: shape.width = Inches(w)
    if h is not None: shape.height = Inches(h)
def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

# ---- 1 표지
s = S[0]
settext(sh(s, 5), "문득클럽 · 3회 프로그램 기획안")
settext(sh(s, 7), "전시 《스페인 미술 500년: 빛과 어둠의 연대기》 · 3회")
kill(s, 19, 20, 21, 25)
for i, (c, n, l, d) in enumerate([(10, 11, 12, 22), (13, 14, 15, 23), (16, 17, 18, 24)]):
    x = 0.8 + 4.0 * i
    move(sh(s, c), x=x); move(sh(s, n), x=x); move(sh(s, l), x=x + 1.05, w=2.7); move(sh(s, d), x=x, w=3.0)
settext(sh(s, 18), "전시 관람 + 발표", "간식 먹으며 3분 해설")
settext(sh(s, 22), "12/9 (수)"); settext(sh(s, 23), "12/23 (수)"); settext(sh(s, 24), "1/6 (수) 저녁")
notes(s, "표지. 1~2회차는 책, 3회차는 전시 관람과 해설 발표입니다. 파일럿으로 운영하고 결과를 보고 내용을 고칩니다. 전시 기간은 예술의전당 공식 안내로 최종 확인하세요.")

# ---- 2 커리큘럼
s = S[1]
settext(sh(s, 2), "3회 커리큘럼: 책 2회 → 전시와 나눔 1회")
settext(sh(s, 5), ["3", "회 과정을 지나면", ", ", "스페인 미술에 대해 설명할 수 있어요", ". "])
settext(sh(s, 9), "1회차 · 책 · 화가 세 명을 구분한다")
settext(sh(s, 17), "2회차 · 책 · 시대와 배경을 연결한다")
settext(sh(s, 12), ["활동  ", "체크인 15분 → 화가별 발제 50분 → ‘그림 한 장 천천히 읽기’ 45분 → 화가 구분 게임 30분"])
settext(sh(s, 20), ["활동  ", "맥락 토론 45분 → 시대 구분 지도 40분 → 관람 위시리스트 40분 → 미션 안내 20분"])
kill(s, 30, 31, 32, 33, 34, 35, 36, 37, 42)
move(sh(s, 22), w=12.15)
settext(sh(s, 25), "3회차 · 전시 + 나눔 · 실물 앞에서 보고, 한 점을 3분 해설한다"); move(sh(s, 25), w=10.5)
settext(sh(s, 26), "전시장에서 보고, 간식 먹으며 3분 해설하기"); move(sh(s, 26), w=10.5)
settext(sh(s, 27), "예술의전당 한가람디자인미술관 · 외부 강사 게스트"); move(sh(s, 27), w=10.5)
settext(sh(s, 28), ["활동  ", "외부 강사 해설 → 미션 카드 관람 → 3분 해설 초안 → 간식 먹으며 해설 발표(1인 3분 + 질문 1분)"]); move(sh(s, 28), w=11.5)
settext(sh(s, 29), ["산출물  ", "나의 한 점 + 3분 해설 + 만족도 설문"]); move(sh(s, 29), w=11.5)
settext(sh(s, 39), "12/9"); settext(sh(s, 40), "12/23"); settext(sh(s, 41), "1/6")
notes(s, "목표 설계: 3회(모임 9시간)로 ‘준전문가’까지는 어렵고, ‘처음 보는 사람 앞에서 그림 한 점을 3분 해설’이 현실적인 도달점입니다. 1회차 시작과 3회차 끝에 같은 자기평가 3문항(화가 구분, 시대 순서 설명, 한 점 3분 해설)을 받아 전후를 비교하세요. 교재: 1회차 《나는 스페인 미술관의 도슨트입니다》(이안), 2회차 《스페인, 역사와 미술을 보다》(김영욱). 심화 선택 《프라도 미술관에서 꼭 봐야 할 그림 100》(김영숙).")

# ---- 3 시간표
s = S[2]
for x in s.shapes:
    if x.has_text_frame and x.text_frame.text.startswith("주차별 시간표"): settext(x, "회차별 시간표: 한 번에 180분")
settext(sh(s, 29), "1회차 · 12/9 (수)"); settext(sh(s, 53), "2회차 · 12/23 (수)"); settext(sh(s, 77), "3회차 · 1/6 (수) 저녁")
settext(sh(s, 51), "정리 · 2회차 책 범위 안내"); settext(sh(s, 72), "3회차 미션 카드 안내 · 동선 공유")
settext(sh(s, 84), "외부 강사 해설 (불발 시 도슨트·오디오가이드)")
settext(sh(s, 100), "1/6 수요일 저녁에 전시장 관람이 가능한지 운영 시간 확인이 필요합니다. 외부 강사는 섭외 확정 전입니다.")

notes(s, "1회차 15분: 체크인, 자기평가 3문항(화가 구분 / 시대 순서 설명 / 한 점 3분 해설). 50분: 화가별 발제, 책 구성에 맞춰 세 명 확정. 45분: 그림 한 장 천천히 읽기, 한 점을 두고 보이는 것부터 말하기. 30분: 화가 구분 게임. 2회차 40분: 시대 구분 지도 만들기, 연표에 화가와 사건 붙이기. 40분: 관람 위시리스트, 도판을 보며 전시에서 꼭 볼 3점 고르기. 20분: 3회차 미션 카드 안내와 동선 공유. 3회차 25분: 3분 해설 초안 쓰기, 구성은 장면 → 빛과 어둠 → 내 한 줄. 40분: 스페인 간식을 먹으며 해설 발표, 1인당 3분 해설과 1분 질문(10명 기준). 마지막 10분: 자기평가 재측정, 만족도 조사, 다음 기수 제안은 설문 문항 하나로 받습니다.")

# ---- 4 3회차 상세 (사용자 작성 슬라이드)
s = S[3]
for x in s.shapes:
    if x.has_text_frame and x.text_frame.text.startswith("3·4"): settext(x, "3회차: 전시 관람과 해설 발표")
settext(sh(s, 4), "관람과 이동 (105분)")
settext(sh(s, 6), "10분"); settext(sh(s, 7), "전시장 집결 · 오리엔테이션")
settext(sh(s, 9), "35분"); settext(sh(s, 10), "외부 강사 해설 (불발 시 도슨트)")
settext(sh(s, 12), "50분"); settext(sh(s, 13), "자유 관람 + 미션 카드")
settext(sh(s, 15), "10분"); settext(sh(s, 16), "모임 장소로 이동 · 휴식")
settext(sh(s, 35), "1/6 (수) 저녁")
settext(sh(s, 19), "스페인 간식과 3분 해설 (75분)")
settext(sh(s, 20), "메뉴 후보 — 관람 중 찾은 ‘그림 속 음식’에서 확정")
r = sh(s, 29).text_frame.paragraphs[0].runs
r[3].text = "3분 해설 초안 25분 → 간식 먹으며 발표 40분 → 자기평가·만족도 설문 10분"
for k in (2, 4, 5): r[k]._r.getparent().remove(r[k]._r)
settext(sh(s, 30), "음식 반입·취식 가능 · 알레르기는 2회차에 조사")
kill(s, 36)
settext(sh(s, 32), "외부 강사 (깜짝 게스트)")
settext(sh(s, 33), "① 전시를 잘 아는 외부 강사가 관람 중 해설합니다 (섭외 확정 전, 이름은 현장에서 공개)", "② 일정이 맞지 않으면 전시장 도슨트나 오디오가이드로 대체합니다")
notes(s, "3회차는 평일 저녁이라 전시장 입장 마감(18:00)·폐관(19:00) 시간과 겹치는지 확인해야 합니다. 외부 강사는 섭외가 확정되기 전이므로 모집글에는 이름을 쓰지 않고 ‘외부 강사 게스트’로만 안내합니다.")

# ---- 5 시장 비교
s = S[4]
move(sh(s, 24), x=3.1, y=4.33, w=0.8, h=0.8); move(sh(s, 25), y=4.43)
settext(sh(s, 26), "문득클럽 (우리)", "약 2.5만"); move(sh(s, 26), y=4.48)
move(sh(s, 21), x=1.95, y=5.14)
settext(sh(s, 31), "3회에 76,000원 (파일럿 특가).", "넷플연가 정규 하단(17만 원)보다 낮게 책정")
notes(s, "포지셔닝 맵: 가로축은 약정 기간·몰입 깊이, 세로축은 회당 가격(막대는 범위)입니다. 우리는 3회, 76,000원이라 회당 약 2.5만 원입니다. 트레바리·넷플연가 가격은 사용자 제공 자료 기준입니다.")

# ---- 6 참가비
s = S[5]
settext(sh(s, 2), "참가비: 파일럿 특가 76,000원, 10명까지")
settext(sh(s, 4), "포함"); settext(sh(s, 5), "관람료·간식"); settext(sh(s, 6), "자료까지 포함")
settext(sh(s, 8), "파일럿 특가"); settext(sh(s, 9), "76,000원"); settext(sh(s, 10), "회당 약 2.5만 원 · 10명 한정")
settext(sh(s, 14), "도서 2권 개별 구입")
settext(sh(s, 16), "10명 기준 수지"); settext(sh(s, 18), "760,000원"); settext(sh(s, 20), "580,000원"); settext(sh(s, 22), "180,000원")
settext(sh(s, 23), "비용: 관람료 23만 · 간식 12만 · 자료 3만 · 외부 강사 20만 (가정).", "남는 돈은 진행자 보상·예비비. 손익분기 6명")
ps = sh(s, 26).text_frame.paragraphs
ps[2].runs[1].text = "관람료 · 간식 · 자료"
notes(s, "파일럿이라 76,000원으로 낮췄습니다(한·스페인 수교 76주년과 맞춘 숫자, 수교 연도는 확인 필요). 비용은 가정값이며 관람료, 외부 강사비 견적을 확인해 갱신하세요. 외부 강사비가 20만 원일 때 손익분기는 6명입니다.")

# ---- 7 모집
s = S[6]
settext(sh(s, 2), "모집 계획: 12월 2일까지 6~10명")
for i, (d, a, b) in zip((29, 33, 37, 41), [("11/25 (수)", "모집 오픈", "파일럿 특가 76,000원"), ("11/30 (월)", "마감 임박 알림", "인스타·당근에 다시 게시"),
                                          ("12/2 (수)", "마감", "최소 6명 확인 · 확정 안내"), ("12/9 (수)", "1회차, 이후 격주", "12/23 · 1/6 (수)")]):
    settext(sh(s, i), d); settext(sh(s, i + 1), a); settext(sh(s, i + 2), b)
settext(sh(s, 47), "동네생활·모임에 서초 인근으로 게시, 11/25 오픈·11/30 마감 알림")
settext(sh(s, 58), "인스타·플랫폼")
settext(sh(s, 59), "클럽 인스타 게시, 문토·소셜링은 보조(수수료 20% 별도, 자료 기준)")
ps = sh(s, 64).text_frame.paragraphs
def body(p, t):
    rs = p.runs; rs[1].text = t
    for r in rs[2:]: r._r.getparent().remove(r._r)
body(ps[1], "격주 수요일 3회: 12/9 · 12/23 · 1/6 (3회차는 전시장, 저녁)")
body(ps[2], "76,000원 파일럿 특가, 10명 한정 (관람료·간식 포함)")
ps[3]._p.getparent().remove(ps[3]._p)
body(ps[4], "진행자 소개·환불 규정·책 두 권 공개. 외부 강사는 깜짝 게스트")
body(ps[5], "이름·연락처·알레르기·1/6 저녁 참석 가능 여부")
settext(sh(s, 65), "기존 회원이 없어 채널별 인원은 모두 가정입니다. 문의 25~30명에서 신청 10명(전환 3분의 1)을 가정하고, 11/30에 채널을 조정합니다.")
notes(s, "모집 기간은 11/25~12/2, 8일입니다. 1회차(12/9)까지 책 읽을 시간이 1주일이라 1회차 책은 앞부분 범위를 정해 신청 확정 때 안내하세요. 클럽 이름은 문득클럽(wakeon.club)입니다. 당근은 거주 동네 기준이라 서초에서 일해도 다른 동네에 살면 게시글을 못 볼 수 있습니다. 채널별 목표 인원과 전환율은 가정값이며, 11/30에 부족하면 채널을 조정하세요.")

# ---- 전체: 이름, 주차→회차, 레이아웃 푸터
def fix(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            t = r.text
            t2 = t.replace("리드온 독서클럽", "문득클럽").replace("리드온", "문득클럽")
            t2 = re.sub(r"(\d)주차", r"\1회차", t2)
            t2 = t2.replace("스페인 미술 4주 과정", "스페인 미술 3회 과정")
            if t2 != t: r.text = t2
for sl in S:
    for x in sl.shapes:
        if x.has_text_frame: fix(x.text_frame)
for l in prs.slide_layouts:
    for x in l.shapes:
        if x.has_text_frame: fix(x.text_frame)
prs.save(sys.argv[2]); print("saved")
