# -*- coding: utf-8 -*-
"""
브랜드북 빌더 v3 (좌표 기반 python-pptx)
- 세원 16장 시안(4번 톤)을 코드로 재현
- 팔레트 3종: teal_blue / navy_amber / green_orange
- 학원별 데이터 편차 스마트 채움:
  * 데이터 없는 섹션 슬라이드 자동 생략
  * 시간표는 그룹 수만큼 동적 생성, 행 많으면 행간 축소 + 자동 페이지 분할
  * 본문 넘침 시 폰트 auto-fit(길이 기반 pt 단계 축소)
슬라이드 크기 13.3 x 7.5 in. 서체 Pretendard(미설치 환경은 대체폰트).
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

EMU_IN = 914400
SLIDE_W = 13.333
SLIDE_H = 7.5
FONT = "Pretendard"

# ---------------- 팔레트 ----------------
PALETTES = {
    "teal_blue":  {"accent": "0EA5B7", "accent2": "2B7FFF", "navy": "22375F",
                   "soft": "E2F6F8", "soft2": "E8F1FF", "pill_bg": "D5EEF2", "pill_ink": "0B7C8A"},
    "navy_amber": {"accent": "22375F", "accent2": "D98A1F", "navy": "22375F",
                   "soft": "E7ECF4", "soft2": "FBF0DC", "pill_bg": "DCE3EF", "pill_ink": "22375F"},
    "green_orange":{"accent": "1F7A55", "accent2": "E07B39", "navy": "1D3F30",
                   "soft": "E3F2EA", "soft2": "FCEBDD", "pill_bg": "D6EBDF", "pill_ink": "196046"},
}

INK = "2D3748"        # 본문 진한먹색
INK_STRONG = "1A2233" # 헤더
LABEL = "5B6B85"      # 라벨 회색
LINE = "E3E8EF"
CARD = "F6F8FB"
CARD_LINE = "E7ECF3"
MUTED = "9AA7BD"
PH_BG = "DBE4EF"      # 사진 자리 배경

# 시간표 학년 태그 색 (팔레트 무관 고정 구분색은 accent 계열로 매핑)
def group_color(pal, group):
    g = (group or "").strip()
    if g.startswith("초"): return pal["accent"]
    if g.startswith("중"): return pal["accent2"]
    if g.startswith("고"): return pal["navy"]
    if "특" in g:          return "7C5CFF"
    return "E08A2B"


# ---------------- 저수준 헬퍼 ----------------
def _rgb(h): return RGBColor.from_string(h)

def _set_bg_white(slide):
    # 슬라이드 배경 흰색
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = _rgb("FFFFFF")

def _no_line(shape):
    shape.line.fill.background()

def _fill(shape, hexc):
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(hexc)
    _no_line(shape)

def _strip_style(shape):
    """python-pptx auto_shape가 자동 삽입하는 <p:style>(테마색/그림자 오염) 제거"""
    sp = shape._element
    for st in sp.findall(qn('p:style')):
        sp.remove(st)

def rect(slide, x, y, w, h, hexc=None, radius=False, line_hex=None, line_w=None):
    shp = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shp, Inches(x), Inches(y), Inches(w), Inches(h))
    _strip_style(s)
    if hexc:
        _fill(s, hexc)
    else:
        s.fill.background()
        _no_line(s)
    if line_hex:
        s.line.color.rgb = _rgb(line_hex)
        s.line.width = Pt(line_w or 1)
    else:
        if not hexc:
            _no_line(s)
    # 그림자 제거
    s.shadow.inherit = False
    return s

def hline(slide, x, y, w, hexc=LINE, weight=1.0):
    ln = slide.shapes.add_connector(2, Inches(x), Inches(y), Inches(x+w), Inches(y))
    ln.line.color.rgb = _rgb(hexc)
    ln.line.width = Pt(weight)
    return ln

def vline(slide, x, y, h, hexc=LINE, weight=1.0):
    ln = slide.shapes.add_connector(2, Inches(x), Inches(y), Inches(x), Inches(y+h))
    ln.line.color.rgb = _rgb(hexc)
    ln.line.width = Pt(weight)
    return ln

def txt(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        line_spacing=1.0, wrap=True, space_after=None):
    """
    runs: list of (text, size_pt, bold, color_hex) 또는 문자열
    여러 문단은 runs 안에 dict {'para':[(...)]}로 넣거나, 리스트의 리스트로 전달
    여기서는 간단히: runs=[(text,size,bold,color), ...] = 한 문단 내 여러 run
    문단 구분이 필요하면 paras() 사용
    """
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing: p.line_spacing = line_spacing
    if space_after is not None: p.space_after = Pt(space_after)
    if isinstance(runs, str):
        runs = [(runs, 18, False, INK)]
    for (t, sz, b, c) in runs:
        r = p.add_run(); r.text = t
        r.font.size = Pt(sz); r.font.bold = b
        r.font.color.rgb = _rgb(c); r.font.name = FONT
    return tb

def paras(slide, x, y, w, h, plist, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          line_spacing=1.0, para_space=4, wrap=True):
    """plist: list of paragraphs; each = list of (text,size,bold,color) runs"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    first = True
    for para in plist:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(para_space)
        for (t, sz, b, c) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.bold = b
            r.font.color.rgb = _rgb(c); r.font.name = FONT
    return tb

def bullets(slide, x, y, w, h, items, size=13, color=INK, dot_hex="0EA5B7",
            line_spacing=1.15, para_space=3):
    """불릿 목록 - 점은 색 있는 마커로. hanging indent로 둘째 줄을 텍스트 시작선에 맞춤."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    # 들여쓰기 폭: 글자 크기에 비례 (점+공백 만큼)
    indent_in = 0.02 + size*0.0135   # 대략 점"• " 폭
    marL = Emu(int(indent_in*EMU_IN))
    ind = Emu(int(-indent_in*EMU_IN))
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.line_spacing = line_spacing
        p.space_after = Pt(para_space)
        # hanging indent: 문단 왼쪽 marL, 첫 줄만 -indent(점이 왼쪽으로)
        pPr = p._pPr if p._pPr is not None else p.get_or_add_pPr()
        pPr.set('marL', str(int(indent_in*EMU_IN)))
        pPr.set('indent', str(int(-indent_in*EMU_IN)))
        r0 = p.add_run(); r0.text = "•\t"
        r0.font.size = Pt(size); r0.font.bold = True; r0.font.color.rgb = _rgb(dot_hex); r0.font.name = FONT
        r1 = p.add_run(); r1.text = it
        r1.font.size = Pt(size); r1.font.color.rgb = _rgb(color); r1.font.name = FONT
    return tb


# ---------------- auto-fit (길이 기반 pt 축소) ----------------
def fit_size(text, base, min_size, cap_per_size):
    """
    text 길이에 따라 base에서 단계적으로 줄임.
    cap_per_size: 해당 pt에서 '한 줄에 편안한 글자수 * 허용 줄수' 상한(대략).
    넘으면 1pt씩 내려 min_size까지.
    """
    n = len(text)
    size = base
    while size > min_size and n > cap_per_size(size):
        size -= 0.5
    return size


# ---------------- 공통 슬라이드 프레임 ----------------
def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    _set_bg_white(s)
    return s

PADX = 0.62      # 좌우 여백(인치)
PADY = 0.58      # 상단 여백

def label(slide, pal, text, x=PADX, y=PADY):
    """작은 라벨 + 세로바 (한글만)"""
    vline(slide, x+0.02, y+0.02, 0.20, pal["accent"], weight=3.2)
    txt(slide, x+0.14, y-0.02, 6.0, 0.3, [(text, 12, True, LABEL)], anchor=MSO_ANCHOR.TOP)

def header(slide, text, x=PADX, y=None, size=27, w=11.9):
    if y is None: y = PADY + 0.30
    txt(slide, x, y, w, 0.9, [(text, size, True, INK_STRONG)], line_spacing=1.12)


# ======================================================================
#  각 슬라이드 빌더
# ======================================================================

def slide_cover(prs, pal, d):
    s = new_slide(prs)
    a = d["academy"]
    # 우측 사진 영역
    rect(s, 7.4, 0, SLIDE_W-7.4, SLIDE_H, PH_BG)
    txt(s, 7.4, SLIDE_H/2-0.2, SLIDE_W-7.4, 0.4,
        [("학생 학습 사진", 12, False, "9FB0C6")], align=PP_ALIGN.CENTER)
    # 좌측
    lx = 0.85
    # 슬로건(한 줄, 넘침 방지: 폭 6.2in에 맞춰 크기 조정)
    slo = a.get("slogan", "")
    ssize = fit_size(slo, 18, 12, lambda sz: int(6.2/(sz*0.017)))
    txt(s, lx, 1.75, 6.3, 0.5, [(slo, ssize, False, pal["accent"])], wrap=False)
    # 학원명
    txt(s, lx, 2.25, 6.4, 1.3, [(a["name"], 52, True, pal["navy"])], line_spacing=1.0)
    # 과목 뱃지
    badge = rect(s, lx, 3.75, _badge_w(a.get("subjects","")), 0.62, pal["soft"], radius=True)
    txt(s, lx, 3.80, _badge_w(a.get("subjects",""))-0.0, 0.5,
        [(a.get("subjects",""), 17, True, pal["pill_ink"])], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 연락처 (각 한 줄)
    paras(s, lx, 5.9, 6.4, 1.2, [
        [("위치   ", 12, True, pal["navy"]), (a.get("location",""), 12, False, "667788")],
        [("연락처   ", 12, True, pal["navy"]), (a.get("phone",""), 12, False, "667788")],
    ], line_spacing=1.2, para_space=6)
    return s

def _badge_w(text):
    return max(2.0, 0.5 + len(text)*0.16)

def slide_intro(prs, pal, d):
    s = new_slide(prs)
    intro = d["intro"]; feats = d.get("features", [])[:6]
    label(s, pal, "학원 소개")
    header(s, intro["head"], size=26)
    # 소개 본문
    body = intro.get("body","")
    bsize = fit_size(body, 15, 12, lambda sz: int((11.9/(sz*0.0165))*3.2))
    txt(s, PADX, 1.75, 11.9, 1.2, [(body, bsize, False, INK)], line_spacing=1.45)
    # 강점 6박스 (3x2)
    gx, gy = PADX, 3.15
    gw, gh = (11.9-2*0.28)/3, 1.75
    gap = 0.28
    for i, f in enumerate(feats):
        r = i//3; c = i%3
        x = gx + c*(gw+gap); y = gy + r*(gh+0.24)
        card = rect(s, x, y, gw, gh, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        # pill 제목
        pw = min(gw-0.5, 0.5 + len(f["title"])*0.20)
        pill = rect(s, x+(gw-pw)/2, y+0.18, pw, 0.42, pal["pill_bg"], radius=True)
        txt(s, x+(gw-pw)/2, y+0.20, pw, 0.38, [(f["title"], 13, True, pal["pill_ink"])],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 설명
        dsize = fit_size(f["desc"], 12.5, 10, lambda sz: int((gw/(sz*0.017))*3.0))
        txt(s, x+0.18, y+0.72, gw-0.36, gh-0.85, [(f["desc"], dsize, False, INK)],
            align=PP_ALIGN.CENTER, line_spacing=1.28, anchor=MSO_ANCHOR.TOP)
    return s

def slide_achievements(prs, pal, d):
    ach = d.get("achievements")
    if not ach or not ach.get("items"): return None
    s = new_slide(prs)
    label(s, pal, "주요 실적")
    header(s, ach.get("head","진학으로 증명하는 결과"), size=26)
    items = ach["items"][:3]
    n = len(items)
    gap = 0.34
    gw = (11.9-(n-1)*gap)/n; gh = 3.9
    gx, gy = PADX, 2.35
    icon_map = {"trophy":"🏆","cap":"🎓","star":"⭐","book":"📘","medal":"🏅"}
    for i, it in enumerate(items):
        x = gx + i*(gw+gap)
        rect(s, x, gy, gw, gh, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        ic = icon_map.get(it.get("icon",""), "⭐")
        txt(s, x, gy+0.5, gw, 0.9, [(ic, 40, False, INK)], align=PP_ALIGN.CENTER)
        txt(s, x, gy+1.6, gw, 0.6, [(it["name"], 24, True, pal["navy"])], align=PP_ALIGN.CENTER)
        txt(s, x+0.3, gy+2.35, gw-0.6, 1.3, [(it["desc"], 15, False, INK)],
            align=PP_ALIGN.CENTER, line_spacing=1.35)
    return s

def slide_targets(prs, pal, d):
    tg = d.get("targets")
    if not tg or not tg.get("items"): return None
    s = new_slide(prs)
    label(s, pal, "수업 대상 · 과목")
    header(s, tg.get("head",""), size=26)
    items = tg["items"][:3]
    n = len(items)
    stage_colors = _stage_colors(pal, n)   # 커리큘럼과 동일: 초→중→고 진해짐
    top = 2.2; avail = 6.9 - top
    rh = min(1.55, avail/n)
    for i, it in enumerate(items):
        y = top + i*rh
        cy = y + rh/2          # 행 세로 중앙
        c = stage_colors[i]
        soft = _tint_for(pal, c)
        subj = it.get("subj","")
        desc = it.get("desc","")
        # 좌 학년 라벨(pill) — 학년별 색, 폭 넓혀 1줄, 행 중앙 정렬
        lw = 2.0; lh = 0.62
        rect(s, PADX, cy-lh/2, lw, lh, soft, radius=True)
        txt(s, PADX, cy-lh/2, lw, lh, [(it["grade"], 14, True, c)],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        # 중앙: 과목/반(학년색 강조) + 배우는 내용
        cx = PADX+lw+0.35; cw = 11.9+PADX - cx - 2.5
        ssize = fit_size(subj, 15, 12, lambda sz: int((cw/(sz*0.016))*1.6))
        subj_h = _subj_h(subj, cw)
        if desc:
            # 과목명 + 설명 두 줄 블록을 행 중앙에
            block_h = subj_h + 0.05 + 0.4
            by = cy - block_h/2
            txt(s, cx, by, cw, 0.6, [(subj, ssize, True, c)], line_spacing=1.2,
                anchor=MSO_ANCHOR.TOP)
            txt(s, cx, by+subj_h+0.05, cw, 0.5, [(desc, 13.5, False, INK)], line_spacing=1.3)
        else:
            # 과목명만 → 라벨과 같은 세로 중앙에 정확히
            txt(s, cx, cy-lh/2, cw, lh, [(subj, ssize, True, c)], line_spacing=1.2,
                anchor=MSO_ANCHOR.MIDDLE)
        # 우 사진 — 행 중앙 정렬
        pw = 2.2; ph_h = rh-0.4
        rect(s, 11.9+PADX-pw, cy-ph_h/2, pw, ph_h, PH_BG, radius=True)
        txt(s, 11.9+PADX-pw, cy-0.15, pw, 0.3,
            [((it["grade"].split()[0])+" 수업", 11, False, "9FB0C6")], align=PP_ALIGN.CENTER)
        if i < n-1:
            hline(s, PADX, y+rh, 11.9, LINE, 1)
    return s

def _subj_h(subj, cw):
    # 과목명이 한 줄이면 0.42, 두 줄이면 0.72 정도 (대략)
    return 0.72 if len(subj) > int(cw/0.017) else 0.42

def slide_curriculum(prs, pal, d):
    cu = d.get("curriculum")
    if not cu or not cu.get("stages"): return None
    s = new_slide(prs)
    label(s, pal, "단계별 커리큘럼")
    header(s, cu.get("head",""), size=26)
    if cu.get("sub"):
        txt(s, PADX, 1.55, 11.9, 0.4, [(cu["sub"], 15, False, "556677")])
    stages = cu["stages"]
    n = len(stages)
    # 단계별 색: 초→중→고로 진해짐
    step_colors = _stage_colors(pal, n)
    gap = 0.35
    bw = (11.9 - (n-1)*gap)/n
    # ── 박스 높이 통일 (내용 최다 단계 기준) ──
    max_items = max(len(st.get("items",[])) for st in stages)
    bh = 0.60 + 0.42*max_items + 0.35   # 헤더 + 항목수*행높이 + 여백
    bh = min(bh, 2.7)
    base_y = 6.6           # 계단 맨 아래 박스의 바닥
    step_up = min(0.62, (6.6-2.55-bh)/(n-1)) if n>1 else 0  # 위치만 계단
    for i, st in enumerate(stages):
        x = PADX + i*(bw+gap)
        y = base_y - bh - i*step_up     # 높이는 같고 위치만 위로
        c = step_colors[i]
        soft = _tint_for(pal, c)
        # 태그라인 (박스 위)
        tag = st.get("tag","")
        if tag:
            tagw = 0.4 + len(tag)*0.24
            rect(s, x+0.05, y-0.52, tagw, 0.40, soft, radius=True)
            txt(s, x+0.05, y-0.50, tagw, 0.36, [(tag, 13, True, c)],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            vline(s, x+0.05+tagw/2, y-0.12, 0.12, c, 2)
        # 단계 박스 (모두 동일 높이)
        rect(s, x, y, bw, bh, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        rect(s, x, y, bw, 0.60, c, radius=True)
        rect(s, x, y+0.30, bw, 0.30, c)
        txt(s, x+0.22, y+0.08, bw-0.4, 0.45, [(st["name"], 16, True, "FFFFFF")], anchor=MSO_ANCHOR.MIDDLE)
        lv = st.get("level", i+1)
        txt(s, x, y+0.08, bw-0.22, 0.45, [("LV."+str(lv), 12, True, "FFFFFF")],
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, x+0.22, y+0.80, bw-0.44, bh-0.9, st["items"], size=11.5,
                color=INK, dot_hex=c, line_spacing=1.2, para_space=4)
        # 상승 화살표 (다음 박스 태그 쪽으로)
        if i < n-1:
            txt(s, x+bw-0.02, y+0.15, 0.7, 0.4, [("↗", 18, True, MUTED)])
    return s

def _stage_colors(pal, n):
    seq = [pal["accent"], pal["accent2"], pal["navy"]]
    if n <= 3:
        return seq[:n] if n>1 else [pal["accent"]]
    # 3단계 초과면 반복/보간
    out = []
    for i in range(n):
        out.append(seq[min(i, len(seq)-1)])
    return out

def _tint_for(pal, c):
    if c == pal["accent"]: return pal["soft"]
    if c == pal["accent2"]: return pal["soft2"]
    return "E7ECF4"

def slide_timetable(prs, pal, d, group, rows, part=None, parts=None):
    s = new_slide(prs)
    label(s, pal, "시간표")
    gc = group_color(pal, group)
    # 학년 태그 + 헤더
    tag_w = 0.5 + len(group)*0.28
    rect(s, PADX, PADY+0.30, tag_w, 0.5, gc, radius=True)
    txt(s, PADX, PADY+0.35, tag_w, 0.42, [(group, 15, True, "FFFFFF")],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ttl = "시간표" + (f"  ({part}/{parts})" if parts and parts>1 else "")
    txt(s, PADX+tag_w+0.25, PADY+0.28, 8.0, 0.6, [(ttl, 26, True, INK_STRONG)])
    # 표
    _draw_table(s, pal, gc, rows, top=1.95)
    return s

def _draw_table(s, pal, header_hex, rows, top):
    cols = [0.30, 0.34, 0.18, 0.18]  # 비율: 반 / 요일시간 / 강사 / 강의실
    tbl_w = 11.9
    xs = [PADX]
    for c in cols[:-1]:
        xs.append(xs[-1] + c*tbl_w)
    widths = [c*tbl_w for c in cols]
    headers = ["반", "요일 · 시간", "강사", "강의실"]
    n = len(rows)
    # 행 높이 자동조절: 세로영역 5.0in ÷ (행수+헤더), 0.30~0.62 제한
    avail = 5.0
    rh = avail/(n+1)
    rh = max(0.30, min(0.62, rh))
    hh = min(0.5, rh+0.04)
    body_font = 15 if n <= 7 else (14 if n <= 10 else 13)
    # 헤더
    y = top
    hdr = rect(s, PADX, y, tbl_w, hh, header_hex)
    for i,htext in enumerate(headers):
        txt(s, xs[i]+0.14, y, widths[i]-0.2, hh, [(htext, 14, True, "FFFFFF")],
            anchor=MSO_ANCHOR.MIDDLE)
    y += hh
    # 바디
    for ri, row in enumerate(rows):
        if ri % 2 == 1:
            rect(s, PADX, y, tbl_w, rh, CARD)
        for ci, cell in enumerate(row):
            col = MUTED if (cell.strip()=="-" ) else INK
            txt(s, xs[ci]+0.14, y, widths[ci]-0.2, rh, [(cell, body_font, False, col)],
                anchor=MSO_ANCHOR.MIDDLE)
        # 행 구분선
        hline(s, PADX, y+rh, tbl_w, LINE, 0.75)
        y += rh

def slide_specials(prs, pal, d):
    sp = d.get("specials")
    if not sp or not sp.get("items"): return None
    s = new_slide(prs)
    label(s, pal, "특별 프로그램")
    header(s, sp.get("head",""), size=26)
    items = sp["items"][:4]
    # 2x2 사분면
    gx, gy = PADX, 2.25
    cw, ch = 11.9/2, 2.15
    # 구분선(십자)
    vline(s, gx+cw, gy, ch*2, LINE, 1)
    hline(s, gx, gy+ch, 11.9, LINE, 1)
    for i, it in enumerate(items):
        r=i//2; c=i%2
        x = gx + c*cw + 0.3; y = gy + r*ch + 0.3
        paras(s, x, y, cw-0.6, 0.5, [[
            (it.get("no","")+"  ", 20, True, pal["accent"]),
            (it["title"], 19, True, INK_STRONG)]])
        txt(s, x, y+0.65, cw-0.6, ch-1.0, [(it["desc"], 14.5, False, INK)], line_spacing=1.35)
    return s

def slide_management(prs, pal, d):
    mg = d.get("management")
    if not mg or not mg.get("columns"): return None
    s = new_slide(prs)
    label(s, pal, "과목별 학생 관리")
    header(s, mg.get("head",""), size=26)
    cols = mg["columns"][:2]
    # 좌우 2단 + 가운데 세로 구분선(카드 배경 없음)
    total = 11.9; gap = 0.7
    colw = (total-gap)/2
    xL = PADX; xR = PADX+colw+gap
    vline(s, PADX+colw+gap/2, 2.15, 4.4, LINE, 1)
    key_colors = [pal["accent"], pal["accent2"]]
    for idx, col in enumerate(cols):
        x = xL if idx==0 else xR
        kc = key_colors[idx]
        txt(s, x, 2.15, colw, 0.5, [(col["name"], 19, True, kc)])
        hline(s, x, 2.72, colw, LINE, 1)
        yy = 2.95
        kw = 0.78
        for row in col["rows"]:
            # 키 박스
            rect(s, x, yy, kw, 0.42, kc, radius=True)
            txt(s, x, yy+0.02, kw, 0.38, [(row["k"], 12.5, True, "FFFFFF")],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            # 설명 (왼쪽 줄맞춤 통일)
            vv = row["v"]
            vsize = fit_size(vv, 14, 11, lambda sz: int(((colw-kw-0.2)/(sz*0.017))*2.4))
            txt(s, x+kw+0.2, yy+0.01, colw-kw-0.2, 0.9, [(vv, vsize, False, INK)], line_spacing=1.28)
            yy += 0.92
    return s

def slide_admission(prs, pal, d):
    ad = d.get("admission")
    if not ad or not ad.get("steps"): return None
    s = new_slide(prs)
    label(s, pal, "입학 절차")
    header(s, ad.get("head",""), size=26)
    steps = ad["steps"]
    top = 2.3
    block_h = 1.7
    for i, st in enumerate(steps):
        y = top + i*(block_h+0.55)
        hline(s, PADX, y, 11.9, LINE, 1)
        # 큰 숫자
        txt(s, PADX, y+0.15, 1.4, 1.4, [(st["no"], 62, True, pal["accent"])], align=PP_ALIGN.CENTER)
        # 제목 + 항목
        txt(s, PADX+1.7, y+0.18, 9.8, 0.5, [(st["title"], 20, True, INK_STRONG)])
        items = ["– "+it for it in st.get("items",[])]
        paras(s, PADX+1.7, y+0.72, 9.8, 0.9,
              [[(it, 15, False, INK)] for it in items], line_spacing=1.4, para_space=2)
        # 화살표 브릿지
        if i < len(steps)-1 and ad.get("bridge"):
            by = y+block_h+0.05
            txt(s, PADX, by, 1.4, 0.4, [("↓", 20, False, MUTED)], align=PP_ALIGN.CENTER)
            txt(s, PADX+1.7, by+0.02, 6.0, 0.4, [(ad["bridge"], 14, True, pal["accent"])])
    return s

def slide_rules(prs, pal, d):
    ru = d.get("rules")
    if not ru or not ru.get("items"): return None
    s = new_slide(prs)
    # 좌 제목 / 우 항목
    label(s, pal, "관리 지침", x=PADX, y=2.3)
    txt(s, PADX, 2.7, 3.0, 1.2, [(ru.get("head","학원 관리 지침"), 34, True, INK_STRONG)], line_spacing=1.15)
    rx = 4.1; rw = 11.9+PADX-rx
    items = ru["items"]
    top = 0.9; gap = (6.0-0)/max(len(items),1)
    gap = min(1.15, gap)
    kw = 0.82
    yy = top
    for it in items:
        rect(s, rx, yy, kw, 0.46, pal["soft"], radius=True)
        txt(s, rx, yy+0.03, kw, 0.4, [(it["k"], 13.5, True, pal["pill_ink"])],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        vv = it["v"]
        vsize = fit_size(vv, 15, 12, lambda sz: int(((rw-kw-0.25)/(sz*0.017))*2.0))
        txt(s, rx+kw+0.25, yy+0.02, rw-kw-0.25, 0.9, [(vv, vsize, False, INK)], line_spacing=1.3)
        yy += gap
    return s

def slide_faq(prs, pal, d):
    fq = d.get("faq")
    if not fq or not fq.get("items"): return None
    s = new_slide(prs)
    label(s, pal, "자주 묻는 질문")
    header(s, fq.get("head","궁금한 점을 미리 확인하세요"), size=26)
    items = fq["items"][:4]
    gx, gy = PADX, 2.15
    gap = 0.3
    cw = (11.9-gap)/2; ch = 2.15
    for i, it in enumerate(items):
        r=i//2; c=i%2
        x = gx + c*(cw+gap); y = gy + r*(ch+0.25)
        rect(s, x, y, cw, ch, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        # Q
        q = it["q"]; a = it["a"]
        qsize = fit_size(q, 15, 12, lambda sz: int(((cw-0.5)/(sz*0.017))*2.0))
        paras(s, x+0.28, y+0.24, cw-0.56, 0.8,
              [[("Q. ", qsize, True, pal["accent2"]), (q, qsize, True, INK_STRONG)]], line_spacing=1.25)
        asize = fit_size(a, 13.5, 11, lambda sz: int(((cw-0.5)/(sz*0.017))*3.0))
        paras(s, x+0.28, y+1.05, cw-0.56, ch-1.2,
              [[("A. ", asize, True, pal["accent"]), (a, asize, False, "4A5568")]], line_spacing=1.3)
    return s

def slide_closing(prs, pal, d):
    s = new_slide(prs)
    cl = d.get("closing", {})
    a = d["academy"]
    # 우 지도
    mx = 6.9
    rect(s, mx, 0.9, 11.9+PADX-mx, 4.35, PH_BG, radius=True)
    txt(s, mx, 2.85, 11.9+PADX-mx, 0.4, [("지도 · 오시는 길", 12, False, "8CA0BB")], align=PP_ALIGN.CENTER)
    # 좌 카피 (쉼표까지 한 줄 / 다음 줄)
    head = cl.get("head","")
    lines = head.split("\n")
    hl = cl.get("highlight","")
    para_runs = []
    for ln in lines:
        runs=[]
        if hl and hl in ln:
            before, after = ln.split(hl,1)
            if before: runs.append((before, 30, True, INK_STRONG))
            runs.append((hl, 30, True, pal["accent"]))
            runs.append((after, 30, True, INK_STRONG))
        else:
            runs.append((ln, 30, True, INK_STRONG))
        para_runs.append(runs)
    paras(s, PADX, 1.0, 5.9, 1.6, para_runs, line_spacing=1.32, para_space=2, wrap=False)
    # 연락처(각 한 줄) — 값 있는 것만
    contact_paras = []
    if a.get("phone"):
        contact_paras.append([("전화 문의    ", 15, True, pal["navy"]), (a.get("phone",""), 15, False, INK)])
    if a.get("address_short") or a.get("location"):
        contact_paras.append([("위치    ", 15, True, pal["navy"]), (a.get("address_short","") or a.get("location",""), 15, False, INK)])
    if a.get("hours"):
        contact_paras.append([("운영시간    ", 15, True, pal["navy"]), (a.get("hours",""), 15, False, INK)])
    if contact_paras:
        paras(s, PADX, 2.75, 6.0, 1.6, contact_paras, line_spacing=1.3, para_space=8)
    # QR
    for i,(lab) in enumerate(["네이버 블로그","카카오 맵"]):
        qx = PADX + i*1.5
        rect(s, qx, 4.7, 0.95, 0.95, "EEF2F7", radius=True, line_hex=LINE, line_w=1)
        txt(s, qx, 5.05, 0.95, 0.3, [("QR", 11, False, "9FB0C6")], align=PP_ALIGN.CENTER)
        txt(s, qx-0.15, 5.72, 1.25, 0.3, [(lab, 11, False, "667788")], align=PP_ALIGN.CENTER)
    # CTA
    if cl.get("cta"):
        rect(s, PADX, 6.55, 11.9, 0.62, pal["accent"], radius=True)
        txt(s, PADX, 6.60, 11.9, 0.52, [(cl["cta"], 18, True, "FFFFFF")],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return s


# ======================================================================
#  메인 빌드
# ======================================================================
MAX_ROWS_PER_TT = 13  # 시간표 한 장 최대 행수

def build(data, palette="teal_blue", out="brandbook.pptx"):
    pal = PALETTES.get(palette, PALETTES["teal_blue"])
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    slide_cover(prs, pal, data)
    slide_intro(prs, pal, data)
    slide_achievements(prs, pal, data)
    slide_targets(prs, pal, data)
    slide_curriculum(prs, pal, data)

    # 시간표: 그룹 수만큼, 행 많으면 자동 분할
    for tt in data.get("timetables", []):
        group = tt.get("group","")
        rows = tt.get("rows", [])
        if not rows: continue
        if len(rows) <= MAX_ROWS_PER_TT:
            slide_timetable(prs, pal, data, group, rows)
        else:
            chunks = [rows[i:i+MAX_ROWS_PER_TT] for i in range(0,len(rows),MAX_ROWS_PER_TT)]
            for pi, ch in enumerate(chunks, 1):
                slide_timetable(prs, pal, data, group, ch, part=pi, parts=len(chunks))

    slide_specials(prs, pal, data)
    slide_management(prs, pal, data)
    slide_admission(prs, pal, data)
    slide_rules(prs, pal, data)
    slide_faq(prs, pal, data)
    slide_closing(prs, pal, data)

    prs.save(out)   # out은 경로(str) 또는 file-like(BytesIO) 모두 가능
    return out


if __name__ == "__main__":
    import json, sys
    src = sys.argv[1] if len(sys.argv)>1 else "sewon.json"
    pal = sys.argv[2] if len(sys.argv)>2 else "teal_blue"
    out = sys.argv[3] if len(sys.argv)>3 else "brandbook.pptx"
    data = json.load(open(src, encoding="utf-8"))
    build(data, pal, out)
    print("saved", out)
