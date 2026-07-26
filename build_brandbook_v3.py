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
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy
import math
import io as _io, urllib.request as _urlreq

EMU_IN = 914400

# URL/경로 이미지를 메모리로 가져오기 (실패 시 None)
_IMG_CACHE = {}
def _fetch_image(url):
    if not url: return None
    if url in _IMG_CACHE: return _IMG_CACHE[url]
    try:
        if url.startswith("data:"):
            # 담당자가 PC에서 올린 이미지(base64 data URL)
            import base64 as _b64
            _, _, b64 = url.partition(",")
            data = _b64.b64decode(b64)
        elif url.startswith("http"):
            req = _urlreq.Request(url, headers={"User-Agent":"Mozilla/5.0"})
            data = _urlreq.urlopen(req, timeout=8).read()
        else:
            with open(url,"rb") as f: data = f.read()
        _IMG_CACHE[url] = data
        return data
    except Exception:
        _IMG_CACHE[url] = None
        return None

def place_image(slide, url, x, y, w, h, cover=True):
    """url 이미지를 (x,y,w,h) 박스에 넣음. cover=True면 박스를 꽉 채우고 넘치는 부분 crop.
    성공 True / 실패 False (실패 시 호출부가 회색 박스 유지)."""
    data = _fetch_image(url)
    if not data: return False
    try:
        from PIL import Image as _PILImage
        im = _PILImage.open(_io.BytesIO(data))
        iw, ih = im.size
        box_ratio = w/h; img_ratio = iw/ih
        pic = slide.shapes.add_picture(_io.BytesIO(data), Inches(x), Inches(y), Inches(w), Inches(h))
        if cover:
            # 박스를 꽉 채우도록 crop (가운데 기준)
            if img_ratio > box_ratio:
                # 이미지가 더 넓음 → 좌우 crop
                crop = (1 - box_ratio/img_ratio)/2
                pic.crop_left = crop; pic.crop_right = crop
            else:
                crop = (1 - img_ratio/box_ratio)/2
                pic.crop_top = crop; pic.crop_bottom = crop
        else:
            # 비율 유지해 박스 안에 전체가 들어가게(로고 찌그러짐 방지). 좌·상단 기준 정렬.
            if img_ratio > box_ratio:
                nw = w; nh = w/img_ratio
            else:
                nh = h; nw = h*img_ratio
            pic.width = Inches(nw); pic.height = Inches(nh)
            pic.left = Inches(x); pic.top = Inches(y)
        return True
    except Exception:
        try:
            slide.shapes.add_picture(_io.BytesIO(data), Inches(x), Inches(y), Inches(w), Inches(h))
            return True
        except Exception:
            return False

SLIDE_W = 13.333
SLIDE_H = 7.5
FONT = "Pretendard SemiBold"

# ★한글 서체 적용: python-pptx 의 font.name 은 <a:latin> 만 채운다.
#   한글은 <a:ea>(East Asian) 슬롯을 보므로 이걸 비워두면 테마 기본서체
#   (Calibri, 한글 글리프 없음)로 떨어져 PowerPoint 가 맑은 고딕 등으로
#   임의 대체한다. 세 슬롯을 모두 지정해야 지정한 서체가 실제로 나온다.
def _apply_font(run, name=None):
    name = name or FONT
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def _force_theme_font(prs, name=None):
    """테마 majorFont/minorFont 를 지정 서체로 바꾼다.
    서체 지정이 누락된 텍스트가 있어도 Calibri(한글 없음)로 떨어지지 않게
    하는 안전장치. theme1.xml 은 python-pptx 가 파싱하지 않는 일반 Part 라
    blob 을 직접 치환한다."""
    name = name or FONT
    try:
        for part in prs.part.package.iter_parts():
            if "theme" not in str(part.partname):
                continue
            xml = part.blob.decode("utf-8", errors="ignore")
            def _fix(b):
                for slot in ("latin", "ea", "cs"):
                    b = re.sub(r'(<a:%s[^/>]*typeface=")[^"]*(")' % slot,
                               r"\1" + name + r"\2", b)
                return b
            for tag in ("majorFont", "minorFont"):
                m = re.search(r"<a:%s>.*?</a:%s>" % (tag, tag), xml, re.S)
                if m:
                    xml = xml[:m.start()] + _fix(m.group(0)) + xml[m.end():]
            part._blob = xml.encode("utf-8")
    except Exception:
        pass

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
        r.font.color.rgb = _rgb(c); _apply_font(r)
    return tb

def paras(slide, x, y, w, h, plist, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          line_spacing=1.0, para_space=4, wrap=True, hang=0.0):
    """plist: list of paragraphs; each = list of (text,size,bold,color) runs.
    hang: >0이면 hanging indent(인치) — 둘째 줄부터 그만큼 들여써서 마커 뒤에 맞춤."""
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
        if hang > 0:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(int(hang*EMU_IN)))
            pPr.set('indent', str(int(-hang*EMU_IN)))
        for (t, sz, b, c) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.bold = b
            r.font.color.rgb = _rgb(c); _apply_font(r)
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
        pPr = p._p.get_or_add_pPr()
        pPr.set('marL', str(int(indent_in*EMU_IN)))
        pPr.set('indent', str(int(-indent_in*EMU_IN)))
        r0 = p.add_run(); r0.text = "•\t"
        r0.font.size = Pt(size); r0.font.bold = True; r0.font.color.rgb = _rgb(dot_hex); _apply_font(r0)
        r1 = p.add_run(); r1.text = it
        r1.font.size = Pt(size); r1.font.color.rgb = _rgb(color); _apply_font(r1)
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


def _disp_width(text):
    """한글=1.0, 영문/숫자/기호=0.55 로 계산한 표시 폭(글자수 환산)."""
    w = 0.0
    for ch in text:
        o = ord(ch)
        # 한글 음절/자모 + CJK
        if 0xAC00 <= o <= 0xD7A3 or 0x3130 <= o <= 0x318F or 0x4E00 <= o <= 0x9FFF:
            w += 1.0
        elif ch in " ·":
            w += 0.4
        else:
            w += 0.55
    return w

def wrap_lines(text, width_in, size_pt):
    """텍스트박스 폭(in)과 폰트 크기(pt)로 실제 줄바꿈 수를 근사한다.
    Pretendard 기준 한글 한 글자 폭 ≈ size_pt*0.0139 in."""
    if not text:
        return 1
    char_in = size_pt * 0.0139
    chars_per_line = max(6, width_in / char_in)
    return max(1, math.ceil(_disp_width(text) / chars_per_line))


def fit_box(text, w, h, base, min_size, line_spacing=1.32, step=0.5):
    """텍스트박스 폭·높이 안에 들어오도록 글자 크기를 낮춘다(문장은 자르지 않는다)."""
    t = str(text or "")
    if not t: return base
    sz = base
    while sz > min_size:
        lines = sum(max(1, wrap_lines(ln, w, sz)) for ln in t.split("\n"))
        if lines * (sz*0.0139*1.62) * line_spacing <= h:
            return sz
        sz -= step
    return min_size


def _chunk(lst, n):
    """n개씩 나눈다. 개수가 많아도 잘라 버리지 않고 슬라이드를 나누기 위한 헬퍼."""
    n = max(1, int(n))
    return [lst[i:i+n] for i in range(0, len(lst), n)]

def _part(base, pi, parts):
    """여러 장으로 나뉠 때 라벨에 (1/2) 표기"""
    return base if parts <= 1 else f"{base} ({pi}/{parts})"


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
    # 두 줄을 넘어가면 아래 본문과 겹친다 → 크기를 낮춰 두 줄 안에 담는다
    t = str(text or "")
    while size > 19 and wrap_lines(t, w-0.14, size) > 2:
        size -= 1
    # 제목 왼쪽을 라벨 글자 시작선(세로바 오른쪽, +0.14)에 맞춤
    txt(slide, x+0.14, y, w-0.14, 0.95, [(t, size, True, INK_STRONG)], line_spacing=1.12)


# ======================================================================
#  각 슬라이드 빌더
# ======================================================================

def slide_cover(prs, pal, d):
    s = new_slide(prs)
    a = d["academy"]
    assets = d.get("assets", {}) or {}
    photos = assets.get("photos", []) or []
    logo = assets.get("logo", "")
    # 우측 사진 영역 — 위아래 여백을 줘 표지가 답답하지 않게
    PH_X, PH_Y = 7.15, 0.62
    PH_W, PH_H = SLIDE_W - PH_X - 0.72, SLIDE_H - 0.62*2
    rect(s, PH_X, PH_Y, PH_W, PH_H, PH_BG)
    if photos and place_image(s, photos[0], PH_X, PH_Y, PH_W, PH_H, cover=True):
        pass
    else:
        txt(s, PH_X, PH_Y+PH_H/2-0.2, PH_W, 0.4,
            [("학생 학습 사진", 12, False, "9FB0C6")], align=PP_ALIGN.CENTER)
    # 좌측
    lx = 0.85
    name_y = 2.25
    # 로고 — 사진과 윗변을 맞추고 크기를 키움
    if logo:
        LOGO_W, LOGO_H = 2.1, 1.0
        if place_image(s, logo, lx, PH_Y, LOGO_W, LOGO_H, cover=False):
            pass  # 비율 유지
    # 슬로건
    slo = a.get("slogan", "")
    ssize = fit_size(slo, 18, 12, lambda sz: int(6.2/(sz*0.017)))
    txt(s, lx, 1.85, 6.3, 0.5, [(slo, ssize, False, pal["accent"])], wrap=False)
    # 학원명
    txt(s, lx, name_y, 6.4, 1.3, [(a["name"], 52, True, pal["navy"])], line_spacing=1.0)
    # 과목 뱃지
    badge = rect(s, lx, 3.75, _badge_w(a.get("subjects","")), 0.62, pal["soft"], radius=True)
    txt(s, lx, 3.80, _badge_w(a.get("subjects",""))-0.0, 0.5,
        [(a.get("subjects",""), 17, True, pal["pill_ink"])], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 연락처 (각 한 줄) — 전화가 두 개면 둘 다, 카카오 채널도 있으면 함께
    _tel = a.get("phone","")
    if a.get("phone2"):
        _tel = f"{_tel} · {a['phone2']}" if _tel else a["phone2"]
    _c_paras = [
        [("위치   ", 12, True, pal["navy"]), (a.get("location",""), 12, False, "667788")],
        [("연락처   ", 12, True, pal["navy"]), (_tel, 12, False, "667788")],
    ]
    if a.get("kakao"):
        _c_paras.append([("카카오   ", 12, True, pal["navy"]), (a.get("kakao",""), 12, False, "667788")])
    paras(s, lx, 5.75, 6.4, 1.5, _c_paras, line_spacing=1.2, para_space=6)
    return s

def _badge_w(text):
    return max(2.0, 0.5 + len(text)*0.16)

def slide_intro(prs, pal, d):
    s = new_slide(prs)
    intro = d["intro"]; feats = d.get("features", [])[:6]
    label(s, pal, "학원 소개")
    header(s, intro["head"], size=26)
    # 소개 본문 — 헤더가 두 줄이면 그만큼 아래에서 시작
    _hl = wrap_lines(str(intro.get("head","")), 11.76, 26)
    body_y = 1.75 if _hl <= 1 else 2.05
    body = intro.get("body","")
    bsize = fit_box(body, 11.8, (3.15-body_y-0.05), 15, 11, line_spacing=1.45)
    txt(s, PADX, body_y, 11.9, 3.15-body_y-0.05, [(body, bsize, False, INK)], line_spacing=1.45)
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

def slide_identity(prs, pal, d):
    """정체성·교육철학 — 소개 슬라이드에 흡수돼 사라지던 문단을 제 자리에 싣는다."""
    idt = d.get("identity")
    if not idt or not str(idt.get("body","")).strip(): return None
    s = new_slide(prs)
    label(s, pal, "학원 정체성")
    header(s, idt.get("head",""), size=26)
    body = str(idt["body"]).strip()
    # 본문 폭·크기 자동 (문단이 길수록 작게, 최소 14pt)
    bw = 11.9 - 0.55
    size = 19.0
    while size > 14.0 and wrap_lines(body, bw, size) > 5:
        size -= 0.5
    lines = wrap_lines(body, bw, size)
    bh = lines * (size*0.0139*1.62) * 1.5 + 0.2
    top = 2.45 + max(0.0, (3.9 - bh)/2)
    # 좌측 강조 세로바 + 문단
    vline(s, PADX+0.03, top, min(bh, 3.9), pal["accent"], 3.2)
    txt(s, PADX+0.45, top, bw, bh, [(body, size, False, INK)], line_spacing=1.5)
    return s


def slide_achievements(prs, pal, d):
    ach = d.get("achievements")
    if not ach or not ach.get("items"): return None
    pages = _chunk(ach["items"], 5)
    last = None
    for pi, items in enumerate(pages, 1):
        last = _achievements_page(prs, pal, ach, items, pi, len(pages))
    return last

def _achievements_page(prs, pal, ach, items, pi=1, parts=1):
    """실적 카드 — 라벨 + 변화(강조) + 한 줄 설명. 문장형으로 들어와도 한 줄로 정리해 보여준다."""
    s = new_slide(prs)
    label(s, pal, _part("주요 실적", pi, parts))
    header(s, ach.get("head","진학으로 증명하는 결과") if pi == 1 else "", size=26)
    n = len(items)
    top, bottom = 2.20, 7.05
    gap = 0.20 if n >= 5 else 0.26
    ch = min(1.35, (bottom - top - gap*(n-1)) / max(n, 1))
    # 카드 안쪽 배치는 카드 높이에서 역산한다(작아져도 글자가 카드를 넘지 않게)
    pad   = 0.16 if ch >= 1.0 else 0.10
    tsize = 19 if ch >= 1.05 else (17 if ch >= 0.88 else 15)
    nsize = 13.5 if ch >= 1.05 else (12.5 if ch >= 0.88 else 11.0)
    th    = tsize*0.0139*1.52
    nh    = max(0.20, ch - pad*2 - th - 0.02)
    # 블록(제목+설명)을 카드 세로 중앙에
    blk   = th + 0.02 + nh
    for i, it in enumerate(items):
        y = top + i*(ch+gap)
        rect(s, PADX, y, 11.9, ch, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        title  = str(it.get("title") or it.get("name") or "").strip()
        change = str(it.get("change") or "").strip()
        note   = str(it.get("note") or "").strip()
        if not change and not note:
            # 구형 데이터({name, desc}) 호환 — 여러 줄이면 첫 줄만 강조, 나머지는 설명
            lines = [x.strip() for x in str(it.get("desc","")).split("\n") if x.strip()]
            change = lines[0] if lines else ""
            note = " / ".join(lines[1:])
        runs = []
        if title:
            runs.append((title + ("   " if change else ""), tsize, True, INK_STRONG))
        if change:
            runs.append((change, tsize, True, pal["accent"]))
        has_note = bool(note)
        ty = y + (pad if has_note else (ch-th)/2)
        if runs:
            paras(s, PADX+0.42, ty, 11.9-0.84, th+0.06, [runs], line_spacing=1.16, wrap=False)
        if has_note:
            nsz = fit_box(note, 11.9-0.94, nh, nsize, 9.5, line_spacing=1.28)
            txt(s, PADX+0.42, ty+th+0.02, 11.9-0.84, nh,
                [(note, nsz, False, "667788")], line_spacing=1.28)
    return s


def slide_targets(prs, pal, d):
    tg = d.get("targets")
    if not tg or not tg.get("items"): return None
    pages = _chunk(tg["items"], 4)      # 초·중·고·특강까지 한 장, 그 이상이면 나눔
    last = None
    base = 0
    for pi, items in enumerate(pages, 1):
        last = _targets_page(prs, pal, tg, items, pi, len(pages), base)
        base += len(items)
    return last

def _targets_page(prs, pal, tg, items, pi, parts, base):
    s = new_slide(prs)
    label(s, pal, _part("수업 대상 · 과목", pi, parts))
    header(s, tg.get("head","") if pi == 1 else "", size=26)
    n = len(items)
    stage_colors = _stage_colors(pal, base + n)[base:base+n]   # 초→중→고 색 이어짐
    area_top = 2.15; area_bottom = 7.05
    avail = area_bottom - area_top
    rh = min(1.55, avail/n)
    # 행 묶음을 세로 중앙 정렬(위아래 여백 균등)
    block = rh*n
    top = area_top + (avail - block)/2
    for i, it in enumerate(items):
        y = top + i*rh
        cy = y + rh/2          # 행 세로 중앙
        c = stage_colors[i]
        soft = _tint_for(pal, c)
        subj = it.get("subj","")
        desc = it.get("desc","")
        # 좌 학년 라벨(pill) — 학년별 색, 크게, 행 중앙 정렬
        lw = 2.0; lh = 0.66
        rect(s, PADX, cy-lh/2, lw, lh, soft, radius=True)
        txt(s, PADX, cy-lh/2, lw, lh, [(it["grade"], 15, True, c)],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        # 중앙: 과목/반(학년색 강조, 크게, 최대 2줄) + 배우는 내용
        cx = PADX+lw+0.35; cw = 11.9+PADX - cx - 2.1
        # 17pt 기준 2줄까지 허용: 1줄 넘으면 16 → 15로만 낮춤(작아지지 않게)
        ssize = fit_size(subj, 17, 15, lambda sz: int((cw/(sz*0.016))*2.0))
        # 과목명 줄 수 추정
        two_line = len(subj) > int(cw/(ssize*0.016))
        subj_box_h = 0.72 if two_line else 0.44
        if desc:
            block_h = subj_box_h + 0.05 + 0.4
            by = cy - block_h/2
            txt(s, cx, by, cw, subj_box_h, [(subj, ssize, True, c)], line_spacing=1.18,
                anchor=MSO_ANCHOR.TOP)
            txt(s, cx, by+subj_box_h+0.05, cw, 0.5, [(desc, 13.5, False, INK)], line_spacing=1.3)
        else:
            # 과목명만 → 라벨과 같은 세로 중앙에
            txt(s, cx, cy-subj_box_h/2, cw, subj_box_h, [(subj, ssize, True, c)],
                line_spacing=1.18, anchor=MSO_ANCHOR.MIDDLE)
        # 우 사진 자리 — 회색 박스만 (수기로 사진 넣을 자리, 자동삽입 안 함)
        pw = 1.85; ph_h = rh-0.4
        rect(s, 11.9+PADX-pw, cy-ph_h/2, pw, ph_h, PH_BG, radius=True)
        if i < n-1:
            hline(s, PADX, y+rh, 11.9, LINE, 1)
    return s

def _subj_h(subj, cw):
    # 과목명이 한 줄이면 0.42, 두 줄이면 0.72 정도 (대략)
    return 0.72 if len(subj) > int(cw/0.017) else 0.42

def slide_classes(prs, pal, d):
    """반 구성 상세 — 반 이름과 설명을 한 줄도 빠뜨리지 않는다.
    학년 그룹 수(초·중·고·특강…)와 반 개수에 제한이 없다. 우선 글자를 줄여 한 장에 담고,
    그래도 넘치면 다음 장으로 넘기되 이어지는 장에도 학년 라벨을 다시 표기한다."""
    cl = d.get("classes")
    if not cl or not cl.get("groups"): return None
    AREA_TOP, AREA_BOTTOM = 2.15, 7.05
    AVAIL = AREA_BOTTOM - AREA_TOP
    NAME_W, GAPX = 2.85, 0.30
    desc_w = 11.9 - NAME_W - GAPX

    groups = [g for g in cl["groups"] if g.get("rows")]
    if not groups: return None
    gcolors = _stage_colors(pal, max(len(groups), 1))

    def _rows_for(size):
        out = []
        for gi, g in enumerate(groups):
            c = gcolors[gi % len(gcolors)]
            if g.get("grade"):
                out.append(("head", (g["grade"], c), 0.44))
            for r in g["rows"]:
                nl = max(1, wrap_lines(str(r.get("name","")), NAME_W-0.2, size))
                dl = max(1, wrap_lines(str(r.get("desc","")), desc_w-0.15, size))
                h = max(nl, dl) * (size*0.0139*1.62) * 1.34 + 0.17
                out.append(("row", (r, c, g.get("grade","")), max(0.46, h)))
        return out

    rows, size = None, 13.5
    for sz in (13.5, 12.5, 11.5, 10.5, 9.5):
        rows, size = _rows_for(sz), sz
        if sum(r[2] for r in rows) <= AVAIL:
            break

    # 페이지 나누기 (넘칠 때만)
    pages, cur, acc = [], [], 0.0
    for it in rows:
        if acc + it[2] > AVAIL and cur:
            pages.append(cur); cur, acc = [], 0.0
        cur.append(it); acc += it[2]
    if cur: pages.append(cur)

    last = None
    for pi, page in enumerate(pages, 1):
        s = new_slide(prs)
        label(s, pal, _part("반 구성", pi, len(pages)))
        header(s, cl.get("head","반 구성과 학습 체계") if pi == 1 else "", size=26)
        y = AREA_TOP
        # 페이지가 학년 도중에 시작하면 그 학년 라벨을 다시 찍는다
        if page and page[0][0] == "row":
            gname = page[0][1][2]
            if gname:
                gc = page[0][1][1]
                _lab = gname + " (이어서)"
                pw = 0.5 + _disp_width(_lab)*0.19
                rect(s, PADX, y+0.02, pw, 0.40, _tint_for(pal, gc), radius=True)
                txt(s, PADX, y+0.02, pw, 0.40, [(_lab, 12, True, gc)],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
                y += 0.48
        for kind, payload, h in page:
            if kind == "head":
                gname, c = payload
                pw = 0.5 + _disp_width(gname)*0.19
                rect(s, PADX, y+0.01, pw, 0.40, _tint_for(pal, c), radius=True)
                txt(s, PADX, y+0.01, pw, 0.40, [(gname, 13, True, c)],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
            else:
                r, c, _g = payload
                txt(s, PADX+0.05, y+0.06, NAME_W-0.1, h-0.10,
                    [(str(r.get("name","")), size, True, c)], line_spacing=1.3)
                if r.get("desc"):
                    txt(s, PADX+NAME_W+GAPX, y+0.06, desc_w, h-0.10,
                        [(str(r["desc"]), size, False, INK)], line_spacing=1.34)
                hline(s, PADX, y+h, 11.9, LINE, 1)
            y += h
        last = s
    return last


def slide_growth(prs, pal, d):
    """성장 방식 — 과목을 열로 나란히. 과목이 1개든 4개든 한 장에, 5개 이상이면 나눈다."""
    gr = d.get("growth")
    if not gr or not gr.get("items"): return None
    subs = [x for x in gr["items"] if (x.get("teach") or x.get("check") or x.get("after"))]
    if not subs: return None
    pages = _chunk(subs, 4)
    last = None
    for pi, group in enumerate(pages, 1):
        last = _growth_page(prs, pal, gr, group, pi, len(pages))
    return last

def _growth_page(prs, pal, gr, subs, pi, parts):
    s = new_slide(prs)
    label(s, pal, _part("성장 방식", pi, parts))
    header(s, gr.get("head","") if pi == 1 else "", size=26)
    n = len(subs)
    gap = 0.30
    cw = (11.9 - (n-1)*gap)/n
    top, bottom = 2.30, 7.05
    colors = [pal["accent"], pal["accent2"], pal["navy"]]
    ROWS = [("이렇게 가르칩니다", "teach"), ("이렇게 확인합니다", "check"), ("3개월 후", "after")]
    # 본문 글자 크기: 가장 긴 칸 기준으로 한 번에 정한다(칸마다 들쭉날쭉하지 않게)
    band = (bottom - top - 0.62) / 3          # 과목명 줄 제외한 한 칸 높이
    vs = 13.5
    def _fits(sz):
        for it in subs:
            for _, k in ROWS:
                v = str(it.get(k,"") or "")
                if not v: continue
                need = wrap_lines(v, cw-0.36, sz) * (sz*0.0139*1.62) * 1.34
                if need > band - 0.34: return False
        return True
    while vs > 9.5 and not _fits(vs):
        vs -= 0.5
    for i, it in enumerate(subs):
        x = PADX + i*(cw+gap)
        c = colors[i % len(colors)]
        nm = str(it.get("name","")).strip()
        # 과목 헤더
        if nm:
            pw = min(cw, 0.6 + _disp_width(nm)*0.26)
            rect(s, x, top, pw, 0.50, c, radius=True)
            txt(s, x, top, pw, 0.50, [(nm, 15, True, "FFFFFF")],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        if i < n-1:
            vline(s, x+cw+gap/2, top, bottom-top, LINE, 1)
        y = top + 0.62
        for ri, (lab_t, k) in enumerate(ROWS):
            v = str(it.get(k,"") or "")
            if not v:
                y += band; continue
            txt(s, x, y, cw, 0.30, [(lab_t, 11.5, True, c)])
            txt(s, x, y+0.32, cw-0.18, band-0.36, [(v, vs, False, INK)], line_spacing=1.34)
            if ri < 2:
                hline(s, x, y+band-0.10, cw-0.15, LINE, 1)
            y += band
    return s


def slide_curriculum(prs, pal, d):
    cu = d.get("curriculum")
    if not cu or not cu.get("stages"): return None
    pages = _chunk(cu["stages"], 4)     # 초·중·고 3단이 기본, 4단까지 한 장
    last = None
    base = 0
    for pi, stages in enumerate(pages, 1):
        last = _curriculum_page(prs, pal, cu, stages, pi, len(pages), base)
        base += len(stages)
    return last

def _curriculum_page(prs, pal, cu, stages, pi, parts, base):
    s = new_slide(prs)
    label(s, pal, _part("단계별 커리큘럼", pi, parts))
    header(s, cu.get("head","") if pi == 1 else "", size=26)
    if cu.get("sub") and pi == 1:
        txt(s, PADX, 1.55, 11.9, 0.4, [(cu["sub"], 15, False, "556677")])
    n = len(stages)
    # 단계별 색: 초→중→고로 진해짐
    step_colors = _stage_colors(pal, base + n)[base:base+n]
    gap = 0.35
    bw = (11.9 - (n-1)*gap)/n
    # ── 항목이 길거나 많으면 글자를 줄여 박스 안에 담는다(문장은 자르지 않는다) ──
    MAX_BH = 3.25
    item_size = 11.5
    def _need(sz):
        mx = 1
        for st in stages:
            ln = sum(max(1, wrap_lines(str(it), bw-0.62, sz)) for it in st.get("items", []) or [""])
            mx = max(mx, ln)
        return 0.60 + mx*(sz*0.0139*1.62)*1.25 + 0.30
    bh = _need(item_size)
    while bh > MAX_BH and item_size > 9.0:
        item_size -= 0.5
        bh = _need(item_size)
    bh = min(bh, MAX_BH)
    base_y = 6.6           # 계단 맨 아래 박스의 바닥
    step_up = max(0.0, min(0.62, (6.6-2.30-bh)/(n-1))) if n>1 else 0  # 위치만 계단
    for i, st in enumerate(stages):
        x = PADX + i*(bw+gap)
        y = base_y - bh - i*step_up     # 높이는 같고 위치만 위로
        c = step_colors[i]
        soft = _tint_for(pal, c)
        # 태그라인 (박스 위)
        tag = st.get("tag","")
        if tag:
            tagw = 0.5 + len(tag)*0.30
            rect(s, x+0.05, y-0.62, tagw, 0.48, soft, radius=True)
            txt(s, x+0.05, y-0.60, tagw, 0.44, [(tag, 16, True, c)],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            vline(s, x+0.05+tagw/2, y-0.14, 0.14, c, 2.5)
        # 단계 박스 (모두 동일 높이)
        rect(s, x, y, bw, bh, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        rect(s, x, y, bw, 0.60, c, radius=True)
        rect(s, x, y+0.30, bw, 0.30, c)
        _nm = st.get("name") or st.get("title") or f"{i+1}단계"
        txt(s, x+0.22, y+0.08, bw-0.44, 0.45, [(_nm, 16, True, "FFFFFF")], anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, x+0.22, y+0.80, bw-0.44, bh-0.9, st.get("items", []), size=item_size,
                color=INK, dot_hex=c, line_spacing=1.2, para_space=4)
        # 상승 화살표 — 두 박스 사이 여백 중앙에 배치(박스에 가리지 않게)
        if i < n-1:
            y_next = base_y - bh - (i+1)*step_up
            mid_y = ((y + bh/2) + (y_next + bh/2)) / 2      # 두 박스 중심의 중간
            txt(s, x+bw+0.02, mid_y-0.22, gap-0.04, 0.44, [("↗", 20, True, MUTED)],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
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
    # ★ 진짜 표(python-pptx table 객체)로 생성 — 원장이 셀 단위로 편집 가능.
    #    (기존 박스+텍스트 레이어 방식 폐기)
    cols_ratio = [0.30, 0.34, 0.18, 0.18]  # 반 / 요일시간 / 강사 / 강의실
    tbl_w = 11.9
    headers = ["반", "요일 · 시간", "강사", "강의실"]
    n = len(rows)
    ncol = len(headers)

    # 행 높이 자동조절: 세로영역 ÷ (행수+헤더), 0.30~0.62 제한
    avail = 5.0
    rh = max(0.30, min(0.62, avail/(n+1)))
    hh = min(0.5, rh+0.04)
    body_font = 15 if n <= 7 else (14 if n <= 10 else 13)
    tbl_h = hh + rh*n

    gf = s.shapes.add_table(n+1, ncol, Inches(PADX), Inches(top), Inches(tbl_w), Inches(tbl_h))
    tbl = gf.table
    # 기본 표 스타일(밴딩/강조 행) 끄기 → 색을 직접 제어
    tbl.first_row = False
    tbl.horz_banding = False
    tbl.first_col = False
    # 기본 테마 표스타일(테두리/밴딩 색 오염) 제거 → "스타일 없음"으로
    try:
        tblPr = tbl._tbl.find(qn('a:tblPr'))
        if tblPr is not None:
            styleId = tblPr.find(qn('a:tableStyleId'))
            if styleId is not None:
                tblPr.remove(styleId)
    except Exception:
        pass

    # 열 너비
    for ci, ratio in enumerate(cols_ratio):
        tbl.columns[ci].width = Inches(ratio*tbl_w)
    # 헤더 행 높이
    tbl.rows[0].height = Inches(hh)
    for ri in range(1, n+1):
        tbl.rows[ri].height = Inches(rh)

    def _style_cell(cell, text, *, fs, bold, color, fill_hex, align_left=True):
        cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid(); cell.fill.fore_color.rgb = _rgb(fill_hex)
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if align_left else PP_ALIGN.CENTER
        run = p.add_run(); run.text = text
        run.font.size = Pt(fs); run.font.bold = bold
        run.font.color.rgb = _rgb(color); _apply_font(run)

    # 헤더
    for ci, htext in enumerate(headers):
        _style_cell(tbl.cell(0, ci), htext, fs=14, bold=True,
                    color="FFFFFF", fill_hex=header_hex)
    # 바디 (지브라: 짝수 흰색 / 홀수 연회색)
    for ri, row in enumerate(rows):
        zebra = CARD if (ri % 2 == 1) else "FFFFFF"
        for ci in range(ncol):
            cell_text = row[ci] if ci < len(row) else ""
            color = MUTED if cell_text.strip() == "-" else INK
            _style_cell(tbl.cell(ri+1, ci), cell_text, fs=body_font, bold=False,
                        color=color, fill_hex=zebra)

    # 표 테두리(가로 구분선만 은은하게) — 셀 하단 라인
    _table_hlines(tbl, LINE)
    return gf


def _table_hlines(tbl, hex_line):
    """각 셀 하단에 가는 가로선 추가(표 편집성 유지하면서 구분감)."""
    from pptx.oxml.ns import qn as _qn
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            # 아래쪽 테두리 lnB
            for tag in ("a:lnB",):
                existing = tcPr.find(_qn(tag))
                if existing is not None:
                    tcPr.remove(existing)
                ln = tcPr.makeelement(_qn(tag), {"w": "6350", "cap": "flat"})  # 6350 EMU ≈ 0.5pt
                fill = ln.makeelement(_qn("a:solidFill"), {})
                clr = fill.makeelement(_qn("a:srgbClr"), {"val": hex_line})
                fill.append(clr); ln.append(fill)
                tcPr.append(ln)

def slide_specials(prs, pal, d):
    sp = d.get("specials")
    if not sp or not sp.get("items"): return None
    pages = _chunk(sp["items"], 4)
    last = None
    for pi, items in enumerate(pages, 1):
        last = _specials_page(prs, pal, sp, items, pi, len(pages))
    return last

def _specials_page(prs, pal, sp, items, pi, parts):
    s = new_slide(prs)
    label(s, pal, _part("특별 프로그램", pi, parts))
    header(s, sp.get("head","") if pi == 1 else "", size=26)
    n = len(items)
    gx, gy = PADX, 2.25
    # 개수에 맞춰 배치: 1개=전폭 / 2개=좌우 / 3개=가로 3단 / 4개=사분면
    if n <= 3:
        gap = 0.34
        cw = (11.9 - (n-1)*gap)/n
        ch = 4.5
        for i, it in enumerate(items):
            x = gx + i*(cw+gap)
            rect(s, x, gy, cw, ch, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
            paras(s, x+0.3, gy+0.35, cw-0.6, 0.6, [[
                (it.get("no","")+"  ", 20, True, pal["accent"]),
                (it["title"], 19, True, INK_STRONG)]], line_spacing=1.2)
            dsize = 14.5
            while dsize > 11.0 and wrap_lines(it.get("desc",""), cw-0.6, dsize)*(dsize*0.0139*1.62)*1.35 > (ch-1.5):
                dsize -= 0.5
            txt(s, x+0.3, gy+1.15, cw-0.6, ch-1.5, [(it.get("desc",""), dsize, False, INK)], line_spacing=1.35)
        return s
    cw, ch = 11.9/2, 2.15
    vline(s, gx+cw, gy, ch*2, LINE, 1)
    hline(s, gx, gy+ch, 11.9, LINE, 1)
    for i, it in enumerate(items):
        r=i//2; c=i%2
        x = gx + c*cw + 0.3; y = gy + r*ch + 0.3
        paras(s, x, y, cw-0.6, 0.5, [[
            (it.get("no","")+"  ", 20, True, pal["accent"]),
            (it["title"], 19, True, INK_STRONG)]])
        txt(s, x, y+0.65, cw-0.6, ch-1.0, [(it.get("desc",""), 14.5, False, INK)], line_spacing=1.35)
    return s

def slide_management(prs, pal, d):
    mg = d.get("management")
    if not mg or not mg.get("columns"): return None
    # ★과목이 3개·4개여도 버리지 않는다. 한 장에 2개씩 싣고 나머지는 다음 장으로.
    _cols = [c for c in mg["columns"] if c.get("rows")]
    # 3과목까지는 한 장(3열), 4과목 이상은 열이 좁아 읽히지 않으므로 2개씩 나눈다
    pages = [_cols] if len(_cols) <= 3 else _chunk(_cols, 2)
    if not pages: return None
    last = None
    for pi, cols in enumerate(pages, 1):
        last = _management_page(prs, pal, mg, cols, pi, len(pages))
    return last

def _management_page(prs, pal, mg, cols, pi, parts):
    s = new_slide(prs)
    label(s, pal, _part("수업 관리", pi, parts))
    header(s, mg.get("head","") if pi == 1 else "", size=26)
    single = (len(cols) == 1)
    key_colors = [pal["accent"], pal["accent2"], pal["navy"], pal["accent"]]

    # 타입 선택: mg["style"] = "row" | "box" | "auto"
    # auto → 과목 1개면 박스(오른쪽 여백 방지), 2개면 줄 타입, 3개 이상이면 열이 좁아지므로 박스
    style = (mg.get("style") or "auto").lower()
    if style == "auto":
        style = "box" if (single or len(cols) >= 3) else "row"
    if len(cols) >= 3:
        style = "box"      # 3열 이상에서 줄 타입은 설명 폭이 나오지 않는다

    if style == "box":
        return _management_box(s, pal, cols, single, key_colors)
    return _management_row(s, pal, cols, single, key_colors)


def _management_row(s, pal, cols, single, key_colors):
    # 줄 타입: 좌 키박스 + 우 설명. 과목 1개면 전체폭 대신 왼쪽 정리(오른쪽 여백 축소)
    total = 11.9; gap = 0.7
    if single:
        colw = 7.6          # 전체폭(11.9) 대신 왼쪽에 모아 오른쪽 여백 줄임
        col_x = [PADX]
    else:
        colw = (total-gap)/2
        col_x = [PADX, PADX+colw+gap]
        vline(s, PADX+colw+gap/2, 2.15, 4.4, LINE, 1)
    max_rows = max(len(c["rows"]) for c in cols)
    rows_top = 2.98; rows_bottom = 7.0
    pitch = min(1.15, (rows_bottom - rows_top)/max(max_rows,1))
    for idx, col in enumerate(cols):
        x = col_x[idx]; kc = key_colors[idx]
        txt(s, x, 2.15, colw, 0.5, [(col["name"], 19, True, kc)])
        hline(s, x, 2.72, colw, LINE, 1)
        kw = 1.15
        for j, row in enumerate(col["rows"]):
            yy = rows_top + j*pitch
            vv = row["v"]
            vsize = fit_size(vv, 14, 11, lambda sz: int(((colw-kw-0.2)/(sz*0.017))*2.4))
            # 본문은 위에서부터 흐르게 하고, 키 단추를 본문 첫 줄 높이에 맞춘다
            LINE_H = vsize * 0.0175 * 1.28          # 한 줄 높이(인치 근사)
            txt(s, x+kw+0.2, yy, colw-kw-0.2, pitch-0.05, [(vv, vsize, False, INK)],
                line_spacing=1.28, anchor=MSO_ANCHOR.TOP)
            pill_h = 0.44
            pill_y = yy + max(0.0, (LINE_H - pill_h)/2)   # 첫 줄 중앙과 단추 중앙 일치
            rect(s, x, pill_y, kw, pill_h, kc, radius=True)
            ktext = row["k"]
            ksize = 12.5 if len(ktext) <= 5 else (11 if len(ktext) <= 7 else 10)
            txt(s, x+0.05, pill_y+0.02, kw-0.1, pill_h-0.04, [(ktext, ksize, True, "FFFFFF")],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    return s


def _management_box(s, pal, cols, single, key_colors):
    # 박스 타입: 항목마다 카드. 과목 1개면 2열 그리드로 채워 오른쪽 여백 방지.
    area_top = 2.20 if len(cols) <= 2 else 2.02
    area_bottom = 7.05 if len(cols) <= 2 else 7.18
    if single:
        col = cols[0]; kc = key_colors[0]
        rows = col["rows"]
        # 제목
        txt(s, PADX, area_top, 11.9, 0.5, [(col["name"], 19, True, kc)])
        hline(s, PADX, area_top+0.55, 11.9, LINE, 1)
        grid_top = area_top + 0.80
        # 2열 그리드
        gcols = 2 if len(rows) >= 3 else 1
        gap = 0.5
        cw = (11.9 - gap*(gcols-1)) / gcols
        nrows = (len(rows)+gcols-1)//gcols
        ch = min(1.9, (area_bottom - grid_top - gap*(nrows-1)) / max(nrows,1))
        for i, row in enumerate(rows):
            r = i // gcols; c = i % gcols
            x = PADX + c*(cw+gap)
            y = grid_top + r*(ch+gap)
            _mgmt_card(s, pal, kc, pal.get("soft","F2F5F4"), x, y, cw, ch, row)
        return s
    # 2~4과목: 과목마다 한 열, 열 안에서 항목 카드 세로 배치
    n = len(cols)
    total = 11.9
    gap = 0.7 if n <= 2 else (0.45 if n == 3 else 0.34)
    colw = (total - gap*(n-1))/n
    col_x = [PADX + i*(colw+gap) for i in range(n)]
    soft_pool = [pal.get("soft","F2F5F4"), pal.get("soft2", pal.get("soft","F2F5F4"))]
    name_size = 19 if n <= 2 else (17 if n == 3 else 15)
    for idx, col in enumerate(cols):
        x = col_x[idx]; kc = key_colors[idx % len(key_colors)]
        txt(s, x, area_top, colw, 0.5, [(col["name"], name_size, True, kc)], line_spacing=1.15)
        hline(s, x, area_top+(0.55 if n <= 2 else 0.48), colw, LINE, 1)
        if idx < n-1 and n >= 3:
            vline(s, x+colw+gap/2, area_top, area_bottom-area_top, LINE, 1)
        rows = col["rows"]
        grid_top = area_top + (0.80 if n <= 2 else 0.62)
        vgap = 0.3 if n <= 2 else 0.14
        ch = min(1.15, (area_bottom - grid_top - vgap*(len(rows)-1)) / max(len(rows),1))
        for j, row in enumerate(rows):
            y = grid_top + j*(ch+vgap)
            _mgmt_card(s, pal, kc, soft_pool[idx % 2], x, y, colw, ch, row)
    return s


def _mgmt_card(s, pal, kc, soft_bg, x, y, w, h, row):
    # 카드 배경 없음: 소제목 pill(컬러)만 + 설명. (연한 박스 배경 제거)
    ktext = row["k"]
    ksize = 13 if len(ktext) <= 5 else (11.5 if len(ktext) <= 7 else 10.5)
    # 키 pill (열이 좁으면 pill과 글자를 함께 줄인다)
    kw = min(w-0.15, 0.42 + _disp_width(ktext)*0.20)
    while kw > w-0.15 and ksize > 9:
        ksize -= 0.5; kw = min(w-0.15, 0.42 + _disp_width(ktext)*0.18)
    ph = 0.34 if h < 0.95 else 0.40
    rect(s, x, y+0.06, kw, ph, kc, radius=True)
    txt(s, x, y+0.06, kw, ph, [(ktext, ksize, True, "FFFFFF")],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    # 설명 — 칸 높이 안에 들어오도록 자동 축소(넘쳐서 아래 칸을 침범하지 않게)
    vv = row["v"]
    top_off = 0.44 if h < 0.95 else 0.54
    vsize = fit_box(vv, w-0.12, max(0.22, h-top_off), 13.5, 9.0, line_spacing=1.26)
    txt(s, x+0.02, y+top_off, w-0.10, max(0.22, h-top_off), [(vv, vsize, False, INK)],
        line_spacing=1.26, anchor=MSO_ANCHOR.TOP)

def slide_admission(prs, pal, d):
    ad = d.get("admission")
    if not ad or not ad.get("steps"): return None
    s = new_slide(prs)
    label(s, pal, "입학 절차")
    header(s, ad.get("head",""), size=26)
    steps = ad["steps"]
    n = len(steps)
    top = 2.20
    bottom = 7.10
    avail = bottom - top
    slot = avail / n
    has_bridge = bool(ad.get("bridge"))

    num_size = 58 if n <= 2 else (48 if n == 3 else 40)
    title_size = 19 if n <= 3 else 17

    text_x = PADX + 1.7
    text_w = 11.9 + PADX - text_x
    title_h = 0.40
    sep_gap = 0.14

    for i, st in enumerate(steps):
        y = top + i*slot
        hline(s, PADX, y, 11.9, LINE, 1)
        # 브릿지 — 구분선 우측 끝에 작은 캡션(왼쪽 설명과 좌우로 분리 → 겹칠 수 없음)
        if i > 0 and has_bridge:
            txt(s, 11.9+PADX-4.4, y-0.30, 4.4, 0.28,
                [("↓ ", 12, False, MUTED), (ad["bridge"], 11.5, True, pal["accent"])],
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, wrap=False)

        items = st.get("items", [])
        body = "  ".join(items) if items else ""
        # 설명 세로 영역: 슬롯에서 제목·여백 제외(브릿지는 다음 슬롯 상단 캡션이라 이 슬롯을 안 먹음)
        desc_top = y + sep_gap + title_h + 0.06
        desc_bottom = y + slot - 0.14
        desc_h = max(0.30, desc_bottom - desc_top)

        # 폰트를 desc_h 안에 들어오도록 자동 축소
        base_size = 14 if n <= 3 else 12.5
        min_size = 10.5
        line_sp = 1.30
        size = base_size
        while size > min_size:
            lines = wrap_lines(body, text_w, size)
            need = lines * (size*0.0139*1.55) * line_sp
            if need <= desc_h:
                break
            size -= 0.5

        # 제목 + 본문이 차지하는 실제 높이를 계산해, 숫자를 그 묶음의 세로 중앙에 맞춘다
        body_h = 0.0
        if body:
            lines = wrap_lines(body, text_w, size)
            body_h = lines * (size*0.0139*1.55) * line_sp
        block_top = y + sep_gap
        block_h = title_h + (0.06 + body_h if body else 0)
        block_mid = block_top + block_h/2

        txt(s, PADX, block_mid-0.62, 1.5, 1.24, [(st["no"], num_size, True, pal["accent"])],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, text_x, y+sep_gap, text_w, 0.5, [(st["title"], title_size, True, INK_STRONG)])
        if body:
            txt(s, text_x, desc_top, text_w, desc_h,
                [("– ", size, False, INK), (body, size, False, INK)],
                line_spacing=line_sp, anchor=MSO_ANCHOR.TOP)
    return s

def slide_rules(prs, pal, d):
    ru = d.get("rules")
    if not ru or not ru.get("items"): return None
    s = new_slide(prs)
    items = ru["items"]
    n = len(items)
    rx = 4.3; rw = 11.9+PADX-rx
    kw = 0.9
    vw = rw - kw - 0.28   # 설명 폭

    # 본문 영역(헤더 없는 슬라이드): 슬라이드 세로 정중앙(7.5in의 절반=3.75)을 기준으로 중앙 정렬
    AREA_TOP = 0.75
    AREA_BOTTOM = 7.4
    AREA_MID = 3.75   # 슬라이드(13.3×7.5) 세로 정중앙 — 항목이 몇 개든 이 지점을 블록 중앙으로

    # 각 항목 실제 높이(설명 줄 수 반영, 타이트하게)
    vsizes = []; text_h = []
    for it in items:
        vv = it["v"]
        vsize = fit_size(vv, 15, 12, lambda sz: int(((vw)/(sz*0.017))*2.0))
        lines = wrap_lines(vv, vw, vsize)
        th = lines * (vsize*0.0139*1.55) * 1.3      # 텍스트 자체 높이
        vsizes.append(vsize)
        text_h.append(max(0.34, th))

    # 항목 간 간격: 항목이 많으면 좁게, 적으면 넉넉하게(상하 중앙은 유지)
    content_h = sum(text_h)
    # 사용 가능한 세로에서 텍스트를 뺀 나머지를 (n+? ) 간격으로 배분하되, 상한/하한
    free = (AREA_BOTTOM - AREA_TOP) - content_h
    gap = free / (n)                     # 항목 사이+양끝을 고르게
    gap = max(0.14, min(0.9, gap))       # 너무 붙거나 너무 벌어지지 않게
    block = content_h + gap*(n-1)

    # 블록이 영역보다 크면(항목 매우 많음) 간격을 더 줄여 강제로 맞춤
    if block > (AREA_BOTTOM - AREA_TOP):
        gap = max(0.08, ((AREA_BOTTOM - AREA_TOP) - content_h) / max(n-1,1))
        block = content_h + gap*(n-1)

    top = AREA_MID - block/2
    if top < AREA_TOP: top = AREA_TOP
    mid = top + block/2

    # 좌 제목 — 블록 세로 중앙에 맞춤
    label(s, pal, "학원 규정", x=PADX, y=mid-1.02)
    _rhead = ru.get("head") or "학원 규정 안내"
    _rsz = fit_box(_rhead, 3.35, 1.7, 31, 19, line_spacing=1.16)
    txt(s, PADX, mid-0.70, 3.35, 1.9,
        [(_rhead, _rsz, True, INK_STRONG)],
        line_spacing=1.16, anchor=MSO_ANCHOR.TOP)

    yy = top
    for i, it in enumerate(items):
        th = text_h[i]
        cyi = yy + th/2
        rect(s, rx, cyi-0.23, kw, 0.46, pal["soft"], radius=True)
        txt(s, rx, cyi-0.23, kw, 0.46, [(it["k"], 13, True, pal["pill_ink"])],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        txt(s, rx+kw+0.28, cyi-th/2-0.05, vw, th+0.1, [(it["v"], vsizes[i], False, INK)],
            line_spacing=1.3, anchor=MSO_ANCHOR.MIDDLE)
        yy += th + gap
    return s

def slide_faq(prs, pal, d):
    fq = d.get("faq")
    if not fq or not fq.get("items"): return None
    pages = _chunk(fq["items"], 6)
    last = None
    for pi, items in enumerate(pages, 1):
        last = _faq_page(prs, pal, fq, items, pi, len(pages))
    return last

def _faq_page(prs, pal, fq, _items, pi, parts):
    fq = dict(fq); fq["items"] = _items
    s = new_slide(prs)
    label(s, pal, _part("자주 묻는 질문", pi, parts))
    header(s, fq.get("head","궁금한 점을 미리 확인하세요") if pi == 1 else "", size=26)
    items = fq["items"][:6]
    gx, gy = PADX, 2.15
    gap = 0.3
    ncol = 2 if len(items) <= 4 else 3
    cw = (11.9-gap*(ncol-1))/ncol
    nrow = (len(items)+ncol-1)//ncol
    ch = min(2.30, (7.12-gy-0.22*(nrow-1))/max(nrow,1))
    for i, it in enumerate(items):
        r=i//ncol; c=i%ncol
        x = gx + c*(cw+gap); y = gy + r*(ch+0.22)
        rect(s, x, y, cw, ch, CARD, radius=True, line_hex=CARD_LINE, line_w=1)
        # Q
        q = it["q"]; a = it["a"]
        qsize = fit_box(q, cw-0.62, 0.74, 15, 11, line_spacing=1.25)
        paras(s, x+0.28, y+0.24, cw-0.56, 0.8,
              [[("Q. ", qsize, True, pal["accent2"]), (q, qsize, True, INK_STRONG)]],
              line_spacing=1.25, hang=0.32)
        asize = fit_box(a, cw-0.62, ch-1.10, 13.5, 9.5, line_spacing=1.3)
        paras(s, x+0.28, y+0.98, cw-0.56, ch-1.10,
              [[("A. ", asize, True, pal["accent"]), (a, asize, False, "4A5568")]],
              line_spacing=1.3, hang=0.30)
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
    _rows = []
    if a.get("phone"):  _rows.append(("상담 문의", a.get("phone","")))
    if a.get("phone2"): _rows.append(("학원 전화", a.get("phone2","")))
    if a.get("kakao"):  _rows.append(("카카오 채널", a.get("kakao","")))
    if a.get("address_short") or a.get("location"):
        _rows.append(("위치", a.get("address_short","") or a.get("location","")))
    if a.get("hours"):  _rows.append(("운영시간", a.get("hours","")))
    csize = 15 if len(_rows) <= 3 else 13
    for k, v in _rows:
        contact_paras.append([(k + "    ", csize, True, pal["navy"]), (v, csize, False, INK)])
    if contact_paras:
        paras(s, PADX, 2.55, 6.0, 2.05, contact_paras, line_spacing=1.28,
              para_space=(8 if len(_rows) <= 3 else 4))
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
    slide_identity(prs, pal, data)      # 정체성·교육철학
    slide_achievements(prs, pal, data)
    slide_targets(prs, pal, data)
    slide_classes(prs, pal, data)       # 반 구성 상세(전량)
    slide_curriculum(prs, pal, data)
    slide_growth(prs, pal, data)        # 과목별 성장 방식

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

    _force_theme_font(prs)
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
