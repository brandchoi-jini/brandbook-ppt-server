# -*- coding: utf-8 -*-
"""좌표 기반 3단 리플렛 빌더 (GLT 템플릿 레이아웃을 코드로 재현).

- 세원 teal/yellow 템플릿의 배치·색·서체를 좌표로 그대로 그린다.
- 템플릿 방식과 달리 '빈 칸'이 없다: 데이터 개수만큼만 카드를 생성 → 데이터가
  얼마든 세원 잔존/빈칸 문제가 구조적으로 발생하지 않는다.
- 로고·표지사진 자리를 명시적으로 둔다(assets 없으면 안내 박스).
- 11.7 x 8.27in (A4 가로), 3등분 패널. Pretendard.

입력: to_schema_v3 가 만든 v3 스키마.
출력: 2슬라이드 PPTX (바깥면/안쪽면), 모든 도형 편집 가능.
"""
from __future__ import annotations

import io
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt

import io as _io
import urllib.request as _urlreq

# URL/경로/base64 이미지를 메모리로 가져오기 (실패 시 None) — v3 빌더와 동일 방식
_IMG_CACHE = {}
def _fetch_image(url):
    if not url:
        return None
    if url in _IMG_CACHE:
        return _IMG_CACHE[url]
    try:
        if url.startswith("data:"):
            import base64 as _b64
            _, _, b64 = url.partition(",")
            data = _b64.b64decode(b64)
        elif url.startswith("http"):
            req = _urlreq.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            data = _urlreq.urlopen(req, timeout=8).read()
        else:
            with open(url, "rb") as f:
                data = f.read()
        _IMG_CACHE[url] = data
        return data
    except Exception:
        _IMG_CACHE[url] = None
        return None

# A4 가로 3단 접지. 297mm = 11.6929in. 3등분이 아니라 접지 규격으로 나눈다.
# 안으로 접혀 들어가는 패널을 3mm 좁게(97mm) 만들어야 접었을 때 맞는다.
MM = 1.0 / 25.4  # mm → inch
W_IN, H_IN = 297 * MM, 210 * MM   # 11.6929 x 8.2677 in

# 패널 폭(mm) — 바깥면과 안쪽면이 좌우 반대다.
#  안쪽면: 좌100 · 중100 · 우97 (우측이 맨 먼저 안으로 접힘)
#  바깥면: 좌97 · 중100 · 우100 (안쪽면 우측 패널의 뒷면이 좌측 97)
PANELS_INSIDE_MM  = [100, 100, 97]
PANELS_OUTSIDE_MM = [97, 100, 100]

def _panel_geom(widths_mm):
    """패널 폭 배열(mm) → [(x_in, w_in), ...] 누적 좌표."""
    out, cx = [], 0.0
    for w in widths_mm:
        w_in = w * MM
        out.append((cx, w_in))
        cx += w_in
    return out

PANEL = 100 * MM  # 기준 패널 폭(3.937in) — 내부 여백 계산 기본값
# 서체는 Pretendard SemiBold 한 가지만 쓴다(Bold/Regular 구분 폐기).
# python-pptx 는 굵기를 이름으로 지정하므로 패밀리명에 SemiBold 를 박고,
# run.font.bold 는 항상 False 로 둔다(True 면 SemiBold 위에 인조 굵게가 겹침).
FONT = "Pretendard SemiBold"

# 좌우 여백 기준(모든 패널 공통). 라벨·헤더·본문·번호칩이 이 세로선에 맞춰 정렬된다.
PAD = 0.30            # 패널 안쪽 좌측 여백
LEFT = PAD            # 콘텐츠 왼쪽 기준(패널 x에 더해 사용: x + LEFT)
# 콘텐츠 폭은 가장 좁은 패널(97mm=3.819in) 기준으로 통일 → 어느 패널에서도 안 넘침
CW = 97 * MM - PAD * 2  # ≈ 3.22in
MAX_ADMISSION = 6     # 입학절차 표시 상한(워드 5단계까지 그대로 수용)
CHIP = 0.34           # 번호칩 지름
CHIP_GAP = 0.14       # 칩과 텍스트 사이 간격
TEXT_X = LEFT + CHIP + CHIP_GAP   # 칩 오른쪽 텍스트 시작
TEXT_W = CW - CHIP - CHIP_GAP     # 칩 있는 줄의 텍스트 폭

# 팔레트: 세원 teal/yellow 실측색 그대로 + 확장용 navy.
PALETTES = {
    "sewon_teal": {
        "cover_bg": "17345C",   # 표지 패널 배경(딥)
        "cover_ink": "FFFFFF",  # 표지 위 글자
        "badge": "0AA6B5",      # 과목 뱃지/번호칩(주 청록)
        "accent2": "2D74DA",    # 보조 블루(커리큘럼 라벨·관리 번호)
        "label": "0AA6B5",      # 섹션 라벨(청록)
        "head": "17345C",       # 헤더(딥네이비)
        "title": "17345C",      # 카드 제목
        "body": "5C6B76",       # 설명 회색
        "soft": "E8F7F8",       # 연한 박스(FAQ답·관리카드)
        "card_line": "E3E9EE",  # 카드 테두리
        "footer": "8A97A0",     # 푸터 회색
        "paper": "FFFFFF",
    },
    "sewon_yellow": {
        "cover_bg": "F5C84C",
        "cover_ink": "5C3B20",
        "badge": "F29B21",
        "accent2": "27A9B7",
        "label": "F29B21",
        "head": "5C3B20",
        "title": "5C3B20",
        "body": "716254",
        "soft": "FFF3C9",
        "card_line": "EDE4CB",
        "footer": "A08A6A",
        "paper": "FFFFFF",
    },
    "navy_amber": {
        "cover_bg": "22375F",
        "cover_ink": "FFFFFF",
        "badge": "22375F",
        "accent2": "D98A1F",
        "label": "D98A1F",
        "head": "22375F",
        "title": "22375F",
        "body": "5C6B76",
        "soft": "EEF2F8",
        "card_line": "E3E9EE",
        "footer": "8A97A0",
        "paper": "FFFFFF",
    },
}
# 약칭
PALETTES["teal"] = PALETTES["sewon_teal"]
PALETTES["yellow"] = PALETTES["sewon_yellow"]

BASE_DIR = Path(__file__).resolve().parent


# ── 유틸 ──────────────────────────────────────────────
def _rgb(h: str) -> RGBColor:
    h = h.strip().lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _set_font(run, name=None, size=None, bold=False, color=None):
    """서체를 latin·ea·cs 세 슬롯에 모두 지정한다.
    ★python-pptx 의 run.font.name 은 <a:latin> 만 채운다.
      한글은 <a:ea>(East Asian) 슬롯을 보므로, 이걸 비워두면
      테마 기본서체(Calibri)로 떨어지고 PowerPoint 가 임의의 한글 폰트
      (맑은 고딕 등)로 대체한다 → 지정한 Pretendard 가 안 나온다."""
    name = name or FONT
    f = run.font
    f.name = name
    if size is not None:
        f.size = Pt(size)
    f.bold = bool(bold)
    if color is not None:
        f.color.rgb = _rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ("ea", "cs"):
        el = rPr.find(_A_NS + tag)
        if el is None:
            el = rPr.makeelement(_A_NS + tag, {})
            rPr.append(el)
        el.set("typeface", name)


def _s(v: Any, default: str = "") -> str:
    if v is None:
        return default
    if isinstance(v, str):
        return re.sub(r"[ \t]+", " ", v).strip() or default
    return str(v).strip() or default


def _first(*vals: Any, default: str = "") -> str:
    for v in vals:
        s = _s(v)
        if s:
            return s
    return default


def _sentences(text):
    """한국어 문장 분리. 종결부호(.!?·。) 뒤에서 끊고 부호는 살린다."""
    s = _s(text)
    if not s:
        return []
    parts = re.split(r"(?<=[.!?。])\s+", s)
    return [p.strip() for p in parts if p.strip()]


def _is_complete_sentence(x):
    """완결된 서술문인지. 표지 카피 판정용.
    종결어미로 끝나야 하고, 마침표만 붙은 명사형("~하는 것.")은 미완성으로 본다."""
    x = _s(x).strip()
    if not x:
        return False
    if not re.search(r"(합니다|입니다|습니다|됩니다|드립니다|십니다|해요|이에요|예요|한다|이다)[.!?]?$", x):
        return False
    body = re.sub(r"[.!?]$", "", x)
    return not re.search(
        r"(것|점|중|뿐|까지|부터|처럼|만큼|위해|통해|대한|위한|하는|되는|있는|없는)$", body)


def _cover_copy(text, max_chars):
    """표지 한마디 전용 필터.
    ★완결된 문장만 남긴다. 하나도 없으면 빈 문자열(표지에서 생략).
      미완성 조각을 표지에 올리느니 슬로건만 보여주는 편이 낫다."""
    s = _s(text)
    if not s:
        return ""
    good = [x for x in _sentences(s) if _is_complete_sentence(x)]
    if not good:
        return ""
    out = good[0]
    for nxt in good[1:]:
        if _dwidth(out) + 1 + _dwidth(nxt) > max_chars:
            break
        out = out + " " + nxt
    return out if _dwidth(out) <= max_chars * 1.35 else ""


def _trim_sentences(text, max_chars):
    """표시폭이 max_chars 를 넘지 않도록 <문장 단위로> 줄인다.
    ★완결된 문장만 남긴다. 명사형으로 끝나는 조각("~하는 것.")은 버린다.
      표지에 미완성 문장이 올라가는 것을 막기 위함(경희궁 사례).
    - 첫 문장이 완결형이면 항상 보존한다.
    - 단어 중간을 자르거나 '…' 를 붙이지 않는다."""
    s = _s(text)
    if not s:
        return s
    sents = _sentences(s)
    if not sents:
        return s

    def _complete(x):
        x = x.strip()
        if not re.search(r"(합니다|입니다|습니다|됩니다|드립니다|십니다|해요|이에요|예요|한다|이다)[.!?]?$", x):
            return False
        # 마침표만 붙은 명사형 종결은 미완성으로 본다
        body = re.sub(r"[.!?]$", "", x)
        return not re.search(r"(것|점|중|뿐|까지|부터|처럼|만큼|위해|통해|대한|위한|하는|되는|있는|없는)$", body)

    good = [x for x in sents if _complete(x)]
    if not good:
        # 완결 문장이 없으면 전체가 max 안에 들어갈 때만 통과, 아니면 비운다
        return s if _dwidth(s) <= max_chars else ""
    out = good[0]
    for nxt in good[1:]:
        if _dwidth(out) + 1 + _dwidth(nxt) > max_chars:
            break
        out = out + " " + nxt
    return out


def _clip(text: Any, n: int, suffix: str = "…") -> str:
    s = _s(text)
    if len(s) <= n:
        return s
    return s[: max(1, n - len(suffix))].rstrip() + suffix


def _items(value: Any) -> List[Dict[str, Any]]:
    """dict({items|stages|steps}) 또는 list 를 dict 리스트로 정규화."""
    if isinstance(value, dict):
        value = value.get("items") or value.get("stages") or value.get("steps") or []
    if not isinstance(value, list):
        return []
    out = []
    for x in value:
        if isinstance(x, dict):
            out.append(x)
        elif isinstance(x, str) and x.strip():
            out.append({"title": x.strip(), "desc": ""})
    return out


def _dwidth(s: str) -> float:
    """표시 폭(한글=1.0, 영숫자·공백=0.5). 한글/영문 혼용 문장의 줄 길이를
    글자 수가 아니라 실제 차지하는 폭으로 재기 위한 근사."""
    w = 0.0
    for ch in s:
        w += 1.0 if ord(ch) > 0x2E80 else 0.5
    return w


def _wrap(text: str, per_line: int, max_lines: int = 0) -> str:
    """어절 단위 줄바꿈. ★글자를 절대 버리지 않는다.
    기존 구현은 max_lines 를 넘으면 남은 어절을 통째로 버려서
    '아이의 현재 수준을' 처럼 문장이 말없이 끊겼다(경희궁 사례).
    이제 max_lines 는 무시하고 전문을 반환한다 — 넘치는 분량은
    _fit_text 가 글자 크기를 줄여 흡수한다."""
    s = _s(text)
    if not s:
        return s
    if "\n" in s:
        return s
    limit = max(4, per_line)
    lines, cur = [], ""
    for w in s.split(" "):
        # 한 어절이 한 줄보다 길면(예: 긴 영문 단어·URL) 어쩔 수 없이 강제로 끊는다.
        while _dwidth(w) > limit:
            take = ""
            for ch in w:
                if _dwidth(take + ch) > limit:
                    break
                take += ch
            if not take:
                break
            if cur:
                lines.append(cur)
                cur = ""
            lines.append(take)
            w = w[len(take):]
        if cur and _dwidth(cur) + 0.5 + _dwidth(w) > limit:
            lines.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur:
        lines.append(cur)
    return "\n".join(lines)


def _char_w(size):
    """1.0폭 글자(한글 1자)가 차지하는 <가로 길이(inch)>.
    ★단위 주의: 포인트당 비율이 아니라 절대 길이를 돌려준다.
      (이전에 비율만 돌려주다 호출부에서 size 를 곱하지 않아
       한 줄에 233자가 들어간다고 계산되는 버그가 있었다.)
    큰 글자일수록 자간이 상대적으로 넓어 계수를 키운다."""
    if size >= 22:
        return size * 0.0162
    if size >= 17:
        return size * 0.0155
    return size * 0.0148


def _fit_block(text, w_in, max_h, size, *, min_size=7.0, line_spacing=1.10,
               pad=0.10):
    """(size, wrapped, height) 를 돌려준다.
    ★height 는 wrapped 를 담기에 <충분한> 값이다. max_h 로 잘라내지 않는다.
      (예전엔 min(max_h, …) 로 깎아서, 줄 수가 max_h 를 넘길 때
       '박스는 작은데 글자는 많은' 상태가 되어 넘침이 났다.)
      호출부는 반환된 height 를 그대로 쓰되, max_h 를 넘겼는지 보고
      글을 줄일지 결정한다."""
    size, wrapped = _fit_size(text, w_in, max_h, size,
                              min_size=min_size, line_spacing=line_spacing)
    n = wrapped.count("\n") + 1
    h = n * (size * line_spacing) / 72.0 + pad
    return size, wrapped, h


def _fit_size(text, w_in, h_in, size, *, min_size=6.5, line_spacing=1.04):
    """박스(w_in × h_in)에 text 전문이 들어가는 최대 글자 크기를 찾는다.
    한 줄에 들어가는 표시 폭 ≈ (박스폭 - 여백) / (글자폭). Pretendard 기준
    글자 1.0폭 ≈ 0.0148in per pt. 잘라내는 대신 크기로 맞춘다.
    ★줄바꿈이 이미 있는 텍스트도 각 줄을 다시 랩해서 실제 줄 수를 센다."""
    s = _s(text)
    if not s:
        return size, s
    usable_w = max(0.4, w_in - 0.10)
    usable_h = max(0.15, h_in)
    hard = s.split("\n")
    cand = size
    while cand >= min_size:
        per_line = max(4, int(usable_w / _char_w(cand)))
        out, n_lines = [], 0
        for seg in hard:
            w = _wrap(seg, per_line)
            out.append(w)
            n_lines += w.count("\n") + 1
        need_h = n_lines * (cand * line_spacing) / 72.0
        if need_h <= usable_h:
            return cand, "\n".join(out)
        cand -= 0.5
    # 작성자가 넣은 줄바꿈을 지키면 안 들어가는 경우 → 줄바꿈을 풀고 다시 흘린다
    if len(hard) > 1:
        joined = " ".join(x.strip() for x in hard if x.strip())
        cand = size
        while cand >= min_size:
            per_line = max(4, int(usable_w / _char_w(cand)))
            w = _wrap(joined, per_line)
            if (w.count("\n") + 1) * (cand * line_spacing) / 72.0 <= usable_h:
                return cand, w
            cand -= 0.5
    per_line = max(4, int(usable_w / _char_w(min_size)))
    return min_size, "\n".join(_wrap(seg, per_line) for seg in hard)


# ── 도형/텍스트 ───────────────────────────────────────
def _strip_style(shape):
    """python-pptx auto_shape가 자동 삽입하는 <p:style> 제거(그림자·테마색 오염 방지)."""
    try:
        el = shape._element
        for st in el.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}style"):
            st.getparent().remove(st)
        sp = el.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}style")
        if sp is not None:
            sp.getparent().remove(sp)
    except Exception:
        pass


def _rect(slide, x, y, w, h, fill=None, line=None, radius=False, line_w=0.75):
    shp = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shp, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = _rgb(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = _rgb(line)
        sh.line.width = Pt(line_w)
    _strip_style(sh)
    return sh


def _text(slide, text, x, y, w, h, *, size=11, color="222222", bold=False,
          align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, line_spacing=1.04,
          autofit=True, cap=None, fit=True, min_size=6.5):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = valign
    txt = _s(text)
    if cap:
        txt = _clip(txt, cap)
    # ★ 잘라내지 않고 크기로 맞춘다 — 문장 잘림(경희궁 사례) 방지.
    if fit and txt:
        size, txt = _fit_size(txt, w, h, size,
                              min_size=min_size, line_spacing=line_spacing)
    lines = txt.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln
        # SemiBold 단일 서체 — 인조 굵게를 걸지 않는다. 위계는 크기·색으로만 준다.
        _set_font(r, size=size, bold=False, color=color)
    if autofit:
        try:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass
    return box


def _chip(slide, x, y, d, number, c):
    """번호 원형 칩."""
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    sh.fill.solid(); sh.fill.fore_color.rgb = _rgb(c["badge"])
    sh.line.fill.background()
    _strip_style(sh)
    tf = sh.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = str(number)
    _set_font(r, size=10.5, bold=False, color="FFFFFF")
    return sh


def _label(slide, x, text, c, dark=False, y=0.33):
    """섹션 라벨(작게) — 한글만."""
    _text(slide, text, x + LEFT, y, CW, 0.25,
          size=9, color=(c["paper"] if dark else c["label"]), bold=True)


def _header(slide, x, text, c, dark=False, y=0.61, max_h=1.30):
    """대형 헤더. ★실제로 차지한 높이(in)를 반환한다.
    3줄짜리 헤더가 들어오면 0.9in 박스를 넘겨 아래 subhead 와 글자가 겹쳤다
    (경희궁 '레벨 테스트부터 반 배정까지, 네 단계로 안내합니다').
    이제 줄 수만큼 높이를 늘려 잡고, 호출부는 반환값 아래에 다음 요소를 놓는다."""
    size, wrapped, h = _fit_block(_s(text), CW, max_h, 19.5,
                                  min_size=14.5, line_spacing=1.16, pad=0.10)
    h = max(0.42, h)
    _text(slide, wrapped, x + LEFT, y, CW, h,
          size=size, color=(c["paper"] if dark else c["head"]), bold=True,
          line_spacing=1.16, fit=False)
    return y + h


def _footer(slide, x, name, phone, c):
    line = f"{name}  ·  {phone}" if phone else name
    _text(slide, line, x + LEFT, 7.92, CW, 0.23,
          size=7.5, color=c["footer"], align=PP_ALIGN.LEFT)


def _place_image(slide, url, x, y, w, h, cover=True):
    """url/base64/경로 이미지를 박스에 배치. cover=True면 꽉 채우고 crop, False면 비율 유지."""
    data = _fetch_image(url)
    if not data:
        return False
    try:
        from PIL import Image as _PILImage
        im = _PILImage.open(_io.BytesIO(data))
        iw, ih = im.size
        box_ratio = w / h
        img_ratio = iw / ih
        pic = slide.shapes.add_picture(_io.BytesIO(data), Inches(x), Inches(y), Inches(w), Inches(h))
        if cover:
            if img_ratio > box_ratio:
                crop = (1 - box_ratio / img_ratio) / 2
                pic.crop_left = crop; pic.crop_right = crop
            else:
                crop = (1 - img_ratio / box_ratio) / 2
                pic.crop_top = crop; pic.crop_bottom = crop
        else:
            if img_ratio > box_ratio:
                nw = w; nh = w / img_ratio
            else:
                nh = h; nw = h * img_ratio
            pic.width = Inches(nw); pic.height = Inches(nh)
            pic.left = Inches(x); pic.top = Inches(y)
        return True
    except Exception:
        try:
            slide.shapes.add_picture(_io.BytesIO(data), Inches(x), Inches(y), Inches(w), Inches(h))
            return True
        except Exception:
            return False


def _photo_or_box(slide, path, x, y, w, h, c, label_txt, cover_mode=True):
    """이미지 있으면 배치, 없으면 점선 안내 박스."""
    if _place_image(slide, path, x, y, w, h, cover=cover_mode):
        return True
    box = _rect(slide, x, y, w, h, fill=c["soft"], line=c["card_line"], radius=True)
    try:
        box.line.dash_style = 4
    except Exception:
        pass
    _text(slide, label_txt, x, y, w, h, size=8, color=c["footer"], bold=True,
          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    return False


# ── 데이터 정규화 ─────────────────────────────────────
def _feature_list(schema):
    out = []
    for it in _items(schema.get("features"))[:6]:
        t = _s(it.get("title") or it.get("name") or "")
        d = _s(it.get("desc") or it.get("description") or "")
        if t or d:
            out.append({"title": t, "desc": d})
    return out


def _curriculum_list(schema):
    cur = schema.get("curriculum") or {}
    stages = _items(cur) or _items(schema.get("targets"))
    out = []
    for it in stages[:3]:
        name = _s(it.get("name") or it.get("title") or it.get("grade"))
        # v3 스키마: 설명은 desc가 아니라 tag(목표 한마디) + items(반·과목 배열)에 있다.
        tag = _s(it.get("tag"))
        items = it.get("items")
        items_txt = ""
        if isinstance(items, list):
            items_txt = " · ".join(_s(x) for x in items if _s(x))
        elif isinstance(items, str):
            items_txt = _s(items)
        # 기존 desc/subj 계열도 폴백으로 지원
        desc_fallback = _s(it.get("desc") or it.get("description") or it.get("subj"))
        # 표시용 설명: tag를 위(굵게), items를 아래(상세)로 합침
        desc = tag or desc_fallback
        detail = items_txt
        if not (name or desc or detail):
            continue
        out.append({
            "name": name.replace("부", ""),
            "desc": desc,
            "detail": detail,
        })
    return out


def _mgmt_list(schema):
    mg = schema.get("management") or {}
    out = []
    # v3 스키마: management.columns = [{name, rows:[{k,v}]}]
    if isinstance(mg, dict) and isinstance(mg.get("columns"), list):
        for col in mg["columns"]:
            if not isinstance(col, dict):
                continue
            for r in (col.get("rows") or []):
                if not isinstance(r, dict):
                    continue
                k = _s(r.get("k") or r.get("key") or r.get("title"))
                v = _s(r.get("v") or r.get("value") or r.get("desc"))
                line = f"{k} — {v}" if (k and v) else (k or v)
                if line:
                    out.append(line)
        if out:
            return out[:4]
    # 폴백: items/steps/math/science 계열
    src = _items(mg)
    if not src and isinstance(mg, dict):
        for key in ("steps", "items", "math", "science"):
            src.extend(_items(mg.get(key)))
    for it in src[:4]:
        t = _s(it.get("title") or it.get("name") or it.get("desc") or it.get("description"))
        if t:
            out.append(t)
    return out[:4]


def _faq_list(schema):
    out = []
    for it in _items(schema.get("faq"))[:4]:
        q = _s(it.get("q") or it.get("question"))
        a = _s(it.get("a") or it.get("answer") or it.get("desc"))
        if q:
            out.append({"q": q, "a": a})
    return out


def _admission_list(schema):
    """입학·상담 절차 전체를 반환한다.
    워드 컨펌본이 5단계면 리플렛도 5단계다 — 3단계 하드캡은 폐기.
    글자수도 여기서 자르지 않는다(잘림의 원인). 넘치는 분량은
    _panel_admission 이 단계 수에 맞춰 크기·간격을 줄여 흡수한다."""
    src = _items(schema.get("admission"))
    out = []
    for it in src[:MAX_ADMISSION]:
        t = _s(it.get("title") or it.get("name"))
        d = _s(it.get("desc") or it.get("body"))
        if not d:
            # v3 admission step: 내용이 items 배열에 있음
            items = it.get("items")
            if isinstance(items, list):
                d = " · ".join(_s(x) for x in items if _s(x))
        if t or d:
            out.append({"title": t, "desc": d})
    return out


_EMOJI_RE = re.compile(
    "[\U0001F300-\U0001FAFF\U00002600-\U000027BF\U0001F000-\U0001F0FF\uFE0F]+")


def _de_emoji(s):
    """이모지 제거(디자인 원칙: 이모지 금지 → 선아이콘·도형으로).
    앞머리 이모지는 통째로 떼고, 문장 중간은 공백으로 바꾼다."""
    t = _EMOJI_RE.sub("", _s(s))
    return re.sub(r"\s{2,}", " ", t).strip(" ·-—")


def _achievement_lines(schema):
    ach = schema.get("achievements") or {}
    lines = []
    for it in _items(ach)[:8]:
        name = _s(it.get("name") or it.get("title"))
        desc = _s(it.get("desc") or it.get("description"))
        line = " · ".join(x for x in [name, desc] if x)
        line = _de_emoji(line)
        if not line:
            continue
        # ★한 항목에 실적이 통째로 뭉쳐 오는 경우가 있다(경희궁: "· 주요 실적 · 2024 KCIA…").
        #   줄바꿈이나 ' · ' 로 이어붙은 덩어리는 항목별로 쪼개 읽기 쉽게 만든다.
        parts = [p.strip(" ·-—") for p in re.split(r"[\n\r]+", line)]
        if len(parts) == 1 and len(line) > 46 and line.count(" · ") >= 2:
            parts = [p.strip() for p in line.split(" · ")]
        for p in parts:
            p = re.sub(r"^주요\s*실적\s*[·:]?\s*", "", p).strip(" ·-—")
            if p:
                lines.append(p)
    head = _s(ach.get("head") if isinstance(ach, dict) else "", "함께 이룬 결과")
    return head, lines[:8]


def _subject_line(schema):
    ac = schema.get("academy") or {}
    subj = _s(ac.get("subjects"))
    subj = subj.replace(" 전문", "")
    grades = []
    for t in _items(schema.get("targets")) or _items(schema.get("curriculum")):
        g = _s(t.get("grade") or t.get("name"))
        if g:
            grades.append(g.replace("부", ""))
    gtext = " · ".join(dict.fromkeys(grades)) if grades else ""
    if gtext and subj:
        return f"{gtext}  |  {subj}"
    return gtext or subj or ""


# ── 바깥면 패널 ───────────────────────────────────────
def _panel_admission(slide, schema, x, c):
    ac = schema.get("academy") or {}
    admission = schema.get("admission") or {}
    _label(slide, x, _first(admission.get("label") if isinstance(admission, dict) else "",
                            default="상담 · 예약 안내"), c)
    hb = _header(slide, x, _first(admission.get("head") if isinstance(admission, dict) else "",
                                  default="아이에게 맞는 시작점을\n함께 찾습니다."), c)
    sub_y = max(1.38, hb + 0.10)          # 헤더가 3줄이어도 겹치지 않게 아래로 흐름
    _text(slide, _first(admission.get("subhead") if isinstance(admission, dict) else "",
                        default="상담 후 적합한 과정을 안내합니다."),
          x + LEFT, sub_y, CW, 0.31, size=9.5, color=c["body"])

    steps = _admission_list(schema)
    if not steps:  # 최소 보장(학원 무관 일반 상담 절차 — 세원 특정 아님)
        steps = [
            {"title": "전화 상담", "desc": "가능 시간과 과정을 먼저 확인합니다."},
            {"title": "학생 상담", "desc": "현재 학습 흐름과 목표를 듣습니다."},
            {"title": "등원 결정", "desc": "적합한 과정과 시간대를 안내합니다."},
        ]
    # ── 단계 수에 맞춰 자동 축소 ──
    # 세로 가용구간: subhead 아래 ~ 5.00(구분선 위). 헤더가 길면 함께 내려간다.
    Y0, Y_END = max(1.92, sub_y + 0.44), 4.98
    n = max(1, len(steps))
    # 내용 길이에 맞춰 단계별 높이를 배분(고정 pitch 는 04번 설명 3줄에서 넘쳤다)
    t_h0 = 0.29
    need = []
    for it in steps:
        dl = max(1, int(_dwidth(it["desc"]) / 26) + 1) if it["desc"] else 0
        need.append(t_h0 + dl * 0.20 + 0.14)
    total = sum(need)
    avail = Y_END - Y0
    scale = min(1.0, avail / total) if total > 0 else 1.0
    tight = scale < 0.92
    t_size = 11 if not tight else 10
    d_size = 8.5 if not tight else 8
    chip_d = CHIP if not tight else 0.30

    y = Y0
    for i, it in enumerate(steps):
        blk = need[i] * scale
        _chip(slide, x + LEFT, y + 0.05, chip_d, f"{i+1:02d}", c)
        t_h = t_h0 if not tight else 0.27
        _text(slide, it["title"], x + TEXT_X, y, TEXT_W, t_h,
              size=t_size, color=c["title"], bold=True)
        # 설명은 자르지 않고 남은 높이 전부를 준다
        _text(slide, it["desc"], x + TEXT_X, y + t_h + 0.02,
              TEXT_W, max(0.20, blk - t_h - 0.06), size=d_size, color=c["body"],
              line_spacing=1.06)
        y += blk

    # 구분선
    _rect(slide, x + LEFT, 5.06, CW, 0.012, fill=c["card_line"])
    phone = _s(ac.get("phone"))
    _text(slide, f"전화  {phone}" if phone else "전화 상담", x + LEFT, 5.29,
          CW, 0.29, size=11, color=c["head"], bold=True)
    loc = _first(ac.get("location"), ac.get("address_short"), default="상담 시 위치를 안내해 드립니다.")
    _text(slide, loc, x + LEFT, 5.65, CW, 0.9,
          size=9, color=c["body"])


def _panel_faq(slide, schema, x, c, name):
    _label(slide, x, "자주 묻는 질문", c)
    hb = _header(slide, x, _first(_s((schema.get("faq") or {}).get("head")
            if isinstance(schema.get("faq"), dict) else ""), default="자주 묻는 질문"), c)
    faqs = _faq_list(schema)
    if not faqs:
        faqs = [
            {"q": "수업은 어떻게 배정되나요?", "a": "상담·진단 후 수준에 맞는 반을 안내합니다."},
            {"q": "결석하면 보강되나요?", "a": "운영 기준에 따라 보강 또는 자료를 제공합니다."},
        ]
    # ── 답변 전문이 들어가도록 높이를 먼저 배분한다 ──
    # 가용 세로: 1.38 ~ 7.80(푸터 위). 질문·답변 길이에 비례해 나눠 갖고,
    # 총합이 넘치면 전체를 비례 축소한다(글자를 자르지 않는다).
    Y0, Y_END = max(1.38, hb + 0.14), 7.78
    blocks = []
    for i, it in enumerate(faqs):
        q = f"Q{i+1}. {it['q']}"
        q_lines = max(1, int(_dwidth(q) / 20) + 1)
        q_h = 0.30 * q_lines
        a_lines = max(1, int(_dwidth(it["a"]) / 24) + 1)
        box_h = 0.26 + 0.21 * a_lines
        blocks.append({"q": q, "a": it["a"], "q_h": q_h, "box_h": box_h})

    GAP = 0.22
    total = sum(b["q_h"] + 0.08 + b["box_h"] for b in blocks) + GAP * max(0, len(blocks) - 1)
    avail = Y_END - Y0
    scale = min(1.0, avail / total) if total > 0 else 1.0
    gap = GAP * scale

    y = Y0
    for b in blocks:
        q_h = b["q_h"] * scale
        box_h = b["box_h"] * scale
        _text(slide, b["q"], x + LEFT, y, CW, q_h + 0.06,
              size=10, color=c["head"], bold=True, line_spacing=1.05)
        by = y + q_h + 0.08 * scale
        _rect(slide, x + LEFT, by, CW, box_h, fill=c["soft"], radius=True)
        _text(slide, b["a"], x + LEFT + 0.15, by, CW - 0.3, box_h,
              size=8.5, color=c["body"], valign=MSO_ANCHOR.MIDDLE, line_spacing=1.06)
        y = by + box_h + gap


def _panel_cover(slide, schema, x, c, assets, pw=None):
    PANEL = pw if pw else globals()["PANEL"]  # 이 패널의 실제 폭(표지=100mm)
    ac = schema.get("academy") or {}
    name = _s(ac.get("name"), "우리학원")
    phone = _s(ac.get("phone"))
    # 슬로건: 최상위 slogan / academy.slogan / brand.slogan / closing.head 모두 수용
    slogan = _first(schema.get("slogan"), ac.get("slogan"),
                    (schema.get("brand") or {}).get("slogan"),
                    (schema.get("closing") or {}).get("head"),
                    default="학생의 성장을\n함께 만들어갑니다.")
    # ※미리 _wrap 하지 않는다 — 아래 _fit_slogan 이 크기와 줄바꿈을 함께 정한다.
    # 한마디: intro.promise / intro.body / intro.head / closing.cta
    promise = _first((schema.get("intro") or {}).get("promise"),
                     (schema.get("intro") or {}).get("body"),
                     (schema.get("intro") or {}).get("head"),
                     (schema.get("closing") or {}).get("cta"),
                     default="정확한 진단 · 맞춤 수업 · 꾸준한 관리")

    # 표지 패널 배경(딥/노랑)
    _rect(slide, x, 0, PANEL, H_IN, fill=c["cover_bg"])

    # 로고 자리(좌상단) + 학원명
    logo_ok = _photo_or_box(slide, (assets or {}).get("logo"),
                            x + 0.19, 0.29, 0.72, 0.67, c, "로고", cover_mode=False)
    _text(slide, name, x + 1.02, 0.39, PANEL - 1.2, 0.5,
          size=13.5, color=c["cover_ink"], bold=True, valign=MSO_ANCHOR.MIDDLE)

    # 슬로건 — 줄 수에 맞춰 높이를 잡고, 아래 요소를 그만큼 밀어낸다.
    # (3줄 슬로건 'From Reading to Your Voice — and Beyond!' 이
    #  1.5in 고정 박스를 넘겨 소개문·사진과 겹쳤다)
    # 슬로건 — 표지의 주인공. 작게 줄이지 않고 3줄까지만 쓴다.
    # 문장부호가 없어 문장 트리밍이 안 되는 경우까지 대비해,
    # <실제로 감싼 줄 수>를 보고 어절을 하나씩 덜어낸다.
    SL_MAX_LINES = 3
    SL_MIN = 18
    SL_MAX = SL_MAX_LINES * (25 * 1.18) / 72.0 + 0.12
    sl_words = _s(slogan).replace("\n", " ").split()

    def _fit_slogan(txt):
        return _fit_block(txt, PANEL - 0.38, SL_MAX, 25,
                          min_size=SL_MIN, line_spacing=1.18, pad=0.12)

    sl_size, sl_wrapped, sl_h = _fit_slogan(slogan)
    while sl_wrapped.count("\n") + 1 > SL_MAX_LINES and len(sl_words) > 1:
        sl_words.pop()
        sl_size, sl_wrapped, sl_h = _fit_slogan(" ".join(sl_words))
    # 어절이 하나뿐인데도 넘치면(끊을 공백이 없음) 글자 단위로 줄인다
    if sl_wrapped.count("\n") + 1 > SL_MAX_LINES and len(sl_words) == 1:
        one = sl_words[0]
        while len(one) > 4 and sl_wrapped.count("\n") + 1 > SL_MAX_LINES:
            one = one[:-2]
            sl_size, sl_wrapped, sl_h = _fit_slogan(one)
    sl_h = max(0.5, sl_h)
    _text(slide, sl_wrapped, x + 0.19, 1.35, PANEL - 0.38, sl_h,
          size=sl_size, color=c["cover_ink"], bold=True,
          line_spacing=1.18, fit=False)

    # 소개 한마디 — 표지는 "읽히는 크기"가 우선이다.
    # ★슬로건과 크기 차이가 너무 벌어지지 않게, 슬로건의 <절반>을 기준으로 잡는다.
    #   (25pt 슬로건 → 12.5pt 소개문. 슬로건이 줄면 소개문도 같은 비율로 준다.)
    # ① 문장 경계에서 먼저 줄이고(첫 문장 보존) ② 그래도 넘치면 어절을 덜어낸다.
    pr_y = 1.35 + sl_h + 0.16
    PR_RATIO = 0.5
    PR_BASE = round(sl_size * PR_RATIO * 2) / 2.0        # 0.5pt 단위
    PR_MIN = max(9.5, round(SL_MIN * PR_RATIO * 2) / 2.0)
    PR_MAX_LINES = 3
    per_line = max(8, int((PANEL - 0.38 - 0.10) / _char_w(PR_BASE)))
    # ★표지에는 완결된 문장만 올린다. 길이와 무관하게 조각은 제외.
    promise = _cover_copy(promise, per_line * PR_MAX_LINES)
    pr_max_h = PR_MAX_LINES * (PR_BASE * 1.16) / 72.0 + 0.10

    def _fit_promise(txt):
        return _fit_block(txt, PANEL - 0.38, pr_max_h, PR_BASE,
                          min_size=PR_MIN, line_spacing=1.16, pad=0.10)

    pr_size, pr_wrapped, pr_h = _fit_promise(promise)
    # 넘치면 <문장 단위로> 덜어낸다(어절을 잘라 미완성 문장을 만들지 않는다)
    _pr_sents = _sentences(pr_wrapped.replace("\n", " "))
    while pr_wrapped.count("\n") + 1 > PR_MAX_LINES and len(_pr_sents) > 1:
        _pr_sents.pop()
        pr_size, pr_wrapped, pr_h = _fit_promise(" ".join(_pr_sents))
    # 한 문장인데도 3줄을 넘으면 표지에서 뺀다(잘린 문장을 올리지 않는다)
    if pr_wrapped.count("\n") + 1 > PR_MAX_LINES:
        pr_wrapped, pr_h = "", 0.0
    pr_h = max(0.0, pr_h)
    if pr_wrapped:
        _text(slide, pr_wrapped, x + 0.19, pr_y, PANEL - 0.38, pr_h,
              size=pr_size, color=c["cover_ink"], line_spacing=1.16, fit=False)

    # 표지 사진 자리 — 소개문 아래에서 시작(겹침 방지). 공간이 부족하면 생략.
    ph_y = pr_y + pr_h + 0.14
    ph_h = 6.04 - ph_y
    if ph_h >= 1.10:
        _photo_or_box(slide, (assets or {}).get("cover") or (assets or {}).get("banner"),
                      x + 0.19, ph_y, PANEL - 0.38, ph_h, c, "학원 사진", cover_mode=True)

    # 과목·학년 뱃지
    subj = _subject_line(schema)
    if subj:
        _rect(slide, x + 0.21, 6.27, PANEL - 0.42, 0.34, fill=c["badge"], radius=True)
        _text(slide, subj, x + 0.21, 6.27, PANEL - 0.42, 0.34,
              size=9.5, color="FFFFFF", bold=True, align=PP_ALIGN.CENTER,
              valign=MSO_ANCHOR.MIDDLE)

    # 표지 하단 강점 요약
    feats = _feature_list(schema)
    if feats:
        ftext = "  ·  ".join(f["title"] for f in feats[:3])
        _text(slide, ftext, x + 0.19, 6.75, PANEL - 0.38, 0.55,
              size=10, color=c["cover_ink"])

    _text(slide, f"{name}  ·  {phone}" if phone else name, x + 0.17, 7.92,
          PANEL - 0.2, 0.23, size=7.5, color=c["cover_ink"], align=PP_ALIGN.LEFT)


# ── 안쪽면 패널 ───────────────────────────────────────
def _panel_features(slide, schema, x, c, name):
    _label(slide, x, f"WHY {name}"[:18] if False else "우리 학원의 강점", c)
    feats = _feature_list(schema)
    head = f"성장을 만드는\n{len(feats)}가지 학습 원칙" if len(feats) >= 2 else "우리 학원의\n학습 원칙"
    hb = _header(slide, x, _first((schema.get("features_head")), default=head), c)
    Y0, Y_END = max(1.54, hb + 0.14), 7.80
    # 제목이 2줄 되는 항목이 있으므로(예: '1:1 레벨 진단 및 성장 경로 설계')
    # 항목마다 필요한 높이를 재서 배분한다.
    need = []
    for it in feats:
        tl = max(1, int(_dwidth(it["title"]) / 19) + 1)
        dl = max(1, int(_dwidth(it["desc"]) / 26) + 1) if it["desc"] else 0
        need.append(tl * 0.26 + dl * 0.20 + 0.20)
    total = sum(need)
    avail = Y_END - Y0
    scale = min(1.0, avail / total) if total > 0 else 1.0
    t_size = 12 if scale >= 0.92 else 11
    y = Y0
    for i, it in enumerate(feats):
        blk = need[i] * scale
        tl = max(1, int(_dwidth(it["title"]) / 19) + 1)
        t_h = max(0.26, tl * 0.26 * scale)
        _chip(slide, x + LEFT, y, CHIP, f"{i+1:02d}", c)
        _text(slide, it["title"], x + TEXT_X, y, TEXT_W, t_h,
              size=t_size, color=c["title"], bold=True, line_spacing=1.08)
        _text(slide, it["desc"], x + TEXT_X, y + t_h + 0.03, TEXT_W,
              max(0.20, blk - t_h - 0.09), size=8.5, color=c["body"],
              line_spacing=1.06)
        y += blk


def _panel_curriculum(slide, schema, x, c):
    cur = _curriculum_list(schema)
    if not cur:
        # 커리큘럼 없으면 패널 라벨·헤더 자체를 안 그림(빈 잔존 없음)
        return
    _label(slide, x, "교육 과정", c)
    hb = _header(slide, x, _first(_s((schema.get("curriculum") or {}).get("head")),
                                  default="단계별로 이어지는\n학습 과정"), c)
    cin = 0.18  # 카드 내부 여백
    y = max(1.73, hb + 0.20)
    # 카드 높이를 내용 길이로 산출하고, 넘치면 비례 축소(글자 안 자름)
    # 아래에 특별 프로그램 블록이 붙으면 그 위에서 멈춘다.
    _has_sp = bool(_items(schema.get("specials")))
    Y_END = 6.80 if _has_sp else 7.78
    txt_w = CW - cin * 2
    plan = []
    for it in cur:
        d_lines = max(1, int(_dwidth(it.get("desc") or "") / 20) + 1) if it.get("desc") else 0
        t_lines = max(1, int(_dwidth(it.get("detail") or "") / 24) + 1) if it.get("detail") else 0
        h = 0.56 + d_lines * 0.24 + t_lines * 0.20 + 0.16
        # ★1.05in 최소높이를 두면 내용 2줄짜리 초등 카드가 과하게 커져
        #   중등 카드와 균형이 깨졌다(경희궁). 내용만큼만 준다.
        plan.append({"it": it, "h": max(0.82, h), "d": d_lines, "t": t_lines})
    GAP = 0.35
    total = sum(p["h"] for p in plan) + GAP * max(0, len(plan) - 1)
    avail = Y_END - y
    scale = min(1.0, avail / total) if total > 0 else 1.0
    gap = GAP * scale

    for p in plan:
        it, ch = p["it"], p["h"] * scale
        _rect(slide, x + LEFT, y, CW, ch, fill="FFFFFF",
              line=c["card_line"], radius=True)
        # 라벨 pill: 글자 폭에 맞춰 너비 산출(고정 0.79in 이면 '초등 단계'가 줄바꿈됨)
        pill_txt = _s(it["name"])
        pill_w = min(CW - cin * 2, max(0.72, _dwidth(pill_txt) * 0.115 + 0.30))
        _rect(slide, x + LEFT + cin, y + 0.19, pill_w, 0.29, fill=c["accent2"], radius=True)
        _text(slide, pill_txt, x + LEFT + cin, y + 0.19, pill_w, 0.29,
              size=9.5, color="FFFFFF", bold=True, align=PP_ALIGN.CENTER,
              valign=MSO_ANCHOR.MIDDLE, fit=False)
        cy = y + 0.56
        # 목표 한마디
        if it.get("desc"):
            dh = max(0.24, p["d"] * 0.24 * scale)
            _text(slide, it["desc"], x + LEFT + cin, cy, txt_w, dh,
                  size=10.5, color=c["head"], bold=True)
            cy += dh
        # 반·과목 상세
        if it.get("detail"):
            th = max(0.20, p["t"] * 0.20 * scale)
            _text(slide, it["detail"], x + LEFT + cin, cy, txt_w, th,
                  size=8.5, color=c["body"])
        y += ch + gap

    # 특별 프로그램(있을 때만)
    specials = _items(schema.get("specials"))
    stext = " · ".join(_s(s.get("title") or s.get("name"))
                       for s in specials if _s(s.get("title") or s.get("name")))
    if stext:
        _rect(slide, x + LEFT, 6.92, CW, 0.012, fill=c["card_line"])  # 구분선
        _text(slide, _first(_s((schema.get("specials") or {}).get("head")), default="특별 프로그램"),
              x + LEFT, 7.08, CW, 0.25, size=9, color=c["label"], bold=True)
        _text(slide, stext, x + LEFT, 7.36, CW, 0.46,
              size=8.5, color=c["body"])


def _panel_management(slide, schema, x, c):
    steps = _mgmt_list(schema)
    _label(slide, x, "학습 관리", c)
    hb = _header(slide, x, _first(_s((schema.get("management") or {}).get("head")),
                                  default="진도보다 이해를\n먼저 확인합니다."), c)
    sub_y = max(1.48, hb + 0.10)
    if steps:
        _text(slide, _first(_s((schema.get("management") or {}).get("subhead")),
                            default=f"학습 관리 {len(steps)} STEP"),
              x + LEFT, sub_y, CW, 0.27, size=9.5, color=c["body"], bold=True)
    # 카드 높이를 내용 길이로 산출(고정 0.82in 이 길면 잘렸다).
    # 실적 블록이 아래에 붙으므로 그 위까지만 쓰고 여유를 둔다.
    Y0 = max(1.95, sub_y + 0.42)
    has_ach = bool(_achievement_lines(schema)[1])
    Y_END = 5.72 if has_ach else 7.80
    plan = []
    for txt in steps:
        lines = max(1, int(_dwidth(txt) / 22) + 1)
        plan.append({"txt": txt, "h": max(0.62, 0.26 + lines * 0.22)})
    GAP = 0.20
    total = sum(p["h"] for p in plan) + GAP * max(0, len(plan) - 1)
    avail = Y_END - Y0
    scale = min(1.0, avail / total) if total > 0 else 1.0
    gap = GAP * scale

    y = Y0
    for i, p in enumerate(plan):
        bh = p["h"] * scale
        _chip(slide, x + LEFT, y + 0.04, CHIP, str(i + 1), c)
        _rect(slide, x + TEXT_X, y, TEXT_W, bh, fill=c["soft"], radius=True)
        _text(slide, p["txt"], x + TEXT_X + 0.15, y, TEXT_W - 0.3, bh,
              size=9, color=c["head"], valign=MSO_ANCHOR.MIDDLE, line_spacing=1.08)
        y += bh + gap

    # 실적(있을 때만)
    head, lines = _achievement_lines(schema)
    if lines:
        _rect(slide, x + LEFT, 5.92, CW, 0.012, fill=c["card_line"])  # 구분선
        _text(slide, _de_emoji(head), x + LEFT, 6.10, CW, 0.28,
              size=10.5, color=c["head"], bold=True)
        # 줄마다 랩된 줄 수를 세어 실제 필요한 높이를 준다(푸터 7.90 위까지).
        body = "\n".join("· " + l for l in lines)
        _text(slide, body, x + LEFT, 6.42, CW, 1.44,
              size=8, color=c["body"], line_spacing=1.08, min_size=6.0)


# ── 예비 패널(대체 섹션) ──────────────────────────────
def _panel_achievements(slide, schema, x, c):
    """실적 단독 패널. 앞 섹션이 비어 면이 남을 때 승격되어 들어온다."""
    head, lines = _achievement_lines(schema)
    _label(slide, x, "주요 실적", c)
    hb = _header(slide, x, _first(head, default="숫자로 남은\n결과입니다."), c)
    y = max(1.95, hb + 0.18)
    for ln in lines[:6]:
        _rect(slide, x + LEFT, y, CW, 0.62, fill=c["soft"], radius=True)
        _text(slide, ln, x + LEFT + 0.15, y, CW - 0.3, 0.62,
              size=9.5, color=c["head"], valign=MSO_ANCHOR.MIDDLE)
        y += 0.76


def _panel_rules(slide, schema, x, c):
    """학원 규정 패널(예비)."""
    rules = _items(schema.get("rules")) or _items(schema.get("policy"))
    _label(slide, x, "학원 안내", c)
    hb = _header(slide, x, "함께 지키는\n약속입니다.", c)
    y = max(1.95, hb + 0.18)
    for it in rules[:6]:
        k = _s(it.get("k") or it.get("title") or it.get("name"))
        v = _s(it.get("v") or it.get("desc") or it.get("body"))
        if not (k or v):
            continue
        if k:
            _text(slide, k, x + LEFT, y, CW, 0.26,
                  size=10, color=c["title"], bold=True)
        _text(slide, v, x + LEFT, y + 0.28, CW, 0.46,
              size=8.5, color=c["body"])
        y += 0.86


def _panel_specials(slide, schema, x, c):
    """특별 프로그램 단독 패널(예비)."""
    sp = _items(schema.get("specials"))
    _label(slide, x, "특별 프로그램", c)
    hb = _header(slide, x, _first(_s((schema.get("specials") or {}).get("head")),
                                  default="정규 수업 밖에서\n더 채웁니다."), c)
    y = max(1.95, hb + 0.18)
    for i, it in enumerate(sp[:5]):
        t = _s(it.get("title") or it.get("name"))
        d = _s(it.get("desc") or it.get("description"))
        _chip(slide, x + LEFT, y + 0.05, CHIP, f"{i+1:02d}", c)
        _text(slide, t, x + TEXT_X, y, TEXT_W, 0.28,
              size=10.5, color=c["title"], bold=True)
        _text(slide, d, x + TEXT_X, y + 0.30, TEXT_W, 0.48,
              size=8.5, color=c["body"])
        y += 0.94


# ── 섹션 풀: 데이터가 있는 섹션만 패널에 배정 ──────────
def _has_admission(s):    return bool(_admission_list(s)) or bool(_s((s.get("academy") or {}).get("phone")))
def _has_faq(s):          return bool(_faq_list(s))
def _has_features(s):     return bool(_feature_list(s))
def _has_curriculum(s):   return bool(_curriculum_list(s))
def _has_management(s):   return bool(_mgmt_list(s))
def _has_achievements(s): return bool(_achievement_lines(s)[1])
def _has_specials(s):     return bool(_items(s.get("specials")))
def _has_rules(s):        return bool(_items(s.get("rules")) or _items(s.get("policy")))

# 우선순위 순. 앞 섹션이 비면 뒤 섹션이 자동으로 앞 면으로 당겨진다(A+C).
# 표지(_panel_cover)는 위치가 고정이라 이 풀에서 제외한다.
SECTION_POOL = [
    ("admission",    _has_admission,    _panel_admission),
    ("features",     _has_features,     None),   # name 인자 필요 → 아래에서 래핑
    ("curriculum",   _has_curriculum,   _panel_curriculum),
    ("management",   _has_management,   _panel_management),
    ("faq",          _has_faq,          None),   # name 인자 필요
    ("achievements", _has_achievements, _panel_achievements),
    ("specials",     _has_specials,     _panel_specials),
    ("rules",        _has_rules,        _panel_rules),
]


def _pick_sections(schema, slots, name):
    """채울 수 있는 섹션을 우선순위대로 slots개 고른다.
    - 데이터 없는 섹션은 건너뛰고 뒤 섹션을 당겨온다 → 빈 면이 생기지 않는다.
    - 후보가 slots보다 적으면 있는 만큼만(억지 확대·빈칸 없음).
    반환: [(key, draw_fn(slide, x, c)), ...]"""
    picked = []
    for key, has, fn in SECTION_POOL:
        if len(picked) >= slots:
            break
        try:
            if not has(schema):
                continue
        except Exception:
            continue
        if key == "features":
            draw = lambda sl, x, c: _panel_features(sl, schema, x, c, name)
        elif key == "faq":
            draw = lambda sl, x, c: _panel_faq(sl, schema, x, c, name)
        else:
            draw = (lambda f: lambda sl, x, c: f(sl, schema, x, c))(fn)
        picked.append((key, draw))
    return picked


def _panel_combo(slide, schema, x, c, parts):
    """여러 섹션을 한 면에 모아 그린다.
    섹션이 3~4개일 때 안쪽면(3면)을 못 채워 콘텐츠를 버리는 것을 막는다.
    각 파트를 [소제목 + 줄목록]으로 압축해 세로로 쌓고, 남는 높이에 맞춘다."""
    _label(slide, x, "한눈에 보기", c)
    hb = _header(slide, x, "학원 안내", c)
    Y0, Y_END = max(1.54, hb + 0.14), 7.80
    blocks = []
    for title, lines in parts:
        lines = [l for l in lines if _s(l)]
        if not lines:
            continue
        blocks.append((title, lines))
    if not blocks:
        return
    # 필요한 높이를 먼저 재고, 넘치면 비례 축소
    def _need(b):
        return 0.30 + sum(max(1, int(_dwidth(l) / 24) + 1) * 0.21 for l in b[1])
    total = sum(_need(b) for b in blocks) + 0.22 * (len(blocks) - 1)
    scale = min(1.0, (Y_END - Y0) / total) if total > 0 else 1.0
    y = Y0
    for title, lines in blocks:
        _text(slide, title, x + LEFT, y, CW, 0.28,
              size=10.5, color=c["label"], bold=True)
        by = y + 0.30 * scale
        body = "\n".join("· " + _s(l) for l in lines)
        bh = max(0.30, (_need((title, lines)) - 0.30) * scale)
        _text(slide, body, x + LEFT, by, CW, bh,
              size=9, color=c["body"], line_spacing=1.10, min_size=6.5)
        y = by + bh + 0.22 * scale


def _combo_parts(schema, keys):
    """섹션 키 목록 → _panel_combo 가 쓸 [(소제목, [줄...])] 로 변환."""
    out = []
    for k in keys:
        if k == "achievements":
            head, lines = _achievement_lines(schema)
            out.append(("주요 실적", lines))
        elif k == "specials":
            out.append(("특별 프로그램",
                        [_s(i.get("title") or i.get("name")) for i in _items(schema.get("specials"))]))
        elif k == "management":
            out.append(("학습 관리", _mgmt_list(schema)))
        elif k == "faq":
            out.append(("자주 묻는 질문",
                        [f"{i['q']} — {i['a']}" for i in _faq_list(schema)]))
        elif k == "rules":
            src = _items(schema.get("rules")) or _items(schema.get("policy"))
            out.append(("학원 규정",
                        [f"{_s(i.get('k') or i.get('title'))} {_s(i.get('v') or i.get('desc'))}".strip()
                         for i in src]))
        elif k == "curriculum":
            out.append(("교육 과정",
                        [f"{i['name']} {i.get('desc') or ''} {i.get('detail') or ''}".strip()
                         for i in _curriculum_list(schema)]))
        elif k == "features":
            out.append(("우리 학원의 강점",
                        [f"{i['title']} — {i['desc']}" for i in _feature_list(schema)]))
        elif k == "admission":
            out.append(("상담 · 등록 안내",
                        [f"{i['title']} {i['desc']}".strip() for i in _admission_list(schema)]))
    return [(t, l) for t, l in out if l]


# ── 빌드 ──────────────────────────────────────────────
def _force_theme_font(prs, name=FONT):
    """테마의 majorFont/minorFont 를 지정 서체로 바꾼다.
    기본 테마는 Calibri(한글 글리프 없음)라서, 혹시 서체 지정이 누락된
    텍스트가 있으면 PowerPoint 가 임의의 한글 폰트로 대체해 버린다.
    테마까지 맞춰두면 그런 경우에도 Pretendard 로 떨어진다.
    ※theme1.xml 은 python-pptx 가 파싱하지 않는 일반 Part 라서
      blob(바이트)을 직접 치환한다."""
    try:
        for part in prs.part.package.iter_parts():
            if "theme" not in str(part.partname):
                continue
            xml = part.blob.decode("utf-8", errors="ignore")

            def _fix(block):
                block = re.sub(r'(<a:latin[^/>]*typeface=")[^"]*(")',
                               r"\1" + name + r"\2", block)
                block = re.sub(r'(<a:ea[^/>]*typeface=")[^"]*(")',
                               r"\1" + name + r"\2", block)
                block = re.sub(r'(<a:cs[^/>]*typeface=")[^"]*(")',
                               r"\1" + name + r"\2", block)
                return block

            for tag in ("majorFont", "minorFont"):
                m = re.search(r"<a:%s>.*?</a:%s>" % (tag, tag), xml, re.S)
                if m:
                    xml = xml[:m.start()] + _fix(m.group(0)) + xml[m.end():]
            part._blob = xml.encode("utf-8")
    except Exception:
        pass  # 테마 수정 실패는 치명적이지 않다(런 단위 지정이 이미 있음)


def build(schema: Dict[str, Any], palette: str = "sewon_teal",
          out: Union[str, os.PathLike, io.BytesIO, None] = None,
          assets: Optional[Dict[str, str]] = None):
    c = PALETTES.get(palette, PALETTES["sewon_teal"])

    prs = Presentation()
    prs.slide_width = Inches(W_IN)
    prs.slide_height = Inches(H_IN)
    blank = prs.slide_layouts[6]

    ac = schema.get("academy") or {}
    name = _s(ac.get("name"), "우리학원")
    phone = _s(ac.get("phone"))

    # assets: 스키마 내장 + 파라미터 병합(파라미터 우선)
    _assets = dict(schema.get("assets") or {})
    if isinstance(assets, dict):
        _assets.update(assets)

    # 접지 좌표 계산
    out_geom = _panel_geom(PANELS_OUTSIDE_MM)   # [(x,w)] 좌97·중100·우100
    in_geom  = _panel_geom(PANELS_INSIDE_MM)    # [(x,w)] 좌100·중100·우97

    # ── 섹션 배정: 표지 1면 + 콘텐츠 면 ──
    # 데이터 있는 섹션만 우선순위대로 채우고, 비면 뒤 섹션을 앞으로 당긴다.
    # → 빈 면이 생기지 않는다(A: 풀+동적배치, C: 예비섹션 승격).
    #
    # 3단 접지는 면 단위로만 쓸 수 있다. 바깥면 [섹션·섹션·표지] 3면,
    # 안쪽면 [섹션·섹션·섹션] 3면. 안쪽면은 3면이 다 찰 때만 만든다.
    # 따라서 실제로 쓸 수 있는 섹션 수는 2개(1장) 또는 5개(2장) 뿐이다.
    # 후보를 5개까지 뽑아보고, 3~4개면 상위 2개만 써서 1장으로 낸다.
    cand = _pick_sections(schema, 8, name)
    if len(cand) >= 5:
        picked = cand[:5]          # 2장(바깥3 + 안쪽3) 꽉 채움
    elif len(cand) >= 3:
        # 3~4개: 1장으로 낸다. 콘텐츠 면이 2개뿐이므로
        # 첫 섹션은 제 면으로, 나머지 전부를 한 면에 합쳐 버리지 않는다.
        rest = [k for k, _ in cand[1:]]
        parts = _combo_parts(schema, rest)
        picked = cand[:1]
        if parts:
            picked = picked + [("combo",
                                lambda sl, x, c, _p=parts: _panel_combo(sl, schema, x, c, _p))]
    else:
        picked = cand              # 1~2개: 바깥면만

    def _draw(slide, idx, x):
        """picked[idx] 를 그린다. 후보가 모자라면 아무것도 안 그림."""
        if 0 <= idx < len(picked):
            picked[idx][1](slide, x, c)
            _footer(slide, x, name, phone, c)
            return True
        return False

    n = len(picked)
    if n >= 5:
        outer_secs, inner_secs = 2, 3    # 2장: 바깥[섹·섹·표지] + 안쪽[섹·섹·섹]
    else:
        # 1장: 바깥[섹·섹·표지]. 표지는 항상 있어야 하므로 콘텐츠는 최대 2면.
        # (3면짜리 combo 케이스는 위에서 이미 2섹션+combo=3 이므로 여기서 2로 잘림 →
        #  combo 가 밀리지 않도록 n==3 이면 combo 를 두 번째 면에 둔다)
        outer_secs, inner_secs = min(2, n), 0

    # ── 슬라이드 1: 바깥면 (섹션 / 섹션 / 표지) ──
    s1 = prs.slides.add_slide(blank)
    (ax, aw), (fx, fw), (cx, cw) = out_geom
    if outer_secs > 0:
        _draw(s1, 0, ax)
    if outer_secs > 1:
        _rect(s1, fx, 0, 0.008, H_IN, fill=c["card_line"])   # 접는 선
        _draw(s1, 1, fx)
    _rect(s1, cx, 0, 0.008, H_IN, fill=c["card_line"])       # 접는 선
    _panel_cover(s1, schema, cx, c, _assets, pw=cw)          # 표지는 위치 고정

    # ── 슬라이드 2: 안쪽면 ──
    # 안쪽면은 3면이 다 차거나(3섹션) 아예 안 만든다 → 백지 면이 남지 않는다.
    if inner_secs > 0:
        s2 = prs.slides.add_slide(blank)
        (ix, iw), (mx, mw), (gx, gw) = in_geom
        _draw(s2, outer_secs, ix)
        if inner_secs > 1:
            _rect(s2, mx, 0, 0.008, H_IN, fill=c["card_line"])
            _draw(s2, outer_secs + 1, mx)
        if inner_secs > 2:
            _rect(s2, gx, 0, 0.008, H_IN, fill=c["card_line"])
            _draw(s2, outer_secs + 2, gx)

    _force_theme_font(prs, FONT)
    prs.core_properties.title = f"{name} 3단 리플렛"
    prs.core_properties.subject = "YouMeanI coordinate leaflet"

    if out is None:
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return buf
    prs.save(out)
    return out
