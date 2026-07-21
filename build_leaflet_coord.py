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
FONT = "Pretendard"

# 좌우 여백 기준(모든 패널 공통). 라벨·헤더·본문·번호칩이 이 세로선에 맞춰 정렬된다.
PAD = 0.30            # 패널 안쪽 좌측 여백
LEFT = PAD            # 콘텐츠 왼쪽 기준(패널 x에 더해 사용: x + LEFT)
# 콘텐츠 폭은 가장 좁은 패널(97mm=3.819in) 기준으로 통일 → 어느 패널에서도 안 넘침
CW = 97 * MM - PAD * 2  # ≈ 3.22in
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


def _wrap(text: str, per_line: int, max_lines: int) -> str:
    """어절 단위 줄바꿈(단어 중간 안 끊음), 최대 줄 수 제한."""
    s = _s(text)
    if not s or "\n" in s:
        return s
    words = s.split(" ")
    lines, cur = [], ""
    for w in words:
        if cur and len(cur) + 1 + len(w) > per_line:
            lines.append(cur)
            cur = w
            if len(lines) >= max_lines:
                break
        else:
            cur = (cur + " " + w).strip()
    if cur and len(lines) < max_lines:
        lines.append(cur)
    return "\n".join(lines)


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
          autofit=True, cap=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = valign
    txt = _s(text)
    if cap:
        txt = _clip(txt, cap)
    lines = txt.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = _rgb(color)
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
    r.font.name = FONT; r.font.size = Pt(10.5); r.font.bold = True
    r.font.color.rgb = _rgb("FFFFFF")
    return sh


def _label(slide, x, text, c, dark=False, y=0.33):
    """섹션 라벨(작게) — 한글만."""
    _text(slide, text, x + LEFT, y, CW, 0.25,
          size=9, color=(c["paper"] if dark else c["label"]), bold=True)


def _header(slide, x, text, c, dark=False, y=0.61):
    _text(slide, _wrap(text, 12, 2), x + LEFT, y, CW, 0.9,
          size=19.5, color=(c["paper"] if dark else c["head"]), bold=True)


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
        t = _clip(it.get("title") or it.get("name") or "", 18)
        d = _clip(it.get("desc") or it.get("description") or "", 42)
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
            "name": _clip(name.replace("부", ""), 8),
            "desc": _clip(_wrap(desc, 16, 2), 45),
            "detail": _clip(_wrap(detail, 18, 2), 50),
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
                    out.append(_clip(line, 34))
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
            out.append(_clip(t, 30))
    return out[:4]


def _faq_list(schema):
    out = []
    for it in _items(schema.get("faq"))[:4]:
        q = _s(it.get("q") or it.get("question"))
        a = _s(it.get("a") or it.get("answer") or it.get("desc"))
        if q:
            out.append({"q": _clip(q, 34), "a": _clip(a, 76)})
    return out


def _admission_list(schema):
    src = _items(schema.get("admission"))
    out = []
    for it in src[:3]:
        t = _s(it.get("title") or it.get("name"))
        d = _s(it.get("desc") or it.get("body"))
        if not d:
            # v3 admission step: 내용이 items 배열에 있음
            items = it.get("items")
            if isinstance(items, list):
                d = " · ".join(_s(x) for x in items if _s(x))
        if t or d:
            out.append({"title": _clip(t, 18), "desc": _clip(d, 44)})
    return out


def _achievement_lines(schema):
    ach = schema.get("achievements") or {}
    lines = []
    for it in _items(ach)[:4]:
        name = _s(it.get("name") or it.get("title"))
        desc = _s(it.get("desc") or it.get("description"))
        line = " · ".join(x for x in [name, desc] if x)
        if line:
            lines.append(line)
    head = _s(ach.get("head") if isinstance(ach, dict) else "", "함께 이룬 결과")
    return head, lines


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
    _header(slide, x, _first(admission.get("head") if isinstance(admission, dict) else "",
                             default="아이에게 맞는 시작점을\n함께 찾습니다."), c)
    _text(slide, _first(admission.get("subhead") if isinstance(admission, dict) else "",
                        default="상담 후 적합한 과정을 안내합니다."),
          x + LEFT, 1.38, CW, 0.31, size=9.5, color=c["body"])

    steps = _admission_list(schema)
    if not steps:  # 최소 보장(학원 무관 일반 상담 절차 — 세원 특정 아님)
        steps = [
            {"title": "전화 상담", "desc": "가능 시간과 과정을 먼저 확인합니다."},
            {"title": "학생 상담", "desc": "현재 학습 흐름과 목표를 듣습니다."},
            {"title": "등원 결정", "desc": "적합한 과정과 시간대를 안내합니다."},
        ]
    y = 1.92
    for i, it in enumerate(steps):
        _chip(slide, x + LEFT, y + 0.06, CHIP, f"{i+1:02d}", c)
        _text(slide, it["title"], x + TEXT_X, y, TEXT_W, 0.29,
              size=11, color=c["title"], bold=True)
        _text(slide, it["desc"], x + TEXT_X, y + 0.31, TEXT_W, 0.44,
              size=8.5, color=c["body"])
        y += 0.96

    # 구분선
    _rect(slide, x + LEFT, 5.06, CW, 0.012, fill=c["card_line"])
    phone = _s(ac.get("phone"))
    _text(slide, f"전화  {phone}" if phone else "전화 상담", x + LEFT, 5.29,
          CW, 0.29, size=11, color=c["head"], bold=True)
    loc = _first(ac.get("location"), ac.get("address_short"), default="상담 시 위치를 안내해 드립니다.")
    _text(slide, _wrap(loc, 23, 3), x + LEFT, 5.65, CW, 0.9,
          size=9, color=c["body"])


def _panel_faq(slide, schema, x, c, name):
    _label(slide, x, "자주 묻는 질문", c)
    _header(slide, x, _first(_s((schema.get("faq") or {}).get("head")
            if isinstance(schema.get("faq"), dict) else ""), default="자주 묻는 질문"), c)
    faqs = _faq_list(schema)
    if not faqs:
        faqs = [
            {"q": "수업은 어떻게 배정되나요?", "a": "상담·진단 후 수준에 맞는 반을 안내합니다."},
            {"q": "결석하면 보강되나요?", "a": "운영 기준에 따라 보강 또는 자료를 제공합니다."},
        ]
    y = 1.38
    for i, it in enumerate(faqs):
        q = f"Q{i+1}. {it['q']}"
        # 질문이 길면 2줄 → 답 박스를 그만큼 내림
        q_lines = 2 if len(q) > 22 else 1
        q_h = 0.30 * q_lines
        _text(slide, _wrap(q, 22, 2), x + LEFT, y, CW, q_h + 0.06,
              size=10, color=c["head"], bold=True, line_spacing=1.05)
        # 답 길이에 맞춰 박스 높이
        a_lines = max(1, min(3, (len(it["a"]) // 24) + 1))
        box_h = 0.30 + 0.22 * a_lines
        by = y + q_h + 0.08
        _rect(slide, x + LEFT, by, CW, box_h, fill=c["soft"], radius=True)
        _text(slide, _wrap(it["a"], 24, 3), x + LEFT + 0.15, by, CW - 0.3, box_h,
              size=8.5, color=c["body"], valign=MSO_ANCHOR.MIDDLE, line_spacing=1.06)
        y = by + box_h + 0.26


def _panel_cover(slide, schema, x, c, assets, pw=None):
    PANEL = pw if pw else globals()["PANEL"]  # 이 패널의 실제 폭(표지=100mm)
    ac = schema.get("academy") or {}
    name = _s(ac.get("name"), "우리학원")
    phone = _s(ac.get("phone"))
    slogan = _first(ac.get("slogan"), (schema.get("closing") or {}).get("head"),
                    default="학생의 성장을\n함께 만들어갑니다.")
    slogan = _wrap(slogan, 11, 3)
    promise = _first((schema.get("intro") or {}).get("body"),
                     (schema.get("closing") or {}).get("cta"),
                     default="정확한 진단 · 맞춤 수업 · 꾸준한 관리")

    # 표지 패널 배경(딥/노랑)
    _rect(slide, x, 0, PANEL, H_IN, fill=c["cover_bg"])

    # 로고 자리(좌상단) + 학원명
    logo_ok = _photo_or_box(slide, (assets or {}).get("logo"),
                            x + 0.19, 0.29, 0.72, 0.67, c, "로고", cover_mode=False)
    _text(slide, name, x + 1.02, 0.39, PANEL - 1.2, 0.5,
          size=13.5, color=c["cover_ink"], bold=True, valign=MSO_ANCHOR.MIDDLE)

    # 슬로건
    _text(slide, slogan, x + 0.19, 1.35, PANEL - 0.38, 1.5,
          size=25, color=c["cover_ink"], bold=True)
    _text(slide, _clip(promise, 34), x + 0.19, 2.79, PANEL - 0.38, 0.4,
          size=10, color=c["cover_ink"])

    # 표지 사진 자리
    _photo_or_box(slide, (assets or {}).get("cover") or (assets or {}).get("banner"),
                  x + 0.19, 3.44, PANEL - 0.38, 2.6, c, "학원 사진", cover_mode=True)

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
        _text(slide, _clip(ftext, 40), x + 0.19, 6.75, PANEL - 0.38, 0.4,
              size=10, color=c["cover_ink"])

    _text(slide, f"{name}  ·  {phone}" if phone else name, x + 0.17, 7.92,
          PANEL - 0.2, 0.23, size=7.5, color=c["cover_ink"], align=PP_ALIGN.LEFT)


# ── 안쪽면 패널 ───────────────────────────────────────
def _panel_features(slide, schema, x, c, name):
    _label(slide, x, f"WHY {name}"[:18] if False else "우리 학원의 강점", c)
    feats = _feature_list(schema)
    head = f"성장을 만드는\n{len(feats)}가지 학습 원칙" if len(feats) >= 2 else "우리 학원의\n학습 원칙"
    _header(slide, x, _first((schema.get("features_head")), default=head), c)
    y = 1.54
    for i, it in enumerate(feats):
        _chip(slide, x + LEFT, y, CHIP, f"{i+1:02d}", c)
        _text(slide, it["title"], x + TEXT_X, y, TEXT_W, 0.35,
              size=12, color=c["title"], bold=True)
        _text(slide, it["desc"], x + TEXT_X, y + 0.38, TEXT_W, 0.4,
              size=8.5, color=c["body"])
        y += 0.92


def _panel_curriculum(slide, schema, x, c):
    cur = _curriculum_list(schema)
    if not cur:
        # 커리큘럼 없으면 패널 라벨·헤더 자체를 안 그림(빈 잔존 없음)
        return
    _label(slide, x, "교육 과정", c)
    _header(slide, x, _first(_s((schema.get("curriculum") or {}).get("head")),
                             default="단계별로 이어지는\n학습 과정"), c)
    y = 1.73
    cin = 0.18  # 카드 내부 여백
    for it in cur:
        _rect(slide, x + LEFT, y, CW, 1.38, fill="FFFFFF",
              line=c["card_line"], radius=True)
        _rect(slide, x + LEFT + cin, y + 0.19, 0.79, 0.29, fill=c["accent2"], radius=True)
        _text(slide, it["name"], x + LEFT + cin, y + 0.19, 0.79, 0.29,
              size=9.5, color="FFFFFF", bold=True, align=PP_ALIGN.CENTER,
              valign=MSO_ANCHOR.MIDDLE)
        # 목표 한마디(굵게)
        if it.get("desc"):
            _text(slide, it["desc"], x + LEFT + cin, y + 0.56, CW - cin * 2, 0.42,
                  size=10.5, color=c["head"], bold=True)
        # 반·과목 상세(회색, 아래)
        if it.get("detail"):
            _text(slide, it["detail"], x + LEFT + cin, y + 0.98, CW - cin * 2, 0.34,
                  size=8.5, color=c["body"])
        y += 1.73

    # 특별 프로그램(있을 때만)
    specials = _items(schema.get("specials"))
    stext = " · ".join(_s(s.get("title") or s.get("name"))
                       for s in specials if _s(s.get("title") or s.get("name")))
    if stext:
        _text(slide, _first(_s((schema.get("specials") or {}).get("head")), default="특별 프로그램"),
              x + LEFT, 7.1, CW, 0.25, size=9, color=c["label"], bold=True)
        _text(slide, _clip(stext, 60), x + LEFT, 7.38, CW, 0.42,
              size=8.5, color=c["body"])


def _panel_management(slide, schema, x, c):
    steps = _mgmt_list(schema)
    _label(slide, x, "학습 관리", c)
    _header(slide, x, _first(_s((schema.get("management") or {}).get("head")),
                             default="진도보다 이해를\n먼저 확인합니다."), c)
    if steps:
        _text(slide, _first(_s((schema.get("management") or {}).get("subhead")),
                            default=f"학습 관리 {len(steps)} STEP"),
              x + LEFT, 1.48, CW, 0.27, size=9.5, color=c["body"], bold=True)
    y = 1.95
    for i, txt in enumerate(steps):
        _chip(slide, x + LEFT, y + 0.04, CHIP, str(i + 1), c)
        _rect(slide, x + TEXT_X, y, TEXT_W, 0.82, fill=c["soft"], radius=True)
        _text(slide, _wrap(txt, 20, 3), x + TEXT_X + 0.15, y, TEXT_W - 0.3, 0.82,
              size=9, color=c["head"], valign=MSO_ANCHOR.MIDDLE, line_spacing=1.08)
        y += 1.02

    # 실적(있을 때만)
    head, lines = _achievement_lines(schema)
    if lines:
        _text(slide, head, x + LEFT, 6.1, CW, 0.29,
              size=10.5, color=c["head"], bold=True)
        _text(slide, "\n".join(lines), x + LEFT, 6.5, CW, 1.3,
              size=8.5, color=c["body"])


# ── 빌드 ──────────────────────────────────────────────
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

    # ── 슬라이드 1: 바깥면 (상담 / FAQ / 표지) ──
    s1 = prs.slides.add_slide(blank)
    (ax, aw), (fx, fw), (cx, cw) = out_geom
    _panel_admission(s1, schema, ax, c)
    _footer(s1, ax, name, phone, c)
    _rect(s1, fx, 0, 0.008, H_IN, fill=c["card_line"])   # 접는 선
    _panel_faq(s1, schema, fx, c, name)
    _footer(s1, fx, name, phone, c)
    _rect(s1, cx, 0, 0.008, H_IN, fill=c["card_line"])   # 접는 선
    _panel_cover(s1, schema, cx, c, _assets, pw=cw)

    # ── 슬라이드 2: 안쪽면 (강점 / 커리큘럼 / 관리) ──
    s2 = prs.slides.add_slide(blank)
    (ix, iw), (mx, mw), (gx, gw) = in_geom
    _panel_features(s2, schema, ix, c, name)
    _footer(s2, ix, name, phone, c)
    _rect(s2, mx, 0, 0.008, H_IN, fill=c["card_line"])
    _panel_curriculum(s2, schema, mx, c)
    _footer(s2, mx, name, phone, c)
    _rect(s2, gx, 0, 0.008, H_IN, fill=c["card_line"])
    _panel_management(s2, schema, gx, c)
    _footer(s2, gx, name, phone, c)

    prs.core_properties.title = f"{name} 3단 리플렛"
    prs.core_properties.subject = "YouMeanI coordinate leaflet"

    if out is None:
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return buf
    prs.save(out)
    return out
