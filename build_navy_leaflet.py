# -*- coding: utf-8 -*-
"""
build_navy_leaflet.py — navy 스킨 리플렛 (양면 3단, 2장)
기준점: 더바른수학전문학원 초등부 3단리플렛(네이비골드)

판형 12.4 x 8.77in, 패널폭 4.133in, 패널 내부여백 0.29in
슬라이드1(외부면): [FAQ] [입학안내+연락처] [표지]
슬라이드2(내부면): [철학] [커리큘럼] [학습관리]

입력: v3 표준스키마 (PPT와 동일). 리플렛은 지면이 좁아 항목 수를 압축한다.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from navy_core import (
    FS, get_palette, blank_slide, add_box, add_text, add_pill,
    place_image, clean, fit_size, gv, as_dicts, wrap_words,
)

PW = 4.133          # 패널 폭
SW = PW * 3         # 12.4
SH = 8.77
PAD = 0.29          # 패널 내부 여백
R = 0.055           # 공통 라운드 반경
CW = PW - PAD * 2   # 콘텐츠 폭 3.553


def panel_x(i):
    return PW * i + PAD


def divider(s, P):
    """패널 경계 세로선."""
    for i in (1, 2):
        add_box(s, PW * i - 0.005, 0, 0.01, SH, fill=P["line"])


def head(s, P, x, eyebrow, title_lines, y=0.36):
    """패널 헤더. 줄 수에 맞춰 박스 높이를 늘리고, 3줄이 되면 크기를 줄인다."""
    if isinstance(title_lines, str):
        title_lines = [title_lines]
    title_lines = [clean(t) for t in title_lines if clean(t)]

    add_text(s, x, y, CW, 0.19, eyebrow.upper(), size=FS.eyebrow, bold=True,
             color=P["primary"], line_spacing=1.0)

    size = 17.25
    if len(title_lines) > 2 or any(len(t) > 16 for t in title_lines):
        size = 15.0
    if len(title_lines) > 3:
        size = 13.5
    h = max(0.54, len(title_lines) * size * 1.3 / 72.0 + 0.06)
    add_text(s, x, y + 0.25, CW, h, title_lines, size=size, bold=True,
             color=P["text"], line_spacing=1.3)
    return y + 0.25 + h


# ── 패널: FAQ ────────────────────────────────────────────────────
def p_faq(s, d, P, x):
    items = as_dicts(d.get("faq"), "q")
    items = [i for i in items if clean(i.get("q") or i.get("question"))][:4]
    if not items:
        return False
    head(s, P, x, "FAQ", ["등록 전에", "많이 물어보세요"])
    y = 1.61
    for it in items:
        q = clean(it.get("q") or it.get("question") or "")
        a = clean(it.get("a") or it.get("answer") or "")
        add_text(s, x, y, CW, 0.36, "Q. " + q, size=FS.body,
                 bold=True, color=P["text"], line_spacing=1.3)
        add_text(s, x, y + 0.42, CW, 0.57, "A. " + a, size=FS.body_sm,
                 color=P["muted"], line_spacing=1.45)
        y += 1.37
    return True


# ── 패널: 입학안내 + 연락처 ─────────────────────────────────────
def p_admission(s, d, P, x):
    steps = as_dicts(d.get("admission"))[:4]
    b = d.get("basic", {})
    head(s, P, x, "ADMISSION", ["입학 안내"])
    y = 1.41
    for i, st in enumerate(steps):
        add_pill(s, x, y, 0.5, 0.29, f"{i+1:02d}", fill=P["primary"],
                 color=P["onDark"], size=FS.small, radius=0.2)
        add_text(s, x + 0.67, y, CW - 0.67, 0.31,
                 clean(st.get("title") or st.get("name") or ""),
                 size=FS.body + 0.75, bold=True, color=P["text"],
                 anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
        y += 0.78

    # 연락처 박스
    by = max(y + 0.3, 4.95)
    add_box(s, x, by, CW, 1.88, fill=P["card"], line=P["line"], radius=R)
    add_text(s, x + 0.26, by + 0.26, CW - 0.52, 0.29, "상담 · 예약",
             size=FS.body, bold=True, color=P["primary"], line_spacing=1.0)
    add_text(s, x + 0.26, by + 0.67, CW - 0.52, 0.35,
             clean(b.get("phone") or ""), size=16.5, bold=True,
             color=P["text"], line_spacing=1.0)
    addr = clean(b.get("address") or "")
    if addr:
        add_text(s, x + 0.26, by + 1.2, CW - 0.52, 0.57,
                 wrap_words(addr, 20)[:2], size=FS.small,
                 color=P["muted"], line_spacing=1.4)

    # 약도 — 연락처 박스 아래 남는 공간에
    mapimg = gv(d, "assets", "map")
    my = by + 1.88 + 0.28
    if mapimg and my + 1.0 < SH - 0.4:
        mh = min(1.9, SH - 0.4 - my - 0.3)
        add_text(s, x, my, 2.0, 0.22, "오시는 길", size=8.0,
                 bold=True, color=P["primary"], line_spacing=1.0)
        my += 0.3
        mh = min(mh, SH - 0.4 - my)
        add_box(s, x, my, CW, mh, fill=P["card"], line=P["line"], radius=R)
        place_image(s, mapimg, x + 0.05, my + 0.05,
                    CW - 0.1, mh - 0.1, cover=False)
    return True


# ── 패널: 표지 (딥네이비) ────────────────────────────────────────
def p_cover(s, d, P):
    b = d.get("basic", {})
    px = PW * 2
    add_box(s, px, 0, PW, SH, fill=P["deep"])
    x = px + PAD

    target = clean(b.get("target") or "")
    if target:
        tw = min(CW, 0.5 + len(target) * 0.14)
        add_pill(s, x, 0.73, tw, 0.31, target, fill=P["accent"],
                 color=P["deep"], size=FS.small)

    slogan = clean(b.get("slogan") or gv(d, "identity", "slogan") or "")
    if slogan:
        add_text(s, x, 1.41, CW, 1.15, wrap_words(slogan, 14)[:3],
                 size=19.5, bold=True, color=P["onDark"], line_spacing=1.35,
                 clip=False)

    logo = gv(d, "assets", "logo")
    if logo:
        place_image(s, logo, x + 0.5, 2.92, CW - 1.0, 1.88, cover=False)

    sub = clean(b.get("subline") or "")
    if sub:
        add_text(s, x, 5.05, CW, 0.29, sub, size=FS.small,
                 color=P["onDarkSub"], align=PP_ALIGN.CENTER, line_spacing=1.0)

    add_box(s, x, 5.83, CW, 1.17, fill=P["accent"], radius=R)
    add_text(s, x + 0.21, 5.83, CW - 0.42, 1.17,
             wrap_words(clean(b.get("name") or ""), 12)[:2],
             size=FS.card_ttl, bold=True, color=P["deep"],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.3)

    add_text(s, x, 7.5, CW, 0.29, clean(b.get("phone") or ""), size=FS.body,
             bold=True, color=P["onDark"], align=PP_ALIGN.CENTER,
             line_spacing=1.0)


# ── 패널: 철학 ───────────────────────────────────────────────────
def p_philosophy(s, d, P, x):
    ph = d.get("philosophy") or {}
    pts = as_dicts(ph.get("points"), "title")[:4]
    intro = clean(ph.get("intro") or gv(d, "identity", "intro") or "")
    if not pts and not intro:
        return False
    hb = head(s, P, x, "ABOUT", wrap_words(
        clean(ph.get("headline") or "공부하는 방법이 결과를 만듭니다"), 13)[:3])
    y = hb + 0.28
    if intro:
        add_text(s, x, y, CW, 1.0, intro, size=FS.body_sm,
                 color=P["muted"], line_spacing=1.55)
        y += 1.25
    for i, p in enumerate(pts):
        add_pill(s, x, y, 0.42, 0.28, f"{i+1:02d}", fill=P["primary"],
                 color=P["onDark"], size=8.0, radius=0.2)
        add_text(s, x + 0.58, y, CW - 0.58, 0.28,
                 clean(p.get("title") or ""), size=FS.body, bold=True,
                 color=P["text"], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
        y += 0.62

    note = (ph.get("note") or {})
    nb = clean(note.get("footer") or note.get("body") or "")
    if nb:
        by = max(y + 0.25, 6.3)
        add_box(s, x, by, CW, 1.5, fill=P["card"], line=P["line"], radius=R)
        add_text(s, x + 0.24, by + 0.24, CW - 0.48, 1.02, nb,
                 size=fit_size(nb, FS.body_sm, CW - 0.48, 1.02),
                 color=P["text"], line_spacing=1.55)
    return True


# ── 패널: 커리큘럼 ───────────────────────────────────────────────
def p_curriculum(s, d, P, x):
    stages = as_dicts(d.get("curriculum"), "name")[:3]
    if not stages:
        return False
    head(s, P, x, "CURRICULUM", ["단계별 학습 설계"])
    y = 1.55
    for st in stages:
        add_box(s, x, y, CW, 1.95, fill=P["card"], line=P["line"], radius=R)
        gname = clean(st.get("name") or st.get("grade") or "")
        gw = min(CW - 0.44, 0.46 + len(gname) * 0.13)
        add_pill(s, x + 0.22, y + 0.24, gw, 0.28, gname,
                 fill=P["primary"], color=P["onDark"], size=8.0)
        ttl = clean(st.get("title") or "")
        add_text(s, x + 0.22, y + 0.62, CW - 0.44, 0.34, ttl,
                 size=fit_size(ttl, FS.card_ttl - 1, CW - 0.44, 0.36),
                 bold=True, color=P["text"], line_spacing=1.15)
        body = clean(st.get("desc") or st.get("body") or "")
        if body:
            add_text(s, x + 0.22, y + 1.05, CW - 0.44, 0.75, body,
                     size=fit_size(body, FS.body_sm, CW - 0.44, 0.75, min_size=8.0),
                     color=P["muted"], line_spacing=1.45)
        y += 2.16

    note = clean(d.get("curriculumNote") or "")
    if note and y < SH - 1.0:
        add_text(s, x, y + 0.05, CW, 0.6, note, size=8.0,
                 color=P["muted"], line_spacing=1.4)
    return True


# ── 패널: 학습관리 (+강점 압축) ─────────────────────────────────
def p_management(s, d, P, x):
    items = as_dicts(d.get("management"))[:4]
    if not items:
        items = as_dicts(d.get("strengths"))[:4]
        for it in items:
            it.setdefault("key", clean(it.get("title", ""))[:4])
    if not items:
        return False
    head(s, P, x, "LEARNING MANAGEMENT", ["매 수업, 확인하고", "채워갑니다"])
    y = 1.75
    for it in items:
        key = clean(it.get("key") or it.get("title") or "")
        body = clean(it.get("desc") or it.get("body") or "")
        add_box(s, x, y, CW, 1.42, fill=P["card"], line=P["line"], radius=R)
        add_pill(s, x + 0.22, y + 0.2, 0.86, 0.3, key, fill=P["primary"],
                 color=P["onDark"], size=FS.small, radius=0.2)
        add_text(s, x + 0.22, y + 0.62, CW - 0.44, 0.66, body,
                 size=fit_size(body, FS.body_sm, CW - 0.44, 0.66, min_size=8.0),
                 color=P["text"], line_spacing=1.45)
        y += 1.62
    return True


def build(data, out=None, palette=None):
    P = get_palette(palette)
    d = data or {}
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    # ── 외부면
    s1 = blank_slide(prs)
    divider(s1, P)
    p_faq(s1, d, P, panel_x(0))
    p_admission(s1, d, P, panel_x(1))
    p_cover(s1, d, P)

    # ── 내부면
    s2 = blank_slide(prs)
    divider(s2, P)
    p_philosophy(s2, d, P, panel_x(0))
    p_curriculum(s2, d, P, panel_x(1))
    p_management(s2, d, P, panel_x(2))

    if out is None:
        import io
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio
    prs.save(out)
    return out


if __name__ == "__main__":
    import json, sys
    src = sys.argv[1] if len(sys.argv) > 1 else "sample.json"
    pal = sys.argv[2] if len(sys.argv) > 2 else "navy_gold"
    with open(src, encoding="utf-8") as f:
        data = json.load(f)
    print(build(data, f"navy_leaflet_{pal}.pptx", pal))
