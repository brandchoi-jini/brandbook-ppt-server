# -*- coding: utf-8 -*-
"""
navy_registry.py — 스킨 × 종류 라우팅 레지스트리

기존 v3 / book 빌더는 건드리지 않는다. import 실패해도 navy 는 동작한다.

사용법 (ppt_server.py 에서):
    from navy_registry import get_builder, list_options
    fn = get_builder(skin, kind)          # 없으면 None
    stream = fn(schema, None, palette)    # BytesIO 반환
"""

# ── navy 스킨 (신규) ─────────────────────────────────────────────
import build_brandbook_navy as _navy_ppt
import build_navy_leaflet as _navy_leaflet
import build_navy_catalog as _navy_catalog
from navy_core import PALETTES as NAVY_PALETTES

_REG = {
    ("navy", "ppt"):      _navy_ppt.build,
    ("navy", "leaflet"):  _navy_leaflet.build,
    ("navy", "catalog"):  _navy_catalog.build,
}

# ── 기존 스킨 (있으면 등록, 없으면 조용히 건너뜀) ───────────────
try:
    import build_brandbook_v3 as _v3
    _REG[("v3", "ppt")] = _v3.build
except Exception:
    pass

try:
    import build_brandbook_book as _book
    _REG[("book", "ppt")] = _book.build
except Exception:
    pass


DEFAULT_SKIN = "v3" if ("v3", "ppt") in _REG else "navy"

# 종류별 기본 스킨.
# 카탈로그·리플렛은 navy 에만 있으므로, PPT 도 navy 로 맞춰야 3종 디자인이 일치한다.
# 앱이 skin 을 안 보내도 kind 만으로 올바른 조합이 나오게 한다.
KIND_DEFAULT_SKIN = {
    "ppt": DEFAULT_SKIN,
    "catalog": "navy",
    "leaflet": "navy",
}

# navy 팔레트 별칭 — 앱에서 넘어오는 옛 팔레트명을 흡수
PALETTE_ALIAS = {
    "teal_blue": "navy_gold",
    "navy_amber": "navy_gold",
    "green_orange": "forest_gold",
    "세이지&골드": "forest_gold",
    "네이비&앰버": "navy_gold",
    "모노크롬": "charcoal_gold",
    "포레스트&골드": "forest_gold",
}


def normalize_palette(skin, palette):
    if skin != "navy":
        return palette
    if not palette:
        return "navy_gold"
    if palette in NAVY_PALETTES:
        return palette
    return PALETTE_ALIAS.get(palette, "navy_gold")


def resolve_skin(skin, kind="ppt"):
    """앱이 skin 을 안 보냈을 때 kind 에 맞는 스킨을 고른다."""
    if skin:
        return skin.lower()
    return KIND_DEFAULT_SKIN.get((kind or "ppt").lower(), DEFAULT_SKIN)


def get_builder(skin=None, kind="ppt"):
    skin = resolve_skin(skin, kind)
    kind = (kind or "ppt").lower()
    fn = _REG.get((skin, kind))
    if fn is None and kind != "ppt":
        # 해당 스킨에 그 종류가 없으면 navy 로 폴백
        fn = _REG.get(("navy", kind))
    if fn is None:
        fn = _REG.get((skin, "ppt")) or _REG.get((DEFAULT_SKIN, "ppt"))
    return fn


def list_options():
    """앱 STEP7 에서 선택지 렌더용."""
    skins = {}
    for (sk, kd) in _REG:
        skins.setdefault(sk, []).append(kd)
    return {
        "skins": {k: sorted(v) for k, v in skins.items()},
        "default_skin": DEFAULT_SKIN,
        "navy_palettes": [
            {"key": k, "label": v["label"],
             "primary": v["primary"], "accent": v["accent"]}
            for k, v in NAVY_PALETTES.items()
        ],
    }


def render(schema, skin=None, kind="ppt", palette=None):
    """단일 진입점. BytesIO 반환."""
    use_skin = resolve_skin(skin, kind)
    fn = get_builder(skin, kind)
    pal = normalize_palette(use_skin, palette)

    data = schema
    if use_skin == "navy":
        # 실제 v3 스키마(academy/features/targets)를 navy 형식으로 변환
        try:
            from navy_adapt import adapt
            data = adapt(schema)
        except Exception:
            data = schema

    try:
        return fn(data, None, pal)
    except TypeError:
        # 옛 빌더가 palette 인자를 안 받는 경우
        return fn(data, None)
