# -*- coding: utf-8 -*-
"""
브랜드북 PPT 서버 v3 (FastAPI)
POST /build  { data: <표준스키마 or 유미니원본>, brand?: {...}, palette: "teal_blue|navy_amber|green_orange", kind: "ppt" }
  -> PPTX 바이너리 반환
GET /        health
"""
from fastapi import FastAPI, Request
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import io, traceback

import build_brandbook_v3 as B
import build_brandbook_book as BOOK
import to_schema_v3 as TS
import build_leaflet_coord as LEAFLET
from navy_registry import render as skin_render, list_options as navy_options

app = FastAPI()
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)

VALID_PAL = {"teal_blue","navy_amber","green_orange"}

def _is_schema(d):
    # 이미 표준 스키마인지(academy 키 + intro/features 등) 판별
    return isinstance(d, dict) and "academy" in d and ("intro" in d or "features" in d or "timetables" in d)

@app.get("/")
def health():
    return {"ok": True, "service": "brandbook-ppt-server-v3",
            "palettes": sorted(VALID_PAL), "skins": ["v3", "book", "navy"],
            "navy": navy_options(),
            "leaflet_palettes": ["sewon_teal", "sewon_yellow", "navy_amber"]}

GRADE_KW = {"elem": "초", "mid": "중", "high": "고"}


def _grade_of(label):
    """'초등'·'초등부'·'초3' 등 어떤 표기든 초/중/고 판별. 못 찾으면 ''.
    '기초'·'초격차' 등 학년과 무관한 '초'는 초등으로 오분류하지 않는다."""
    import re as _re
    s = str(label or "").strip()
    if not s:
        return ""
    if _re.search(r"초[1-6]|초등", s):
        return "초"
    if _re.search(r"중[1-3]|중등|중학|예비중", s):
        return "중"
    if _re.search(r"고[1-3]|고등|고교|수능|정시|수시|재수", s):
        return "고"
    return ""


def _apply_filter(schema, filt):
    """앱 STEP7의 학년·시간표 그룹 선택을 서버에서 최종 적용.
    - grade가 elem/mid/high면 그 학년만 남김(시간표·수업대상·커리큘럼)
    - ttOff[그룹]=True면 그 시간표 그룹 제거
    - ttOff_all=True면 시간표 전체 제거"""
    if not isinstance(filt, dict) or not filt:
        return schema
    grade = (filt.get("grade") or "all")
    tt_off = filt.get("ttOff") or {}
    tt_off_all = bool(filt.get("ttOff_all"))

    # 시간표 전체 끄기
    if tt_off_all:
        schema["timetables"] = []
        return schema

    tts = schema.get("timetables") or []

    # 학년 필터
    kw = GRADE_KW.get(grade, "")
    if kw:
        tts = [t for t in tts if _grade_of(t.get("group", "")) == kw]
        # 수업 대상
        tg = schema.get("targets") or {}
        if isinstance(tg, dict) and isinstance(tg.get("items"), list):
            tg["items"] = [it for it in tg["items"] if _grade_of(it.get("grade", "")) == kw]
            if not tg["items"]:
                schema["targets"] = {}
        # 커리큘럼 (name 또는 title에 학년이 있을 수 있음)
        cu = schema.get("curriculum") or {}
        if isinstance(cu, dict) and isinstance(cu.get("stages"), list):
            st = [s for s in cu["stages"]
                  if not _grade_of(s.get("name", "") or s.get("title", ""))
                  or _grade_of(s.get("name", "") or s.get("title", "")) == kw]
            cu["stages"] = st
            if not st:
                schema["curriculum"] = {}

    # 그룹별 끄기 (라벨 표기 차이 허용)
    if tt_off:
        def _off(group):
            g = str(group or "")
            if tt_off.get(g):
                return True
            for k, v in tt_off.items():
                if v and (str(k).strip() == g.strip() or _grade_of(k) and _grade_of(k) == _grade_of(g)):
                    return True
            return False
        tts = [t for t in tts if not _off(t.get("group", ""))]

    schema["timetables"] = tts
    return schema


@app.post("/build")
async def build(req: Request):
    try:
        payload = await req.json()
    except Exception:
        return JSONResponse({"error":"invalid json"}, status_code=400)

    data = payload.get("data") or payload.get("schema") or payload
    raw  = payload.get("raw")   # 유미니 원본(시간표 보강용)
    brand = payload.get("brand") or {}
    filt = payload.get("filter") or {}   # {grade, ttOff, ttOff_all} — 앱 STEP7 선택
    palette = payload.get("palette") or payload.get("design") or "teal_blue"
    if palette not in VALID_PAL:
        # 앱 팔레트명 매핑(6종 → 3종)
        MAP = {"forest":"green_orange","sage":"green_orange","navy":"navy_amber",
               "violet":"navy_amber","crimson":"green_orange","mono":"navy_amber",
               "green":"green_orange","blue":"teal_blue","orange":"green_orange",
               "teal":"teal_blue"}
        palette = MAP.get(palette, "teal_blue")

    # 스키마 완성본이면 그대로, 아니면 유미니 원본으로 보고 변환
    try:
        if _is_schema(data):
            schema = data
            # 유미니 원본(raw)이 있으면 빈 값들을 raw에서 보강 (시간표·연락처 독립 처리)
            if raw:
                try:
                    conv = TS.convert(raw, brand)
                    # 시간표: 비어 있으면 채움
                    if not schema.get("timetables") and conv.get("timetables"):
                        schema["timetables"] = conv["timetables"]
                    # 연락처·위치·시간·과목: 비어 있으면 항상 채움 (시간표 유무와 무관)
                    ac = schema.setdefault("academy", {})
                    cac = conv.get("academy", {})
                    for k in ("phone","location","hours","address_short","subjects"):
                        if not ac.get(k) and cac.get(k):
                            ac[k] = cac[k]
                except Exception:
                    pass
        else:
            schema = TS.convert(data, brand)
    except Exception as e:
        return JSONResponse({"error":"schema convert failed","detail":str(e),
                             "trace":traceback.format_exc()}, status_code=500)

    # ── STEP7 선택(학년·시간표 그룹) 적용 ──
    # 앱이 raw를 함께 보내면 서버가 원본에서 시간표를 채우는데,
    # 그때 학년 필터가 빠져 전체 시간표가 나오던 문제를 여기서 바로잡는다.
    try:
        schema = _apply_filter(schema, filt)
    except Exception:
        pass

    # 스킨 선택 (기본 v3, book 선택 시 책펼침 렌더러)
    skin = (payload.get("skin") or "v3").lower()

    # 빌드
    try:
        if skin == "navy":
            kind = (payload.get("kind") or "ppt").lower()
            buf = skin_render(schema, skin="navy", kind=kind,
                              palette=payload.get("palette"))
        else:
            buf = io.BytesIO()
            if skin == "book":
                BOOK.build(schema, palette=palette, out=buf)
            else:
                B.build(schema, palette=palette, out=buf)   # build가 file-like도 받게
            buf.seek(0)
        fname = (schema.get("academy",{}).get("name","brandbook")) + f"_{palette}.pptx"
        return Response(
            content=buf.read(),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="brandbook.pptx"'}
        )
    except Exception as e:
        return JSONResponse({"error":"build failed","detail":str(e),
                             "trace":traceback.format_exc()}, status_code=500)


@app.post("/build-leaflet")
async def build_leaflet(req: Request):
    """3단 리플렛(2슬라이드) 생성. /build 와 같은 payload 사용.
    payload: { data, raw?, brand?, filter?, palette?, assets? }
      palette: sewon_teal | sewon_yellow | navy_amber (약칭 teal/yellow/blue 등 허용)
      assets : {"logo": "/서버경로", "cover": "/서버경로"} (외부 URL은 받지 않음)
    """
    try:
        payload = await req.json()
    except Exception:
        return JSONResponse({"error": "invalid json"}, status_code=400)

    data = payload.get("data") or payload.get("schema") or payload
    raw = payload.get("raw")
    brand = payload.get("brand") or {}
    filt = payload.get("filter") or {}

    # 팔레트: 리플렛 전용(teal/yellow/navy). 어떤 키로 와도 받음.
    pal = (payload.get("palette") or payload.get("template")
           or payload.get("design") or "sewon_teal").lower()
    alias = {"teal": "sewon_teal", "yellow": "sewon_yellow", "blue": "sewon_teal",
             "green": "sewon_teal", "orange": "sewon_yellow", "navy": "navy_amber"}
    if pal not in LEAFLET.PALETTES:
        pal = alias.get(pal, "sewon_teal")

    # 스키마 변환/보강 (브랜드북 /build 와 동일 로직)
    try:
        if _is_schema(data):
            schema = data
            if raw:
                try:
                    conv = TS.convert(raw, brand)
                    for key in ("timetables", "faq", "management", "admission",
                                "curriculum", "targets", "features", "achievements",
                                "specials", "intro"):
                        if not schema.get(key) and conv.get(key):
                            schema[key] = conv[key]
                    ac = schema.setdefault("academy", {})
                    for k, v in (conv.get("academy") or {}).items():
                        if not ac.get(k) and v:
                            ac[k] = v
                except Exception:
                    pass
        else:
            schema = TS.convert(data, brand)
    except Exception as e:
        return JSONResponse({"error": "schema convert failed", "detail": str(e),
                             "trace": traceback.format_exc()}, status_code=500)

    try:
        schema = _apply_filter(schema, filt)
    except Exception:
        pass

    try:
        buf = io.BytesIO()
        LEAFLET.build(schema, palette=pal, out=buf, assets=payload.get("assets") or {})
        buf.seek(0)
        return Response(
            content=buf.read(),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="leaflet.pptx"'},
        )
    except Exception as e:
        return JSONResponse({"error": "leaflet build failed", "detail": str(e),
                             "trace": traceback.format_exc()}, status_code=500)
