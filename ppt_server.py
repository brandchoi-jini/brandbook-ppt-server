#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
브랜드북 PPT 서버 (FastAPI)
앱(index.html) → POST /build → PPTX 반환
payload: 유미니 원본 JSON + brand + design.palette + richContent + kind
"""
import io, re, tempfile, os
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from build_report import Report

app = FastAPI()
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)

# ── 앱 payload → build_report content 구조로 변환 ──
DAYK = {"MONDAY":"월","TUESDAY":"화","WEDNESDAY":"수","THURSDAY":"목",
        "FRIDAY":"금","SATURDAY":"토","SUNDAY":"일"}

def _split_kv(s):
    """'제목: 설명' → (제목, 설명)"""
    m = re.match(r"^\s*([^:：]{1,30})\s*[:：]\s*(.+)$", s.strip())
    if m: return m.group(1).strip(), m.group(2).strip()
    return s.strip(), ""

def adapt(payload):
    """앱 richContent(+원본)를 build_report content dict로."""
    rc = payload.get("richContent") or {}
    raw = payload  # 원본 필드(basic, timetable 등)가 최상위에 있음
    brand = payload.get("brand") or {}

    content = {}
    # 브랜드
    content["brand"] = {
        "slogan":   rc.get("slogan") or brand.get("slogan") or "",
        "identity": rc.get("identity") or brand.get("identity") or "",
        "intro":    rc.get("intro") or brand.get("intro") or "",
    }
    # 강점: [{t,d}]
    strengths = []
    for it in (rc.get("strengths") or rc.get("features") or []):
        if isinstance(it, dict):
            t = it.get("title") or it.get("t") or it.get("name") or ""
            d = it.get("desc") or it.get("d") or ""
            if not t and d: t, d = _split_kv(d)
            elif not d and t: t, d = _split_kv(t)
        else:
            t, d = _split_kv(str(it))
        if t: strengths.append({"t": t, "d": d})
    content["strengths"] = strengths[:6]

    # 실적: [{tag,text}]
    ach = []
    for a in (rc.get("achievements") or []):
        if isinstance(a, dict):
            ach.append({"tag": a.get("tag",""), "text": a.get("text") or a.get("desc","")})
        else:
            s = str(a)
            m = re.match(r"^\[([^\]]+)\]\s*(.+)$", s)
            if m: ach.append({"tag": m.group(1), "text": m.group(2)})
            else: ach.append({"tag":"", "text": s})
    content["achievements"] = ach

    # 반 구성(수업대상): divisions{elem,mid,high} → classes{초등,중등,고등}
    div = rc.get("divisions") or {}
    classes = {}
    for key, ko in [("elem","초등"),("mid","중등"),("high","고등")]:
        arr = div.get(key) or []
        if arr:
            classes[ko] = [{"n": (x.get("name") if isinstance(x,dict) else str(x)),
                            "d": (x.get("desc","") if isinstance(x,dict) else "")} for x in arr]
    content["classes"] = classes

    # 관리: management[] → {과목: [{k,v}]} 또는 단일
    mng = rc.get("management") or []
    subjects_ko = []
    smap = {"math":"수학","science":"과학","korean":"국어","english":"영어","social":"사회"}
    for k,v in (raw.get("basic",{}).get("subjects") or {}).items():
        if v: subjects_ko.append(smap.get(k,k))
    # management가 [{title,desc}]이면 항목으로. 과목 구분이 없으면 단일 그룹.
    mg_items = []
    for m in mng:
        if isinstance(m, dict):
            kk = m.get("title") or m.get("k") or ""
            vv = m.get("desc") or m.get("v") or ""
            if not kk and vv: kk, vv = _split_kv(vv)
        else:
            kk, vv = _split_kv(str(m))
        if kk: mg_items.append({"k": kk, "v": vv})
    # 과목이 2개 이상이고 항목이 과목별로 반복되면 그대로, 아니면 첫 과목에
    if mg_items:
        if len(subjects_ko) >= 2 and len(mg_items) >= 8:
            half = len(mg_items)//len(subjects_ko)
            management = {}
            for i, sub in enumerate(subjects_ko):
                management[sub] = mg_items[i*half:(i+1)*half]
        else:
            management = {(subjects_ko[0] if subjects_ko else "학습"): mg_items}
    else:
        management = {}
    content["management"] = management

    # 특강
    specials = []
    for p in (rc.get("specials") or []):
        if isinstance(p, dict):
            t = p.get("title") or p.get("t") or ""
            d = p.get("desc") or p.get("d") or ""
            if not t and d: t, d = _split_kv(d)
        else:
            t, d = _split_kv(str(p))
        if t: specials.append({"t": t, "d": d})
    content["specials"] = specials

    # 입학절차
    admission = []
    for a in (rc.get("admission") or []):
        if isinstance(a, dict):
            admission.append({"step": a.get("step",""), "desc": a.get("desc","")})
        else:
            st, de = _split_kv(str(a))
            admission.append({"step": st, "desc": de})
    content["admission"] = admission

    # 규정
    rules = []
    for r in (rc.get("rules") or []):
        if isinstance(r, dict):
            rules.append({"k": r.get("k") or r.get("title",""), "v": r.get("v") or r.get("desc","")})
        else:
            kk, vv = _split_kv(str(r))
            rules.append({"k": kk, "v": vv})
    content["rules"] = rules

    # FAQ
    faq = []
    for f in (rc.get("faq") or []):
        if isinstance(f, dict):
            faq.append({"q": f.get("q") or f.get("question",""), "a": f.get("a") or f.get("answer","")})
    content["faq"] = faq

    # 커리큘럼: richContent.curriculum 형태에 따라. 없으면 생략(표지엔 영향 없음)
    cur = rc.get("curriculum")
    if isinstance(cur, dict) and cur.get("초등"):
        content["curriculum_math"] = cur

    # 컨택 (원본에서)
    ic = raw.get("introChannel") or {}
    ops = raw.get("operations") or {}
    b = raw.get("basic") or {}
    hours = []
    ot = ops.get("operatingTime") or {}
    if ot.get("weekdays"): hours.append({"k":"평일","v":f"{ot['weekdays']['start']} – {ot['weekdays']['end']}"})
    if ot.get("saturday"): hours.append({"k":"토·일","v":f"{ot['saturday']['start']} – {ot['saturday']['end']}"})
    links = []
    if ic.get("naverMapUrl"): links.append({"k":"네이버 지도","v":ic["naverMapUrl"].replace("https://","")})
    if ic.get("kakaoMapUrl"): links.append({"k":"카카오 지도","v":ic["kakaoMapUrl"].replace("https://","")})
    if ic.get("blogUrl"): links.append({"k":"블로그","v":ic["blogUrl"].replace("https://","")})
    content["contact"] = {
        "phone": b.get("phoneNumber") or ops.get("phone") or "",
        "address": ic.get("howToCome") or "",
        "parking": ic.get("parking") or "",
        "hours": hours, "links": links,
    }
    return content

@app.get("/")
def health():
    return {"status":"ok","service":"brandbook-ppt-server"}

@app.post("/build")
async def build(request: Request):
    try:
        payload = await request.json()
    except Exception as e:
        return JSONResponse({"error":f"잘못된 요청: {e}"}, status_code=400)

    kind = payload.get("kind","ppt")
    palette = ((payload.get("design") or {}).get("palette")) or "navy"
    # build_report는 navy/navy2 지원. 다른 색이면 navy로 폴백(추후 팔레트 확장).
    if palette not in ("navy","navy2"):
        # 색상 id를 build_report 팔레트로 매핑 (임시: 전부 navy 계열)
        palette = "navy"

    try:
        content = adapt(payload)
        # raw(원본)는 시간표·basic 위해 그대로 전달
        raw = payload
        rep = Report(content, raw, palette)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        rep.build(tmp.name)
        tmp.close()
        data = open(tmp.name,"rb").read()
        os.unlink(tmp.name)
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition":'attachment; filename="brandbook.pptx"'}
        )
    except Exception as e:
        import traceback
        return JSONResponse({"error":f"생성 실패: {e}","trace":traceback.format_exc()[:500]}, status_code=500)
