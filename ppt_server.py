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
import to_schema_v3 as TS

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
    return {"ok": True, "service": "brandbook-ppt-server-v3", "palettes": sorted(VALID_PAL)}

@app.post("/build")
async def build(req: Request):
    try:
        payload = await req.json()
    except Exception:
        return JSONResponse({"error":"invalid json"}, status_code=400)

    data = payload.get("data") or payload.get("schema") or payload
    brand = payload.get("brand") or {}
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
        else:
            schema = TS.convert(data, brand)
    except Exception as e:
        return JSONResponse({"error":"schema convert failed","detail":str(e),
                             "trace":traceback.format_exc()}, status_code=500)

    # 빌드
    try:
        buf = io.BytesIO()
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
