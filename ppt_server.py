#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt_server.py — 브랜드북 산출물 생성 FastAPI 서버.
POST /build : 표준 콘텐츠 스키마(JSON)를 받아 kind에 따라 PPT/카탈로그/리플렛 PPTX 반환.
GET  /      : health check.

payload.kind: 'ppt' | 'catalog' | 'leaflet' (기본 ppt)
템플릿 파일은 templates/ 폴더에 위치.
"""
import os, io, tempfile
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import build_brandbook_v2 as B

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)

TPL_DIR = os.path.join(os.path.dirname(__file__), 'templates')

@app.get('/')
def health():
    return {"status": "ok", "service": "brandbook-ppt-server", "kinds": ["ppt", "catalog", "leaflet"]}

@app.post('/build')
async def build(req: Request):
    content = await req.json()
    kind = (content.get('kind') or 'ppt').lower()
    # 임시 출력 파일
    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tf:
        out_path = tf.name
    try:
        if kind == 'catalog':
            B.build_catalog(content, out_path, TPL_DIR)
            fname = 'catalog.pptx'
        elif kind == 'leaflet':
            B.build_leaflet(content, out_path, TPL_DIR)
            fname = 'leaflet.pptx'
        else:
            B.build(content, out_path, TPL_DIR)
            fname = 'brandbook.pptx'
        with open(out_path, 'rb') as f:
            data = f.read()
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "kind": kind})
    finally:
        try: os.remove(out_path)
        except Exception: pass
    return StreamingResponse(
        io.BytesIO(data),
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        headers={'Content-Disposition': f'attachment; filename="{fname}"'}
    )
