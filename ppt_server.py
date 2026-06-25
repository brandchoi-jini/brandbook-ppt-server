#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt_server.py — 브랜드북 PPT 생성 서버 (FastAPI)

앱(STEP 7)이 콘텐츠 JSON을 POST /build 로 보내면,
sky 템플릿을 편집해 만든 PPTX 파일을 돌려준다.

배포 (Railway / Render 등):
    pip install fastapi uvicorn python-pptx pillow
    uvicorn ppt_server:app --host 0.0.0.0 --port 8000

필요 파일(같은 폴더):
    build_brandbook.py   빌더
    templates/sky.pptx   템플릿
    assets/              로고·사진 (요청에 base64로 받지 않을 경우의 기본값)
"""
import os, io, json, base64, tempfile, shutil
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import build_brandbook

app = FastAPI(title="Brandbook PPT Builder")

# 앱(Vercel)에서 호출하므로 CORS 허용
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],      # 운영 시 앱 도메인으로 제한 권장
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE, "templates")
DEFAULT_ASSETS = os.path.join(BASE, "assets")

@app.get("/")
def health():
    return {"ok": True, "service": "brandbook-ppt-builder"}

@app.post("/build")
async def build_ppt(request: Request):
    content = await request.json()

    # 임시 작업 폴더 (요청별 격리)
    workdir = tempfile.mkdtemp(prefix="bb_")
    try:
        # 1) 에셋 준비: 요청에 base64 이미지가 있으면 저장, 없으면 기본 에셋 복사
        assets_in = content.pop("_assets", None)  # {"logo.png":"<base64>", ...}
        if assets_in:
            for fname, b64 in assets_in.items():
                safe = os.path.basename(fname)
                with open(os.path.join(workdir, safe), "wb") as f:
                    f.write(base64.b64decode(b64.split(",")[-1]))
        elif os.path.isdir(DEFAULT_ASSETS):
            for fn in os.listdir(DEFAULT_ASSETS):
                shutil.copy(os.path.join(DEFAULT_ASSETS, fn), os.path.join(workdir, fn))

        # 2) 템플릿 선택
        tpl_name = (content.get("design") or {}).get("template", "sky.pptx")
        template_path = os.path.join(TEMPLATE_DIR, os.path.basename(tpl_name))
        if not os.path.exists(template_path):
            return JSONResponse({"error": f"template not found: {tpl_name}"}, status_code=400)

        # 3) 빌드
        out_path = os.path.join(workdir, "brandbook.pptx")
        n = build_brandbook.build(content, template_path, out_path, workdir)

        # 4) 파일 스트리밍 반환
        data = open(out_path, "rb").read()
        academy = (content.get("academy") or {}).get("name", "brandbook")
        headers = {
            "Content-Disposition": f'attachment; filename="brandbook.pptx"',
            "X-Slide-Count": str(n),
        }
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers=headers,
        )
    except Exception as e:
        import traceback; traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=500)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
