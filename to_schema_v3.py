# -*- coding: utf-8 -*-
"""
유미니 원본 JSON + 앱 STEP2~5 생성물(brand) → v3 표준 스키마 변환
build_brandbook_v3.py 가 먹는 스키마(sewon.json 형태)를 만든다.
빈 섹션은 빈 값으로 두면 빌더가 알아서 슬라이드 생략.
"""
import re
from content_quality import normalize_raw

def _g(d, *keys, default=""):
    cur = d
    for k in keys:
        if isinstance(cur, dict) and k in cur and cur[k] is not None:
            cur = cur[k]
        else:
            return default
    return cur

_SUBJECT_KO = {"korean":"국어","math":"수학","english":"영어","science":"과학",
               "social":"사회","other":"기타","consulting":"컨설팅"}
def _subjects_str(subj):
    """유미니 subjects는 {math:true, science:false} 형태일 수 있음 → '수학 · 과학 전문'"""
    if isinstance(subj, str):
        return subj
    if isinstance(subj, dict):
        names = [_SUBJECT_KO.get(k, k) for k, v in subj.items() if v is True]
        if names:
            return " · ".join(names) + " 전문"
    if isinstance(subj, list):
        return " · ".join(_SUBJECT_KO.get(x, x) for x in subj)
    return ""

def _grade_of(name):
    import re as _re
    n = name or ""
    # '기초', '초격차' 등 학년과 무관한 '초'는 초등으로 오분류하지 않는다.
    # 초1~6 / 초등 / 초등부 / 예비중 등 명확한 학년 표기만 매칭.
    if _re.search(r"초[1-6]|초등|elem", n, _re.I): return "초등"
    if _re.search(r"중[1-3]|중등|중학|예비중", n): return "중등"
    if _re.search(r"고[1-3]|고등|고교|수능|정시|수시|재수|N수", n, _re.I): return "고등"
    if "특강" in n or "방학" in n: return "특강"
    return "기타"

def convert(raw, brand=None):
    """raw = 유미니 원본 JSON, brand = 앱 STEP2~5 생성물(slogan/intro 등)"""
    raw = normalize_raw(raw)
    brand = brand or {}
    basic = raw.get("basic", raw) if isinstance(raw, dict) else {}
    content = raw.get("content", {}) if isinstance(raw, dict) else {}
    ai = raw.get("aiProfile", {}) if isinstance(raw, dict) else {}

    intro_ch = raw.get("introChannel", {}) if isinstance(raw, dict) else {}
    ops = raw.get("operations", {}) if isinstance(raw, dict) else {}

    name = brand.get("name") or _g(basic,"academyName") or _g(basic,"name") or "우리학원"
    subjects = brand.get("subjects") or _subjects_str(basic.get("subjects")) or ""

    phone = (brand.get("phone") or _g(basic,"phoneNumber") or _g(basic,"phone")
             or _g(basic,"tel") or "")
    location = (brand.get("location") or _g(intro_ch,"howToCome")
                or _g(basic,"businessAddress") or _g(basic,"address") or _g(basic,"location") or "")
    hours = (brand.get("hours") or _g(ops,"operatingHours") or _g(basic,"operatingHours") or "")
    addr_short = (brand.get("address_short") or _g(basic,"businessAddress")
                  or location or "")

    academy = {
        "name": name,
        "subjects": subjects,
        "slogan": brand.get("slogan") or _g(ai,"positioning") or "",
        "location": location,
        "phone": phone,
        "hours": hours,
        "address_short": addr_short,
    }

    intro = {
        "head": brand.get("intro_head") or _g(ai,"oneLineDef") or f"{name} 소개",
        "body": brand.get("intro_body") or _g(basic,"intro") or _g(content,"intro") or "",
    }

    # 강점(features): brand 우선, 없으면 aiProfile.strengths
    features = brand.get("features") or []
    if not features:
        for st in (_g(ai,"strengths",default=[]) or []):
            if isinstance(st, dict):
                features.append({"title": st.get("title",""), "desc": st.get("desc","") or st.get("description","")})
            elif isinstance(st, str):
                features.append({"title": st[:12], "desc": st})
    # 원본 항목은 여기서 자르지 않는다. 페이지 수용량 판단은 빌더가 맡는다.

    # 실적
    achievements = brand.get("achievements") or {}
    if not achievements:
        ach_items = _g(content,"achievements",default=[]) or _g(basic,"achievements",default=[])
        items=[]
        for it in (ach_items or []):
            if isinstance(it, dict):
                items.append({"icon": it.get("icon","star"),
                              "name": it.get("name","") or it.get("title","") or it.get("label",""),
                              "desc": it.get("desc","") or it.get("description","") or it.get("text","")})
            elif isinstance(it, str) and it.strip():
                items.append({"icon": "star", "name": "", "desc": it.strip()})
        if items:
            achievements = {"head":"주요 실적", "items": items}

    # 수업 대상(targets): classProfiles/divisions에서 학년별로
    targets = brand.get("targets") or {}
    if not targets:
        by_grade = {}
        for cp in (_g(raw,"classProfiles",default=[]) or _g(content,"divisions",default=[]) or []):
            if not isinstance(cp, dict): continue
            g = _grade_of(cp.get("className","") or cp.get("name",""))
            if g in ("특강","기타"): continue
            by_grade.setdefault(g, []).append(cp.get("className","") or cp.get("name",""))
        # classProfiles가 불완전해도 실제 개설 수업은 빠뜨리지 않는다.
        for course in (_g(raw, "courses", default=[]) or []):
            if not isinstance(course, dict): continue
            course_name = course.get("name","") or course.get("className","")
            g = _grade_of(course_name)
            if g not in ("특강", "기타") and course_name:
                by_grade.setdefault(g, []).append(course_name)
        order=[("초등","초등부"),("중등","중등부"),("고등","고등부")]
        items=[]
        for key,lab in order:
            if by_grade.get(key):
                items.append({"grade":lab, "subj":" · ".join(dict.fromkeys(by_grade[key]))[:60], "desc":""})
        if items:
            targets = {"head":"학년에 꼭 맞는 과정","items":items}

    # 커리큘럼(단계): content.curriculum 또는 brand
    curriculum = brand.get("curriculum") or {}
    if curriculum and "stages" in curriculum:
        # tag/level 보정
        for i,st in enumerate(curriculum["stages"]):
            st.setdefault("level", i+1)
            # 근거 없는 학년별 상투어를 자동 생성하지 않는다.
            st.setdefault("tag", "")

    # 시간표(courses → 학년 그룹). 유미니 실제 구조:
    #   course = {name, staffName, room, weeklySchedule:{slots:[{day,start,duration_minutes}]}}
    timetables = brand.get("timetables") or []
    if not timetables:
        DAY_KO = {"MONDAY":"월","TUESDAY":"화","WEDNESDAY":"수","THURSDAY":"목",
                  "FRIDAY":"금","SATURDAY":"토","SUNDAY":"일"}
        rows_raw = _g(raw,"courses",default=[]) or _g(content,"timetables",default=[])
        groups = {}
        for c in (rows_raw or []):
            if not isinstance(c, dict): continue
            cls = c.get("name","") or c.get("className","") or c.get("courseName","")
            if not cls: continue
            g = _grade_of(cls)
            # 시간 문자열 조립: weeklySchedule.slots → "월 14:00, 금 14:00"
            time_str = c.get("schedule","") or c.get("time","")
            if not time_str:
                ws = c.get("weeklySchedule") or {}
                slots = ws.get("slots") if isinstance(ws, dict) else (ws if isinstance(ws, list) else [])
                if not slots and isinstance(c.get("slots"), list):
                    slots = c.get("slots")
                parts = []
                for s in (slots or []):
                    if not isinstance(s, dict): continue
                    d = DAY_KO.get(s.get("day",""), s.get("day",""))
                    st = s.get("start","") or s.get("startTime","")
                    if d or st:
                        parts.append(f"{d} {st}".strip())
                time_str = ", ".join(parts)
            teacher = c.get("staffName","") or c.get("teacher","") or c.get("instructor","")
            room = c.get("room","") or c.get("classroom","") or "-"
            groups.setdefault(g, []).append([cls, time_str, teacher, room or "-"])
        order = ["초등","중등","고등","특강","기타"]
        for g in order:
            if groups.get(g):
                timetables.append({"group": g, "rows": groups[g]})

    # 특별프로그램
    specials = brand.get("specials") or {}
    if not specials:
        sp = _g(content,"specialPrograms",default=[]) or _g(ai,"specials",default=[])
        items=[]
        for i,it in enumerate(sp or []):
            if isinstance(it, dict):
                items.append({"no": f"{i+1:02d}", "title": it.get("title","") or it.get("name",""),
                              "desc": it.get("desc","") or it.get("description","")})
            elif isinstance(it, str) and it.strip():
                items.append({"no": f"{i+1:02d}", "title": it.strip(), "desc": ""})
        if items:
            specials={"head":"특별 프로그램","items":items}

    # 과목별 관리
    management = brand.get("management") or {}
    if not management:
        # ★학습관리 = 과제·평가·기준 미달 조치. 출결·결석·퇴원·환불(운영규정)과 섞지 않는다.
        #   과제(operatingRules.homework)를 규정에서 빼고 여기 넣지 않아 통째로 사라졌었다.
        _ai_p = _g(raw, "aiProfile", default={}) or {}
        _cps  = [c for c in (_g(raw, "classProfiles", default=[]) or []) if isinstance(c, dict)]
        _op   = _g(content, "operatingRules", default={}) or {}
        _hw   = str(_op.get("homework") or "").strip() if isinstance(_op, dict) else ""
        _low  = [str(x).strip() for x in (_ai_p.get("lowScorePolicies") or []) if str(x).strip()]
        _cols = []
        for cp in _cps:
            subj = str(cp.get("subject") or "").strip()
            rows = []
            hp = cp.get("homeworkPolicy")
            hp = ([str(x).strip() for x in hp if str(x).strip()] if isinstance(hp, list)
                  else ([str(hp).strip()] if str(hp or "").strip() else []))
            if not hp and _hw: hp = [_hw]
            if hp: rows.append({"k": "과제", "v": " · ".join(hp)})
            asm = [str(a_.get("type") or "").strip() for a_ in (cp.get("assessments") or [])
                   if isinstance(a_, dict) and str(a_.get("type") or "").strip()]
            if asm: rows.append({"k": "평가", "v": " · ".join(dict.fromkeys(asm))})
            if _low: rows.append({"k": "기준 미달 시", "v": " · ".join(_low)})
            if rows: _cols.append({"name": subj or "학습관리", "rows": rows})
        if not _cols and (_hw or _low):
            rows = []
            if _hw:  rows.append({"k": "과제", "v": _hw})
            if _low: rows.append({"k": "기준 미달 시", "v": " · ".join(_low)})
            _cols = [{"name": "학습관리", "rows": rows}]
        if _cols:
            management = {"head": "", "columns": _cols, "style": "auto"}

    # 입학절차
    admission = brand.get("admission") or {}
    if not admission:
        steps = _g(content, "admissionSteps", default=[]) or []
        items = []
        for i, it in enumerate(steps):
            if isinstance(it, dict):
                items.append({"no": i+1, "step": it.get("step","") or it.get("title","") or it.get("name",""),
                              "desc": it.get("desc","") or it.get("description","")})
            elif isinstance(it, str) and it.strip():
                items.append({"no": i+1, "step": it.strip(), "desc": ""})
        if items:
            admission = {"head": "입학 절차", "items": items}

    # 규정
    rules = brand.get("rules") or {}
    if not rules:
        op = _g(content,"operatingRules",default={})
        if isinstance(op, dict) and op:
            # 과제·테스트·피드백은 학습관리이며 운영규정에 넣지 않는다.
            keymap=[("attendance","출결"),("absence","결석"),("withdrawal","퇴원"),("refund","환불")]
            items=[{"k":lab,"v":op[k]} for k,lab in keymap if op.get(k)]
            if items: rules={"head":"학원 관리 지침","items":items}

    # FAQ
    faq = brand.get("faq") or {}
    if not faq:
        fq = _g(content,"faq",default=[]) or _g(content,"faqs",default=[])
        items=[]
        for it in (fq or []):
            if isinstance(it, dict) and (it.get("q") or it.get("question")):
                items.append({"q":it.get("q","") or it.get("question",""),
                              "a":it.get("a","") or it.get("answer","")})
        if items: faq={"head":"","items":items}

    # 마무리
    closing = brand.get("closing") or {
        "head": "",
        "highlight": name,
        "cta": "지금 바로 상담을 예약하세요"
    }

    return {
        "academy": academy, "intro": intro, "features": features,
        "achievements": achievements, "targets": targets, "curriculum": curriculum,
        "timetables": timetables, "specials": specials, "management": management,
        "admission": admission, "rules": rules, "faq": faq, "closing": closing,
        "_dataAudit": raw.get("_quality") or {},
    }
