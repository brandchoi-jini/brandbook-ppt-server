# -*- coding: utf-8 -*-
"""
유미니 원본 JSON + 앱 STEP2~5 생성물(brand) → v3 표준 스키마 변환
build_brandbook_v3.py 가 먹는 스키마(sewon.json 형태)를 만든다.
빈 섹션은 빈 값으로 두면 빌더가 알아서 슬라이드 생략.
"""
# ★[헤더 정책] 슬라이드 <한마디>(head)는 "학원 소개"에만 둔다.
#   예전에는 이 변환기가 "주요 실적"·"학년에 꼭 맞는 과정"·"특별 프로그램"·
#   "입학 절차"·"학원 관리 지침" 같은 고정 문구를 넣어서,
#   앱이 head를 비워 보내도 슬라이드에 한마디가 찍혔다.
#   라벨과 같은 말이라 정보가 없고 어색해서 전부 빈 값으로 바꿨다.
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

# ── 실적 파싱 ────────────────────────────────────────────
# ★한 줄에 "A | B | C" 로 뭉쳐 들어오는 실적이 통째로 note 에 박혀
#   PPT 에서 한 덩어리 회색 글씨로 뭉개졌다. 여기서 잘라 유형별로 나눈다.
# ★가운뎃점(·)은 '서울대·고려대' 처럼 <병렬 나열>이라 구분자로 쓰면 안 된다.
#   파이프·세미콜론·줄바꿈만 항목 구분자로 본다.
_ACH_SPLIT = re.compile(r"\s*[|‖]\s*|\s*[;；]\s*|\n+")

_ACH_KIND = [
    ("대입 실적",     r"수능|정시|수시|대학|대입|서울대|연세|고려|의대|약대|교대|카이스트|KAIST|SKY|학과\)|최종\s*합격"),
    ("고입 실적",     r"고입|특목고|자사고|영재고|과학고|외고|국제고|자공고|고교\s*진학"),
    ("내신·성적 향상", r"내신|등급|점수|평균|만점|100점|\d+\s*점|모의고사|전교|상위권|최상위"),
    ("성장 사례",     r"성장|입학\s*후|재원|레벨\s*상향|올라|향상|이동|진급"),
]

def _ach_kind(text):
    t = text or ""
    for name, pat in _ACH_KIND:
        if re.search(pat, t, re.I):
            return name
    return "주요 실적"

def _ach_split_line(s):
    """'2026 서울대 최종 합격 | 동방고 영어 전교 1등 | ...' → 개별 항목 리스트"""
    parts = [p.strip(" .·-—") for p in _ACH_SPLIT.split(s or "") if p and p.strip(" .·-—")]
    return [p for p in parts if len(p) > 1]

def _ach_item(text):
    """한 항목 문자열 → {title, change, note}
    '동방고 수학 전교 1등 유지 → 직전 시험 100점' 처럼 화살표가 있으면 앞/뒤로 나눈다."""
    t = (text or "").strip()
    if not t:
        return None
    change = ""
    m = re.split(r"\s*(?:→|->|⇒|=>)\s*", t, maxsplit=1)
    if len(m) == 2:
        t, change = m[0].strip(), m[1].strip()
    # 괄호 보충 설명은 note 로 뺀다
    note = ""
    mb = re.match(r"^(.*?)\s*[(（]([^)）]{4,})[)）]\s*$", t)
    if mb:
        t, note = mb.group(1).strip(), mb.group(2).strip()
    # 제목이 너무 길면 앞부분만 제목, 나머지는 note
    if len(t) > 26 and not note:
        cut = t.rfind(" ", 0, 26)
        if cut > 8:
            t, note = t[:cut].strip(), (t[cut:].strip() + ((" " + note) if note else ""))
    return {"title": t, "change": change, "note": note}

def _build_ach_groups(raw_items):
    """평면 실적 입력 → {head, groups:[{name, items:[{title,change,note}]}]}"""
    flat = []
    for it in (raw_items or []):
        if isinstance(it, dict):
            base = " ".join(str(it.get(k, "")) for k in
                            ("name", "title", "label", "desc", "description", "text") if it.get(k))
            src = base.strip()
        elif isinstance(it, str):
            src = it.strip()
        else:
            continue
        if not src:
            continue
        for piece in (_ach_split_line(src) or [src]):
            flat.append(piece)
    order, buckets = [], {}
    for piece in flat:
        item = _ach_item(piece)
        if not item:
            continue
        k = _ach_kind(piece)
        if k not in buckets:
            buckets[k] = []
            order.append(k)
        # 같은 유형 내 중복 제거
        if any(x["title"] == item["title"] and x["change"] == item["change"] for x in buckets[k]):
            continue
        buckets[k].append(item)
    groups = [{"name": k, "items": buckets[k]} for k in order if buckets[k]]
    return groups


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
    # ★brand 쪽에 있어도 groups 가 없으면(=평면 items 뭉치) 여기서 분해·분류한다.
    #   실적이 PPT 에 안 보이거나 한 덩어리로 뭉개지던 원인.
    _ach_src = []
    if achievements:
        if achievements.get("groups"):
            _ach_src = []
        else:
            _ach_src = achievements.get("items") or []
    if not _ach_src and not (achievements.get("groups") if achievements else None):
        _ach_src = _g(content,"achievements",default=[]) or _g(basic,"achievements",default=[])
    if _ach_src:
        _groups = _build_ach_groups(_ach_src)
        if _groups:
            achievements = {"head": (achievements.get("head") if achievements else "") or "",
                            "groups": _groups}

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
            targets = {"head":"","items":items}

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
            specials={"head":"","items":items}

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
            admission = {"head": "", "items": items}

    # 규정
    rules = brand.get("rules") or {}
    if not rules:
        op = _g(content,"operatingRules",default={})
        if isinstance(op, dict) and op:
            # 과제·테스트·피드백은 학습관리이며 운영규정에 넣지 않는다.
            keymap=[("attendance","출결"),("absence","결석"),("withdrawal","퇴원"),("refund","환불")]
            items=[{"k":lab,"v":op[k]} for k,lab in keymap if op.get(k)]
            if items: rules={"head":"","items":items}

    def _faq_short(v, limit=90):
        """★FAQ 답변이 300~580자로 들어와 PPT·리플렛에서 안 읽혔다.
        문장 단위로 잘라 limit 안으로 줄인다(문장 중간은 자르지 않는다).
        원문은 STEP6·컨펌 문서에 그대로 남는다."""
        t = re.sub(r"\s+", " ", str(v or "")).strip()
        if len(t) <= limit:
            return t
        out = ""
        for sent in re.split(r"(?<=[.!?])\s+", t):
            if not sent.strip():
                continue
            if out and len(out) + len(sent) + 1 > limit:
                break
            out = (out + " " + sent).strip()
        if not out:                      # 한 문장이 이미 길면 절 단위로
            for sep in ("니다", ",", "·"):
                idx = t.rfind(sep, 0, limit)
                if idx > 20:
                    return t[:idx + len(sep)].strip()
            return t[:limit].strip()
        return out

    # ── 시간표 요약 ─────────────────────────────────────────
    # ★강좌가 104개(더케이)면 표를 그대로 실을 수 없다.
    #   학년별로 <주N회 · 요일 · 수업시간> 대표 패턴을 뽑아 한 줄로 안내한다.
    #   예) 중등  주3회 월수금 1시간 30분 · 주2회 화목 2시간
    def _schedule_summary(raw_):
        DAY = {"MONDAY":"월","TUESDAY":"화","WEDNESDAY":"수","THURSDAY":"목",
               "FRIDAY":"금","SATURDAY":"토","SUNDAY":"일"}
        ORD = {"월":0,"화":1,"수":2,"목":3,"금":4,"토":5,"일":6}
        def _grade(nm):
            if re.search(r"초등|초[1-6]|파닉스", nm): return "초등"
            if re.search(r"중등|중[1-3]|예비중", nm): return "중등"
            if re.search(r"고등|고[1-3]|수능", nm): return "고등"
            return ""
        def _hhm(m):
            m = int(m or 0)
            if m <= 0: return ""
            if m < 60: return f"{m}분"
            return f"{m//60}시간" + (f" {m%60}분" if m % 60 else "")
        from collections import Counter, defaultdict
        buckets = defaultdict(list)
        for c in (raw_.get("courses") or []):
            nm = str(c.get("name") or "").strip()
            if not nm: continue
            slots = ((c.get("weeklySchedule") or {}).get("slots") or [])
            days = sorted({DAY.get(s.get("day"), "") for s in slots if s.get("day")},
                          key=lambda d: ORD.get(d, 9))
            mins = [s.get("duration_minutes") for s in slots if s.get("duration_minutes")]
            if not days: continue
            g = _grade(nm)
            if not g: continue                      # 학년을 못 읽는 내부용 강좌는 제외
            buckets[g].append(("".join(days), mins[0] if mins else 0))
        out = []
        for g in ("초등", "중등", "고등"):
            if not buckets.get(g): continue
            pat = Counter(buckets[g])
            parts = []
            for (d, m), _n in pat.most_common(3):
                t = _hhm(m)
                parts.append(f"주{len(d)}회 {d}" + (f" {t}" if t else ""))
            if parts:
                out.append({"k": f"{g}부", "v": " · ".join(parts)})
        return out

    _sched = _schedule_summary(raw) if isinstance(raw, dict) else []

    # FAQ
    faq = brand.get("faq") or {}
    if not faq:
        fq = _g(content,"faq",default=[]) or _g(content,"faqs",default=[])
        items=[]
        for it in (fq or []):
            if isinstance(it, dict) and (it.get("q") or it.get("question")):
                items.append({"q": _faq_short(it.get("q","") or it.get("question",""), 40),
                              "a": _faq_short(it.get("a","") or it.get("answer",""), 90)})
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
        "schedule": ({"head": "", "items": _sched} if _sched else {}),
        "admission": admission, "rules": rules, "faq": faq, "closing": closing,
        "_dataAudit": raw.get("_quality") or {},
    }
