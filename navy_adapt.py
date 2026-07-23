# -*- coding: utf-8 -*-
"""
navy_adapt.py — 실제 v3 스키마(to_schema_v3 출력) → navy 빌더 입력 형식 변환

실제 v3 스키마:
    academy{name,subjects,slogan,location,phone,hours,address_short}
    intro{head,body}, features[{title,desc}]
    achievements{head,items[{icon,name,desc}]}
    targets{head,items[{grade,subj,desc}]}
    curriculum{stages[{name,title,desc,tag,level,items}]}
    timetables[{group,rows:[[반,시간,강사,강의실], ...]}]   ← rows 가 배열!
    specials{head,items[{no,title,desc}]}
    management{...}, admission{...}, rules{head,items[{k,v}]}
    faq{head,items[{q,a}]}, closing{head,highlight,cta}

navy 빌더가 읽는 형식:
    basic{name,slogan,target,subline,phone,address}
    identity{headline,lead,intro}
    strengths[{title,desc}], philosophy{...}
    curriculum[{name,title,desc,tags}]
    timetables[{group,rows:[{name,time}]}]                  ← rows 가 dict!
    ...

navy 빌더는 원래 형식도 그대로 받으므로, 이미 변환된 것은 손대지 않는다.
"""
import re


def _s(v):
    return "" if v is None else str(v).strip()


def _items(x):
    """{head, items:[...]} 또는 [...] 둘 다 리스트로."""
    if isinstance(x, dict):
        it = x.get("items")
        return it if isinstance(it, list) else []
    if isinstance(x, list):
        return x
    return []


def _head(x, default=""):
    if isinstance(x, dict):
        return _s(x.get("head")) or default
    return default


def is_v3_schema(d):
    """to_schema_v3 가 낸 실제 v3 스키마인가?"""
    return isinstance(d, dict) and isinstance(d.get("academy"), dict)


def is_navy_schema(d):
    """navy 빌더 네이티브 형식인가?"""
    return isinstance(d, dict) and isinstance(d.get("basic"), dict)


def _target_from(schema):
    """'초등 · 중등' 같은 대상 문자열 만들기."""
    labs = []
    for it in _items(schema.get("targets")):
        if isinstance(it, dict):
            g = _s(it.get("grade"))
            if g:
                labs.append(g)
    if not labs:
        for t in (schema.get("timetables") or []):
            if isinstance(t, dict):
                g = _s(t.get("group"))
                if g and g not in ("특강", "기타"):
                    labs.append(g)
    seen, out = set(), []
    for l in labs:
        if l not in seen:
            seen.add(l)
            out.append(l)
    return " · ".join(out[:3])


def _rows_to_dicts(rows):
    """
    시간표 rows 정규화.
    실제 v3: [[반, 시간, 강사, 강의실], ...]  (배열)
    navy   : [{name, time}, ...]              (딕셔너리)
    """
    out = []
    for r in (rows or []):
        if isinstance(r, dict):
            nm = _s(r.get("name") or r.get("class") or r.get("className"))
            tm = _s(r.get("time") or r.get("when") or r.get("schedule"))
            if nm or tm:
                out.append({"name": nm, "time": tm})
        elif isinstance(r, (list, tuple)) and r:
            nm = _s(r[0])
            tm = _s(r[1]) if len(r) > 1 else ""
            if nm or tm:
                out.append({"name": nm, "time": tm})
        elif isinstance(r, str) and r.strip():
            out.append({"name": r.strip(), "time": ""})
    return out


def _curriculum(schema):
    cu = schema.get("curriculum")
    stages = []
    if isinstance(cu, dict):
        stages = cu.get("stages") or []
    elif isinstance(cu, list):
        stages = cu
    out = []
    for st in stages:
        if not isinstance(st, dict):
            if isinstance(st, str) and st.strip():
                out.append({"name": st.strip(), "title": "", "desc": "",
                            "tags": []})
            continue
        tags = st.get("tags")
        if not isinstance(tags, list) or not tags:
            tags = [t for t in [_s(st.get("tag"))] if t]
        # items 가 있으면 설명으로 합침
        desc = _s(st.get("desc") or st.get("body"))
        if not desc and isinstance(st.get("items"), list):
            desc = " · ".join(_s(x) for x in st["items"] if _s(x))[:200]
        out.append({
            "name": _s(st.get("name") or st.get("grade")),
            "title": _s(st.get("title")),
            "desc": desc,
            "tags": [_s(t) for t in tags if _s(t)][:3],
        })
    return out


def _achievements(schema):
    out = []
    for it in _items(schema.get("achievements")):
        if not isinstance(it, dict):
            continue
        # 실제 v3 는 {icon,name,desc} — 숫자 값이 따로 없다
        name = _s(it.get("name") or it.get("title") or it.get("label"))
        val = _s(it.get("value") or it.get("num"))
        desc = _s(it.get("desc") or it.get("body"))
        if not val:
            # 설명에서 수치를 뽑아 큰 숫자로.
            # '최근 3개년' 같은 기간 표현은 실적 수치가 아니므로 제외한다.
            src = f"{desc} {name}"
            src = re.sub(r"(최근\s*)?\d+\s*개?년", " ", src)
            m = re.search(r"(\d[\d,]*\s*(?:명|건|개교|개|%|위|등))", src)
            val = m.group(1).replace(" ", "") if m else ""
        if not (name or val or desc):
            continue
        out.append({"value": val or name[:6], "title": name, "desc": desc})
    return out


def _kv_items(x):
    """rules/management 의 {k,v} 형태를 {key,desc} 로."""
    out = []
    for it in _items(x):
        if isinstance(it, dict):
            k = _s(it.get("k") or it.get("key") or it.get("title"))
            v = _s(it.get("v") or it.get("desc") or it.get("body"))
            if k or v:
                out.append({"key": k[:6], "title": k, "desc": v})
        elif isinstance(it, str) and it.strip():
            out.append({"key": it.strip()[:6], "title": it.strip(), "desc": ""})
    return out


def adapt(schema):
    """실제 v3 스키마 → navy 빌더 입력. 이미 navy 형식이면 그대로 반환."""
    if not isinstance(schema, dict):
        return {}
    if is_navy_schema(schema) and not is_v3_schema(schema):
        return schema
    if not is_v3_schema(schema):
        return schema

    ac = schema.get("academy") or {}
    intro = schema.get("intro") or {}
    closing = schema.get("closing") or {}

    name = _s(ac.get("name")) or "학원"
    slogan = _s(ac.get("slogan")) or _s(intro.get("head"))

    d = {
        "basic": {
            "name": name,
            "slogan": slogan,
            "target": _target_from(schema),
            "subline": _s(ac.get("subjects")),
            "phone": _s(ac.get("phone")),
            "address": _s(ac.get("address_short") or ac.get("location")),
        },
        "identity": {
            "headline": _s(intro.get("head")) or f"{name} 소개",
            "lead": "",
            "intro": _s(intro.get("body")),
        },
    }

    # 강점
    feats = []
    for f in (schema.get("features") or []):
        if isinstance(f, dict):
            t = _s(f.get("title"))
            ds = _s(f.get("desc") or f.get("description"))
            if t or ds:
                feats.append({"title": t or ds[:14], "desc": ds})
        elif isinstance(f, str) and f.strip():
            feats.append({"title": f.strip()[:14], "desc": f.strip()})
    d["strengths"] = feats

    # 수업 대상 — 학년별 과목/반 구성. 별도 섹션으로 넘긴다.
    tgts = []
    for it in _items(schema.get("targets")):
        if isinstance(it, dict):
            g = _s(it.get("grade"))
            sj = _s(it.get("subj"))
            ds = _s(it.get("desc"))
            if g or sj or ds:
                tgts.append({"grade": g, "subj": sj, "desc": ds})
    d["targets"] = tgts
    d["targetsHead"] = _head(schema.get("targets"), "학년별 수업 대상")

    # 학원 소개 — 유미니 JSON의 소개문구(intro.body)를 그대로 쓴다.
    # 대부분의 학원이 소개문구를 가지고 있으므로 이 페이지는 거의 항상 만들어진다.
    body = _s(intro.get("body"))
    if body:
        ihead = _s(intro.get("head"))
        # 표지 슬로건이나 강점 페이지 헤드라인과 겹치면 중복이므로
        # 일반 소개 제목을 쓴다 (강점 헤드라인 = identity.headline)
        used = {slogan, _s(d.get("identity", {}).get("headline"))}
        if not ihead or ihead in used:
            ihead = f"{name} 소개"
        d["philosophy"] = {
            "headline": ihead,
            "intro": body,
            "points": [],
            "note": {},
        }

    d["curriculum"] = _curriculum(schema)
    d["curriculumLead"] = _head(schema.get("curriculum"), "")

    # 특별 프로그램
    sps = []
    for it in _items(schema.get("specials")):
        if isinstance(it, dict):
            t = _s(it.get("title"))
            ds = _s(it.get("desc"))
            if t or ds:
                sps.append({"title": t, "desc": ds})
    d["specials"] = sps

    # 학습관리: management 우선, 없으면 rules
    mg = _kv_items(schema.get("management"))
    if not mg:
        mg = _kv_items(schema.get("rules"))
    d["management"] = mg

    d["achievements"] = _achievements(schema)
    d["achievementsHead"] = _head(schema.get("achievements"), "주요 실적")

    # 시간표 (rows 배열 → dict)
    tts = []
    for t in (schema.get("timetables") or []):
        if not isinstance(t, dict):
            continue
        rows = _rows_to_dicts(t.get("rows"))
        if rows:
            tts.append({"group": _s(t.get("group")) or "수업", "rows": rows})
    d["timetables"] = tts

    # 입학절차
    adm = []
    for it in _items(schema.get("admission")):
        if isinstance(it, dict):
            t = _s(it.get("title") or it.get("name") or it.get("k"))
            ds = _s(it.get("desc") or it.get("body") or it.get("v"))
            if t or ds:
                adm.append({"title": t, "desc": ds})
        elif isinstance(it, str) and it.strip():
            adm.append({"title": it.strip(), "desc": ""})
    d["admission"] = adm

    # FAQ
    fq = []
    for it in _items(schema.get("faq")):
        if isinstance(it, dict) and _s(it.get("q")):
            fq.append({"q": _s(it.get("q")), "a": _s(it.get("a"))})
    d["faq"] = fq

    # 마무리/연락처
    d["contact"] = {
        "headline": _s(closing.get("head")).replace("\n", " "),
        "title": "상담 전에 알려주시면 좋아요",
        "asks": [],
        "closing": _s(closing.get("cta")),
    }

    # 이미지
    assets = schema.get("assets")
    if isinstance(assets, dict):
        d["assets"] = assets

    return d
