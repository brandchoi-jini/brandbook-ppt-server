# -*- coding: utf-8 -*-
"""유미니 원본의 외부 공개용 정규화와 완전성 검사를 담당한다.

이 모듈은 문구를 창작하지 않는다. 원본 사실을 보존하면서 JSON 문자열을
구조화하고, 완전히 같은 시간표만 합치며, 확인이 필요한 충돌을 표시한다.
"""
from __future__ import annotations

import json
import re
from copy import deepcopy


def parse_json_value(value, fallback=None):
    if not isinstance(value, str):
        return value
    text = value.strip()
    if not text:
        return fallback if fallback is not None else value
    if text[:1] not in ("[", "{"):
        return value
    try:
        return json.loads(text)
    except (TypeError, ValueError):
        return value


def _slots(course):
    weekly = course.get("weeklySchedule") or {}
    slots = weekly.get("slots") if isinstance(weekly, dict) else weekly
    if not isinstance(slots, list):
        slots = course.get("slots") if isinstance(course.get("slots"), list) else []
    return slots


def course_key(course):
    slots = sorted(
        (str(s.get("day") or ""), str(s.get("start") or s.get("startTime") or ""),
         str(s.get("duration_minutes") or s.get("durationMinutes") or ""))
        for s in _slots(course) if isinstance(s, dict)
    )
    return (
        str(course.get("name") or course.get("className") or course.get("courseName") or "").strip(),
        str(course.get("staffName") or course.get("teacher") or course.get("instructor") or "").strip(),
        str(course.get("room") or course.get("classroom") or "").strip(),
        tuple(slots),
    )


def normalize_raw(raw):
    out = deepcopy(raw if isinstance(raw, dict) else {})
    content = out.get("content")
    if not isinstance(content, dict):
        content = {}
        out["content"] = content
    for key in (
        "achievements", "faq", "faqs", "managementItems", "specialPrograms",
        "admissionSteps", "operatingRules", "curriculum", "divisions", "timetables",
    ):
        if key in content:
            content[key] = parse_json_value(content[key], [] if key not in ("operatingRules", "curriculum") else {})

    unique, duplicate_count, seen = [], 0, set()
    for course in out.get("courses") or []:
        if not isinstance(course, dict):
            continue
        key = course_key(course)
        if key in seen:
            duplicate_count += 1
            continue
        seen.add(key)
        unique.append(course)
    out["courses"] = unique
    out["_quality"] = audit(out, raw_count=len((raw or {}).get("courses") or []),
                            duplicate_count=duplicate_count)
    return out


def _items(value):
    value = parse_json_value(value, [])
    return value if isinstance(value, list) else []


def audit(raw, raw_count=None, duplicate_count=0):
    content = raw.get("content") or {}
    courses = raw.get("courses") or []
    issues = []

    # ★비용 충돌 — 금액이 0보다 큰데 방식이 "무료"인 경우를 직접 본다.
    #   (더포스둔산: testFee 30000 + testFeeMethod "무료")
    ops = raw.get("operations") or {}
    for fee_k, method_k, lab in (("consultingFee", "consultingFeeMethod", "상담"),
                                 ("testFee", "testFeeMethod", "테스트")):
        try:
            amount = float(str(ops.get(fee_k) or 0).replace(",", "") or 0)
        except Exception:
            amount = 0
        method = str(ops.get(method_k) or "").strip()
        if amount > 0 and "무료" in method:
            issues.append({"code": "FEE_CONFLICT",
                           "value": f"{lab} {int(amount):,}원 / 방식 '{method}'",
                           "message": f"{lab} 비용이 금액과 '무료'로 동시에 기록돼 있습니다. 확인이 필요합니다."})

    for course in courses:
        name = str(course.get("name") or "").strip()
        if re.fullmatch(r"(?:class|반)\s*[A-Z0-9]+|테스트|test", name, re.I):
            issues.append({"code": "POSSIBLE_PLACEHOLDER_COURSE", "value": name,
                           "message": "외부 공개용 강좌명인지 확인이 필요합니다."})
        # ★한글 뒤에 붙은 Tr(이우경Tr), 레벨코드(M3·T2), 언더바 강사명(중3 수학_도훈)
        if re.search(r"(?:[가-힣A-Za-z]Tr\b|\bTr\b|^[A-Z]\d+\b|\s[A-Z]\d+\s|_[-가-힣A-Za-z]+$)", name):
            issues.append({"code": "POSSIBLE_INTERNAL_COURSE_NAME", "value": name,
                           "message": "강사명·내부 코드가 포함된 강좌명인지 확인이 필요합니다."})

    rules = content.get("operatingRules") or {}
    rule_count = len([v for v in rules.values() if str(v or "").strip()]) if isinstance(rules, dict) else 0
    return {
        "counts": {
            "coursesRaw": raw_count if raw_count is not None else len(courses),
            "coursesUnique": len(courses),
            "courseDuplicates": duplicate_count,
            "classProfiles": len(raw.get("classProfiles") or []),
            "achievements": len(_items(content.get("achievements"))),
            "managementItems": len(_items(content.get("managementItems"))),
            "specialPrograms": len(_items(content.get("specialPrograms"))),
            "admissionSteps": len(_items(content.get("admissionSteps"))),
            "rules": rule_count,
            "faq": len(_items(content.get("faq") or content.get("faqs"))),
        },
        "issues": issues,
    }

