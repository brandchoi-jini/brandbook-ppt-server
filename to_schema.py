# 유미니 원본 JSON + (STEP2~5 생성물) -> 서버 표준 스키마 변환
import json

def parse_maybe(v):
    if isinstance(v,str):
        try: return json.loads(v)
        except: return v
    return v

def convert(raw, brand=None):
    """raw = 유미니 export JSON, brand = STEP2~5에서 만든 브랜드 요소(dict)"""
    brand = brand or {}
    basic = raw.get('basic',{})
    ai = raw.get('aiProfile') or {}
    content = raw.get('content') or {}
    cps = raw.get('classProfiles') or []
    courses = raw.get('courses') or []

    subjects = [k for k,v in basic.get('subjects',{}).items() if v]
    grades = [k for k,v in basic.get('focusGrades',{}).items() if v]
    GRADE_KO = {'elemLow':'초등 저학년','elemHigh':'초등 고학년','middle':'중등','high':'고등','retake':'재수'}
    grade_labels = [GRADE_KO.get(g,g) for g in grades]

    # 강점 (제목=strengths 라벨, 설명=differentiationDetail)
    strengths=[]
    diff = ai.get('differentiationDetail') or []
    st_titles = ai.get('strengths') or []
    for i,det in enumerate(diff[:3]):
        title = st_titles[i] if i<len(st_titles) else ''
        strengths.append({'title':title,'desc':det})
    while len(strengths)<len(st_titles):
        i=len(strengths); strengths.append({'title':st_titles[i],'desc':''})

    # features: 교육특징 6칸용. 강점3 + 차별화(제목 자동 생성)로 최대 6개 확보
    features = list(strengths)
    # 차별화 상세를 추가 특징으로 (강점에 이미 쓴 것 제외)
    used_desc = {s['desc'] for s in strengths}
    extra_titles = ['맞춤 피드백','개별 첨삭','학습 습관 관리','정기 상담','성적 관리','진로 설계']
    ei=0
    for det in diff:
        if len(features)>=6: break
        if det in used_desc: continue
        features.append({'title':extra_titles[ei%len(extra_titles)],'desc':det}); ei+=1
    # copyBank.trust 로도 보충
    trust = (ai.get('copyBank') or {}).get('trust') or []
    for tv in trust:
        if len(features)>=6: break
        features.append({'title':extra_titles[ei%len(extra_titles)],'desc':tv[:60]}); ei+=1

    # FAQ
    faq = parse_maybe(content.get('faq')) or []
    faq = [f for f in faq if isinstance(f, dict) and f.get('q')]
    faq = [{'q':f.get('q',''),'a':f.get('a','')} for f in faq if isinstance(f,dict) and f.get('q')]

    # 커리큘럼: content.curriculum 있으면 우선, 없으면 classProfiles/courses에서 학년별 유도
    curric = parse_maybe(content.get('curriculum')) or []
    curric = [c for c in curric if isinstance(c, dict)]
    if not curric:
        # classProfiles subject별 levels -> 학년 라벨과 결합
        by_grade=[]
        subj_ko = {'math':'수학','korean':'국어','english':'영어','science':'과학','social':'사회'}
        subj_str = ' · '.join(subj_ko.get(s,s) for s in subjects) or '수학'
        for gl in grade_labels:
            # 해당 학년 반들
            desc_bits=[]
            for cp in cps:
                for lv in cp.get('levels',[]):
                    pts = lv.get('counselingPoints') or []
                    if pts: desc_bits.append(pts[0])
            curric.append({
                'grade': gl,
                'title': f'{gl} 과정',
                'subject': subj_str,
                'content': (desc_bits[0][:60] if desc_bits else ''),
                'goal': (ai.get('goalGroup') or [''])[0]
            })

    # 관리항목: content.managementItems + classProfiles(평가·과제·상담·출결) 보강
    # 빌더 PPT는 [출결, 과제, 성취도, 상담] 순서로 4칸을 채운다 → 그 순서에 맞춘다
    mgmt = parse_maybe(content.get('managementItems')) or []
    mgmt = [{'title':m.get('title',''),'desc':m.get('desc','')} for m in mgmt
            if isinstance(m,dict) and ((m.get('title') or '').strip() or (m.get('desc') or '').strip())]
    if len(mgmt) < 4 and cps:
        evals, hws, cpts, atts = [], [], [], []
        for cp in cps:
            for a in (cp.get('assessments') or []):
                if isinstance(a, dict):
                    t = ' · '.join(str(x) for x in [a.get('type'), a.get('frequency'), a.get('belowPolicy')] if x)
                    if t: evals.append(t)
            hp = cp.get('homeworkPolicy')
            if isinstance(hp, list): hws += [str(x) for x in hp if x]
            elif hp: hws.append(str(hp))
            for lv in (cp.get('levels') or []):
                cpts += [str(x) for x in (lv.get('counselingPoints') or []) if x]
        # 출결: 운영규칙의 출결/결석 규정에서
        _or = parse_maybe(content.get('operatingRules')) or {}
        if isinstance(_or, dict):
            for k in ('attendance','absence'):
                if _or.get(k): atts.append(str(_or[k]))
        lsp = ai.get('lowScorePolicies') or []
        def _u(a):
            seen=set(); out=[]
            for x in a:
                x=str(x).strip()
                if x and x not in seen: seen.add(x); out.append(x)
            return out
        # 빌더 순서: 출결 → 과제 → 성취도 → 상담
        slots = [
            ('출결',   _u(atts)),
            ('과제',   _u(hws)),
            ('성취도', _u(evals) + ([' → '.join(str(x) for x in lsp)] if lsp else [])),
            ('상담',   _u(cpts)),
        ]
        have = {(m.get('title') or '').strip() for m in mgmt}
        filled = []
        for title, items in slots:
            if title in have: continue
            if items:
                filled.append({'title': title, 'desc': ' / '.join(items[:3])})
        mgmt = (mgmt + filled)[:6]

    # 입학단계
    adm = parse_maybe(content.get('admissionSteps')) or []
    adm = [{'step':a.get('step',''),'desc':a.get('desc','')} for a in adm if isinstance(a,dict)]

    # 특별프로그램 → {title, desc} 정규화 (빌더가 desc 키를 요구)
    special = parse_maybe(content.get('specialPrograms')) or []
    _sp = []
    for s0 in special:
        if not isinstance(s0, dict): continue
        title = (s0.get('title') or s0.get('name') or '').strip()
        desc = s0.get('desc') or s0.get('description') or ''
        if not desc:
            it = s0.get('items')
            if isinstance(it, list): desc = ' / '.join(str(x) for x in it if x)
            elif it: desc = str(it)
        if title or desc:
            _sp.append({'title': title, 'desc': str(desc)})
    special = _sp

    # 운영 규칙 (카탈로그·리플렛 필수)
    orules = parse_maybe(content.get('operatingRules')) or {}
    if not isinstance(orules, dict): orules = {}

    # 시간표 + 학년별 반 구성 (courses에서)
    import re as _re
    courses = raw.get('courses') or (raw.get('timetable') or {}).get('courses') or []
    DAY = {'MONDAY':'월','TUESDAY':'화','WEDNESDAY':'수','THURSDAY':'목',
           'FRIDAY':'금','SATURDAY':'토','SUNDAY':'일'}
    DAY_ORDER = ['월','화','수','목','금','토','일']
    def _grade_of(n):
        n = (n or '').replace(' ', '')
        if _re.search(r'고[123]|고등|수능|예비고', n): return 'high'
        if _re.search(r'중[123]|중등|예비중', n): return 'mid'
        if _re.search(r'초[1-6]|초등|예비초', n): return 'elem'
        return None

    # divisions: 학년별 [{name, desc}] — PPT 반구성 슬라이드 형식
    divisions = {'elem': [], 'mid': [], 'high': []}
    seen_div = {'elem': set(), 'mid': set(), 'high': set()}
    # 반 설명은 classProfiles의 상담포인트 첫 줄에서 보충
    lv_desc = {}
    for cp in cps:
        for lv in (cp.get('levels') or []):
            nm = (lv.get('className') or lv.get('name') or '').strip()
            pts = lv.get('counselingPoints') or []
            if nm and pts: lv_desc[nm] = str(pts[0])[:60]

    # timetable: {headers:[요일], rows:[[시간대, 월, 화, ...]]} — PPT 표 형식
    slot_map = {}   # 시작시각 -> {요일: 반이름}
    for c in courses:
        nm = (c.get('name') or '').strip()
        if not nm: continue
        g = _grade_of(nm)
        if g and nm not in seen_div[g]:
            seen_div[g].add(nm)
            divisions[g].append({'name': nm, 'desc': lv_desc.get(nm, '')})
        ws = c.get('weeklySchedule') or {}
        for s in (ws.get('slots') or []):
            d = DAY.get(s.get('day'), '')
            st = (s.get('start') or '').strip()
            if not d or not st: continue
            slot_map.setdefault(st, {})
            if d not in slot_map[st]:
                slot_map[st][d] = nm

    # 시간표: [구분 | 반 이름 | 요일 | 시간 | 강사] 반 목록 표
    # (요일×시각 격자는 셀이 대부분 비고 반 이름이 잘려서 읽기 어렵다)
    DAY_ORDER = ['월','화','수','목','금','토','일']

    def _subj_of(n):
        n = (n or '')
        if '수학' in n or _re.search(r'\b[MT]\d|MQ', n): return '수학'
        if _re.search(r'과학|물리|화학|생명|지구|통합', n): return '과학'
        if '영어' in n: return '영어'
        if '국어' in n or '논술' in n or '독해' in n or '문학' in n: return '국어'
        return '기타'

    def _grade2(n):
        """반 이름으로 학년 판정. 못 찾으면 과목 기본값."""
        g = _grade_of(n)
        if g: return {'elem':'초등','mid':'중등','high':'고등'}[g]
        nn = (n or '').replace(' ', '')
        # 통합과학·수능·모의고사 등은 고등
        if _re.search(r'통합과학|수능|모의|정시|논술|학평', nn): return '고등'
        return '기타'

    def _mkrow(c):
        nm = (c.get('name') or '').strip()
        slots = (c.get('weeklySchedule') or {}).get('slots') or []
        if not nm or not slots: return None
        days, times = [], []
        for s0 in slots:
            d = DAY.get(s0.get('day'), '')
            st = (s0.get('start') or '').strip()
            if d and d not in days: days.append(d)
            if st and st not in times: times.append(st)
        days.sort(key=lambda x: DAY_ORDER.index(x) if x in DAY_ORDER else 9)
        dur = ''
        if slots[0].get('duration_minutes'):
            dur = f"{slots[0]['duration_minutes']}분"
        return [
            _grade2(nm), nm, '·'.join(days),
            ' / '.join(times[:2]) + (f" ({dur})" if dur else ''),
            (c.get('staffName') or '').strip(),
        ]

    rows_tt = []
    tt_groups = {}          # "수학 · 초등" -> [행...]
    for c in courses:
        r = _mkrow(c)
        if not r: continue
        rows_tt.append(r)
        nm = (c.get('name') or '')
        key = f"{_subj_of(nm)} · {r[0]}"
        tt_groups.setdefault(key, []).append(r)

    _ord = {'초등':0, '중등':1, '고등':2, '기타':3}
    rows_tt.sort(key=lambda r: (_ord.get(r[0], 4), r[1]))
    for k in tt_groups:
        tt_groups[k].sort(key=lambda r: r[1])

    HDR_TT = ['구분', '반 이름', '요일', '시간', '강사']
    # 담당자가 결과물 화면에서 끈 시간표 그룹(과목·학년)은 제외한다
    _ttoff = raw.get('ttOff') or {}
    if isinstance(_ttoff, dict) and _ttoff:
        for k in list(tt_groups.keys()):
            if _ttoff.get(k):
                tt_groups.pop(k, None)
        keep = {tuple(r) for v in tt_groups.values() for r in v}
        rows_tt = [r for r in rows_tt if tuple(r) in keep]

    # 대상 학년이 지정되면 그 학년만 (담당자가 '초등'을 고르면 초등 시간표만)
    _g = (raw.get('gradeFilter') or '').strip()
    _GMAP = {'elem': '초등', 'mid': '중등', 'high': '고등'}
    if _g in _GMAP:
        rows_tt = [r for r in rows_tt if r[0] == _GMAP[_g]]
        tt_groups = {k: v for k, v in tt_groups.items() if k.endswith(_GMAP[_g])}

    # ★ 템플릿 표는 5행(헤더+4행)까지만 감당한다. 그 이상 넣으면 슬라이드를 뚫는다.
    #   학년별로 골고루 뽑아 최대 12행. (표 확장은 _add_row가 처리하되 상한을 둔다)
    TT_MAX = 12
    if len(rows_tt) > TT_MAX:
        by_g = {}
        for r in rows_tt:
            by_g.setdefault(r[0], []).append(r)
        picked, i = [], 0
        while len(picked) < TT_MAX:
            added = False
            for g in ['초등', '중등', '고등', '기타']:
                lst = by_g.get(g) or []
                if i < len(lst) and len(picked) < TT_MAX:
                    picked.append(lst[i]); added = True
            if not added:
                break
            i += 1
        picked.sort(key=lambda r: (_ord.get(r[0], 4), r[1]))
        rows_tt = picked

    timetable = {'headers': HDR_TT, 'rows': rows_tt} if rows_tt else {}
    # 과목×학년별로 나눈 시간표 (담당자가 목차에서 고를 수 있게)
    timetable_groups = [
        {'label': k, 'headers': HDR_TT, 'rows': v[:TT_MAX]}
        for k, v in sorted(tt_groups.items(), key=lambda x: (x[0].split(' · ')[0], _ord.get(x[0].split(' · ')[-1], 4)))
        if v
    ]

    ic = raw.get('introChannel') or {}
    ops = raw.get('operations') or {}

    # 학원 로고 · 사진 (유미니 S3 공개 URL) — PPT에 삽입
    images = {
        'logo': (basic.get('logoUrl') or '').strip(),
        'photos': [u for u in (ic.get('imageUrls') or []) if isinstance(u, str) and u.strip()][:6],
    }

    schema = {
        'academy': {
            'name': basic.get('name',''),
            'slogan': brand.get('slogan') or content.get('catchphrase',''),
            'phone': basic.get('phoneNumber',''),
            'address': basic.get('businessAddress',''),
            'kakao': (ic.get('kakaoChannel') or {}).get('chatUrl','') if isinstance(ic.get('kakaoChannel'),dict) else '',
            'subjects': subjects, 'grades': grade_labels,
        },
        'copy': {
            'catch': brand.get('catch') or content.get('catchphrase',''),
            'identity': brand.get('identity') or ai.get('oneLiner',''),
            'positioning': ai.get('positioning',''),
        },
        'strengths': strengths,
        'features': features,
        'curriculum': curric,
        'management': mgmt,
        'admission': adm,
        'special': special,
        'specials': special,
        'faq': faq,
        'operatingRules': orules,
        'timetable': timetable,
        'timetableGroups': timetable_groups,
        'divisions': divisions,
        'images': images,
        'achievements': [a for a in (parse_maybe(content.get('achievements')) or []) if isinstance(a,dict) and (a.get('text') or a.get('tag'))],
    }

    # 빌더가 요구하는 title/desc 키 보장 (KeyError 방지)
    def _fix(lst, tk='title'):
        out=[]
        for x in (lst or []):
            if isinstance(x, dict):
                x = dict(x)
                x.setdefault(tk, x.get('name') or x.get('step') or '')
                x.setdefault('desc', '')
                out.append(x)
            elif isinstance(x, str) and x.strip():
                out.append({tk: x.strip(), 'desc': ''})
        return out
    # academy 문자열 필드 None 방어 (PPT edit()가 None을 못 받음)
    for _k in ('name','slogan','phone','address','kakao'):
        if schema['academy'].get(_k) is None:
            schema['academy'][_k] = ''
        else:
            schema['academy'][_k] = str(schema['academy'][_k])
    for _k in ('catch','identity','positioning'):
        schema['copy'][_k] = str(schema['copy'].get(_k) or '')

    schema['management'] = _fix(schema.get('management'))
    schema['special'] = _fix(schema.get('special'))
    schema['specials'] = schema['special']
    schema['admission'] = _fix(schema.get('admission'), 'step')
    for _g in ('elem','mid','high'):
        schema['divisions'][_g] = _fix(schema['divisions'].get(_g), 'name')

    # ★STEP6 리치 콘텐츠 병합: 카피/구조는 원장이 확정한 리치콘텐츠 우선,
    #  연락처·과목·학년 등 구조 데이터는 원본(raw) 우선.
    rich = raw.get('richContent')
    if isinstance(rich, dict):
        schema = _merge_rich(schema, rich)

    return schema


def _nonempty(v):
    if v is None: return False
    if isinstance(v, (list, dict, str)): return len(v) > 0
    return True

def _merge_rich(base, rich):
    """리치콘텐츠(STEP6 확정본)를 base 스키마 위에 얹는다.
    - 카피/콘텐츠(강점·반구성·관리·특강·FAQ·커리큘럼·실적): 리치 우선
    - 연락처/과목/학년: base(원본) 우선 (리치가 비었을 때만 보충)
    - divisions: 리치에 있으면 그대로 실어 서버 PPT가 학년별로 채우게 함
    """
    out = dict(base)

    # copy 병합: 리치 카피가 있으면 우선
    rc = rich.get('copy') or {}
    bc = out.get('copy') or {}
    out['copy'] = {
        'catch': rc.get('catch') or bc.get('catch',''),
        'identity': rc.get('identity') or bc.get('identity',''),
        'positioning': bc.get('positioning','') or rc.get('positioning',''),
        'specialsLead': rc.get('specialsLead','') or bc.get('specialsLead',''),
    }

    # academy: 리치엔 연락처가 비어있을 수 있으니 base(원본) 우선, 리치로 보충
    ra = rich.get('academy') or {}
    ba = out.get('academy') or {}
    out['academy'] = dict(ba)
    for k in ('name','slogan','region'):
        if not out['academy'].get(k) and ra.get(k):
            out['academy'][k] = ra[k]

    # 콘텐츠 필드: 리치가 비어있지 않으면 리치로 교체(원장 확정본 우선)
    for k in ('strengths','features','achievements','growth','management',
              'specials','faq','divisions','admission','operatingRules'):
        rv = rich.get(k)
        if _nonempty(rv):
            out[k] = rv
    # 'special'과 'specials' 둘 다 쓰는 빌더 대응
    if _nonempty(out.get('specials')):
        out['special'] = out['specials']

    # 커리큘럼: 서버 카탈로그/리플렛은 list[{grade,...}] 형태를 기대.
    #  리치의 curriculum이 {과목:{rows}} 형태면 base(list) 유지.
    rcur = rich.get('curriculum')
    if isinstance(rcur, list) and _nonempty(rcur):
        out['curriculum'] = rcur
    # rich curriculum이 dict(과목별표)면 카탈로그 표에 쓰도록 별도 키로도 전달
    if isinstance(rcur, dict) and _nonempty(rcur):
        out['curriculumTable'] = rcur

    # 시간표: 리치에 있으면 우선(원장이 STEP6에서 정리한 표)
    if _nonempty(rich.get('timetable')):
        out['timetable'] = rich['timetable']

    return out

if __name__=='__main__':
    raw=json.load(open('/mnt/user-data/uploads/academy_ic-sg-jalpoom_export.json',encoding='utf-8'))
    # STEP2~5 생성물 예시 (실제로는 앱에서 옴)
    brand={'slogan':'오늘 수업이 내일 실력으로','catch':'잘 풀리는 수학, 잘품수학','identity':'중등부터 대입까지, 아이의 수학 흐름을 끊지 않는 곳'}
    s=convert(raw,brand)
    print(json.dumps(s,ensure_ascii=False,indent=1)[:2000])
