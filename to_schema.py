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

    # 관리항목
    mgmt = parse_maybe(content.get('managementItems')) or []
    mgmt = [{'title':m.get('title',''),'desc':m.get('desc','')} for m in mgmt if isinstance(m,dict)]

    # 입학단계
    adm = parse_maybe(content.get('admissionSteps')) or []
    adm = [{'step':a.get('step',''),'desc':a.get('desc','')} for a in adm if isinstance(a,dict)]

    # 특별프로그램
    special = parse_maybe(content.get('specialPrograms')) or []
    special = [s for s in special if isinstance(s, dict)]

    ic = raw.get('introChannel') or {}
    ops = raw.get('operations') or {}

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
        'faq': faq,
        'achievements': [a for a in (parse_maybe(content.get('achievements')) or []) if isinstance(a,dict) and (a.get('text') or a.get('tag'))],
    }

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
              'specials','faq','divisions'):
        rv = rich.get(k)
        if _nonempty(rv):
            out[k] = rv

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
