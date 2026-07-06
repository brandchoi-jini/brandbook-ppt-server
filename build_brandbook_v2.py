#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_brandbook_v2.py — 유미니 상담용 PPT(신규 12장) 팔레트별 빌더.
텍스트 자리표시자만 교체(디자인·이미지·그라데이션 100% 보존).
규모별 스마트 채움: 데이터 없는 학년/섹션 슬라이드는 통째 삭제, 카드는 있는 만큼만.

표준 콘텐츠 스키마(content):
  academy{name,slogan,region,subjects,phone,kakao,address}
  copy{catch,identity}
  strengths[{title,desc}]           핵심 강점
  achievements[str]  growth[str]    주요 실적 / 성장 사례
  divisions{elem[],mid[],high[]}    각 [{name,desc}]
  management[{title,desc}]          학습관리 요소(출결·과제·성취도·상담)
  timetable{headers[],rows[[]]}
  specials[{title,desc}]            특별 프로그램
  faq[{q,a}]
  design{palette: green|blue|orange}
"""
import os, copy, json, argparse
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml.ns import qn

TEMPLATES = {
    'green':  '유미니_상담용_ppt_그린.pptx',
    'blue':   '유미니_상담용_ppt_블루.pptx',
    'orange': '유미니_상담용_ppt_주황.pptx',
}
# 팔레트별 카드 수용량 / 슬라이드 인덱스
CAP = {
    'green':  {'strength':3, 'div':3, 'mng':4, 'faq':3, 'special':3},
    'blue':   {'strength':4, 'div':4, 'mng':4, 'faq':4, 'special':4},
    'orange': {'strength':3, 'div':3, 'mng':4, 'faq':5, 'special':3},
}
# 슬라이드 역할 인덱스(세 팔레트 공통 순서): 0표지 1정체성 2강점 3실적 4초 5중 6고 7관리 8시간표 9특강 10FAQ 11등록
IDX = dict(cover=0, ident=1, strength=2, ach=3, elem=4, mid=5, high=6, mng=7, tt=8, special=9, faq=10, info=11)

def walk(shapes):
    for x in shapes:
        yield x
        if x.shape_type == 6:
            yield from walk(x.shapes)

def _set(tf, txt):
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = txt
        for r in p0.runs[1:]: r._r.getparent().remove(r._r)
    else:
        r = p0.add_run(); r.text = txt
    for p in tf.paragraphs[1:]: p._p.getparent().remove(p._p)

def _set_multi(tf, lines):
    from pptx.text.text import _Paragraph
    lines = [l for l in lines if l is not None] or ['']
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = lines[0]
        for r in p0.runs[1:]: r._r.getparent().remove(r._r)
    else:
        r = p0.add_run(); r.text = lines[0]
    for p in tf.paragraphs[1:]: p._p.getparent().remove(p._p)
    for ln in lines[1:]:
        newp = copy.deepcopy(p0._p); p0._p.getparent().append(newp)
        para = _Paragraph(newp, p0._parent)
        if para.runs:
            para.runs[0].text = ln
            for r in para.runs[1:]: r._r.getparent().remove(r._r)

def norm(t):
    return t.strip().replace('\n', ' / ').replace('\xa0', ' ')

def edit(slide, mapping=None, multimap=None):
    mapping = mapping or {}; multimap = multimap or {}
    # 키도 정규화(\xa0 제거) 해서 비교
    mp = {norm(k): v for k, v in mapping.items()}
    mm = {norm(k): v for k, v in multimap.items()}
    seen = {}
    for x in walk(slide.shapes):
        if not (x.has_text_frame and x.text_frame.text.strip()): continue
        t = norm(x.text_frame.text)
        if t in mm:
            i = seen.get(t, 0); blocks = mm[t]
            blk = blocks[i] if (blocks and isinstance(blocks[0], list)) else blocks
            if blocks and isinstance(blocks[0], list) and i >= len(blocks): blk = ['']
            _set_multi(x.text_frame, blk); seen[t] = i + 1
        elif t in mp:
            i = seen.get(t, 0); v = mp[t]
            if isinstance(v, list): _set(x.text_frame, v[i] if i < len(v) else '')
            else: _set(x.text_frame, v)
            seen[t] = i + 1

def remove_slide(prs, slide):
    lst = prs.slides._sldIdLst
    slides = list(prs.slides)
    xmls = list(lst)
    for idx, sl in enumerate(slides):
        if sl is slide:
            lst.remove(xmls[idx]); return True
    return False

def fill_table(slide, headers, rows):
    for sh in slide.shapes:
        if not sh.has_table: continue
        tbl = sh.table; ncol = len(tbl.columns)
        for c in range(min(ncol, len(headers))):
            _set(tbl.cell(0, c).text_frame, str(headers[c]))
        for ri, row in enumerate(rows, start=1):
            if ri >= len(tbl.rows): _add_row(tbl)
            for c in range(ncol):
                _set(tbl.cell(ri, c).text_frame, str(row[c]) if c < len(row) else '')
        # 데이터보다 많은 잔여 행 비우기
        for ri in range(len(rows)+1, len(tbl.rows)):
            for c in range(ncol): _set(tbl.cell(ri, c).text_frame, '')
        return True
    return False

def _add_row(tbl):
    last = tbl._tbl.findall(qn('a:tr'))[-1]
    new = copy.deepcopy(last)
    for tc in new.findall(qn('a:tc')):
        for t in tc.iter(qn('a:t')): t.text = ''
    tbl._tbl.append(new)

def _wrap(text, n=18):
    text = (text or '').strip()
    if not text: return ['']
    if '\n' in text: return [l for l in text.split('\n') if l.strip()] or ['']
    out, cur = [], ''
    for ch in text:
        cur += ch
        if len(cur) >= n: out.append(cur); cur = ''
    if cur: out.append(cur)
    return out or ['']

def _titles(items, n):   return [(items[i]['title'] if i < len(items) else '') for i in range(n)]
def _names(items, n):    return [(items[i]['name']  if i < len(items) else '') for i in range(n)]
def _descs(items, n):    return [(_wrap(items[i]['desc']) if i < len(items) else ['']) for i in range(n)]

def build(content, out_path, template_dir='.'):
    design = content.get('design') or {}
    pal = design.get('palette', 'green')
    if pal not in TEMPLATES: pal = 'green'
    prs = Presentation(os.path.join(template_dir, TEMPLATES[pal]))
    S = list(prs.slides)
    cap = CAP[pal]
    ac = content.get('academy', {}); cp = content.get('copy', {})
    name = ac.get('name', '학원명')
    st = content.get('strengths', [])

    # [0] 표지
    edit(S[IDX['cover']], {
        '학원명': name, '제목을': name, '제목을\n입력해 주세요': name,
        '입력해 주세요': cp.get('catch',''), '제목을 / 입력해 주세요': name,
        '슬로건을 입력해주세요': ac.get('slogan',''),
        '슬로건을 입력해 주세요': ac.get('slogan',''),
        '슬로건을 입력해주세요 슬로건을 입력해주세요 슬로건을 입력해주세요': ac.get('slogan',''),
    })

    # [1] 정체성 — 팔레트별 자리표시자 상이
    ident_common = {
        '학원의 정체성을 보여주는  / 문구를 한줄로 작성해주세요.': cp.get('identity',''),
        '학원의 정체성을 보여주는 / 문구를 한줄로 작성해주세요.': cp.get('identity',''),
        '학원의 교육 방향과 / 정체성을 / 한 줄로 정의해 주세요': cp.get('catch',''),
    }
    if pal == 'green':
        edit(S[1], {**ident_common,
            '교육 역량 1': st[0]['title'] if len(st)>0 else '',
            '교육 역량 2': st[1]['title'] if len(st)>1 else '',
            '교육 역량 3': st[2]['title'] if len(st)>2 else '',
            '[학원만의 교육 방향을 / 적어주세요': [ (st[i]['desc'] if i<len(st) else '') for i in range(3) ],
        })
    elif pal == 'orange':
        edit(S[1], {**ident_common,
            '교육 방향 1': st[0]['title'] if len(st)>0 else '',
            '교육 방향 2': st[1]['title'] if len(st)>1 else '',
            '교육 방향 3': st[2]['title'] if len(st)>2 else '',
            '[학원만의 / 교육 방향을 / 적어주세요': [ (st[i]['desc'] if i<len(st) else '') for i in range(3) ],
            '학원만의 / 교육 방향을 / 적어주세요': [ (st[i]['desc'] if i<len(st) else '') for i in range(3) ],
        })
    else:  # blue: 정체성 슬라이드는 "한 줄 정의 / 교육철학 2" 2열
        edit(S[1], {
            '학원에 대한 한 줄 정의': cp.get('catch','') or '한 줄 정의',
            '교육 철학 2': (st[0]['title'] if st else '교육 철학'),
        }, multimap={
            '세부 내용을 입력해 주세요 / 세부 내용을 입력해 주세요 / 세부 내용을 입력해 주세요 / 세부 내용을 입력해 주세요':
                [ _wrap(cp.get('identity','')), _wrap(st[0]['desc'] if st else '') ],
        })

    # [2] 핵심 강점
    if pal == 'green':
        edit(S[2], {'강점을 적어주세요': _titles(st,3)},
             multimap={'학원의 핵심 강정을 설명해 주세요 / 학원의 핵심 강정을 설명해 주세요': _descs(st,3)})
    elif pal == 'orange':
        edit(S[2], {'핵심 강점 1': st[0]['title'] if len(st)>0 else '',
                    '핵심 강점 2': st[1]['title'] if len(st)>1 else '',
                    '핵심 강점 3': st[2]['title'] if len(st)>2 else ''},
             multimap={'학원의 핵심 강정을 설명해 주세요 / 학원의 핵심 강정을 설명해 주세요 / 학원의 핵심 강정을 설명해 주세요': _descs(st,3)})
    else:  # blue 4개
        edit(S[2], {'핵심 강점 1': st[0]['title'] if len(st)>0 else '',
                    '핵심 강점 2': st[1]['title'] if len(st)>1 else '',
                    '핵심 강점 3': st[2]['title'] if len(st)>2 else '',
                    '핵심 강점 4': st[3]['title'] if len(st)>3 else ''},
             multimap={'강점에 대한 세부 설명을 적어주세요 / 강점에 대한 세부 설명을 적어주세요': _descs(st,4)})

    # [3] 주요 실적 + 성장 사례
    ach = content.get('achievements', []); growth = content.get('growth', [])
    if ach or growth:
        if pal == 'green':
            details = [([ach[i]] if i<len(ach) else ['']) for i in range(3)] + \
                      [([growth[i]] if i<len(growth) else ['']) for i in range(3)]
            edit(S[3], multimap={'세부 내용을 입력해 주세요': details})
        else:  # blue/orange: "입시 성과를 구체적으로 적어주세요" 6줄
            lines = (ach + growth)[:6]
            edit(S[3], multimap={
                '입시 성과를 구체적으로 적어주세요 / 입시 성과를 구체적으로 적어주세요 / 입시 성과를 구체적으로 적어주세요 / 입시 성과를 구체적으로 적어주세요 / 입시 성과를 구체적으로 적어주세요 / 입시 성과를 구체적으로 적어주세요':
                    lines if lines else ['']})
    else:
        remove_slide(prs, S[3])

    # [4][5][6] 학년별 반구성 — 스마트: 데이터 있는 학년만
    div = content.get('divisions', {})
    for key, si in [('elem', IDX['elem']), ('mid', IDX['mid']), ('high', IDX['high'])]:
        items = div.get(key, [])
        if not items:
            remove_slide(prs, S[si]); continue
        ncap = cap['div']
        if pal == 'green':
            edit(S[si], {'반 이름을 적어주세요': _names(items, 3)},
                 multimap={'학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요': _descs(items, 3)})
        elif pal == 'blue':
            edit(S[si], {'반 이름을 적어주세요': _names(items, 4)},
                 multimap={'세부적인 내용을 적어주세요': [(items[i].get('desc','') if i<len(items) else '') for i in range(4)]})
        else:  # orange
            edit(S[si], {'반 이름을 적어주세요': _names(items, 3),
                         '반 레벨을 설명해 주세요': [(items[i].get('level', items[i].get('name','')) if i<len(items) else '') for i in range(3)]},
                 multimap={'학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요 / 학습체계를 구체적으로 적어주세요': _descs(items,3)})

    # [7] 학습 관리
    mng = content.get('management', [])
    if mng:
        labels = ['출결','과제','성취도','상담']
        m = {}
        for i, lb in enumerate(labels):
            m[lb] = mng[i]['title'] if i < len(mng) else lb
        if pal == 'green':
            m['요소 1'] = ''; m['요소 2'] = ''; m['요소 3'] = ''; m['요소 4'] = ''
            edit(S[7], m, multimap={'구체적인 설명을 적어주세요 / 구체적인 설명을 적어주세요': _descs(mng,4)})
        elif pal == 'blue':
            edit(S[7], m, multimap={'구체적인 설명을 적어주세요 / 구체적인 설명을 적어주세요 / 구체적인 설명을 적어주세요': _descs(mng,4)})
        else:
            m2 = {'한 줄로 정의해주세요': [ (mng[i]['desc'].split('.')[0] if i<len(mng) else '') for i in range(4) ]}
            edit(S[7], {**m, **{}}, multimap={'구체적인 설명을 적어주세요 / 구체적인 설명을 적어주세요 / 구체적인 설명을 적어주세요 / 구체적인 설명을 적어주세요': _descs(mng,4)})
    else:
        remove_slide(prs, S[IDX['mng']])

    # [8] 시간표
    tt = content.get('timetable', {})
    if tt and tt.get('rows'):
        fill_table(S[IDX['tt']], tt.get('headers', ['','월','화','수','목','금','토','일']), tt['rows'])
    else:
        remove_slide(prs, S[IDX['tt']])

    # [9] 특별 프로그램
    sp = content.get('specials', [])
    if sp:
        lead = {'프로그램의 한 줄 정의를 설명해주세요': cp.get('specialsLead',''),
                '학원의 특별 프로그램을 / 적어주세요': cp.get('specialsLead','') or '학원의 특별 프로그램'}
        if pal == 'green':
            edit(S[9], {**lead, '프로그램명': _titles(sp,3)},
                 multimap={'세부적인 설명을 적어주세요 / 세부적인 설명을 적어주세요 / 세부적인 설명을 적어주세요 / 세부적인 설명을 적어주세요': _descs(sp,3)})
        elif pal == 'blue':
            edit(S[9], {'프로그램명': _titles(sp,4)},
                 multimap={'세부 설명을 적어주세요 / 세부 설명을 적어주세요 / 세부 설명을 적어주세요': _descs(sp,4)})
        else:
            edit(S[9], {**lead, '핵심 프로그램명': _titles(sp,3)},
                 multimap={'세부적인 설명을 적어주세요 / 세부적인 설명을 적어주세요 / 세부적인 설명을 적어주세요 / 세부적인 설명을 적어주세요': _descs(sp,3)})
    else:
        remove_slide(prs, S[IDX['special']])

    # [10] FAQ
    faq = content.get('faq', [])
    if faq:
        n = cap['faq']
        if pal == 'green':
            edit(S[10], {'질문을 입력해 주세요': [(faq[i]['q'] if i<len(faq) else '') for i in range(3)],
                         'A. 답변을 적어주세요': [('A. '+faq[i]['a'] if i<len(faq) else '') for i in range(3)]},
                 multimap={'세부 설명을 추가해 주세요 / 세부 설명을 추가해 주세요 / 세부 설명을 추가해 주세요 / 세부 설명을 추가해 주세요':
                           [(_wrap(faq[i].get('detail','')) if i<len(faq) else ['']) for i in range(3)]})
        elif pal == 'blue':
            edit(S[10], {'질문을 적어주세요': [(faq[i]['q'] if i<len(faq) else '') for i in range(4)]},
                 multimap={'답변을 적어주세요 / 답변을 적어주세요': [(_wrap(faq[i]['a']) if i<len(faq) else ['']) for i in range(4)]})
        else:  # orange 5개
            edit(S[10], {'질문의 내용을 적어주세요': [(faq[i]['q'] if i<len(faq) else '') for i in range(5)],
                         '질문에 대한 상세한 답변을 적어주세요': [(faq[i]['a'] if i<len(faq) else '') for i in range(5)]})
    else:
        remove_slide(prs, S[IDX['faq']])

    # [11] 등록 안내
    edit(S[IDX['info']], {
        '상담 전화번호를 적어주세요': ac.get('phone',''),
        '상담 전화번호를 / 적어주세요': ac.get('phone',''),
        '카카오톡 채널을 적어주세요': ac.get('kakao',''),
        '카카오톡 채널을 / 적어주세요': ac.get('kakao',''),
        '주소를 적어주세요': ac.get('address',''),
        '주소를 / 적어주세요': ac.get('address',''),
    })

    prs.save(out_path)
    return out_path

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('content_json'); ap.add_argument('out_pptx')
    ap.add_argument('--templates', default='.')
    a = ap.parse_args()
    content = json.load(open(a.content_json, encoding='utf-8'))
    print('저장:', build(content, a.out_pptx, a.templates))


# ============================================================
# 카탈로그(6장) / 리플렛(2장) 빌더
# ============================================================
CATALOG_TEMPLATES = {
    'gray_gray':  '유미니_카탈로그_디자인_-_그레이__그레이.pptx',
    'gray_green': '유미니_카탈로그_디자인_-_그레이__그린.pptx',
    'blue':       '유미니_카탈로그_디자인_-_블루.pptx',
}
LEAFLET_TEMPLATES = {
    'gray_green': '그레이_그린.pptx',
    'gray_blue':  '그레이_블루.pptx',
    'blue_blue':  '블루_블루.pptx',
}
# 앱 팔레트 → 카탈로그/리플렛 색
PAL_TO_CATALOG = {'forest':'gray_green','sage':'gray_green','navy':'blue','violet':'blue','crimson':'gray_gray','mono':'gray_gray'}
PAL_TO_LEAFLET = {'forest':'gray_green','sage':'gray_green','navy':'blue_blue','violet':'blue_blue','crimson':'gray_blue','mono':'gray_blue'}

# 교육 특징 6개 자리표시자(제목→설명) — 카탈로그/리플렛 공통
FEATURE_SLOTS = [
    ('학년 담임 통합', '전 과목 진도와 생활을 함께 관리합니다.'),
    ('과목별 전담 강사', '5년 이상 경력의 전문 강사진입니다.'),
    ('주간 통합 리포트', '전 과목 학습 현황을 한 페이지로 공유합니다.'),
    ('과목 간 연계 학습', '독해 · 지문을 교차로 활용합니다.'),
    ('시험 4주 대비', '전 과목 동시 부스터를 가동합니다.'),
    ('초등 · 중등 · 고등 진학 연결', '기초부터 시작해 고등까지 이어갑니다.'),
]

def _fill_curriculum_tables(slide, curriculum):
    """커리큘럼 표(과목별)를 데이터로 채움. 각 표 위에 있는 과목 라벨로 매칭.
       curriculum={subject:{rows:[[영역,초,중,고]...]}}. 데이터 없는 과목표는 비움."""
    from pptx.util import Emu
    # 과목 라벨 위치 수집
    labels = []
    for sh in walk(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip() in ['국어','영어','수학','과학','사회']:
            labels.append((sh.text_frame.text.strip(), Emu(sh.left).inches, Emu(sh.top).inches))
    tables = [sh for sh in walk(slide.shapes) if sh.has_table]
    for sh in tables:
        tl, tt = Emu(sh.left).inches, Emu(sh.top).inches
        # 이 표 바로 위(같은 열, top이 표보다 작고 가장 가까운)에 있는 과목 라벨 찾기
        subj = None; best = 999
        for (lb, ll, ltop) in labels:
            if abs(ll - tl) < 2.0 and ltop < tt and (tt - ltop) < best:
                best = tt - ltop; subj = lb
        tbl = sh.table
        data = (curriculum or {}).get(subj) if subj else None
        if not data or not data.get('rows'):
            for r in range(1, len(tbl.rows)):
                for c in range(len(tbl.columns)):
                    _set(tbl.cell(r, c).text_frame, '')
            continue
        rows = data['rows']
        for ri in range(1, len(tbl.rows)):
            row = rows[ri-1] if ri-1 < len(rows) else []
            for c in range(len(tbl.columns)):
                _set(tbl.cell(ri, c).text_frame, str(row[c]) if c < len(row) else '')

def build_catalog(content, out_path, template_dir='.'):
    design = content.get('design') or {}
    pal = design.get('palette', 'sage')
    color = PAL_TO_CATALOG.get(pal, 'gray_green')
    prs = Presentation(os.path.join(template_dir, CATALOG_TEMPLATES[color]))
    S = list(prs.slides)
    ac = content.get('academy', {}); cp = content.get('copy', {})
    feats = content.get('features') or content.get('strengths', [])
    faq = content.get('faq', [])
    curriculum = content.get('curriculum', {})

    # S0 표지
    edit(S[0], {
        '성적이 달라지는 / 전략적 공부 방법': cp.get('catch','') or ac.get('slogan',''),
        '상담문의': '상담문의', '123-123-1234': ac.get('phone',''),
    })
    # 강점 아이콘 3개 (내신·수능대비 / 소수정원 / 성적관리 자리)
    icon_slots = ['내신 · 수능 대비 / 전문 학원','소수 정원으로 / 집중 수업','성적 관리 / 시스템']
    m0 = {}
    for i, slot in enumerate(icon_slots):
        if i < len(feats): m0[slot] = feats[i].get('title', slot)
    edit(S[0], m0)

    # S1 PHILOSOPHY + 교육특징 6
    edit(S[1], {'왜 성장학원인가 ?': f"왜 {ac.get('name','')}인가 ?",
                '단순한 문제 풀이가 아닌 체계적인 학습 시스템으로 지도 합니다.': cp.get('identity','')})
    fm = {}
    for i, (t, d) in enumerate(FEATURE_SLOTS):
        if i < len(feats):
            fm[t] = feats[i].get('title', t)
            fm[d] = feats[i].get('desc', d)
    edit(S[1], fm)

    # S2 커리큘럼 표 4개
    _fill_curriculum_tables(S[2], curriculum)

    # S5 FAQ 6
    faq_slots = [f'Q.   {q}' for q in [
        '한 과목만 등록할 수 있나요?','중간에 과목을 추가할 수 있나요?','진단 테스트는 어떻게 진행되나요?',
        '학년 담임이 모든 과목을 가르치나요?','학부모 리포트는 얼마나 자주 받나요?','시험 기간에는 어떻게 운영되나요?']]
    # FAQ는 자리표시자 텍스트가 특정 질문이라 순서로 교체
    fmap = {}
    ans_slots = ['가능합니다. 다만 전 과목 패키지를 추천드립니다.',
                 '네, 언제든 담임선생님과 상담 후 추가 등록이 가능합니다.',
                 '테스트는 과목별 30분 내외로 진행되며, 학생 개인별 수준과 약점 등을 파악합니다',
                 '담임선생님은 통합 관리만 하시며, 각 과목은 전담 강사가 / 수업합니다.',
                 '주간 통합 리포트가 담임선생님 단일 채널로 발송되므로 매주 받아보실 수 있습니다.',
                 '시험 4주 전부터 전 과목 동시 진행 부스터가 가동됩니다.']
    for i in range(min(6, len(faq))):
        fmap[faq_slots[i]] = f"Q.   {faq[i]['q']}"
        fmap[ans_slots[i]] = faq[i]['a']
    edit(S[5], fmap)

    # S4 입학/상담 안내
    edit(S[4], {'주소 : 대구광역시 달서구 조암로 14 (6F)': '주소 : ' + ac.get('address',''),
                '카카오톡 ID / @sungjang-academy': ac.get('kakao','')})

    prs.save(out_path)
    return out_path

def build_leaflet(content, out_path, template_dir='.'):
    design = content.get('design') or {}
    pal = design.get('palette', 'sage')
    color = PAL_TO_LEAFLET.get(pal, 'gray_green')
    prs = Presentation(os.path.join(template_dir, LEAFLET_TEMPLATES[color]))
    S = list(prs.slides)
    ac = content.get('academy', {}); cp = content.get('copy', {})
    feats = content.get('features') or content.get('strengths', [])
    faq = content.get('faq', [])
    curriculum = content.get('curriculum', {})

    # S0 외부면: 표지 카피 + 입학 + FAQ
    icon_slots = ['내신 · 수능 대비 / 전문 학원','소수 정원으로 / 집중 수업','성적 관리 / 시스템']
    m0 = {'성적이 달라지는 / 전략적 공부 방법': cp.get('catch','') or ac.get('slogan',''),
          '123-123-1234': ac.get('phone',''),
          '주소 : 대구광역시 달서구 조암로 14 (6F)': '주소 : ' + ac.get('address',''),
          '카카오톡 ID  :  @sungjang-academy': ac.get('kakao','')}
    for i, slot in enumerate(icon_slots):
        if i < len(feats): m0[slot] = feats[i].get('title', slot)
    # FAQ 6
    faq_slots = ['한 과목만 등록할 수 있나요?','중간에 과목을 추가할 수 있나요?','진단 테스트는 어떻게 진행되나요?',
                 '학년 담임이 모든 과목을 가르치나요?','학부모 리포트는 얼마나 자주 받나요?','시험 기간에는 어떻게 운영되나요?']
    ans_slots = ['가능합니다. 다만 전 과목 패키지를 추천드립니다.',
                 '네, 언제든 담임선생님과 상담 후 추가 등록이 가능합니다.',
                 '테스트는 과목별 30분 내외로 진행되며, 학생 개인별 수준과 약점 등을 파악합니다.',
                 '담임선생님은 통합 관리만 하시며, 각 과목은 전담 강사가 / 수업합니다.',
                 '주간 통합 리포트가 담임선생님 단일 채널로 발송되므로 매주 받아보실 수 있습니다.',
                 '시험 4주 전부터 전 과목 동시 진행 부스터가 가동됩니다.']
    for i in range(min(6, len(faq))):
        m0[f'Q.   {faq_slots[i]}'] = f"Q.   {faq[i]['q']}"
        m0[ans_slots[i]] = faq[i]['a']
    edit(S[0], m0)

    # S1 내부면: PHILOSOPHY + 교육특징6 + 커리큘럼표4
    edit(S[1], {'왜 성장학원인가 ?': f"왜 {ac.get('name','')}인가 ?",
                '단순한 문제 풀이가 아닌 / 체계적인 학습 시스템으로 지도 합니다.': cp.get('identity','')})
    fm = {}
    for i, (t, d) in enumerate(FEATURE_SLOTS):
        if i < len(feats):
            fm[t] = feats[i].get('title', t); fm[d] = feats[i].get('desc', d)
    edit(S[1], fm)
    _fill_curriculum_tables(S[1], curriculum)

    prs.save(out_path)
    return out_path
