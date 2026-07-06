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


# ============================================================
# v3 좌표 기반 빌더 (카탈로그·리플렛) — 예시 완전 제거 + 학년 편차 스마트 채움
# ============================================================

CATALOG_TEMPLATES = {
    'gray':  '유미니_카탈로그_디자인_-_그레이__그레이.pptx',
    'green': '유미니_카탈로그_디자인_-_그레이__그린.pptx',
    'blue':  '유미니_카탈로그_디자인_-_블루.pptx',
}
LEAFLET_TEMPLATES = {
    'green': '그레이_그린.pptx',
    'blue':  '그레이_블루.pptx',
    'navy':  '블루_블루.pptx',
}

def _in(v): return round(Emu(v).inches, 1)

def _find_pos(slide, left, top, tol=0.35):
    best, bestd = None, 999
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        l, t = _in(sh.left), _in(sh.top)
        d = abs(l-left) + abs(t-top)
        if d <= tol*2 and d < bestd:
            best, bestd = sh, d
    return best

def _put(shape, txt):
    if shape is None or txt is None: return
    tf = shape.text_frame
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = txt
        for r in p.runs[1:]: r.text = ''
    else:
        p.add_run().text = txt
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)

def _put_multi(shape, lines):
    if shape is None: return
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    ref = p0.runs[0] if p0.runs else None
    if p0.runs:
        p0.runs[0].text = lines[0] if lines else ''
        for r in p0.runs[1:]: r.text = ''
    else:
        p0.add_run().text = lines[0] if lines else ''
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    for ln in lines[1:]:
        np = tf.add_paragraph(); r = np.add_run(); r.text = ln
        if ref is not None:
            try:
                r.font.size = ref.font.size; r.font.bold = ref.font.bold
                if ref.font.color and ref.font.color.type is not None:
                    r.font.color.rgb = ref.font.color.rgb
            except Exception: pass

def _clear_region(slide, xmin, xmax, tmin, tmax):
    for sh in list(slide.shapes):
        if not sh.has_text_frame: continue
        l, t = _in(sh.left), _in(sh.top)
        if xmin <= l <= xmax and tmin <= t <= tmax:
            _put(sh, '')

def _grade_map(curric):
    gm = {}
    for c in curric:
        g = c.get('grade','')
        key = '초등' if '초' in g else '중등' if '중' in g else '고등' if '고' in g else g
        gm[key] = c
    return gm


def build_catalog(content, out_path, template_dir='.'):
    design = content.get('design') or {}
    color = design.get('palette', 'green')
    if color not in CATALOG_TEMPLATES: color = 'green'
    prs = Presentation(os.path.join(template_dir, CATALOG_TEMPLATES[color]))
    S = list(prs.slides)
    ac = content.get('academy', {}); cp = content.get('copy', {})
    feats = content.get('features', []) or content.get('strengths', [])
    strengths = content.get('strengths', [])
    faq = content.get('faq', [])
    curric = content.get('curriculum', [])
    mgmt = content.get('management', [])
    special = content.get('special', [])
    ach = content.get('achievements', [])
    rules = content.get('operatingRules') or {}
    name = ac.get('name', '')

    # S0 표지
    _put(_find_pos(S[0], 1.0, 0.8), cp.get('catch','') or ac.get('slogan',''))
    for i,(l,t) in enumerate([(4.3,8.5),(6.8,8.5),(9.3,8.5)]):
        _put(_find_pos(S[0], l, t), strengths[i]['title'] if i < len(strengths) else '')
    _put(_find_pos(S[0], 6.0,10.1), ac.get('phone',''))

    # S1 PHILOSOPHY + 교육특징 6
    _put(_find_pos(S[1], 1.2,1.4), f"왜 {name}인가 ?")
    _put(_find_pos(S[1], 1.2,2.4), cp.get('identity','') or cp.get('positioning',''))
    for i,(l,t) in enumerate([(2.3,4.6),(2.3,6.3),(2.3,8.0),(9.0,4.6),(9.0,6.3),(9.0,8.0)]):
        ts=_find_pos(S[1], l, t); ds=_find_pos(S[1], l, t+0.4)
        if i < len(feats):
            _put(ts, feats[i].get('title','')); _put(ds, feats[i].get('desc',''))
        else:
            _put(ts, ''); _put(ds, '')

    # S2 CURRICULUM 3컬럼 (초/중/고) — 학년 편차 스마트
    col_x  = {'초등':1.0, '중등':5.8, '고등':10.8}
    col_x2 = {'초등':2.2, '중등':7.0, '고등':12.0}
    gm = _grade_map(curric)
    for gk in ['초등','중등','고등']:
        x1, x2 = col_x[gk], col_x2[gk]
        c = gm.get(gk)
        title_sh = _find_pos(S[2], x1, 3.05, tol=0.4)
        if c:
            _put(title_sh, c.get('title', f'{gk} 과정'))
            _clear_region(S[2], x2-0.6, x2+0.6, 3.6, 4.95)
            cont = _find_pos(S[2], x2, 3.7, tol=0.3)
            lines=[]
            if c.get('content'):
                first=c['content'].split('.')[0]
                if len(first)>60: first=c['content'][:60]
                cur=''
                for w in first.split(' '):
                    if len(cur)+len(w)+1<=20: cur=(cur+' '+w).strip()
                    else: lines.append(cur); cur=w
                if cur: lines.append(cur)
                lines=lines[:4]
            _put_multi(cont, lines or [''])
            _put(_find_pos(S[2], x2, 5.2, tol=0.3), c.get('subject',''))
            _put(_find_pos(S[2], x2, 5.65, tol=0.3), c.get('goal',''))
        else:
            _clear_region(S[2], x1-0.3, x2+0.5, 2.9, 5.9)

    # 특별프로그램 2블록 → special 있으면, 없으면 관리로 대체
    sp_titles=[_find_pos(S[2],5.0,7.1),_find_pos(S[2],10.7,7.1)]
    sp_bodies=[_find_pos(S[2],4.7,7.7,tol=0.5),_find_pos(S[2],11.0,7.7,tol=0.5)]
    if special:
        for i in range(2):
            if i < len(special):
                sp=special[i]
                _put(sp_titles[i], sp.get('title',''))
                _put_multi(sp_bodies[i], sp.get('items',[]) if isinstance(sp.get('items'),list) else [sp.get('desc','')])
            else:
                _put(sp_titles[i],''); _put_multi(sp_bodies[i],[''])
    elif mgmt:
        _put(sp_titles[0], '학습 관리')
        _put_multi(sp_bodies[0], [m.get('title','') for m in mgmt[:4]])
        _put(sp_titles[1], ''); _put_multi(sp_bodies[1], [''])
    else:
        for t in sp_titles: _put(t,'')
        for b in sp_bodies: _put_multi(b,[''])

    # S3 관리시스템 6칸 (management + features 보충)
    _put(_find_pos(S[3], 1.7,3.1,tol=0.5), cp.get('positioning','') or '한 명도 놓치지 않는 관리 시스템으로 함께합니다.')
    mgmt_fill=list(mgmt)
    for f in feats:
        if len(mgmt_fill)>=6: break
        if not any(m.get('title')==f.get('title') for m in mgmt_fill):
            mgmt_fill.append({'title':f.get('title',''),'desc':f.get('desc','')})
    mslots=[(2.2,4.5),(5.3,4.6),(2.2,6.3),(5.3,6.3),(2.2,8.1),(5.3,8.1)]
    mdesc =[(1.6,5.1),(4.8,5.1),(1.6,6.9),(4.8,6.9),(1.6,8.7),(4.8,8.7)]
    for i,((tl,tt),(dl,dt)) in enumerate(zip(mslots,mdesc)):
        ts=_find_pos(S[3],tl,tt,tol=0.4); ds=_find_pos(S[3],dl,dt,tol=0.5)
        if i < len(mgmt_fill):
            _put(ts,mgmt_fill[i].get('title','')); _put(ds,mgmt_fill[i].get('desc',''))
        else:
            _put(ts,''); _put(ds,'')
    # 실적
    for i,(l,t) in enumerate([(8.7,3.7),(11.3,3.7),(8.7,4.3),(11.3,4.3)]):
        sh=_find_pos(S[3],l,t,tol=0.4)
        if sh and sh.text_frame.text.strip() not in ('초등','중등','고등'):
            _put(sh, ach[i].get('text','') if i<len(ach) else '')
    if not ach:
        for l,t in [(8.1,3.7),(10.7,3.7),(8.1,4.3)]:
            sh=_find_pos(S[3],l,t,tol=0.25)
            if sh and sh.text_frame.text.strip() in ('초등','중등','고등'): _put(sh,'')
    # 규정
    for (l,t),v in {(9.1,6.0):rules.get('homework'),(9.1,6.4):rules.get('absence'),
                    (9.1,6.7):rules.get('withdrawal'),(9.1,8.2):rules.get('refundBefore'),
                    (9.1,8.6):rules.get('refundAfter'),(9.1,8.9):rules.get('attendance')}.items():
        if v: _put(_find_pos(S[3], l, t, tol=0.3), v)

    # S4 입학안내
    _put(_find_pos(S[4], 8.4,2.6,tol=0.6), '주소 : ' + ac.get('address',''))
    kk=_find_pos(S[4], 9.4,3.6,tol=0.6)
    if kk and ac.get('kakao'): _put(kk, ac.get('kakao'))

    # S5 FAQ 6칸
    for i,(l,t) in enumerate([(1.2,3.3),(1.2,5.4),(1.2,7.4),(8.2,3.3),(8.2,5.4),(8.2,7.4)]):
        q=_find_pos(S[5], l, t, tol=0.3); a=_find_pos(S[5], l+0.5, t+0.5, tol=0.5)
        if i < len(faq):
            _put(q, f"Q.  {faq[i]['q']}"); _put(a, faq[i]['a'].replace('\n',' '))
        else:
            _put(q, ''); _put(a, '')

    prs.save(out_path)
    return out_path


def build_leaflet(content, out_path, template_dir='.'):
    design = content.get('design') or {}
    color = design.get('palette', 'green')
    if color not in LEAFLET_TEMPLATES: color = 'green'
    prs = Presentation(os.path.join(template_dir, LEAFLET_TEMPLATES[color]))
    S = list(prs.slides)
    ac = content.get('academy', {}); cp = content.get('copy', {})
    feats = content.get('features', []) or content.get('strengths', [])
    strengths = content.get('strengths', [])
    faq = content.get('faq', [])
    curric = content.get('curriculum', [])
    special = content.get('special', [])
    adm = content.get('admission', [])
    mgmt = content.get('management', [])
    ach = content.get('achievements', [])
    rules = content.get('operatingRules') or {}
    name = ac.get('name', '')

    s0, s1 = S[0], S[1]
    # S0 FAQ 6
    for i,qt in enumerate([2.0,3.3,4.6,5.9,7.2,8.5]):
        q=_find_pos(s0,0.7,qt,tol=0.3); a=_find_pos(s0,1.0,qt+0.3,tol=0.4)
        if i < len(faq):
            _put(q, f"Q.  {faq[i]['q']}"); _put(a, faq[i]['a'].replace('\n',' '))
        else:
            _put(q, ''); _put(a, '')
    # S0 입학안내 3
    for i,at in enumerate([2.1,3.2,4.4]):
        if i < len(adm):
            _put(_find_pos(s0,6.5,at,tol=0.3), adm[i].get('step',''))
            _put(_find_pos(s0,8.2,at,tol=0.4), adm[i].get('desc','')[:20])
    _put(_find_pos(s0,5.7,6.7,tol=0.5), '주소 : '+ac.get('address',''))
    kk=_find_pos(s0,6.0,7.4,tol=0.5)
    if kk and ac.get('kakao'): _put(kk, '카카오톡 : '+ac.get('kakao'))
    # S0 표지
    _put(_find_pos(s0,11.0,0.9,tol=0.4), (cp.get('catch') or ac.get('slogan','')).replace(' ','\n',1))
    for i,(l,t) in enumerate([(11.0,8.2),(12.5,8.2),(13.9,8.2)]):
        sh=_find_pos(s0,l,t,tol=0.4)
        if sh and i < len(strengths): _put(sh, strengths[i].get('title','').replace(' ','\n'))
    _put(_find_pos(s0,11.7,9.4,tol=0.4), ac.get('phone',''))

    # S1 PHILOSOPHY + 교육특징 6
    _put(_find_pos(s1,0.5,0.8,tol=0.3), f"왜 {name}인가 ?")
    _put(_find_pos(s1,0.5,1.7,tol=0.4), cp.get('identity','') or cp.get('positioning','')[:40])
    for i,ft in enumerate([3.4,4.6,5.7,6.8,8.0,9.1]):
        ts=_find_pos(s1,1.3,ft,tol=0.25); ds=_find_pos(s1,1.3,ft+0.3,tol=0.25)
        if i < len(feats):
            _put(ts,feats[i].get('title','')); _put(ds,feats[i].get('desc','')[:30])
        else:
            _put(ts,''); _put(ds,'')

    # S1 CURRICULUM 세로 3단
    col_tops={'초등':(1.9,2.4,3.4,3.7),'중등':(4.4,4.8,5.8,6.1),'고등':(6.8,7.2,8.2,8.5)}
    gm=_grade_map(curric)
    for gk,(tt,ct,st,gt) in col_tops.items():
        c=gm.get(gk)
        _clear_region(s1,6.5,7.5,ct-0.1,ct+0.7)
        title_sh=_find_pos(s1,6.0,tt,tol=0.3)
        if c:
            _put(title_sh, c.get('title',f'{gk} 과정'))
            if c.get('content'):
                _put(_find_pos(s1,6.8,ct,tol=0.2), c['content'].split('.')[0][:40])
            _put(_find_pos(s1,6.8,st,tol=0.2), c.get('subject',''))
            _put(_find_pos(s1,6.8,gt,tol=0.2), c.get('goal',''))
        else:
            _clear_region(s1,5.9,7.5,tt-0.1,gt+0.3)

    # S1 특별프로그램
    sp=_find_pos(s1,5.8,9.6,tol=0.4)
    if special:
        items=special[0].get('items',[]) if isinstance(special[0],dict) else []
        _put(sp, '\n'.join(items[:4]) if items else '')
    elif strengths:
        _put(sp, '\n'.join(s.get('title','') for s in strengths[:4]))
    else:
        _put(sp, '')

    # S1 세번째 패널: 학습관리
    mgmt_fill=list(mgmt)
    for f in feats:
        if len(mgmt_fill)>=6: break
        if not any(m.get('title')==f.get('title') for m in mgmt_fill):
            mgmt_fill.append({'title':f.get('title',''),'desc':f.get('desc','')})
    mg=[(11.6,2.9,11.2,3.3),(13.7,2.9,13.3,3.3),(11.6,4.1,11.2,4.5),
        (13.7,4.1,13.3,4.5),(11.6,5.3,11.2,5.7),(13.7,5.3,13.3,5.7)]
    for i,(tl,tt,dl,dt) in enumerate(mg):
        ts=_find_pos(s1,tl,tt,tol=0.3); ds=_find_pos(s1,dl,dt,tol=0.4)
        if i < len(mgmt_fill):
            _put(ts,mgmt_fill[i].get('title','')); _put(ds,mgmt_fill[i].get('desc','')[:34])
        else:
            _put(ts,''); _put(ds,'')
    _put(_find_pos(s1,11.2,1.9,tol=0.4), cp.get('positioning','')[:50] or '한 명도 놓치지 않는 관리 시스템으로 함께합니다.')
    for i,(l,t) in enumerate([(11.6,6.9),(13.2,6.9),(11.6,7.4)]):
        sh=_find_pos(s1,l,t,tol=0.3)
        if sh and sh.text_frame.text.strip() not in ('초등','중등','고등'):
            _put(sh, ach[i].get('text','') if i<len(ach) else '')
    if not ach:
        for l,t in [(11.2,6.9),(12.9,6.9),(11.2,7.4)]:
            sh=_find_pos(s1,l,t,tol=0.2)
            if sh and sh.text_frame.text.strip() in ('초등','중등','고등'): _put(sh,'')
    for (l,t),v in {(11.9,8.4):rules.get('homework'),(11.9,8.6):rules.get('absence'),
                    (11.9,8.8):rules.get('withdrawal'),(11.9,9.8):rules.get('refundBefore'),
                    (11.9,10.0):rules.get('refundAfter'),(11.9,10.2):rules.get('attendance')}.items():
        if v: _put(_find_pos(s1,l,t,tol=0.2), v)

    prs.save(out_path)
    return out_path
