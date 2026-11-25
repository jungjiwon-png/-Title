# -*- coding: utf-8 -*-
"""
2023 산업연관표 분석 모듈 (브라우저 Pyodide와 일반 Python에서 모두 동작)

자료: 한국은행 2023년 산업연관표(연장표), 생산자가격 투입산출표
분류: 대분류(33) · 중분류(83) · 소분류(165) · 기본부문(380)
금액 단위: 백만원. 계수는 무단위(원당 원), 취업·고용계수는 명/10억원.
"""
import json, io, re
import numpy as np
import openpyxl

LEVELS = {
    '33':  {'label': '대분류', 'n': 33},
    '83':  {'label': '중분류', 'n': 83},
    '165': {'label': '소분류', 'n': 165},
    '380': {'label': '기본부문', 'n': 380},
}
ORDER = ['33', '83', '165', '380']
FD_ITEMS = ['9111', '9112', '9121', '9122', '9131', '9132', '9140']
FD_NAMES = {'9111': '민간소비지출', '9112': '정부소비지출', '9121': '민간고정자본형성',
            '9122': '정부고정자본형성', '9131': '재고증감', '9132': '귀중품순취득',
            '9140': '수출', '9190': '최종수요계'}
TOL = 1e-8

store = {}       # 분류 → 분석용 자료
mapping = None   # 기본부문 코드 → {'380':..,'165':..,'83':..,'33':..}
names_by = {}    # 분류 → {코드: 이름} (부문분류표 기준)
employ = None    # 대분류 코드 → {'emp': 취업계수, 'wage': 고용계수}
trends = None    # 대분류 코드 → {지표: [2020, 2021, 2022, 2023]} (주요표)
TREND_SHEETS = [('3. 중간투입률', '중간투입률(%)'), ('4. 국산화율', '국산화율(%)'),
                ('5. 수입의존도', '수입의존도(%)'), ('6. 부가가치율', '부가가치율(%)'),
                ('18. 취업계수 및 고용계수', '취업계수(명/10억원)')]


def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)


def clean(s):
    return re.sub(r'\s+', ' ', str(s)).strip() if s is not None else ''


def sdiv(a, b):
    """0으로 나누는 경우(국내 생산이 없는 부문 등)는 0으로 둔다."""
    a = np.asarray(a, dtype=float); b = np.asarray(b, dtype=float)
    out = np.zeros(np.broadcast(a, b).shape)
    np.divide(a, b, out=out, where=(b != 0))
    return out


def need(level):
    if level not in store:
        raise ValueError(LEVELS[level]['label'] + '(' + level + ') 자료를 먼저 불러오세요.')
    return store[level]


# ---------------------------------------------------------------- 불러오기
def load_level(level, raw):
    """투입산출표 워크북(12개 시트)을 읽고 계수를 검증한다."""
    level = str(level)
    if level not in LEVELS:
        raise ValueError('알 수 없는 분류: ' + level)
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: [tuple(r) for r in wb[s].values] for s in wb.sheetnames}
    wb.close()
    head = book['국산투입계수(Ad)'][4]
    codes = []
    for c in head[2:]:
        c = clean(c)
        if c == '9090':
            break
        if c:
            codes.append(c)
    n = LEVELS[level]['n']
    if len(codes) != n or len(set(codes)) != n:
        raise ValueError('%s 파일의 부문 수가 %d개가 아닙니다(%d개 발견). 올바른 파일인지 확인하세요.'
                         % (LEVELS[level]['label'], n, len(codes)))
    rows0 = book['국산투입계수(Ad)']
    nm = {clean(r[0]): clean(r[1]) for r in rows0[6:] if r and r[0] is not None}
    names = [nm[c] for c in codes]

    def table(sheet):
        rows = book[sheet]
        ri = {clean(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None and i >= 6}
        ci = {clean(c): i for i, c in enumerate(rows[4]) if c is not None}
        return rows, ri, ci

    def matrix(sheet, rcodes=None, ccodes=None):
        rows, ri, ci = table(sheet)
        rc = rcodes or codes
        cc = ccodes or codes
        a = np.array([[rows[ri[r]][ci[c]] if ci[c] < len(rows[ri[r]]) else None for c in cc] for r in rc],
                     dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(sheet + ': 비수치 또는 결측 데이터')
        return a

    def vec_row(sheet, rcode, ccodes=None):
        return matrix(sheet, [rcode], ccodes)[0]

    def vec_col(sheet, ccode, rcodes=None):
        return matrix(sheet, rcodes, [ccode])[:, 0]

    T, IMP, DOM = 'A표_총거래표(생산자)', 'A표_수입거래표(생산자)', 'A표_국산거래표(생산자)'
    d = {'book': book, 'codes': codes, 'names': names, 'index': {c: i for i, c in enumerate(codes)}}
    d['L'] = matrix('생산유발계수')
    d['M'] = matrix('수입유발계수')
    d['V'] = matrix('부가가치유발계수')
    d['A'] = matrix('총투입계수(A)')
    d['Ad'] = matrix('국산투입계수(Ad)')
    d['Am'] = matrix('수입투입계수(Am)')
    d['Z'] = matrix(T)                                      # 총 중간거래 (백만원)
    d['Zm'] = matrix(IMP)                                   # 수입 중간거래
    d['X'] = vec_row(T, '9790')                             # 총투입 = 총산출
    d['va'] = vec_row(T, '9690')                            # 부가가치계
    d['va_parts'] = {k: vec_row(T, k) for k in ['9610', '9621', '9622', '9630']}
    d['mid_in'] = vec_row(T, '9590')                        # 중간투입계
    d['mid_in_m'] = vec_row(IMP, '9590')                    # 수입 중간투입계
    d['Yt'] = matrix(T, None, FD_ITEMS + ['9190'])          # 총 최종수요 (상품별)
    d['Yd'] = matrix(DOM, None, FD_ITEMS + ['9190'])        # 국산 최종수요
    d['Ym'] = matrix(IMP, None, FD_ITEMS + ['9190'])        # 수입 최종수요
    d['mid_out'] = vec_col(T, '9090')                       # 중간수요계
    d['Xrow'] = vec_col(T, '9310')                          # 총산출 (행)
    d['imports'] = vec_col(T, '9329')                       # 수입계
    d['supply'] = vec_col(T, '9390')                        # 총공급계
    d['fd_sheets'] = {k: matrix(s, None, FD_ITEMS + ['9190']) for k, s in
                      [('production', '최종수요 항목별 생산유발액'), ('imports', '최종수요 항목별 수입유발액'),
                       ('valueAdded', '최종수요 항목별 부가가치유발액')]}
    I = np.eye(n)
    v = sdiv(d['va'], d['X'])
    # 원표 영향력·감응도 계수 (생산유발계수 시트 9801 행, 9802 열)
    rows, ri, ci = table('생산유발계수')
    d['bl_sheet'] = np.array([rows[ri['9801']][ci[c]] for c in codes], dtype=float)
    d['fl_sheet'] = np.array([rows[ri[c]][ci['9802']] for c in codes], dtype=float)
    L = d['L']
    d['bl'] = L.sum(0) / L.sum(0).mean()
    d['fl'] = L.sum(1) / L.sum(1).mean()
    induced_fd = L @ d['Yd'][:, :7]
    residuals = {
        '생산유발계수 = (I-Ad)^-1': float(np.max(np.abs((I - d['Ad']) @ L - I))),
        '수입유발계수 = Am·L': float(np.max(np.abs(d['Am'] @ L - d['M']))),
        '부가가치유발계수 = Av·L': float(np.max(np.abs(v[:, None] * L - d['V']))),
        '총투입계수 = Z/X': float(np.max(np.abs(sdiv(d['Z'], d['X']) - d['A']))),
        '행·열 총산출 일치(상대오차)': float(np.max(sdiv(np.abs(d['Xrow'] - d['X']), d['X']))),
        '영향력계수 재계산': float(np.max(np.abs(d['bl'] - d['bl_sheet']))),
        '감응도계수 재계산': float(np.max(np.abs(d['fl'] - d['fl_sheet']))),
        '항목별 생산유발액 = L·Yd (상대오차)':
            float(np.max(np.abs(induced_fd - d['fd_sheets']['production'][:, :7])) /
                  max(1.0, np.max(np.abs(d['fd_sheets']['production'])))),
    }
    bad = {k: r for k, r in residuals.items() if not (r <= 1e-6)}
    if bad:
        raise ValueError('계수 검증 실패: ' + str(bad))
    d['zero'] = [c for c, x in zip(codes, d['X']) if x == 0]
    d['residuals'] = residuals
    d['X_check'] = float(d['X'].sum())
    store[level] = d
    return emit({'level': level, 'label': LEVELS[level]['label'],
                 'sectors': [{'code': c, 'name': nme} for c, nme in zip(codes, names)],
                 'sheets': list(book), 'residuals': residuals,
                 'totalOutput': float(d['X'].sum()),
                 'zeroOutput': [c + ' ' + names[d['index'][c]] for c in d['zero']]})


def load_aux(map_raw, main_raw):
    """부문분류표(380→165→83→33)와 주요표(취업·고용계수)를 읽는다."""
    global mapping, employ, trends
    wb = openpyxl.load_workbook(io.BytesIO(bytes(map_raw)), read_only=True, data_only=True)
    rows = [tuple(r) for r in wb.worksheets[0].values]
    wb.close()
    mp, cur = {}, [None] * 8
    for k in ORDER:
        names_by.setdefault(k, {})
    for r in rows[3:]:
        r = list(r) + [None] * (8 - len(r))
        if r[0] is None or clean(r[0]).startswith('9'):
            continue
        for j in range(8):
            if r[j] is not None:
                cur[j] = clean(r[j])
        code = cur[0]
        mp[code] = {'380': cur[0], '165': cur[2], '83': cur[4], '33': cur[6]}
        for k, j in zip(['380', '165', '83', '33'], [0, 2, 4, 6]):
            names_by[k][cur[j]] = cur[j + 1]
    if len(mp) != 380:
        raise ValueError('부문분류표의 기본부문이 380개가 아닙니다: %d' % len(mp))
    mapping = mp

    wb = openpyxl.load_workbook(io.BytesIO(bytes(main_raw)), read_only=True, data_only=True)
    main = {s: [tuple(r) for r in wb[s].values] for s, _ in TREND_SHEETS}
    wb.close()

    def norm(s):
        return re.sub(r'[\s\d\)\(·․,]|및', '', str(s))

    def by_label(rows):
        out = {}
        for r in rows[6:]:
            label = next((r[j] for j in (3, 2, 1) if j < len(r) and isinstance(r[j], str) and clean(r[j])), None)
            nums = [x for x in r[4:] if isinstance(x, (int, float))]
            if label is None or len(nums) < 4:
                continue
            out.setdefault(norm(label), (r, nums))
        return out
    tables = {s: by_label(main[s]) for s, _ in TREND_SHEETS}
    emp, tr = {}, {}
    for code, name in names_by['33'].items():
        key = norm(name)
        hit = tables['18. 취업계수 및 고용계수'].get(key)
        if hit is None:
            raise ValueError('주요표에서 대분류 "%s"의 취업계수를 찾지 못했습니다.' % name)
        r = hit[0]
        emp[code] = {'emp': float(r[8]), 'wage': float(r[13])}
        tr[code] = {label: [float(x) for x in tables[s][key][1][:4]] if key in tables[s] else None
                    for s, label in TREND_SHEETS}
        tr[code]['고용계수(명/10억원)'] = [float(x) for x in r[10:14]]
    employ, trends = emp, tr
    return emit({'basic': len(mapping), 'employment': len(employ)})


def parent_map(level, target):
    """분류 level의 코드 → 상위 분류 target 코드."""
    if mapping is None:
        raise ValueError('부문분류표를 먼저 불러오세요.')
    out = {}
    for m in mapping.values():
        out[m[level]] = m[target]
    return out


def employment_coefs(level):
    """부문별 취업·고용계수(명/10억원). 대분류보다 세분된 경우 소속 대분류 계수를 적용(근사)."""
    if employ is None or mapping is None:
        return None
    d = need(level)
    pm = parent_map(level, '33')
    e = np.array([employ[pm[c]]['emp'] for c in d['codes']])
    w = np.array([employ[pm[c]]['wage'] for c in d['codes']])
    return e, w


# ---------------------------------------------------------------- 원자료 탐색
def sheet_unit(sheet, row, written):
    """시트·행별 단위. 원표에 적힌 단위가 있으면 쓰고, 요약 행은 따로 표기한다."""
    if '유발액' in sheet:
        if row == '9811':
            return '유발계수: 최종수요 항목 1원당 원(무단위)'
        if row == '9812':
            return '구성비: %'
        return (written or '단위 : 백만원') + ' (9811 유발계수 행은 원/원, 9812 구성비 행은 %)'
    if '유발계수' in sheet:
        if row in ('9801', '9802') or row.startswith('98'):
            return '지수: 무단위(전 산업 평균 = 1)'
        return '계수: 해당 열 부문 최종수요 1원당 원(무단위). 9801 영향력계수 행·9802 감응도계수 열은 평균 1 기준 지수'
    if '투입계수' in sheet:
        return '계수: 해당 열 부문 산출(총투입) 1원당 투입액, 원(무단위)'
    return written or '단위 : 백만원'



def explore(level, sheet, rowcode):
    d = need(level)
    rows = d['book'][sheet]
    choices = [{'code': clean(r[0]), 'name': clean(r[1])} for r in rows[6:]
               if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if r and clean(r[0]) == rowcode and r[1] is not None), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': clean(rows[4][i]), 'name': clean(rows[5][i]), 'value': val})
    unit = sheet_unit(sheet, clean(chosen[0]), clean(rows[3][0]) if rows[3] and rows[3][0] else '')
    return emit({'rows': choices, 'selected': clean(chosen[0]), 'name': clean(chosen[1]),
                 'title': clean(rows[0][0]), 'unit': unit, 'cells': cells})


# ---------------------------------------------------------------- 산업구조
def structure(level):
    d = need(level)
    X, va = d['X'], d['va']
    out = []
    tot_x, tot_va = X.sum(), va.sum()
    for i, (c, nm) in enumerate(zip(d['codes'], d['names'])):
        mid = d['mid_in'][i]
        out.append({
            'code': c, 'name': nm,
            'output': float(X[i]), 'outputShare': float(X[i] / tot_x * 100),
            'valueAdded': float(va[i]), 'vaShare': float(va[i] / tot_va * 100),
            'midRate': float(mid / X[i] * 100) if X[i] else 0.0,
            'vaRate': float(va[i] / X[i] * 100) if X[i] else 0.0,
            'localRate': float((mid - d['mid_in_m'][i]) / mid * 100) if mid else 0.0,
            'importDep': float(d['mid_in_m'][i] / X[i] * 100) if X[i] else 0.0,
            'exportRate': float(d['Yt'][i, 6] / X[i] * 100) if X[i] else 0.0,
            'importRate': float(d['imports'][i] / d['supply'][i] * 100) if d['supply'][i] else 0.0,
            'midDemandRate': float(d['mid_out'][i] / d['supply'][i] * 100) if d['supply'][i] else 0.0,
        })
    totals = {
        'output': float(tot_x), 'valueAdded': float(tot_va),
        'midRate': float(d['mid_in'].sum() / tot_x * 100),
        'vaRate': float(tot_va / tot_x * 100),
        'localRate': float((d['mid_in'].sum() - d['mid_in_m'].sum()) / d['mid_in'].sum() * 100),
        'importDep': float(d['mid_in_m'].sum() / tot_x * 100),
        'exportRate': float(d['Yt'][:, 6].sum() / tot_x * 100),
        'importRate': float(d['imports'].sum() / d['supply'].sum() * 100),
    }
    return emit({'rows': out, 'totals': totals})


def sector_detail(level, code):
    """한 부문의 투입구조(어디서 사오는가)와 배분구조(어디에 파는가)."""
    d = need(level)
    if code not in d['index']:
        raise ValueError('부문 코드를 찾을 수 없습니다: ' + code)
    j = d['index'][code]
    X = d['X'][j]
    inputs = sorted(({'code': d['codes'][i], 'name': d['names'][i], 'value': float(d['Z'][i, j]),
                      'coef': float(d['A'][i, j]), 'importValue': float(d['Zm'][i, j])}
                     for i in range(len(d['codes']))), key=lambda r: -r['value'])[:10]
    sales = sorted(({'code': d['codes'][k], 'name': d['names'][k], 'value': float(d['Z'][j, k])}
                    for k in range(len(d['codes']))), key=lambda r: -r['value'])[:10]
    fd = [{'code': k, 'name': FD_NAMES[k], 'value': float(d['Yt'][j, t])} for t, k in enumerate(FD_ITEMS)]
    vaparts = [{'name': nm, 'value': float(d['va_parts'][k][j])} for k, nm in
               [('9610', '피용자보수'), ('9621', '영업잉여'), ('9622', '고정자본소모'), ('9630', '생산세(보조금공제)')]]
    L = d['L']
    trend = None
    if trends is not None:
        p33 = parent_map(level, '33')[code]
        trend = {'code': p33, 'name': names_by['33'].get(p33, p33), 'years': [2020, 2021, 2022, 2023],
                 'series': trends[p33]}
    return emit({'code': code, 'name': d['names'][j], 'output': float(X), 'trend': trend,
                 'midInput': float(d['mid_in'][j]), 'valueAdded': float(d['va'][j]),
                 'midDemand': float(d['mid_out'][j]), 'supply': float(d['supply'][j]),
                 'imports': float(d['imports'][j]),
                 'inputs': inputs, 'sales': sales, 'finalDemand': fd, 'vaParts': vaparts,
                 'multipliers': {'production': float(L[:, j].sum()), 'imports': float(d['M'][:, j].sum()),
                                 'valueAdded': float(d['V'][:, j].sum()),
                                 'backward': float(d['bl'][j]), 'forward': float(d['fl'][j])}})


# ---------------------------------------------------------------- 연관효과
def linkage(level):
    d = need(level)
    ec = employment_coefs(level)
    jobs = (ec[0] @ d['L']) if ec is not None else None   # 명/10억원
    out = []
    for i, (c, nm) in enumerate(zip(d['codes'], d['names'])):
        b, f = float(d['bl'][i]), float(d['fl'][i])
        if b >= 1 and f >= 1:
            g = '핵심(주도)산업'
        elif b >= 1:
            g = '후방연관 높음'
        elif f >= 1:
            g = '전방연관 높음'
        else:
            g = '독립(저연관)산업'
        out.append({'code': c, 'name': nm, 'backward': b, 'forward': f, 'group': g,
                    'production': float(d['L'][:, i].sum()), 'valueAdded': float(d['V'][:, i].sum()),
                    'imports': float(d['M'][:, i].sum()), 'output': float(d['X'][i]),
                    'jobs': float(jobs[i]) if jobs is not None else None})
    return emit({'rows': out, 'approxEmployment': level != '33'})


# ---------------------------------------------------------------- 최종수요 항목별
def final_demand(level):
    """최종수요 항목별 유발액·유발계수·구성비(원표 9819·9811·9812 행)와 부문별 생산유발의존도."""
    d = need(level)
    S = d['fd_sheets']
    sheets = {'production': '최종수요 항목별 생산유발액', 'imports': '최종수요 항목별 수입유발액',
              'valueAdded': '최종수요 항목별 부가가치유발액'}
    cols = FD_ITEMS + ['9190']
    summ = {}
    for key, sh in sheets.items():
        rows = d['book'][sh]
        ci = {clean(c): i for i, c in enumerate(rows[4]) if c is not None}
        ri = {clean(r[0]): r for r in rows[6:] if r and r[0] is not None}
        summ[key] = {k: {'sum': float(ri['9819'][ci[k]]), 'share': float(ri['9812'][ci[k]]),
                         'coef': float(ri['9811'][ci[k]])} for k in cols}
    ec = employment_coefs(level)
    items = []
    for t, k in enumerate(cols):
        jobs = jcoef = None
        if ec is not None:
            jobs = float(ec[0] @ S['production'][:, t] / 1000.0)        # 명
            p_c = summ['production'][k]['coef']
            denom = summ['production'][k]['sum'] / p_c if p_c else 0  # 유발계수 산정 기준 최종수요액
            jcoef = jobs / denom * 1000.0 if denom else None          # 명/10억원
        items.append({'jobs': jobs, 'jCoef': jcoef,'code': k, 'name': FD_NAMES[k],
                      'domestic': float(d['Yd'][:, t].sum()), 'imported': float(d['Ym'][:, t].sum()),
                      'production': summ['production'][k]['sum'], 'imports': summ['imports'][k]['sum'],
                      'valueAdded': summ['valueAdded'][k]['sum'],
                      'pShare': summ['production'][k]['share'], 'mShare': summ['imports'][k]['share'],
                      'vShare': summ['valueAdded'][k]['share'],
                      'pCoef': summ['production'][k]['coef'], 'mCoef': summ['imports'][k]['coef'],
                      'vCoef': summ['valueAdded'][k]['coef']})
    tot = S['production'][:, 7]
    dep = []
    for i, (c, nm) in enumerate(zip(d['codes'], d['names'])):
        row = {'code': c, 'name': nm, 'total': float(tot[i])}
        for t, k in enumerate(FD_ITEMS):
            row[k] = float(S['production'][i, t] / tot[i] * 100) if tot[i] else 0.0
        dep.append(row)
    return emit({'items': items, 'dependency': dep})


# ---------------------------------------------------------------- 시뮬레이션
def simulate(level, payload):
    """국산품 최종수요 변화 Δf(백만원)에 대한 파급효과. Δx = LΔf, Δm = MΔf, Δv = VΔf."""
    d = need(level)
    p = json.loads(payload)
    n = len(d['codes'])
    f = np.zeros(n)
    resolved = []
    cols_all = FD_ITEMS + ['9190']
    for s in p['shocks']:
        if s['code'] == '*':
            # 거시 충격: 모든 부문의 선택 최종수요 항목을 같은 비율로 변화
            item = s.get('item') or '9190'
            if item not in cols_all or s.get('pct') is None:
                raise ValueError('전체 산업 충격은 비율(%)과 최종수요 항목이 필요합니다.')
            pct = float(s['pct'])
            base_v = d['Yd'][:, cols_all.index(item)]
            f += base_v * pct / 100.0
            resolved.append({'code': '*', 'name': '전체 산업', 'item': FD_NAMES[item], 'pct': pct,
                             'base': float(base_v.sum()), 'amount': float(base_v.sum() * pct / 100.0)})
            continue
        if s['code'] not in d['index']:
            raise ValueError('부문 코드를 찾을 수 없습니다: ' + s['code'])
        i = d['index'][s['code']]
        if s.get('pct') is not None:
            # 기존 국산 최종수요(항목별) 대비 비율 충격
            item = s.get('item') or '9190'
            cols = FD_ITEMS + ['9190']
            if item not in cols:
                raise ValueError('알 수 없는 최종수요 항목: ' + item)
            pct = float(s['pct'])
            base = float(d['Yd'][i, cols.index(item)])
            a = base * pct / 100.0
            if not np.isfinite(a) or a == 0:
                raise ValueError('%s %s의 국산 %s이(가) 0이라 비율 충격을 줄 수 없습니다.'
                                 % (s['code'], d['names'][i], FD_NAMES[item]))
            resolved.append({'code': s['code'], 'name': d['names'][i], 'item': FD_NAMES[item],
                             'pct': pct, 'base': base, 'amount': a})
        else:
            a = float(s['amount'])
            if not np.isfinite(a) or a == 0:
                raise ValueError('변화액은 0이 아닌 유한한 수여야 합니다.')
            resolved.append({'code': s['code'], 'name': d['names'][i], 'amount': a})
        f[i] += a
    if not np.any(f):
        raise ValueError('한 부문 이상의 변화액을 입력하세요.')
    x, m, v = d['L'] @ f, d['M'] @ f, d['V'] @ f
    ec = employment_coefs(level)
    emp = wage = None
    if ec is not None:
        emp = ec[0] * x / 1000.0     # 명/10억원 × 백만원/1000
        wage = ec[1] * x / 1000.0
    rows = []
    for i, (c, nm) in enumerate(zip(d['codes'], d['names'])):
        r = {'code': c, 'name': nm, 'shock': float(f[i]), 'production': float(x[i]),
             'indirect': float(x[i] - f[i]), 'imports': float(m[i]), 'valueAdded': float(v[i])}
        if emp is not None:
            r['jobs'] = float(emp[i]); r['wageJobs'] = float(wage[i])
        rows.append(r)
    agg = None
    target = p.get('aggregate') or ''
    if target and target != level and mapping is not None and ORDER.index(target) < ORDER.index(level):
        pm = parent_map(level, target)
        acc = {}
        for r in rows:
            k = pm[r['code']]
            a = acc.setdefault(k, {'code': k, 'name': names_by[target].get(k, k), 'shock': 0.0,
                                   'production': 0.0, 'indirect': 0.0, 'imports': 0.0, 'valueAdded': 0.0,
                                   'jobs': 0.0, 'wageJobs': 0.0})
            for key in ['shock', 'production', 'indirect', 'imports', 'valueAdded', 'jobs', 'wageJobs']:
                a[key] += r.get(key, 0.0) or 0.0
        agg = list(acc.values())
    key = 'production'
    rows.sort(key=lambda r: -abs(r[key]))
    if agg:
        agg.sort(key=lambda r: -abs(r[key]))
    tot = float(f.sum())
    res = {'level': level, 'demand': tot, 'production': float(x.sum()), 'imports': float(m.sum()),
           'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / tot) if tot else None,
           'vaMultiplier': float(v.sum() / tot) if tot else None,
           'rows': rows, 'aggregate': agg, 'aggregateLevel': target if agg else None,
           'approxEmployment': level != '33', 'resolved': resolved}
    if emp is not None:
        res['jobs'] = float(emp.sum()); res['wageJobs'] = float(wage.sum())
    # 부가가치 구성별 유발: 피용자보수(가계 근로소득), 영업잉여, 고정자본소모, 순생산세(세수)
    res['vaParts'] = {k: float((sdiv(d['va_parts'][k], d['X']) * x).sum()) for k in ['9610', '9621', '9622', '9630']}
    res['vaShareOfTotal'] = float(v.sum() / d['va'].sum() * 100)
    return emit(res)


# ---------------------------------------------------------------- 분류 수준 비교
def compare(basic_code):
    """기본부문 하나를 골라, 그 부문이 속한 상위 부문의 승수를 분류 수준별로 비교한다."""
    if mapping is None:
        raise ValueError('부문분류표를 먼저 불러오세요.')
    if basic_code not in mapping:
        raise ValueError('기본부문 코드를 찾을 수 없습니다: ' + basic_code)
    out = []
    for lv in ORDER:
        code = mapping[basic_code][lv]
        row = {'level': lv, 'label': LEVELS[lv]['label'], 'code': code,
               'name': names_by[lv].get(code, code), 'loaded': lv in store}
        if lv in store:
            d = store[lv]
            j = d['index'][code]
            row.update({'production': float(d['L'][:, j].sum()), 'imports': float(d['M'][:, j].sum()),
                        'valueAdded': float(d['V'][:, j].sum()), 'backward': float(d['bl'][j]),
                        'forward': float(d['fl'][j]), 'own': float(d['L'][j, j]),
                        'output': float(d['X'][j])})
            ec = employment_coefs(lv)
            if ec is not None:
                row['jobs'] = float((ec[0] * d['L'][:, j]).sum())  # 명/10억원
        out.append(row)
    return emit({'basic': basic_code, 'rows': out})


def basic_list():
    if mapping is None:
        raise ValueError('부문분류표를 먼저 불러오세요.')
    return emit([{'code': c, 'name': names_by['380'][c], 'levels': mapping[c],
                  'path': ' > '.join(names_by[k][mapping[c][k]] for k in ['33', '83', '165'])}
                 for c in mapping])


# ---------------------------------------------------------------- 물가파급효과
def price_effect(level, payload):
    """비용 인상의 가격 파급(레온티예프 가격모형).
    외생 부문 K(가격 변화율 지정), 나머지 O는 내생:
      Δp_O' = c_O' (I - Ad_OO)^-1,  c = Σ Δpm_i·Am(i,·) + Σ Δp_k·Ad(k,·) + w·(피용자보수/총투입)
    소비자물가 근사 = 민간소비(국산+수입) 구성비로 가중한 가격 변화율."""
    d = need(level)
    p = json.loads(payload)
    n = len(d['codes'])
    idx = d['index']
    c = np.zeros(n)
    dpm = np.zeros(n)
    exo = {}
    for s in p.get('imports', []):
        if s['code'] not in idx:
            raise ValueError('부문 코드를 찾을 수 없습니다: ' + s['code'])
        r = float(s['pct']) / 100.0
        dpm[idx[s['code']]] += r
    c += dpm @ d['Am']
    for s in p.get('domestic', []):
        if s['code'] not in idx:
            raise ValueError('부문 코드를 찾을 수 없습니다: ' + s['code'])
        exo[idx[s['code']]] = float(s['pct']) / 100.0
    w = float(p.get('wage') or 0) / 100.0
    comp = sdiv(d['va_parts']['9610'], d['X'])
    c += w * comp
    for k, r in exo.items():
        c += r * d['Ad'][k, :]
    if not np.any(c) and not exo:
        raise ValueError('가격 변화를 하나 이상 입력하세요.')
    K = sorted(exo)
    O = [j for j in range(n) if j not in exo]
    dp = np.zeros(n)
    for k in K:
        dp[k] = exo[k]
    AdOO = d['Ad'][np.ix_(O, O)]
    dp[O] = np.linalg.solve((np.eye(len(O)) - AdOO).T, c[O])
    Cd, Cm = d['Yd'][:, 0], d['Ym'][:, 0]
    cpi = float((Cd @ dp + Cm @ dpm) / (Cd.sum() + Cm.sum()) * 100)
    ppi = float(d['X'] @ dp / d['X'].sum() * 100)
    rows = [{'code': cc, 'name': nm, 'price': float(dp[i] * 100), 'exogenous': i in exo,
             'consWeight': float((Cd[i] + Cm[i]) / (Cd.sum() + Cm.sum()) * 100),
             'cpiContrib': float((Cd[i] * dp[i] + Cm[i] * dpm[i]) / (Cd.sum() + Cm.sum()) * 100)}
            for i, (cc, nm) in enumerate(zip(d['codes'], d['names']))]
    rows.sort(key=lambda r: -abs(r['price']))
    return emit({'level': level, 'cpi': cpi, 'ppi': ppi, 'rows': rows,
                 'approx': '민간소비 구성비 가중 근사치'})


# ---------------------------------------------------------------- 산업 간 거래 히트맵
def heatmap(level, kind):
    """투입계수 행렬(총/국산/수입). 행: 공급(투입) 부문, 열: 수요(구매) 부문."""
    d = need(level)
    if len(d['codes']) > 83:
        raise ValueError('히트맵은 대분류·중분류에서만 그립니다.')
    key = {'A': 'A', 'Ad': 'Ad', 'Am': 'Am'}.get(kind, 'A')
    m = d[key]
    return emit({'codes': d['codes'], 'names': d['names'], 'kind': key,
                 'matrix': [[round(float(v), 6) for v in row] for row in m], 'max': float(m.max())})
