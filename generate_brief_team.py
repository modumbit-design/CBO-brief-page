"""
월간 사업부 숫자 브리프 MVP v3
6개 실 자동 생성 (퍼블리싱1/2/3/4실 + 마케팅실 + 게임운영실)
실행: python generate_brief_v3.py
"""
import pandas as pd
import os

# ============================================================
# CONFIG
# ============================================================
COST_FILE = "/home/claude/run/부서별_비용현황_2026_03.xlsx"
PNL_FILE = "/home/claude/run/_26Y프로젝트별_손익_행열전환Ver__260414_3월__25_12월_업데이트_.xlsx"
ALLOC_01 = "/home/claude/run/SGP_배부결과_26_02.xlsx"  # 전월 (팀 Card2 전월값)
ALLOC_02 = "/home/claude/run/SGP_배부결과_26_03.xlsx"  # 당월 (팀 Card2 당월값)
LEDGER_FILE = "/home/claude/run/26년_3월.XLSX"
INDEX_FILE = "/home/claude/run/인덱스.xlsx"
# WORKLOAD: 전월(02) vs 당월(03)
WORKLOAD_01 = "/home/claude/run/SGP_업무투입율_현황_2026_02.xlsx"
WORKLOAD_01_SHEET = "SGP 업무투입율 현황"
WORKLOAD_02 = "/home/claude/run/SGP_업무투입율_현황_2026_03_v2.xlsx"
WORKLOAD_02_SHEET = "SGP 업무투입율 현황"
# 12개월 트렌드용 (없으면 공란으로 폴백)
WORKLOAD_HISTORY_DIR = "/home/claude/run/wl_history"
OUTPUT_DIR = "/home/claude/run/output_team"

CURR_MONTH = 202603; PREV_MONTH = 202602; PREV_YEAR = 202503
MONTH_LABEL = "2026년 3월"
M_COL = 3; P_COL = 2

# ============================================================
# 실 정의 (시트명, 표시명, 손익 프로젝트 리스트)
# 손익 프로젝트가 None이면 손익 카드들을 스킵
# ============================================================
def load_division_config():
    """팀 버전: 인덱스 파일 기반으로 퍼블리싱 산하 9개 팀별 프로젝트 매핑"""
    idx = pd.read_excel(INDEX_FILE, header=0)
    idx.columns = ['프로젝트명','26년신작여부','25년신작여부','팀','실']
    # 프로젝트명 별칭 정규화
    idx['프로젝트명'] = idx['프로젝트명'].astype(str).map(_normalize_project_name)
    
    def _dedup(lst):
        seen = set(); out = []
        for x in lst:
            if x not in seen:
                seen.add(x); out.append(x)
        return out
    
    # 퍼블리싱 산하 9개 팀 (요청: 게임사업1~6팀 + PC·패키지·샌드박스)
    TEAMS = [
        ('게임사업1팀', '퍼블리싱사업2실', '미래시 외 준비중 신작'),
        ('게임사업2팀', '퍼블리싱사업1실', '에픽세븐 IP'),
        ('게임사업3팀', '퍼블리싱사업2실', 'PJT D 외'),
        ('게임사업4팀', '퍼블리싱사업4실', '로드나인 IP'),
        ('게임사업5팀', '퍼블리싱사업4실', '이클립스'),
        ('게임사업6팀', '퍼블리싱사업3실', '카오스제로 외'),
        ('PC사업팀',   '퍼블리싱사업4실', '크로스파이어 외 PC 라이브'),
        ('패키지사업팀', '퍼블리싱사업4실', 'PC 패키지'),
        ('샌드박스사업팀','퍼블리싱사업2실','버블리즈'),
    ]
    
    out = []
    for team_name, parent_sil, sub in TEAMS:
        projs = _dedup(idx[idx['팀']==team_name]['프로젝트명'].tolist())
        out.append({
            'sheet': team_name,  # 구분자 용도
            'name' : team_name,
            'sub'  : f"{parent_sil} · {sub}",
            'projects': projs if projs else None,
            'teams': [team_name],  # 자기 자신 1개 팀
            'parent_sil': parent_sil,
        })
    return out


# ============================================================
# 팀용 Card 2: 원장 기반 집계 (2월/3월 원장)
# = 팀 코스트센터 비용 (해당 팀이 쓴 비용)
# + WBS가 팀 프로젝트인 타 부서 집행분 (타 부서가 팀 프로젝트에 쓴 비용)
# 비용계정(계정번호 5로 시작)만 집계, 금액 단순합
# ============================================================
LEDGER_PREV = "/home/claude/run/26년_2월.XLSX"
LEDGER_CURR = "/home/claude/run/26년_3월.XLSX"

# 팀 직접비 집계 시 제외할 계정번호 (매출원가성: R/S·수수료류)
# 이 비용들은 손익상 이미 매출(Net) 산정 시 서비스직접비로 차감되어
# 팀이 통제 가능한 "운영비용"과 구분되어야 함
CARD2_EXCLUDE_ACCTS = {
    '53090600',  # 지급수수료-마켓수수료
    '53090700',  # 지급수수료-PG수수료
    '53090800',  # 지급수수료-게임서비스RS
    '53090801',  # 지급수수료-게임RS-온라인
    '53090802',  # 지급수수료-게임RS-모바일
    '53090809',  # 지급수수료-게임RS-기타
    '53091501',  # 지급수수료-가정산-마켓수수료
    '53091503',  # 지급수수료-가정산-게임RS-온라인
    '53091504',  # 지급수수료-가정산-게임RS-모바일
    '53091509',  # 지급수수료-가정산-게임RS-기타
    '53091511',  # 지급수수료-수익이연-마켓수수료
    '53091512',  # 지급수수료-수익이연-PG수수료
    '53091513',  # 지급수수료-수익이연-게임RS-온라인
    '53091514',  # 지급수수료-수익이연-게임RS-모바일
}

def _load_ledger_cached(path, _cache={}):
    if path not in _cache:
        df = pd.read_excel(path, sheet_name="Sheet1", header=0)
        df['계정 번호'] = df['계정 번호'].astype(str)
        # 비용계정(5xxx)만 유지
        df = df[df['계정 번호'].str.startswith('5')].copy()
        # 매출원가성 계정(R/S·수수료류) 제외 → 팀이 통제 가능한 운영비만 남김
        df = df[~df['계정 번호'].isin(CARD2_EXCLUDE_ACCTS)].copy()
        df['금액(회사 코드 통화)'] = pd.to_numeric(df['금액(회사 코드 통화)'], errors='coerce').fillna(0)
        _cache[path] = df
    return _cache[path]

def _count_team_headcount(team_name, team_projects):
    """팀의 당월 총 MM값을 반환 (Man-Month, 인력 규모 프록시).
    투입율 시트에 개인별 행이 없어서 실제 인원수를 셀 수 없으므로,
    팀 단위 집계값인 MM을 대신 사용.
    
    반환: float (소수점 1자리 MM). 예: 8.4 → "8.4MM"
    """
    try:
        df = pd.read_excel(WORKLOAD_02, sheet_name=WORKLOAD_02_SHEET, header=None)
    except Exception:
        return None
    for i in range(df.shape[0]):
        if str(df.iloc[i,1]).strip() == team_name:
            # 해당 행에서 숫자값들을 전부 더함 (프로젝트 컬럼들)
            row = df.iloc[i]
            total = 0.0
            for v in row[2:]:
                if pd.notna(v) and isinstance(v, (int, float)):
                    total += float(v)
            return round(total, 1) if total > 0 else None
    return None


def extract_card2_team(team_name, team_projects):
    """
    team_name: 코스트센터명 (예: '게임사업4팀')
    team_projects: 해당 팀 귀속 프로젝트 WBS명 리스트 (인덱스 기준)
    """
    curr = _load_ledger_cached(LEDGER_CURR)
    prev = _load_ledger_cached(LEDGER_PREV)
    team_projects = team_projects or []
    
    def _agg(df):
        """(A) 코스트센터=팀 + (B) WBS=팀 프로젝트 & 코스트센터!=팀"""
        a = df[df['코스트센터명'] == team_name].copy()
        a['_src'] = 'CC'
        if team_projects:
            b = df[(df['WBS명'].isin(team_projects)) & (df['코스트센터명'] != team_name)].copy()
            b['_src'] = 'WBS'
        else:
            b = df.iloc[0:0].copy()
            b['_src'] = 'WBS'
        return pd.concat([a, b], ignore_index=True)
    
    c = _agg(curr)
    p = _agg(prev)
    
    dt_c = c[c['_src']=='CC']['금액(회사 코드 통화)'].sum()
    dt_p = p[p['_src']=='CC']['금액(회사 코드 통화)'].sum()
    pt_c = c[c['_src']=='WBS']['금액(회사 코드 통화)'].sum()
    pt_p = p[p['_src']=='WBS']['금액(회사 코드 통화)'].sum()
    
    gt_c = dt_c + pt_c
    gt_p = dt_p + pt_p
    
    NOISE = 5_000_000
    
    # CC(부서 직접비) 계정 top
    cc_curr = c[c['_src']=='CC'].groupby('계정명')['금액(회사 코드 통화)'].sum()
    cc_prev = p[p['_src']=='CC'].groupby('계정명')['금액(회사 코드 통화)'].sum()
    accts = []
    for name in set(cc_curr.index) | set(cc_prev.index):
        cv = int(cc_curr.get(name, 0))
        pv = int(cc_prev.get(name, 0))
        if abs(cv) < NOISE and abs(pv) < NOISE: continue
        accts.append({'name': name, 'curr': cv, 'prev': pv, 'delta': cv - pv})
    accts.sort(key=lambda x: -abs(x['curr']))
    
    # WBS(타부서 집행) 계정 top
    wb_curr = c[c['_src']=='WBS'].groupby('계정명')['금액(회사 코드 통화)'].sum()
    wb_prev = p[p['_src']=='WBS'].groupby('계정명')['금액(회사 코드 통화)'].sum()
    proj_accts = []
    for name in set(wb_curr.index) | set(wb_prev.index):
        cv = int(wb_curr.get(name, 0))
        pv = int(wb_prev.get(name, 0))
        if abs(cv) < NOISE and abs(pv) < NOISE: continue
        proj_accts.append({'name': name, 'curr': cv, 'prev': pv, 'delta': cv - pv})
    proj_accts.sort(key=lambda x: -abs(x['curr']))
    
    # headcount (투입율 시트에서 팀 프로젝트에 MM>0인 인원)
    headcount = _count_team_headcount(team_name, team_projects)
    
    return {
        'hc': headcount,
        'dt': {'c': int(dt_c), 'p': int(dt_p), 'd': int(dt_c - dt_p)},
        'pt': {'c': int(pt_c), 'p': int(pt_p), 'd': int(pt_c - pt_p)},
        'gt': {'c': int(gt_c), 'p': int(gt_p), 'd': int(gt_c - gt_p)},
        'da': accts[:6], 'pa': proj_accts,
        'sheet_total': int(gt_c),
        'sheet_total_prev': int(gt_p),
    }

# ============================================================
# CARD 2: 부서 총지출 (시트별로 동작)
# ============================================================
def extract_card2(sheet_name, mkt_addon=0):
    """sheet_name: 부서별 시트명
    mkt_addon: 원장 기준 마케팅비 (시트에 안 잡힌 분, 총액에 추가)"""
    df = pd.read_excel(COST_FILE, sheet_name=sheet_name, header=None)
    accts, proj_accts, in_proj = [], [], False
    headcount = None
    sheet_total_curr = 0
    sheet_total_prev = 0
    
    for i, row in df.iterrows():
        v = str(row[1]).strip() if pd.notna(row[1]) else ""
        if v == '총 비용':
            sheet_total_curr = int(row[M_COL]) if pd.notna(row[M_COL]) and isinstance(row[M_COL],(int,float)) else 0
            sheet_total_prev = int(row[P_COL]) if pd.notna(row[P_COL]) and isinstance(row[P_COL],(int,float)) else 0
    
    EXCLUDE = ['지급임차료', '건물관리비']
    
    for i, row in df.iterrows():
        v = str(row[1]).strip() if pd.notna(row[1]) else ""
        if '인원' in v and headcount is None:
            headcount = int(row[M_COL]) if pd.notna(row[M_COL]) else None
        if '직접 프로젝트 비용' in v: in_proj = True; continue
        if v and len(v)>=3 and v[:2].isdigit() and v[2]=='.':
            if not in_proj and any(ex in v for ex in EXCLUDE):
                continue
            c = int(row[M_COL]) if pd.notna(row[M_COL]) and isinstance(row[M_COL],(int,float)) else 0
            p = int(row[P_COL]) if pd.notna(row[P_COL]) and isinstance(row[P_COL],(int,float)) else 0
            e = {'name':v,'curr':c,'prev':p,'delta':c-p}
            (proj_accts if in_proj else accts).append(e)
    
    dt = sum(a['curr'] for a in accts); dp = sum(a['prev'] for a in accts)
    pt = sum(a['curr'] for a in proj_accts); pp = sum(a['prev'] for a in proj_accts)
    
    # 500만원 미만 노이즈 제거
    NOISE_THRESHOLD = 5_000_000
    def is_significant(a):
        return abs(a['curr']) >= NOISE_THRESHOLD or abs(a['prev']) >= NOISE_THRESHOLD
    
    ad = sorted([a for a in accts if is_significant(a)], key=lambda x:-abs(x['curr']))[:6]
    ap = sorted([a for a in proj_accts if is_significant(a)], key=lambda x:-abs(x['curr']))
    
    sheet_total_curr = sheet_total_curr if sheet_total_curr else 0
    sheet_total_prev = sheet_total_prev if sheet_total_prev else 0
    
    # 마케팅비 addon: 총액에 더하고, 직접 프로젝트 비용 영역 맨 위에 가상 행 삽입
    if mkt_addon and abs(mkt_addon) >= NOISE_THRESHOLD:
        # 시트에 이미 있는 마케팅비 0인 행은 제거 (중복 방지)
        ap = [a for a in ap if not ('마케팅비' in a['name'] and abs(a['curr']) < NOISE_THRESHOLD)]
        ap.insert(0, {
            'name': '01. 마케팅비 (원장)',
            'curr': mkt_addon,
            'prev': 0,
            'delta': mkt_addon,
        })
        pt += mkt_addon
        sheet_total_curr += mkt_addon
    
    return {'hc':headcount,
            'dt':{'c':dt,'p':dp,'d':dt-dp},
            'pt':{'c':pt,'p':pp,'d':pt-pp},
            'gt':{'c':sheet_total_curr,'p':sheet_total_prev,'d':sheet_total_curr-sheet_total_prev},
            'da':ad,'pa':ap,
            'sheet_total': sheet_total_curr,
            'sheet_total_prev': sheet_total_prev,
            }


def analyze_marketing_from_ledger(teams):
    """원장에서 해당 실 팀들이 직접 사용한 마케팅비를 WBS별로 추출 (코스트센터 기준)"""
    if not teams: return None
    try:
        df = pd.read_excel(LEDGER_FILE, header=0)
        mk = df[df['코스트센터명'].isin(teams) & df['계정명'].str.contains('광고선전|마케팅', na=False)]
        if len(mk) == 0: return None
        total = mk['금액(회사 코드 통화)'].sum()
        if abs(total) < 1e6: return None
        result = []
        g = mk.groupby('WBS명')['금액(회사 코드 통화)'].sum().sort_values(key=abs, ascending=False)
        for wbs, val in g.items():
            if abs(val) < 1e6: continue
            pct = val/total*100
            result.append({'wbs':wbs,'val':int(val),'pct':pct})
        return {'total':int(total),'items':result}
    except Exception as e:
        print(f"  ⚠ 마케팅비 분석 실패 (코스트센터): {e}")
        return None


def analyze_marketing_by_projects(projects):
    """원장에서 우리 실 담당 프로젝트(WBS)에 들어간 마케팅비를 추출 (WBS 기준)
    어느 부서에서 일으켰든 상관없이 우리 프로젝트에 귀속된 광고선전비 전체.
    WBS 단위로 합산 (코스트센터 양수/음수 상계 후 순액)
    
    리턴: {'total': int, 'items': [{'wbs','val','pct'}], 'by_wbs': {wbs: val}}
    """
    if not projects: return None
    try:
        df = pd.read_excel(LEDGER_FILE, header=0)
        mk = df[df['WBS명'].isin(projects) & df['계정명'].str.contains('광고선전|마케팅', na=False)]
        if len(mk) == 0: return None
        # WBS 단위로 합산 (양수/음수 상계 → 순액)
        g = mk.groupby('WBS명')['금액(회사 코드 통화)'].sum()
        # 0 또는 음수 제외
        g_pos = g[g > 0].sort_values(ascending=False)
        total = g_pos.sum()
        if total < 1e6: return None
        result = []
        for wbs, val in g_pos.items():
            if val < 1e5: continue  # 10만원 미만 노이즈 제외
            pct = val/total*100
            result.append({'wbs':wbs,'val':int(val),'pct':pct})
        return {
            'total':int(total),
            'items':result,
            'by_wbs': {wbs: int(v) for wbs, v in g.items()},  # 프로젝트 단건 조회용 (음수 포함 원본)
        }
    except Exception as e:
        print(f"  ⚠ 마케팅비 분석 실패 (WBS): {e}")
        return None


def analyze_other_from_ledger(projects):
    """원장에서 '기타' 카테고리 추출"""
    if not projects: return None
    try:
        df = pd.read_excel(LEDGER_FILE, header=0)
        f = df[df['WBS명'].isin(projects)].copy()
        cost = f[~f['계정명'].str.contains('매출|수익', na=False)]
        def is_other(name):
            if '광고선전' in name or '마케팅' in name: return False
            if 'IT서비스' in name: return False
            if '외주용역' in name: return False
            if '지급수수' in name: return False
            return True
        others = cost[cost['계정명'].apply(is_other)]
        if len(others) == 0: return None
        result = []
        g = others.groupby(['WBS명','계정명'])['금액(회사 코드 통화)'].sum().sort_values(key=abs, ascending=False)
        for (wbs, acct), val in g.items():
            if abs(val) < 1e6: continue
            result.append({'wbs':wbs, 'acct':acct, 'val':int(val)})
        return result
    except Exception as e:
        print(f"  ⚠ 원장 분석 실패: {e}")
        return None


# ============================================================
# CARD 3/4: 관리손익 + 프로젝트별
# ============================================================
def extract_card3_4(projects):
    if not projects: return None, None
    pnl = pd.read_excel(PNL_FILE, sheet_name="기초", header=0)
    # 손익 파일의 프로젝트명도 정규화 (Project TT → 미래시 등)
    pnl['프로젝트'] = pnl['프로젝트'].astype(str).map(_normalize_project_name)
    h = pnl[pnl['구분2']=='합산'].copy()
    ms = ['매출(Net)','공헌이익','영업이익']
    
    # 매출 흐름용 추가 컬럼
    extra_cols = ['매출','서비스직접비','R/S','수수료','사내직접비용']
    
    def gs(mo,t,pj):
        f=h[(h['년월']==mo)&(h['구분1']==t)&(h['프로젝트'].isin(pj))]
        result = {m:f[m].sum() for m in ms}
        for c in extra_cols:
            if c in f.columns:
                result[c] = f[c].sum()
            else:
                result[c] = 0
        return result
    
    # 최근 6개월 (현재 포함, 역순) - 실적 기준
    def recent_6_months(curr_month):
        months = []
        y = curr_month // 100
        m = curr_month % 100
        for _ in range(6):
            months.append(y*100 + m)
            m -= 1
            if m == 0:
                m = 12; y -= 1
        return list(reversed(months))
    
    # 25년 1월부터 현재월까지 전체 (슬라이더용)
    def months_from_2501(curr_month):
        months = []
        y, m = 2025, 1
        while y*100 + m <= curr_month:
            months.append(y*100 + m)
            m += 1
            if m == 13:
                m = 1; y += 1
        return months
    
    months_full = months_from_2501(CURR_MONTH)
    
    def trend(pj, months_list):
        per_month = [gs(mo, '실적', pj) for mo in months_list]
        # 매출(Net) 100만원 이하인 월은 제거 (매출 0인데 영업이익만 잡힌 이상치 차단)
        SALES_THRESHOLD = 1_000_000
        kept = [(mo, d) for mo, d in zip(months_list, per_month)
                if abs(d['매출(Net)']) > SALES_THRESHOLD]
        
        # 영업이익 전용 폴백 시계열 (매출 없어도 영업이익이 있는 월은 포함)
        # 단, 5월 +13억 같은 이상치(매출 0 + 영업이익 양수)는 여기서도 제외
        # → 영업이익이 음수인 월만 신뢰 (준비중 신작은 -비용으로 잡히는 게 정상)
        op_kept = [(mo, d) for mo, d in zip(months_list, per_month)
                   if d['영업이익'] < 0 or abs(d['매출(Net)']) > SALES_THRESHOLD]
        
        if not kept:
            # 매출 시계열 빈 경우: 영업이익 단일 라인용 폴백 제공
            return {
                'months': [], '매출(Net)': [], '영업이익': [], '공헌이익': [], '매출(그로스)': [],
                '_op_only_months': [mo for mo, _ in op_kept],
                '_op_only_values': [d['영업이익'] for _, d in op_kept],
            }
        return {
            'months': [mo for mo, _ in kept],
            '매출(Net)': [d['매출(Net)'] for _, d in kept],
            '매출(그로스)': [d.get('매출', d['매출(Net)']) for _, d in kept],
            '영업이익': [d['영업이익'] for _, d in kept],
            '공헌이익': [d['공헌이익'] for _, d in kept],
            '_op_only_months': [mo for mo, _ in op_kept],
            '_op_only_values': [d['영업이익'] for _, d in op_kept],
        }
    
    a=gs(CURR_MONTH,'실적',projects); p=gs(CURR_MONTH,'계획',projects)
    pm=gs(PREV_MONTH,'실적',projects); py=gs(PREV_YEAR,'실적',projects)
    
    c3 = {}
    for m in ms:
        c3[m] = {
            'a':a[m], 'vp':a[m]-p[m], 'vm':a[m]-pm[m], 'vy':a[m]-py[m],
        }
    # 매출 흐름 추가 (그로스, 서비스직접비)
    c3['매출(그로스)'] = {
        'a': a.get('매출', 0),
        'vp': a.get('매출', 0) - p.get('매출', 0),
        'vm': a.get('매출', 0) - pm.get('매출', 0),
        'vy': a.get('매출', 0) - py.get('매출', 0),
    }
    c3['서비스직접비'] = {
        'a': a.get('서비스직접비', 0),
        'vp': a.get('서비스직접비', 0) - p.get('서비스직접비', 0),
        'vm': a.get('서비스직접비', 0) - pm.get('서비스직접비', 0),
        'vy': a.get('서비스직접비', 0) - py.get('서비스직접비', 0),
    }
    # R/S, 수수료, 사내직접 분해 (참고용 라벨)
    c3['_svc_breakdown'] = {
        'R/S': a.get('R/S', 0),
        '수수료': a.get('수수료', 0),
        '사내직접': a.get('사내직접비용', 0),
    }
    c3['_trend'] = trend(projects, months_full)
    
    months_6 = recent_6_months(CURR_MONTH)
    
    # ─── 별칭 그룹화: 인덱스에 "(P)_중국_운영"과 "_중국_운영" 같은 중복이 있으면
    # 손익 파일에서 계획/실적이 분리되어 들어있을 수 있으므로 하나의 사업으로 묶어서 합산
    groups = []  # [(대표명, [별칭들])]
    used = set()
    for i, pj in enumerate(projects):
        if i in used: continue
        group = [pj]
        used.add(i)
        for j in range(i+1, len(projects)):
            if j in used: continue
            if _proj_match(projects[j], [pj]):
                group.append(projects[j])
                used.add(j)
        # 대표명: 가장 짧지 않고 "(P)"가 있는 걸 우선
        rep = sorted(group, key=lambda x: (0 if '(P)' in str(x) else 1, -len(str(x))))[0]
        groups.append((rep, group))
    
    c4=[]
    for rep, group in groups:
        pa=gs(CURR_MONTH,'실적',group); pp=gs(CURR_MONTH,'계획',group)
        ppm=gs(PREV_MONTH,'실적',group); ppy=gs(PREV_YEAR,'실적',group)
        if all(pa[m]==0 for m in ms) and all(pp[m]==0 for m in ms): continue
        c4.append({
            'n':rep,
            'm':{m:{'a':pa[m],'vp':pa[m]-pp[m],'vm':pa[m]-ppm[m],'vy':pa[m]-ppy[m]} for m in ms},
            'trend': trend(group, months_6),
        })
    c4.sort(key=lambda x:-abs(x['m']['매출(Net)']['a']))
    return c3,c4


# ============================================================
# CARD 5: 실제 투입 MM (업무투입율)
# ============================================================
def _load_workload(fpath, sheet):
    df = pd.read_excel(fpath, sheet_name=sheet, header=None)
    projects = df.iloc[7, 2:47].tolist()  # 행 7 = 프로젝트명
    # 별칭 정규화 (Project TT → 미래시 등)
    projects = [_normalize_project_name(p) if pd.notna(p) else p for p in projects]
    return df, projects

def _get_div_mm(df, projects, teams, filter_projects=None):
    by_proj = {}
    total = 0
    for i in range(len(df)):
        team = df.iloc[i, 1]
        # 총계행 이후는 카테고리 집계표
        if pd.notna(team) and '투입 MM 총계' in str(team):
            break
        if pd.notna(team) and str(team).strip() in teams:
            for c in range(2, 47):
                v = df.iloc[i, c]
                if pd.notna(v) and isinstance(v,(int,float)) and v != 0:
                    pname = projects[c-2]
                    if pd.notna(pname):
                        # 담당 사업 프로젝트 필터 (있으면)
                        if filter_projects is not None:
                            if not _proj_match(pname, filter_projects):
                                continue
                        by_proj[pname] = by_proj.get(pname, 0) + v
                        total += v
    return total, by_proj


# ---- 12개월 MM 트렌드용 캐싱 ----
_WL_MONTH_CACHE = {}

def _last_12_months(curr_month):
    """현재월 포함 과거 12개월 yyyymm 리스트 (오래된 순)"""
    months = []
    y = curr_month // 100
    m = curr_month % 100
    for _ in range(12):
        months.append(y*100 + m)
        m -= 1
        if m == 0:
            m = 12; y -= 1
    return list(reversed(months))

# 월별 워크로드 파일 직접 매핑 (업로드된 실제 파일)
WORKLOAD_BY_MONTH = {
    202501: "/home/claude/run/SGP_업무투입율_현황_2025_01.xlsx",
    202502: "/home/claude/run/SGP_업무투입율_현황_2025_02.xlsx",
    202503: "/home/claude/run/SGP_업무투입율_현황_2025_03.xlsx",
    202504: "/home/claude/run/SGP_업무투입율_현황_2025_04.xlsx",
    202505: "/home/claude/run/SGP_업무투입율_현황_2025_05.xlsx",
    202506: "/home/claude/run/SGP_업무투입율_현황_2025_06.xlsx",
    202507: "/home/claude/run/SGP_업무투입율_현황_2025_07.xlsx",
    202508: "/home/claude/run/SGP_업무투입율_현황_2025_08.xlsx",
    202509: "/home/claude/run/SGP_업무투입율_현황_2025_09.xlsx",
    202510: "/home/claude/run/SGP_업무투입율_현황_2025_10.xlsx",
    202511: "/home/claude/run/SGP_업무투입율_현황_2025_11.xlsx",
    202512: "/home/claude/run/SGP_업무투입율_현황_2025_12.xlsx",
    202601: "/home/claude/run/SGP_업무투입율_현황_2026_01.xlsx",
    202602: "/home/claude/run/SGP_업무투입율_현황_2026_02.xlsx",
    202603: "/home/claude/run/SGP_업무투입율_현황_2026_03_v2.xlsx",
}

def _load_workload_for_month(yyyymm):
    """yyyymm(예: 202603) → (df, projects). 파일 없으면 (None, None) 반환. 캐싱."""
    if yyyymm in _WL_MONTH_CACHE:
        return _WL_MONTH_CACHE[yyyymm]
    # 1) 직접 매핑 우선
    fpath = WORKLOAD_BY_MONTH.get(yyyymm)
    # 2) 폴백: 히스토리 디렉토리 규칙
    if not fpath or not os.path.exists(fpath):
        y = yyyymm // 100
        m = yyyymm % 100
        fpath = os.path.join(WORKLOAD_HISTORY_DIR, f"WL_{y}_{m:02d}.xlsx")
    if not os.path.exists(fpath):
        _WL_MONTH_CACHE[yyyymm] = (None, None)
        return None, None
    try:
        df = pd.read_excel(fpath, sheet_name="SGP 업무투입율 현황", header=None)
        projects = df.iloc[7, 2:47].tolist()
        # 별칭 정규화 (Project TT → 미래시 등)
        projects = [_normalize_project_name(p) if pd.notna(p) else p for p in projects]
        _WL_MONTH_CACHE[yyyymm] = (df, projects)
        return df, projects
    except Exception as e:
        print(f"  ⚠ {fpath} 로드 실패: {e}")
        _WL_MONTH_CACHE[yyyymm] = (None, None)
        return None, None


def extract_mm_trend_for_project(target_proj, my_teams, months_list):
    """특정 프로젝트의 월별 (총 MM, 우리 실 MM) 시계열 추출.
    파일 없는 월은 0으로 채움.
    리턴: {'months': [...], 'total': [...], 'inside': [...]}
    """
    totals = []
    insides = []
    for ym in months_list:
        df, projects = _load_workload_for_month(ym)
        if df is None:
            totals.append(0); insides.append(0)
            continue
        # 컬럼 매칭 (alias)
        col_idx = None
        for i, p in enumerate(projects):
            if pd.notna(p) and _proj_match(p, [target_proj]):
                col_idx = i + 2
                break
        if col_idx is None:
            totals.append(0); insides.append(0)
            continue
        # 합산
        total = 0
        inside = 0
        for r in range(8, len(df)):
            team = df.iloc[r, 1]
            if pd.notna(team) and '투입 MM 총계' in str(team):
                break
            if not _is_real_team(team): continue
            v = df.iloc[r, col_idx]
            if not (pd.notna(v) and isinstance(v,(int,float)) and v != 0): continue
            total += v
            if str(team).strip() in my_teams:
                inside += v
        totals.append(total)
        insides.append(inside)
    return {'months': months_list, 'total': totals, 'inside': insides}


def _proj_match(workload_name, target_projects):
    """업무투입율 시트의 프로젝트명과 인덱스의 프로젝트명 매칭 (alias 포함)"""
    wn = _normalize_project_name(str(workload_name).strip())
    for tp in target_projects:
        tpn = _normalize_project_name(str(tp).strip())
        if wn == tpn:
            return True
        # 핵심 키워드 추출 (괄호, 운영 등 제거)
        wn_key = wn.replace('(P)','').replace('_운영','').replace('_',' ').strip()
        tp_key = tpn.replace('(P)','').replace('_운영','').replace('_',' ').strip()
        if wn_key == tp_key:
            return True
        # 부분 매칭 (예: "이클립스" ↔ "이클립스(P)_글로벌")
        if wn_key in tp_key or tp_key in wn_key:
            # 길이 차이 너무 크면 거부 (오매칭 방지)
            short = min(len(wn_key), len(tp_key))
            long = max(len(wn_key), len(tp_key))
            if short >= 3 and short / long >= 0.4:
                return True
    return False


# ============================================================
# 프로젝트명 별칭 정규화
# 인덱스/손익/워크로드 파일에 같은 사업이 다른 이름으로 들어가는 경우 통합
# ============================================================
PROJECT_ALIASES = {
    # 내부 코드명 → 실제 사업명 (인덱스/손익 기준으로 통일)
    # ── 미래시 (Project TT)
    'Project TT_운영': '미래시_운영',
    'Project TT': '미래시_운영',
    'PJT TT_운영': '미래시_운영',
    'PJT TT': '미래시_운영',
    '미래시': '미래시_운영',
    # ── 버블리즈 (Project B / 파티게임)
    '파티게임_운영': '버블리즈_운영',
    '파티게임': '버블리즈_운영',
    'Project B_운영': '버블리즈_운영',
    'Project B': '버블리즈_운영',
    'PJT B_운영': '버블리즈_운영',
    'PJT B': '버블리즈_운영',
    '버블리즈': '버블리즈_운영',
    # ── 데드어카운트 (Project D / PJT D)
    'Project D_운영': '데드어카운트_운영',
    'Project D': '데드어카운트_운영',
    'PJT D_운영': '데드어카운트_운영',
    'PJT D': '데드어카운트_운영',
    '데드어카운트': '데드어카운트_운영',
    # ── 이클립스 (글로벌 서픽스 다양)
    '이클립스': '이클립스_운영',
    '이클립스(P)_글로벌': '이클립스_운営'.replace('営','영'),  # 안전망
    '이클립스(P)_글로벌_운영': '이클립스_운영',
    # 추가 별칭은 여기에 계속 쌓아나감
}

def _normalize_project_name(name):
    """프로젝트명 별칭 정규화. 별칭이 없으면 원본 그대로 반환."""
    if not name: return name
    s = str(name).strip()
    return PROJECT_ALIASES.get(s, s)



def _load_team_classification():
    """팀분류 시트 → {팀명: 분류} dict"""
    try:
        clf = pd.read_excel(WORKLOAD_02, sheet_name="팀분류", header=0)
        # 앞 2개 컬럼(팀명, 분류)만 사용. 추가 컬럼은 무시
        clf = clf.iloc[:, :2].copy()
        clf.columns = ['팀명','분류']
        clf = clf.dropna(subset=['팀명','분류'])
        return dict(zip(clf['팀명'].astype(str).str.strip(), clf['분류'].astype(str).str.strip()))
    except Exception as e:
        print(f"  ⚠ 팀분류 로드 실패: {e}")
        return {}

_SKIP_TEAM_KEYWORDS = ['투입 MM 총계','총합','소계','합계','부문','구분','연도','프로젝트']

def _is_real_team(team_str):
    if pd.isna(team_str): return False
    t = str(team_str).strip()
    if not t: return False
    for k in _SKIP_TEAM_KEYWORDS:
        if k in t: return False
    return True

def _get_project_breakdown(df, projects, target_proj, my_teams, team2cat):
    """특정 프로젝트의 분류별 MM 분포 + 우리 실 본인 MM 분리"""
    # alias 매칭으로 컬럼 찾기
    col_idx = None
    for i, p in enumerate(projects):
        if pd.notna(p) and _proj_match(p, [target_proj]):
            col_idx = i + 2
            break
    if col_idx is None:
        return None
    
    by_cat = {}
    my_inside = 0
    unmapped = 0
    for i in range(8, len(df)):
        team = df.iloc[i, 1]
        # 총계행 이후는 카테고리 집계표라 중복 카운트되므로 중단
        if pd.notna(team) and '투입 MM 총계' in str(team):
            break
        if not _is_real_team(team): continue
        v = df.iloc[i, col_idx]
        if not (pd.notna(v) and isinstance(v,(int,float)) and v != 0): continue
        t = str(team).strip()
        cat = team2cat.get(t)
        if cat:
            by_cat[cat] = by_cat.get(cat, 0) + v
            if t in my_teams:
                my_inside += v
        else:
            unmapped += v
    
    if unmapped > 0:
        by_cat['기타'] = by_cat.get('기타', 0) + unmapped
    
    total = sum(by_cat.values())
    if total == 0: return None
    items = sorted(by_cat.items(), key=lambda x:-x[1])
    return {
        'project': target_proj,
        'col_idx': col_idx,
        'total': total,
        'inside': my_inside,
        'outside': total - my_inside,
        'categories': [{'name':k,'mm':v,'pct':v/total*100} for k,v in items],
    }


def extract_workload(teams, filter_projects=None):
    """팀 리스트 기반 MM 추출. filter_projects 주면 담당 사업만 필터 + breakdown 생성"""
    if not teams: return None
    try:
        df02, p02 = _load_workload(WORKLOAD_02, WORKLOAD_02_SHEET)
        df01, p01 = _load_workload(WORKLOAD_01, WORKLOAD_01_SHEET)
        team2cat = _load_team_classification()
        
        t02, by02 = _get_div_mm(df02, p02, teams, filter_projects)
        t01, by01 = _get_div_mm(df01, p01, teams, filter_projects)
        team_detail = []
        for tm in teams:
            tt02, _ = _get_div_mm(df02, p02, [tm], filter_projects)
            tt01, _ = _get_div_mm(df01, p01, [tm], filter_projects)
            if tt02 > 0 or tt01 > 0:
                team_detail.append({'team':tm,'curr':tt02,'prev':tt01,'delta':tt02-tt01})
        proj_list = []
        all_projs = set(by02.keys()) | set(by01.keys())
        for pj in all_projs:
            c = by02.get(pj, 0)
            p = by01.get(pj, 0)
            proj_list.append({'name':pj,'curr':c,'prev':p,'delta':c-p})
        proj_list.sort(key=lambda x:-x['curr'])
        
        # 프로젝트별 분류 분포 (담당 사업 있는 실만)
        breakdowns = []
        if filter_projects:
            seen_cols = set()
            for pj in filter_projects:
                bd = _get_project_breakdown(df02, p02, pj, teams, team2cat)
                if not bd: continue
                if bd['col_idx'] in seen_cols:
                    print(f"   ⚠ 중복 별칭 스킵: '{pj}' (워크로드 열 {bd['col_idx']} 이미 매칭됨)")
                    continue
                seen_cols.add(bd['col_idx'])
                breakdowns.append(bd)
            breakdowns.sort(key=lambda x:-x['total'])
            # 각 breakdown에 12개월 MM 시계열 추가 (현재월 포함, 과거 12개월)
            mm_trend_months = _last_12_months(CURR_MONTH)
            for bd in breakdowns:
                bd['mm_trend'] = extract_mm_trend_for_project(
                    bd['project'], teams, mm_trend_months
                )
        
        return {'curr':t02,'prev':t01,'delta':t02-t01,
                'teams':team_detail,'projects':proj_list,
                'filtered': filter_projects is not None,
                'breakdowns':breakdowns}
    except Exception as e:
        print(f"  ⚠ MM 추출 실패: {e}")
        import traceback; traceback.print_exc()
        return None


# ============================================================
# CARD 6: 배부 직접비/간접비
# ============================================================
def extract_card6(projects, card1_total=None, sales_net=None, op_income=None):
    """
    카드 4 (배부 현황):
    - 직접비 = 카드 1의 부서비 총액 (사업부 통제 가능 비용)
    - 간접비 배부 = (매출 - 영업이익) - 직접비 = 손익에서 빠진 모든 비용 - 직접비
    - 프로젝트별 분해는 배부결과 파일의 raw 비율을 그대로 활용해서 안분
    """
    if not projects: return None
    
    # 1. 손익 정합성 기반 총액 계산 (제공된 경우)
    if card1_total is not None and sales_net is not None and op_income is not None:
        direct_total = max(0, card1_total)
        total_cost = sales_net - op_income
        indirect_total = max(0, total_cost - direct_total)
    else:
        direct_total = None
        indirect_total = None
    
    # 2. 배부결과 파일 raw 데이터로 프로젝트별 비율 계산
    by_proj_raw = {}
    raw_totals = {'cd_raw':0,'ci_raw':0,'pd_raw':0,'pi_raw':0}
    for lb,fp in [('01',ALLOC_01),('02',ALLOC_02)]:
        d=pd.read_excel(fp,sheet_name="직접비",header=0)
        ind=pd.read_excel(fp,sheet_name="간접비",header=0)
        for pj in projects:
            df=d[d['프로젝트(WBS)명']==pj]['재전기금액'].sum()
            ii=ind[ind['프로젝트(WBS)명']==pj]['배부금액'].sum()
            if pj not in by_proj_raw: by_proj_raw[pj]={'cd':0,'ci':0,'pd':0,'pi':0}
            if lb=='02':
                by_proj_raw[pj]['cd']=df; by_proj_raw[pj]['ci']=ii
                raw_totals['cd_raw']+=df; raw_totals['ci_raw']+=ii
            else:
                by_proj_raw[pj]['pd']=df; by_proj_raw[pj]['pi']=ii
                raw_totals['pd_raw']+=df; raw_totals['pi_raw']+=ii
    
    # 3. 손익 정합성 적용: 프로젝트별 raw → 보정 비율 → 손익 기준 총액에 안분
    proj_list = []
    for pj, v in by_proj_raw.items():
        if v['cd']==0 and v['ci']==0 and v['pd']==0 and v['pi']==0: continue
        # raw → 보정값으로 변환
        if direct_total is not None and raw_totals['cd_raw'] > 0:
            cd_adj = v['cd'] / raw_totals['cd_raw'] * direct_total
        else:
            cd_adj = v['cd']
        if indirect_total is not None and raw_totals['ci_raw'] > 0:
            ci_adj = v['ci'] / raw_totals['ci_raw'] * indirect_total
        else:
            ci_adj = v['ci']
        # 전월도 동일 보정 (전월 직접비/간접비 총액은 raw 그대로 — 비교만 가능)
        proj_list.append({
            'name':pj,
            'cd':cd_adj,'ci':ci_adj,
            'pd':v['pd'],'pi':v['pi'],
            'dd':cd_adj-v['pd'],'di':ci_adj-v['pi'],
        })
    proj_list.sort(key=lambda x:-(x['cd']+x['ci']))
    
    return {
        'cd': direct_total if direct_total is not None else raw_totals['cd_raw'],
        'ci': indirect_total if indirect_total is not None else raw_totals['ci_raw'],
        'pd': raw_totals['pd_raw'],
        'pi': raw_totals['pi_raw'],
        'projects': proj_list,
        'raw_cd': raw_totals['cd_raw'],
        'raw_ci': raw_totals['ci_raw'],
        'adjusted': direct_total is not None,
    }


# ============================================================
# FORMATTERS
# ============================================================
def fmt(v):
    b=v/1e8
    if abs(b)>=10: return f"{b:,.0f}억"
    if abs(b)>=1: return f"{b:,.1f}억"
    if abs(v)>=1e4: return f"{v/1e4:,.0f}만"
    if abs(v) > 0: return f"{v:,.0f}원"
    return "0"

def fmtd(v):
    s="+" if v>0 else ""
    b=v/1e8
    if abs(b)>=1: return f"{s}{b:,.1f}억"
    return f"{s}{v/1e4:,.0f}만"

def make_big_trend_chart(months, sales_series, profit_series, width=350, height=220, y_range=None, header_label="최근 6개월"):
    """카드 2 상단용 큰 6개월 추이 차트 (매출 + 영업이익 2 라인)
    y_range: (y_min, y_max) 튜플로 Y축 고정. None이면 시리즈 자체에서 산출"""
    if not sales_series or not profit_series or len(sales_series) < 2:
        return ""
    
    # 월 라벨: 첫 달이나 1월인 경우 "YY.M월", 나머지는 "M월"
    def month_label(m, is_first):
        y = (m // 100) % 100
        mm = m % 100
        if is_first or mm == 1:
            return f"{y}.{mm}월"
        return f"{mm}월"
    labels = [month_label(m, i == 0) for i, m in enumerate(months)]
    
    # Y축: 외부 고정값 우선, 없으면 시리즈에서 산출
    if y_range is not None:
        y_min, y_max = y_range
        y_rng = y_max - y_min if y_max != y_min else 1
    else:
        all_vals = sales_series + profit_series
        vmin = min(all_vals); vmax = max(all_vals)
        if vmin == vmax: vmax = vmin + 1
        rng = vmax - vmin
        pad_top = rng * 0.25
        pad_bot = rng * 0.20
        y_min = vmin - pad_bot
        y_max = vmax + pad_top
        y_rng = y_max - y_min
    
    # 영역 (라벨 공간 확보)
    pad_l, pad_r, pad_t, pad_b = 12, 12, 32, 32
    inner_w = width - pad_l - pad_r
    inner_h = height - pad_t - pad_b
    
    def to_xy(i, v):
        x = pad_l + (i / (len(sales_series)-1)) * inner_w
        y = pad_t + (1 - (v - y_min) / y_rng) * inner_h
        return x, y
    
    # 0 기준선 위치
    zero_y = pad_t + (1 - (0 - y_min) / y_rng) * inner_h if y_min < 0 < y_max else None
    
    # 시리즈 path 생성
    def series_path(values):
        pts = [to_xy(i, v) for i, v in enumerate(values)]
        return pts, "M " + " L ".join(f"{x:.1f},{y:.1f}" for x,y in pts)
    
    sales_pts, sales_path = series_path(sales_series)
    profit_pts, profit_path = series_path(profit_series)
    
    # 매출 fill area
    sales_fill = sales_path + f" L {sales_pts[-1][0]:.1f},{height-pad_b:.1f} L {sales_pts[0][0]:.1f},{height-pad_b:.1f} Z"
    
    # 색상
    sales_color = "#3182F6"
    NEG_COLOR = "#F04452"
    POS_COLOR = "#00C471"
    # 영업이익 라인 기본 색상 (트렌드 기준)
    profit_color = POS_COLOR
    early_profit = sum(profit_series[:3])/3
    late_profit = sum(profit_series[-3:])/3
    base = max(abs(early_profit), abs(late_profit), 1)
    if (late_profit - early_profit) / base < -0.10:
        profit_color = NEG_COLOR
    # 영업이익 점/라벨은 부호별 색상 (음수 = 빨강)
    def pt_color(v): return NEG_COLOR if v < 0 else POS_COLOR
    
    def fmt_short(v):
        b = v/1e8
        if abs(b) >= 10: return f"{b:.0f}억"
        if abs(b) >= 1: return f"{b:.1f}억"
        if abs(v) >= 1e4: return f"{v/1e4:.0f}만"
        return "0"
    
    # x축 라벨 (월)
    x_labels = ""
    for i, lb in enumerate(labels):
        x = pad_l + (i / (len(labels)-1)) * inner_w
        fw = "700" if ("." in lb) else "500"  # 연도 붙은 라벨 강조
        fc = "#4E5968" if ("." in lb) else "#8B95A1"
        x_labels += f'<text x="{x:.1f}" y="{height-pad_b+18}" font-size="10" font-weight="{fw}" fill="{fc}" text-anchor="middle" font-family="Pretendard,sans-serif">{lb}</text>'
    
    # 모든 포인트에 금액 라벨 (매출은 점 위, 영업이익은 점 아래 — 겹침 회피)
    # 단, 두 점이 너무 가까우면 위치 조정
    value_labels = ""
    for i in range(len(sales_series)):
        sx, sy = sales_pts[i]
        px, py = profit_pts[i]
        sv = sales_series[i]
        pv = profit_series[i]
        
        # 기본: 매출 위, 영업이익 아래
        sales_ly = sy - 7
        profit_ly = py + 14
        # 두 점 사이 간격 검사
        if abs(sy - py) < 20:
            # 가까우면 매출은 더 위, 영업이익은 더 아래
            if sy <= py:
                sales_ly = sy - 8
                profit_ly = py + 15
            else:
                sales_ly = sy + 15
                profit_ly = py - 8
        
        value_labels += f'<text x="{sx:.1f}" y="{sales_ly:.1f}" font-size="9" font-weight="700" fill="{sales_color}" text-anchor="middle" font-family="Pretendard,sans-serif">{fmt_short(sv)}</text>'
        value_labels += f'<text x="{px:.1f}" y="{profit_ly:.1f}" font-size="9" font-weight="700" fill="{pt_color(pv)}" text-anchor="middle" font-family="Pretendard,sans-serif">{fmt_short(pv)}</text>'
    
    # 0선
    zero_line = ""
    if zero_y is not None:
        zero_line = f'<line x1="{pad_l}" y1="{zero_y:.1f}" x2="{width-pad_r}" y2="{zero_y:.1f}" stroke="#E5E8EB" stroke-width="0.5" stroke-dasharray="2,2"/>'
    
    return f'''<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" style="display:block">
<text x="{pad_l}" y="14" font-size="11" font-weight="600" fill="#3182F6" font-family="Pretendard,sans-serif">● 매출</text>
<text x="{pad_l + 78}" y="14" font-size="11" font-weight="600" fill="{profit_color}" font-family="Pretendard,sans-serif">● 영업이익</text>
<text x="{width-pad_r}" y="14" font-size="10" fill="#8B95A1" text-anchor="end" font-family="Pretendard,sans-serif">{header_label}</text>
{zero_line}
<path d="{sales_fill}" fill="{sales_color}" fill-opacity="0.08" stroke="none"/>
<path d="{sales_path}" fill="none" stroke="{sales_color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
<path d="{profit_path}" fill="none" stroke="{profit_color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
{"".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="2.5" fill="{sales_color}"/>' for x,y in sales_pts)}
{"".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="2.5" fill="{pt_color(profit_series[i])}"/>' for i,(x,y) in enumerate(profit_pts))}
{value_labels}
{x_labels}
</svg>'''


_SLIDER_SEQ = [0]
def make_trend_slider(months_full, sales_full, profit_full, window=6):
    """25.1~현재월 풀데이터를 받아 6개월 윈도우 슬라이더 HTML 생성.
    Y축은 전체 데이터 기준으로 고정."""
    if not months_full or len(months_full) < 2:
        return ""
    n = len(months_full)
    # 윈도우들: 가장 오래된 [0:6], [1:7], ..., [n-6:n] (n<6이면 단일 윈도우)
    if n <= window:
        windows = [(0, n)]
    else:
        windows = [(i, i + window) for i in range(n - window + 1)]
    # 마지막(최근) 윈도우가 기본 표시
    default_idx = len(windows) - 1
    # Y축 고정: 전체 데이터 min/max + 패딩
    all_vals = sales_full + profit_full
    vmin, vmax = min(all_vals), max(all_vals)
    if vmin == vmax: vmax = vmin + 1
    rng = vmax - vmin
    y_min = vmin - rng * 0.20
    y_max = vmax + rng * 0.25
    y_range = (y_min, y_max)

    _SLIDER_SEQ[0] += 1
    sid = f"trsl{_SLIDER_SEQ[0]}"

    # 각 윈도우별 차트 SVG 생성
    slides_html = []
    labels = []
    for wi, (s, e) in enumerate(windows):
        m_w = months_full[s:e]
        sa_w = sales_full[s:e]
        pr_w = profit_full[s:e]
        # 헤더에 윈도우 범위 표시
        first_m = m_w[0]; last_m = m_w[-1]
        hdr = f"{(first_m//100)%100}.{first_m%100}월 ~ {(last_m//100)%100}.{last_m%100}월"
        chart = make_big_trend_chart(m_w, sa_w, pr_w, y_range=y_range, header_label=hdr)
        active = " trsl-active" if wi == default_idx else ""
        slides_html.append(f'<div class="trsl-slide{active}" data-idx="{wi}">{chart}</div>')
        labels.append(hdr)

    dots = "".join(
        f'<span class="trsl-dot{(" trsl-dot-active" if i == default_idx else "")}" data-idx="{i}"></span>'
        for i in range(len(windows))
    )

    # 인라인 CSS/JS — 슬라이더 단일 인스턴스 기준으로 격리
    return f'''
<div class="trsl-wrap" id="{sid}" data-default="{default_idx}" data-count="{len(windows)}">
  <div class="trsl-stage">
    {''.join(slides_html)}
  </div>
  <div class="trsl-ctrl">
    <button type="button" class="trsl-btn trsl-prev" aria-label="이전 6개월">‹</button>
    <div class="trsl-dots">{dots}</div>
    <button type="button" class="trsl-btn trsl-next" aria-label="다음 6개월">›</button>
  </div>
</div>
<style>
#{sid} {{ position:relative; }}
#{sid} .trsl-stage {{ position:relative; }}
#{sid} .trsl-slide {{ display:none; }}
#{sid} .trsl-slide.trsl-active {{ display:block; }}
#{sid} .trsl-ctrl {{ display:flex; align-items:center; justify-content:center; gap:10px; margin-top:6px; }}
#{sid} .trsl-btn {{ width:28px; height:28px; border-radius:50%; border:1px solid #E5E8EB; background:#fff; color:#4E5968; font-size:16px; line-height:1; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:0; }}
#{sid} .trsl-btn:disabled {{ opacity:0.35; cursor:default; }}
#{sid} .trsl-btn:not(:disabled):hover {{ background:#F2F4F6; }}
#{sid} .trsl-dots {{ display:flex; gap:5px; }}
#{sid} .trsl-dot {{ width:6px; height:6px; border-radius:50%; background:#D1D6DB; }}
#{sid} .trsl-dot.trsl-dot-active {{ background:#3182F6; width:18px; border-radius:3px; }}
</style>
<script>
(function(){{
  var root=document.getElementById('{sid}');
  if(!root) return;
  var slides=root.querySelectorAll('.trsl-slide');
  var dots=root.querySelectorAll('.trsl-dot');
  var prev=root.querySelector('.trsl-prev');
  var next=root.querySelector('.trsl-next');
  var n=parseInt(root.getAttribute('data-count'),10);
  var idx=parseInt(root.getAttribute('data-default'),10);
  function show(i){{
    if(i<0||i>=n) return;
    idx=i;
    slides.forEach(function(s){{ s.classList.toggle('trsl-active', parseInt(s.getAttribute('data-idx'),10)===i); }});
    dots.forEach(function(d){{ d.classList.toggle('trsl-dot-active', parseInt(d.getAttribute('data-idx'),10)===i); }});
    prev.disabled=(i===0);
    next.disabled=(i===n-1);
  }}
  prev.addEventListener('click',function(){{ show(idx-1); }});
  next.addEventListener('click',function(){{ show(idx+1); }});
  dots.forEach(function(d){{
    d.addEventListener('click',function(){{ show(parseInt(d.getAttribute('data-idx'),10)); }});
  }});
  show(idx);
}})();
</script>
'''


def make_op_only_chart(months, op_series, width=350, height=200, y_range=None, header_label=""):
    """영업이익 단일 라인 차트 (매출 없는 신작 준비중 조직용). 음수 영역만 표시되는 경우가 많음"""
    if not op_series or len(op_series) < 2:
        return ""
    
    def month_label(m, is_first):
        y = (m // 100) % 100
        mm = m % 100
        if is_first or mm == 1:
            return f"{y}.{mm}월"
        return f"{mm}월"
    labels = [month_label(m, i == 0) for i, m in enumerate(months)]
    
    if y_range is not None:
        y_min, y_max = y_range
    else:
        all_vals = op_series + [0]
        vmin = min(all_vals); vmax = max(all_vals)
        if vmin == vmax: vmax = vmin + 1
        rng = vmax - vmin
        y_min = vmin - rng*0.20
        y_max = vmax + rng*0.25
    # BEP(0) 가시성 보장: 0이 차트 영역 밖이면 강제로 포함
    if y_max < 0:
        # 영업이익 전부 음수 → 위쪽 패딩 늘려서 0 보이게
        y_max = abs(y_min) * 0.10
    if y_min > 0:
        y_min = -abs(y_max) * 0.10
    y_rng = y_max - y_min if y_max != y_min else 1
    
    pad_l, pad_r, pad_t, pad_b = 12, 12, 32, 32
    inner_w = width - pad_l - pad_r
    inner_h = height - pad_t - pad_b
    
    def to_xy(i, v):
        x = pad_l + (i / (len(op_series)-1)) * inner_w
        y = pad_t + (1 - (v - y_min) / y_rng) * inner_h
        return x, y
    
    pts = [to_xy(i, v) for i, v in enumerate(op_series)]
    path = "M " + " L ".join(f"{x:.1f},{y:.1f}" for x,y in pts)
    
    NEG = "#F04452"; POS = "#00C471"
    # 트렌드 색상 (전반부 vs 후반부 비교)
    early = sum(op_series[:3])/3
    late = sum(op_series[-3:])/3
    base = max(abs(early), abs(late), 1)
    line_color = NEG if (late - early)/base < -0.10 else POS
    
    def pt_color(v): return NEG if v < 0 else POS
    
    def fmt_short(v):
        b = v/1e8
        if abs(b) >= 10: return f"{b:.0f}억"
        if abs(b) >= 1: return f"{b:.1f}억"
        if abs(v) >= 1e4: return f"{v/1e4:.0f}만"
        return "0"
    
    # BEP(손익분기점, 영업이익 = 0) 가로 점선 — 강조 표시
    zero_y = pad_t + (1 - (0 - y_min) / y_rng) * inner_h
    bep_line = ""
    if pad_t <= zero_y <= height - pad_b:
        bep_line = (
            f'<line x1="{pad_l}" y1="{zero_y:.1f}" x2="{width-pad_r}" y2="{zero_y:.1f}" '
            f'stroke="#8B95A1" stroke-width="1" stroke-dasharray="4,3"/>'
            f'<text x="{width-pad_r-2}" y="{zero_y-3:.1f}" font-size="9" font-weight="700" '
            f'fill="#8B95A1" text-anchor="end" font-family="Pretendard,sans-serif">BEP</text>'
        )
    
    # x축 라벨
    x_labels = ""
    for i, lb in enumerate(labels):
        x = pad_l + (i / (len(labels)-1)) * inner_w
        fw = "700" if ("." in lb) else "500"
        fc = "#4E5968" if ("." in lb) else "#8B95A1"
        x_labels += f'<text x="{x:.1f}" y="{height-pad_b+18}" font-size="10" font-weight="{fw}" fill="{fc}" text-anchor="middle" font-family="Pretendard,sans-serif">{lb}</text>'
    
    # 값 라벨
    value_labels = ""
    for i, (x, y) in enumerate(pts):
        v = op_series[i]
        # 음수면 점 위, 양수면 점 아래 (라인과 안 겹치게)
        ly = y - 7 if v < 0 else y + 14
        value_labels += f'<text x="{x:.1f}" y="{ly:.1f}" font-size="9" font-weight="700" fill="{pt_color(v)}" text-anchor="middle" font-family="Pretendard,sans-serif">{fmt_short(v)}</text>'
    
    return f'''<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" style="display:block">
<text x="{pad_l}" y="14" font-size="11" font-weight="600" fill="{line_color}" font-family="Pretendard,sans-serif">● 영업이익 (매출 없음)</text>
<text x="{width-pad_r}" y="14" font-size="10" fill="#8B95A1" text-anchor="end" font-family="Pretendard,sans-serif">{header_label}</text>
{bep_line}
<path d="{path}" fill="none" stroke="{line_color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
{"".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="2.5" fill="{pt_color(op_series[i])}"/>' for i,(x,y) in enumerate(pts))}
{value_labels}
{x_labels}
</svg>'''


def make_op_only_slider(months_full, op_full, window=6):
    """영업이익 단일 라인 슬라이더 (카드 1 매출 없는 조직용)"""
    if not months_full or len(months_full) < 2:
        return ""
    n = len(months_full)
    if n <= window:
        windows = [(0, n)]
    else:
        windows = [(i, i + window) for i in range(n - window + 1)]
    default_idx = len(windows) - 1
    all_vals = op_full + [0]
    vmin, vmax = min(all_vals), max(all_vals)
    if vmin == vmax: vmax = vmin + 1
    rng = vmax - vmin
    y_min = vmin - rng*0.20
    y_max = vmax + rng*0.25
    y_range = (y_min, y_max)
    
    _SLIDER_SEQ[0] += 1
    sid = f"opsl{_SLIDER_SEQ[0]}"
    
    slides_html = []
    for wi, (s, e) in enumerate(windows):
        m_w = months_full[s:e]
        o_w = op_full[s:e]
        first_m = m_w[0]; last_m = m_w[-1]
        hdr = f"{(first_m//100)%100}.{first_m%100}월 ~ {(last_m//100)%100}.{last_m%100}월"
        chart = make_op_only_chart(m_w, o_w, y_range=y_range, header_label=hdr)
        active = " trsl-active" if wi == default_idx else ""
        slides_html.append(f'<div class="trsl-slide{active}" data-idx="{wi}">{chart}</div>')
    
    dots = "".join(
        f'<span class="trsl-dot{(" trsl-dot-active" if i == default_idx else "")}" data-idx="{i}"></span>'
        for i in range(len(windows))
    )
    
    return f'''
<div class="trsl-wrap" id="{sid}" data-default="{default_idx}" data-count="{len(windows)}">
  <div class="trsl-stage">{''.join(slides_html)}</div>
  <div class="trsl-ctrl">
    <button type="button" class="trsl-btn trsl-prev" aria-label="이전">‹</button>
    <div class="trsl-dots">{dots}</div>
    <button type="button" class="trsl-btn trsl-next" aria-label="다음">›</button>
  </div>
</div>
<style>
#{sid} {{ position:relative; }}
#{sid} .trsl-stage {{ position:relative; }}
#{sid} .trsl-slide {{ display:none; }}
#{sid} .trsl-slide.trsl-active {{ display:block; }}
#{sid} .trsl-ctrl {{ display:flex; align-items:center; justify-content:center; gap:10px; margin-top:6px; }}
#{sid} .trsl-btn {{ width:28px; height:28px; border-radius:50%; border:1px solid #E5E8EB; background:#fff; color:#4E5968; font-size:16px; line-height:1; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:0; }}
#{sid} .trsl-btn:disabled {{ opacity:0.35; cursor:default; }}
#{sid} .trsl-dots {{ display:flex; gap:5px; }}
#{sid} .trsl-dot {{ width:6px; height:6px; border-radius:50%; background:#D1D6DB; }}
#{sid} .trsl-dot.trsl-dot-active {{ background:#3182F6; width:18px; border-radius:3px; }}
</style>
<script>
(function(){{
  var root=document.getElementById('{sid}');
  if(!root) return;
  var slides=root.querySelectorAll('.trsl-slide');
  var dots=root.querySelectorAll('.trsl-dot');
  var prev=root.querySelector('.trsl-prev');
  var next=root.querySelector('.trsl-next');
  var n=parseInt(root.getAttribute('data-count'),10);
  var idx=parseInt(root.getAttribute('data-default'),10);
  function show(i){{
    if(i<0||i>=n) return;
    idx=i;
    slides.forEach(function(s){{ s.classList.toggle('trsl-active', parseInt(s.getAttribute('data-idx'),10)===i); }});
    dots.forEach(function(d){{ d.classList.toggle('trsl-dot-active', parseInt(d.getAttribute('data-idx'),10)===i); }});
    prev.disabled=(i===0);
    next.disabled=(i===n-1);
  }}
  prev.addEventListener('click',function(){{ show(idx-1); }});
  next.addEventListener('click',function(){{ show(idx+1); }});
  dots.forEach(function(d){{
    d.addEventListener('click',function(){{ show(parseInt(d.getAttribute('data-idx'),10)); }});
  }});
  show(idx);
}})();
</script>
'''



def make_mm_trend_chart(months, total_series, inside_series, width=350, height=180, y_range=None, header_label=""):
    """MM 추이 차트 (총 MM 1 라인). inside_series는 호환성 위해 받지만 표시 안 함"""
    if not total_series or len(total_series) < 2:
        return ""
    
    def month_label(m, is_first):
        y = (m // 100) % 100
        mm = m % 100
        if is_first or mm == 1:
            return f"{y}.{mm}월"
        return f"{mm}월"
    labels = [month_label(m, i == 0) for i, m in enumerate(months)]
    
    if y_range is not None:
        y_min, y_max = y_range
    else:
        all_vals = total_series + [0]
        vmin = min(all_vals); vmax = max(all_vals)
        if vmin == vmax: vmax = vmin + 1
        rng = vmax - vmin
        y_min = max(0, vmin - rng*0.15)
        y_max = vmax + rng*0.30
    y_rng = y_max - y_min if y_max != y_min else 1
    
    pad_l, pad_r, pad_t, pad_b = 12, 12, 32, 32
    inner_w = width - pad_l - pad_r
    inner_h = height - pad_t - pad_b
    
    def to_xy(i, v):
        x = pad_l + (i / (len(total_series)-1)) * inner_w
        y = pad_t + (1 - (v - y_min) / y_rng) * inner_h
        return x, y
    
    def series_path(values):
        pts = [to_xy(i, v) for i, v in enumerate(values)]
        return pts, "M " + " L ".join(f"{x:.1f},{y:.1f}" for x,y in pts)
    
    total_pts, total_path = series_path(total_series)
    total_fill = total_path + f" L {total_pts[-1][0]:.1f},{height-pad_b:.1f} L {total_pts[0][0]:.1f},{height-pad_b:.1f} Z"
    
    TOTAL_COLOR = "#3182F6"
    
    def fmt_mm(v):
        if v == 0: return "0"
        return f"{v:.1f}"
    
    # x축 라벨
    x_labels = ""
    for i, lb in enumerate(labels):
        x = pad_l + (i / (len(labels)-1)) * inner_w
        fw = "700" if ("." in lb) else "500"
        fc = "#4E5968" if ("." in lb) else "#8B95A1"
        x_labels += f'<text x="{x:.1f}" y="{height-pad_b+18}" font-size="10" font-weight="{fw}" fill="{fc}" text-anchor="middle" font-family="Pretendard,sans-serif">{lb}</text>'
    
    # 값 라벨
    value_labels = ""
    for i in range(len(total_series)):
        tx, ty = total_pts[i]
        tv = total_series[i]
        if tv > 0:
            value_labels += f'<text x="{tx:.1f}" y="{ty-7:.1f}" font-size="9" font-weight="700" fill="{TOTAL_COLOR}" text-anchor="middle" font-family="Pretendard,sans-serif">{fmt_mm(tv)}</text>'
    
    return f'''<svg width="100%" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" style="display:block">
<text x="{pad_l}" y="14" font-size="11" font-weight="600" fill="{TOTAL_COLOR}" font-family="Pretendard,sans-serif">● 총 MM</text>
<text x="{width-pad_r}" y="14" font-size="10" fill="#8B95A1" text-anchor="end" font-family="Pretendard,sans-serif">{header_label}</text>
<path d="{total_fill}" fill="{TOTAL_COLOR}" fill-opacity="0.08" stroke="none"/>
<path d="{total_path}" fill="none" stroke="{TOTAL_COLOR}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
{"".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="2.5" fill="{TOTAL_COLOR}"/>' for x,y in total_pts)}
{value_labels}
{x_labels}
</svg>'''


def make_mm_trend_slider(months_full, total_full, inside_full, window=6):
    """MM 시계열 슬라이더 (카드 4 프로젝트 블록용)"""
    if not months_full or len(months_full) < 2:
        return ""
    n = len(months_full)
    if n <= window:
        windows = [(0, n)]
    else:
        windows = [(i, i + window) for i in range(n - window + 1)]
    default_idx = len(windows) - 1
    all_vals = total_full + [0]
    vmin, vmax = min(all_vals), max(all_vals)
    if vmin == vmax: vmax = vmin + 1
    rng = vmax - vmin
    y_min = max(0, vmin - rng*0.15)
    y_max = vmax + rng*0.30
    y_range = (y_min, y_max)
    
    _SLIDER_SEQ[0] += 1
    sid = f"mmsl{_SLIDER_SEQ[0]}"
    
    slides_html = []
    for wi, (s, e) in enumerate(windows):
        m_w = months_full[s:e]
        t_w = total_full[s:e]
        i_w = inside_full[s:e]
        first_m = m_w[0]; last_m = m_w[-1]
        hdr = f"{(first_m//100)%100}.{first_m%100}월 ~ {(last_m//100)%100}.{last_m%100}월"
        chart = make_mm_trend_chart(m_w, t_w, i_w, y_range=y_range, header_label=hdr)
        active = " trsl-active" if wi == default_idx else ""
        slides_html.append(f'<div class="trsl-slide{active}" data-idx="{wi}">{chart}</div>')
    
    dots = "".join(
        f'<span class="trsl-dot{(" trsl-dot-active" if i == default_idx else "")}" data-idx="{i}"></span>'
        for i in range(len(windows))
    )
    
    return f'''
<div class="trsl-wrap" id="{sid}" data-default="{default_idx}" data-count="{len(windows)}">
  <div class="trsl-stage">{''.join(slides_html)}</div>
  <div class="trsl-ctrl">
    <button type="button" class="trsl-btn trsl-prev" aria-label="이전">‹</button>
    <div class="trsl-dots">{dots}</div>
    <button type="button" class="trsl-btn trsl-next" aria-label="다음">›</button>
  </div>
</div>
<style>
#{sid} {{ position:relative; margin-top:10px; padding:8px; background:#F8F9FA; border-radius:10px; }}
#{sid} .trsl-stage {{ position:relative; }}
#{sid} .trsl-slide {{ display:none; }}
#{sid} .trsl-slide.trsl-active {{ display:block; }}
#{sid} .trsl-ctrl {{ display:flex; align-items:center; justify-content:center; gap:8px; margin-top:4px; }}
#{sid} .trsl-btn {{ width:24px; height:24px; border-radius:50%; border:1px solid #E5E8EB; background:#fff; color:#4E5968; font-size:14px; line-height:1; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:0; }}
#{sid} .trsl-btn:disabled {{ opacity:0.35; cursor:default; }}
#{sid} .trsl-dots {{ display:flex; gap:4px; }}
#{sid} .trsl-dot {{ width:5px; height:5px; border-radius:50%; background:#D1D6DB; }}
#{sid} .trsl-dot.trsl-dot-active {{ background:#8B5CF6; width:14px; border-radius:3px; }}
</style>
<script>
(function(){{
  var root=document.getElementById('{sid}');
  if(!root) return;
  var slides=root.querySelectorAll('.trsl-slide');
  var dots=root.querySelectorAll('.trsl-dot');
  var prev=root.querySelector('.trsl-prev');
  var next=root.querySelector('.trsl-next');
  var n=parseInt(root.getAttribute('data-count'),10);
  var idx=parseInt(root.getAttribute('data-default'),10);
  function show(i){{
    if(i<0||i>=n) return;
    idx=i;
    slides.forEach(function(s){{ s.classList.toggle('trsl-active', parseInt(s.getAttribute('data-idx'),10)===i); }});
    dots.forEach(function(d){{ d.classList.toggle('trsl-dot-active', parseInt(d.getAttribute('data-idx'),10)===i); }});
    prev.disabled=(i===0);
    next.disabled=(i===n-1);
  }}
  prev.addEventListener('click',function(){{ show(idx-1); }});
  next.addEventListener('click',function(){{ show(idx+1); }});
  dots.forEach(function(d){{
    d.addEventListener('click',function(){{ show(parseInt(d.getAttribute('data-idx'),10)); }});
  }});
  show(idx);
}})();
</script>
'''


def pcol(v): return "#00C471" if v>0 else "#F04452" if v<0 else "#8B95A1"
def ncol(v): return "#F04452" if v>0 else "#00C471" if v<0 else "#8B95A1"


def make_sparkline(values, width=70, height=24, profit_metric=False):
    """6개월 추이 sparkline SVG 생성
    profit_metric: True면 영업이익 계열(상승=좋음), False면 매출 계열도 동일 처리"""
    if not values or len(values) < 2:
        return ""
    
    vmin = min(values)
    vmax = max(values)
    rng = vmax - vmin if vmax != vmin else 1
    
    # 추세 판정: 최근 3개월 평균 vs 처음 3개월 평균
    early = sum(values[:3]) / 3
    late = sum(values[-3:]) / 3
    
    if abs(early) < 1e6 and abs(late) < 1e6:
        trend = 'flat'
    else:
        base = max(abs(early), abs(late), 1)
        change_ratio = (late - early) / base
        if change_ratio > 0.10:
            trend = 'up'
        elif change_ratio < -0.10:
            trend = 'down'
        else:
            trend = 'flat'
    
    color_map = {
        'up':   '#00C471',
        'down': '#F04452',
        'flat': '#B0B8C1',
    }
    color = color_map[trend]
    
    # 좌표 변환
    pad_x, pad_y = 2, 3
    inner_w = width - pad_x*2
    inner_h = height - pad_y*2
    points = []
    for i, v in enumerate(values):
        x = pad_x + (i / (len(values)-1)) * inner_w
        y = pad_y + (1 - (v - vmin) / rng) * inner_h
        points.append((x, y))
    
    path = "M " + " L ".join(f"{x:.1f},{y:.1f}" for x,y in points)
    
    # 채움 영역 (path 아래)
    fill_path = path + f" L {points[-1][0]:.1f},{height-pad_y} L {points[0][0]:.1f},{height-pad_y} Z"
    
    # 마지막 점 하이라이트
    last_x, last_y = points[-1]
    
    return f'''<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" style="display:block">
<path d="{fill_path}" fill="{color}" fill-opacity="0.10" stroke="none"/>
<path d="{path}" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
<circle cx="{last_x:.1f}" cy="{last_y:.1f}" r="2" fill="{color}"/>
</svg>'''


def trend_label(values, is_profit=True):
    """추세 텍스트 라벨 (영업이익 우하향만 경고)"""
    if not values or len(values) < 6:
        return ""
    early = sum(values[:3]) / 3
    late = sum(values[-3:]) / 3
    base = max(abs(early), abs(late), 1)
    change_ratio = (late - early) / base
    if is_profit and change_ratio < -0.10:
        return '<span class="spark-warn">⚠ 6개월 하락</span>'
    return ""


# ============================================================
# HTML GENERATOR
# ============================================================
def gen_html(div_cfg, c2, c3, c4, c6, c5):
    DIVISION = div_cfg['name']
    SUB = div_cfg['sub']
    has_pnl = c3 is not None
    has_mm = c5 is not None and c5['curr'] > 0
    
    # 카드 개수 계산
    n_cards = 1  # 카드 1: 부서비용
    if has_pnl: n_cards += 3  # 손익, 프로젝트별, 배부
    
    # ─── 카드1 - 부서비용 행
    dept_rows=""
    for a in c2['da']:
        nm=a['name'].split('. ')[1] if '. ' in a['name'] else a['name']
        pc=(a['curr']/c2['dt']['c']*100) if c2['dt']['c'] else 0
        maxv = c2['da'][0]['curr'] if c2['da'] else 1
        bw = (a['curr']/maxv*100) if maxv else 0
        dept_rows+=f"""
        <div class="bar-row">
          <div class="bar-info"><span class="bar-name">{nm}</span><span class="bar-val">{fmt(a['curr'])}<span class="bar-pct">{pc:.0f}%</span></span></div>
          <div class="bar-track"><div class="bar-fill" style="width:{bw}%"></div></div>
          <div class="bar-delta" style="color:{ncol(a['delta'])}">{fmtd(a['delta'])}</div>
        </div>"""
    
    proj_rows=""
    for a in c2['pa']:
        nm=a['name'].split('. ')[1] if '. ' in a['name'] else a['name']
        proj_rows+=f"""
        <div class="list-row">
          <span class="list-name">{nm}</span>
          <span class="list-val">{fmt(a['curr'])}</span>
          <span class="list-delta" style="color:{ncol(a['delta'])}">{fmtd(a['delta'])}</span>
        </div>"""
    
    proj_section = ""
    if c2['pa']:
        proj_section = f"""
          <div class="sec-div"><span>직접 프로젝트 비용</span><span>{fmt(c2['pt']['c'])}</span></div>
          {proj_rows}"""
    
    # 기타 상세
    other_section = ""
    has_other = any('기타' in a['name'] and a['curr']>0 for a in c2['pa'])
    if has_other and c2.get('other_detail'):
        rows = ""
        for d in c2['other_detail'][:5]:
            wbs = d['wbs'].replace('(P)_','·').replace('_운영','').replace('_',' ')
            rows += f"""
        <div class="other-row">
          <div class="other-top">
            <span class="other-wbs">{wbs}</span>
            <span class="other-val">{fmt(d['val'])}</span>
          </div>
          <div class="other-acct">{d['acct']}</div>
        </div>"""
        other_section = f"""
        <div class="other-box">
          <div class="other-title">💡 '기타' 상세 (원장 기준)</div>
          {rows}
        </div>"""
    
    # 마케팅비 상세 (원장 기준)
    # 손익 있는 실: WBS 기준 (담당 프로젝트에 들어간 마케팅비 전체)
    # 손익 없는 실(마케팅실 등): 코스트센터 기준
    mkt_section = ""
    if c2.get('mkt_detail') and c2['mkt_detail'].get('items'):
        md = c2['mkt_detail']
        is_wbs_based = 'cc' in md['items'][0]  # WBS 기준이면 cc 필드 있음
        rows = ""
        for d in md['items'][:6]:
            wbs = d['wbs'].replace('(P)_','·').replace('_운영','').replace('_',' ')
            bw = d['pct']
            cc_label = ""
            if is_wbs_based and d.get('cc'):
                cc_label = f'<span class="mkt-cc">{d["cc"]}</span>'
            rows += f"""
        <div class="mkt-row">
          <div class="mkt-top">
            <span class="mkt-wbs">{wbs}{cc_label}</span>
            <span class="mkt-val">{fmt(d['val'])}<span class="mkt-pct">{d['pct']:.0f}%</span></span>
          </div>
          <div class="mkt-bar"><div class="mkt-fill" style="width:{bw}%"></div></div>
        </div>"""
        title_label = "📣 마케팅비" if is_wbs_based else "📣 마케팅비"
        mkt_section = f"""
        <div class="mkt-box">
          <div class="mkt-title">{title_label} · 총 {fmt(md['total'])}</div>
          {rows}
        </div>"""
    
    # ─── 카드2/3/4 (손익 있는 경우만)
    cards_pnl = ""
    card_pnl1 = ""
    card_pnl2 = ""
    if has_pnl:
        trend_data = c3.get('_trend', {})
        
        # 카드1 상단 큰 차트
        big_chart = ""
        if trend_data.get('months'):
            # 정상 케이스: 매출 + 영업이익 2라인
            big_chart = make_trend_slider(
                trend_data['months'],
                trend_data.get('매출(그로스)', trend_data.get('매출(Net)', [])),
                trend_data.get('영업이익', []),
            )
        elif trend_data.get('_op_only_months') and len(trend_data['_op_only_months']) >= 2:
            # 폴백: 매출 없는 신작 준비중 조직 → 영업이익 단일 라인
            big_chart = make_op_only_slider(
                trend_data['_op_only_months'],
                trend_data['_op_only_values'],
            )
        
        def mk_metric(label, d, emoji, metric_key, is_profit=False):
            return f"""
            <div class="metric-card">
              <div class="metric-top">
                <span class="metric-emoji">{emoji}</span>
                <span class="metric-label">{label}</span>
              </div>
              <div class="metric-num">{fmt(d['a'])}</div>
              <div class="chip-row">
                <div class="chip" style="color:{pcol(d['vp'])};background:{pcol(d['vp'])}0D">계획 {fmtd(d['vp'])}</div>
                <div class="chip" style="color:{pcol(d['vm'])};background:{pcol(d['vm'])}0D">전월 {fmtd(d['vm'])}</div>
                <div class="chip" style="color:{pcol(d['vy'])};background:{pcol(d['vy'])}0D">전년 {fmtd(d['vy'])}</div>
              </div>
            </div>"""
        
        # ─── 카드 2 메트릭 그리드: 매출 그로스 → 매출연동수수료 → Net → 직접비 → 배부 → 영업이익
        sales_gross = c3['매출(그로스)']['a']
        svc_direct = c3['서비스직접비']['a']
        sales_net = c3['매출(Net)']['a']
        op = c3['영업이익']['a']
        direct = c2['gt']['c']  # 카드1 총액 그대로
        indirect = sales_net - op - direct
        
        # 카드 1에서 가장 큰 비용 항목 1개 추출 (직접비 sub 라벨로)
        # da/pa 외에 마케팅비(원장)도 후보에 포함 (main에서 별도 처리되어 pa에 안 들어감)
        all_card1_items = list(c2['da']) + list(c2['pa'])
        if c2.get('mkt_detail') and c2['mkt_detail'].get('total'):
            mkt_total = c2['mkt_detail']['total']
            # 시트의 마케팅비와 중복되지 않는 분만 (extra_mkt) 추가 후보로
            extra_mkt = c2.get('extra_mkt', mkt_total)
            if extra_mkt > 0:
                all_card1_items.append({
                    'name': '마케팅비',
                    'curr': mkt_total,
                })
        top_item_label = "카드1 총액"
        if all_card1_items:
            top_item = max(all_card1_items, key=lambda x: abs(x['curr']))
            # 계정명 클린업 ("01. 인건비성 경비" → "인건비성 경비", "(원장)" 제거)
            nm = top_item['name'].split('. ', 1)[-1].replace(' (원장)', '').strip()
            top_item_label = f"최대 {nm} {fmt(top_item['curr'])}"
        
        # 매출 그로스가 없거나 (Net과 같으면) 단순 4단으로, 있으면 6단으로
        has_gross = sales_gross > 0 and abs(sales_gross - sales_net) > 1e6
        
        # 서비스직접비 분해 (참고 라벨)
        svc_bd = c3.get('_svc_breakdown', {})
        svc_parts = []
        if svc_bd.get('R/S', 0) > 1e6: svc_parts.append(f"R/S {fmt(svc_bd['R/S'])}")
        if svc_bd.get('수수료', 0) > 1e6: svc_parts.append(f"수수료 {fmt(svc_bd['수수료'])}")
        if svc_bd.get('사내직접', 0) > 1e6: svc_parts.append(f"사내직접 {fmt(svc_bd['사내직접'])}")
        svc_sub = " · ".join(svc_parts) if svc_parts else "R/S · 수수료 · 사내직접"
        
        if has_gross:
            pnl_metric_grid = f"""
            <div class="pnl-flow">
              <div class="pnl-row pnl-sales">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">💰</span>
                  <span class="pnl-label">매출 (그로스)</span>
                </div>
                <div class="pnl-num">{fmt(sales_gross)}</div>
              </div>
              
              <div class="pnl-minus">−</div>
              
              <div class="pnl-row pnl-cost">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">💳</span>
                  <span class="pnl-label">매출 연동 수수료</span>
                  <span class="pnl-sub">{svc_sub}</span>
                </div>
                <div class="pnl-num">{fmt(svc_direct)}</div>
              </div>
              
              <div class="pnl-equals pnl-equals-net">=</div>
              
              <div class="pnl-row pnl-net">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">📊</span>
                  <span class="pnl-label">순매출 (매출총이익)</span>
                </div>
                <div class="pnl-num">{fmt(sales_net)}</div>
                <div class="chip-row">
                  <div class="chip" style="color:{pcol(c3['매출(Net)']['vp'])};background:{pcol(c3['매출(Net)']['vp'])}0D">계획 {fmtd(c3['매출(Net)']['vp'])}</div>
                  <div class="chip" style="color:{pcol(c3['매출(Net)']['vm'])};background:{pcol(c3['매출(Net)']['vm'])}0D">전월 {fmtd(c3['매출(Net)']['vm'])}</div>
                  <div class="chip" style="color:{pcol(c3['매출(Net)']['vy'])};background:{pcol(c3['매출(Net)']['vy'])}0D">전년 {fmtd(c3['매출(Net)']['vy'])}</div>
                </div>
              </div>
              
              <div class="pnl-minus">−</div>
              
              <div class="pnl-row pnl-cost">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">🏢</span>
                  <span class="pnl-label">우리 팀 직접비</span>
                  <span class="pnl-sub">{top_item_label}</span>
                </div>
                <div class="pnl-num">{fmt(direct)}</div>
              </div>
              
              <div class="pnl-minus">−</div>
              
              <div class="pnl-row pnl-cost">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">📦</span>
                  <span class="pnl-label">배부 비용</span>
                  <span class="pnl-sub">순매출 − 영업이익 − 직접비</span>
                </div>
                <div class="pnl-num">{fmt(indirect)}</div>
              </div>
              
              <div class="pnl-equals">=</div>
              
              <div class="pnl-row pnl-profit">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">🏆</span>
                  <span class="pnl-label">영업이익</span>
                </div>
                <div class="pnl-num pnl-profit-num">{fmt(op)}</div>
                <div class="chip-row">
                  <div class="chip" style="color:{pcol(c3['영업이익']['vp'])};background:{pcol(c3['영업이익']['vp'])}0D">계획 {fmtd(c3['영업이익']['vp'])}</div>
                  <div class="chip" style="color:{pcol(c3['영업이익']['vm'])};background:{pcol(c3['영업이익']['vm'])}0D">전월 {fmtd(c3['영업이익']['vm'])}</div>
                  <div class="chip" style="color:{pcol(c3['영업이익']['vy'])};background:{pcol(c3['영업이익']['vy'])}0D">전년 {fmtd(c3['영업이익']['vy'])}</div>
                </div>
              </div>
            </div>"""
        else:
            # 매출 그로스가 없는 경우 (퍼블2실 같은 신작 준비) - 기존 4단
            pnl_metric_grid = f"""
            <div class="pnl-flow">
              <div class="pnl-row pnl-sales">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">💰</span>
                  <span class="pnl-label">순매출 (매출총이익)</span>
                </div>
                <div class="pnl-num">{fmt(sales_net)}</div>
                <div class="chip-row">
                  <div class="chip" style="color:{pcol(c3['매출(Net)']['vp'])};background:{pcol(c3['매출(Net)']['vp'])}0D">계획 {fmtd(c3['매출(Net)']['vp'])}</div>
                  <div class="chip" style="color:{pcol(c3['매출(Net)']['vm'])};background:{pcol(c3['매출(Net)']['vm'])}0D">전월 {fmtd(c3['매출(Net)']['vm'])}</div>
                  <div class="chip" style="color:{pcol(c3['매출(Net)']['vy'])};background:{pcol(c3['매출(Net)']['vy'])}0D">전년 {fmtd(c3['매출(Net)']['vy'])}</div>
                </div>
              </div>
              
              <div class="pnl-minus">−</div>
              
              <div class="pnl-row pnl-cost">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">🏢</span>
                  <span class="pnl-label">우리 팀 직접비</span>
                  <span class="pnl-sub">{top_item_label}</span>
                </div>
                <div class="pnl-num">{fmt(direct)}</div>
              </div>
              
              <div class="pnl-minus">−</div>
              
              <div class="pnl-row pnl-cost">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">📦</span>
                  <span class="pnl-label">배부 비용</span>
                  <span class="pnl-sub">순매출 − 영업이익 − 직접비</span>
                </div>
                <div class="pnl-num">{fmt(indirect)}</div>
              </div>
              
              <div class="pnl-equals">=</div>
              
              <div class="pnl-row pnl-profit">
                <div class="pnl-row-top">
                  <span class="pnl-emoji">🏆</span>
                  <span class="pnl-label">영업이익</span>
                </div>
                <div class="pnl-num pnl-profit-num">{fmt(op)}</div>
                <div class="chip-row">
                  <div class="chip" style="color:{pcol(c3['영업이익']['vp'])};background:{pcol(c3['영업이익']['vp'])}0D">계획 {fmtd(c3['영업이익']['vp'])}</div>
                  <div class="chip" style="color:{pcol(c3['영업이익']['vm'])};background:{pcol(c3['영업이익']['vm'])}0D">전월 {fmtd(c3['영업이익']['vm'])}</div>
                  <div class="chip" style="color:{pcol(c3['영업이익']['vy'])};background:{pcol(c3['영업이익']['vy'])}0D">전년 {fmtd(c3['영업이익']['vy'])}</div>
                </div>
              </div>
            </div>"""
        
        proj_cards=""
        for p in c4:
            nm=p['n'].replace('(P)_','·').replace('_운영','').replace('_',' ')
            rev=p['m']['매출(Net)']; oi=p['m']['영업이익']
            status = "🟢" if oi['a']>0 else "🔴" if oi['a']<0 else "⚪"
            # 영업이익 6개월 추이 sparkline (우측 하단, 약간 큼)
            oi_series = p.get('trend', {}).get('영업이익', [])
            oi_spark = make_sparkline(oi_series, width=110, height=36) if oi_series else ""
            mkt_val = p.get('mkt', 0)
            # 6개월 최고/최저 (금액 표기)
            oi_minmax = ""
            if oi_series and len(oi_series) >= 2:
                oi_max = max(oi_series)
                oi_min = min(oi_series)
                oi_minmax = f'<span class="proj-spark-minmax">최고 {fmt(oi_max)} · 최저 {fmt(oi_min)}</span>'
            proj_cards+=f"""
            <div class="proj-item">
              <div class="proj-header">
                <span class="proj-status">{status}</span>
                <span class="proj-title">{nm}</span>
              </div>
              <div class="proj-main-row">
                <div class="proj-main-cell">
                  <span class="pm-label">순매출(매출총이익)</span>
                  <span class="pm-val-lg">{fmt(rev['a'])}</span>
                </div>
                <div class="proj-main-cell">
                  <span class="pm-label">영업이익</span>
                  <span class="pm-val-lg" style="color:{pcol(oi['a'])}">{fmt(oi['a'])}</span>
                </div>
                <div class="proj-main-cell">
                  <span class="pm-label">마케팅비</span>
                  <span class="pm-val-lg">{fmt(mkt_val) if mkt_val > 0 else '-'}</span>
                </div>
              </div>
              <div class="proj-sub-row">
                <span class="proj-sub-item">계획비 영업이익 <span style="color:{pcol(oi['vp'])};font-weight:700">{fmtd(oi['vp'])}</span></span>
                <span class="proj-sub-item">전월비 영업이익 <span style="color:{pcol(oi['vm'])};font-weight:700">{fmtd(oi['vm'])}</span></span>
              </div>
              <div class="proj-spark-row">
                <span class="proj-spark-label">영업이익 6개월</span>
                {oi_minmax}
                <span class="proj-spark-svg">{oi_spark}</span>
              </div>
            </div>"""
        
        if not proj_cards:
            proj_cards = '<div class="empty">표시할 프로젝트 데이터가 없습니다</div>'
        
        alloc_proj_rows=""
        for p in c6['projects']:
            nm=p['name'].replace('(P)_','·').replace('_운영','').replace('_',' ')
            alloc_proj_rows+=f"""
            <div class="ap-item">
              <div class="ap-name">{nm}</div>
              <div class="ap-grid">
                <div class="ap-cell"><span class="ap-lb">직접비</span><span class="ap-vl">{fmt(p['cd'])}</span><span class="ap-dl" style="color:{ncol(p['dd'])}">{fmtd(p['dd'])}</span></div>
                <div class="ap-cell"><span class="ap-lb">간접비</span><span class="ap-vl">{fmt(p['ci'])}</span><span class="ap-dl" style="color:{ncol(p['di'])}">{fmtd(p['di'])}</span></div>
              </div>
            </div>"""
        
        if not alloc_proj_rows:
            alloc_proj_rows = '<div class="empty">표시할 배부 데이터가 없습니다</div>'
        
        card_pnl1 = f"""
      <!-- CARD 1: 관리손익 -->
      <div class="slide">
        <div class="card">
          <div class="card-label">CARD 1 · 관리손익</div>
          <div class="card-title">담당 사업 관리손익</div>
          <div class="big-chart">{big_chart}</div>
          {pnl_metric_grid}
        </div>
      </div>"""
        card_pnl2 = f"""
      <!-- CARD 3: 프로젝트별 -->
      <div class="slide">
        <div class="card">
          <div class="card-label">CARD 3 · 프로젝트별</div>
          <div class="card-title">프로젝트별 주요 수치</div>
          {proj_cards}
        </div>
      </div>"""
        cards_pnl = card_pnl1 + card_pnl2  # 호환성 (다른 곳 참조 대비)
    
    # ─── 카드 5: 실제 투입 MM
    card_mm = ""
    if has_mm:
        # 팀별 행
        team_rows = ""
        for t in c5['teams']:
            maxv = max((tt['curr'] for tt in c5['teams']), default=1) or 1
            bw = (t['curr']/maxv*100)
            team_rows += f"""
            <div class="bar-row">
              <div class="bar-info"><span class="bar-name">{t['team']}</span><span class="bar-val">{t['curr']:.1f} MM</span></div>
              <div class="bar-track"><div class="bar-fill" style="width:{bw}%;background:linear-gradient(90deg,#8B5CF6,#A78BFA)"></div></div>
              <div class="bar-delta" style="color:{ncol(t['delta'])}">{('+' if t['delta']>0 else '')}{t['delta']:.1f} MM</div>
            </div>"""
        # 프로젝트별 행
        proj_mm_rows = ""
        for p in c5['projects'][:8]:
            sign = "+" if p['delta']>0 else ""
            proj_mm_rows += f"""
            <div class="list-row">
              <span class="list-name">{p['name']}</span>
              <span class="list-val">{p['curr']:.1f} MM</span>
              <span class="list-delta" style="color:{ncol(p['delta'])}">{sign}{p['delta']:.1f}</span>
            </div>"""
        
        card_mm = f"""
      <!-- CARD 5: 실제 투입 현황 -->
      <div class="slide">
        <div class="card">
          <div class="card-label">CARD 5 · 실제 투입 현황</div>
          <div class="card-title">{'담당 사업 주요 프로젝트 MM' if c5.get('filtered') else '실제 투입 MM'}</div>
          <div class="hero">
            <div class="hero-label">{'담당 프로젝트 총 MM' if c5.get('filtered') else '총 투입 MM'}</div>
            <div class="hero-num">{c5['curr']:.1f}</div>
            <div class="hero-delta" style="color:{ncol(c5['delta'])}">전월 대비 {('+' if c5['delta']>0 else '')}{c5['delta']:.1f} MM</div>
          </div>
          <div style="font-size:13px;color:#8B95A1;font-weight:600;margin-bottom:12px">팀별 투입</div>
          {team_rows}
          <div class="sec-div"><span>프로젝트별 투입 (상위 8)</span><span></span></div>
          {proj_mm_rows}
        </div>
      </div>"""
    
    # ─── 카드 6: 프로젝트별 전체 부서 MM 구성 (협업 가시화)
    card_breakdown = ""
    has_breakdown = has_mm and c5.get('breakdowns') and len(c5['breakdowns']) > 0
    if has_breakdown:
        # 분류별 색상 매핑
        CAT_COLORS = {
            '사업': '#8B5CF6',
            '마케팅': '#EC4899',
            '운영': '#F59E0B',
            'QA': '#10B981',
            '플랫폼': '#3B82F6',
            '데이터': '#06B6D4',
            '인프라': '#6366F1',
            '개발': '#14B8A6',
            '디자인': '#F97316',
            '로컬': '#84CC16',
            'TPM': '#A855F7',
            '신사업': '#EF4444',
            '기타': '#9CA3AF',
        }
        
        bd_blocks = ""
        for bd in c5['breakdowns']:
            nm = bd['project'].replace('(P)_','·').replace('_운영','').replace('_',' ')
            
            # 스택 바 (가로형)
            stack = ""
            for cat in bd['categories']:
                color = CAT_COLORS.get(cat['name'], '#9CA3AF')
                stack += f'<div class="bd-seg" style="width:{cat["pct"]}%;background:{color}" title="{cat["name"]} {cat["mm"]:.1f}"></div>'
            
            # 분류 리스트
            cat_rows = ""
            for cat in bd['categories']:
                color = CAT_COLORS.get(cat['name'], '#9CA3AF')
                is_mine = (cat['name'] == '사업' and bd['inside'] > 0)
                inside_label = ""
                if is_mine and bd['inside'] > 0:
                    inside_label = f'<span class="bd-inside">우리 팀 {bd["inside"]:.1f}</span>'
                cat_rows += f"""
              <div class="bd-cat-row">
                <span class="bd-dot" style="background:{color}"></span>
                <span class="bd-cat-name">{cat['name']}</span>
                {inside_label}
                <span class="bd-cat-mm">{cat['mm']:.1f} MM</span>
                <span class="bd-cat-pct">{cat['pct']:.0f}%</span>
              </div>"""
            
            inside_pct = (bd['inside']/bd['total']*100) if bd['total'] else 0
            outside_pct = 100 - inside_pct
            
            # MM 트렌드 슬라이더 (12개월)
            mm_slider = ""
            if bd.get('mm_trend') and bd['mm_trend'].get('total'):
                mm_slider = make_mm_trend_slider(
                    bd['mm_trend']['months'],
                    bd['mm_trend']['total'],
                    bd['mm_trend']['inside'],
                )
            
            bd_blocks += f"""
        <div class="bd-block">
          <div class="bd-header">
            <span class="bd-title">{nm}</span>
            <span class="bd-total">{bd['total']:.1f} MM</span>
          </div>
          <div class="bd-summary">
            <div class="bd-sum-item">
              <span class="bd-sum-lb">우리 팀</span>
              <span class="bd-sum-vl" style="color:#8B5CF6">{bd['inside']:.1f}<span class="bd-sum-pct">{inside_pct:.0f}%</span></span>
            </div>
            <div class="bd-sum-divider"></div>
            <div class="bd-sum-item">
              <span class="bd-sum-lb">타 부서 협업</span>
              <span class="bd-sum-vl" style="color:#3182F6">{bd['outside']:.1f}<span class="bd-sum-pct">{outside_pct:.0f}%</span></span>
            </div>
          </div>
          {mm_slider}
          <div class="bd-stack">{stack}</div>
          <div class="bd-cats">
            {cat_rows}
          </div>
        </div>"""
        
        card_breakdown = f"""
      <!-- CARD 4: 담당 프로젝트 MM 구성 -->
      <div class="slide">
        <div class="card">
          <div class="card-label">CARD 4 · 담당 프로젝트 MM 구성</div>
          <div class="card-title">프로젝트별 투입 MM 상세</div>
          {bd_blocks}
        </div>
      </div>"""
    
    # 카드 1 라벨 (전체 카드 수에 따라)
    card1_label = f"CARD 2 · 총지출"
    
    return f"""<!DOCTYPE html>
<html lang="ko"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no">
<title>{DIVISION} · {MONTH_LABEL}</title>
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Pretendard',system-ui,-apple-system,sans-serif;background:#F2F3F6;color:#191F28;-webkit-font-smoothing:antialiased;overflow:hidden;height:100dvh}}
.wrap{{max-width:430px;margin:0 auto;height:100dvh;display:flex;flex-direction:column}}
.hdr{{padding:20px 24px 14px;background:#fff;border-bottom:1px solid #F2F3F6}}
.hdr-badge{{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;color:#3182F6;background:#EBF3FE;padding:4px 10px;border-radius:100px;margin-bottom:8px;letter-spacing:-0.2px}}
.hdr-title{{font-size:22px;font-weight:700;letter-spacing:-0.6px;line-height:1.2}}
.hdr-sub{{font-size:13px;color:#8B95A1;margin-top:4px;font-weight:400}}
.carousel{{flex:1;overflow:hidden;position:relative}}
.track{{display:flex;height:100%;transition:transform .3s cubic-bezier(.25,.1,.25,1);will-change:transform}}
.slide{{min-width:100%;height:100%;overflow-y:auto;padding:12px 16px 24px;-webkit-overflow-scrolling:touch}}
.slide::-webkit-scrollbar{{display:none}}
.card{{background:#fff;border-radius:20px;padding:24px;margin-bottom:12px}}
.card-label{{font-size:12px;font-weight:600;color:#8B95A1;margin-bottom:4px;letter-spacing:-0.2px}}
.card-title{{font-size:18px;font-weight:700;letter-spacing:-0.4px;margin-bottom:20px}}
.hero{{text-align:center;padding:28px 0;margin-bottom:20px;background:#F8F9FA;border-radius:16px}}
.hero-label{{font-size:13px;color:#8B95A1;font-weight:500}}
.hero-num{{font-size:36px;font-weight:800;letter-spacing:-1.5px;margin:6px 0 4px;color:#191F28}}
.hero-delta{{font-size:14px;font-weight:600}}
.hero-breakdown{{font-size:11px;color:#8B95A1;margin-top:6px;font-weight:500}}
.bar-row{{margin-bottom:14px}}
.bar-info{{display:flex;justify-content:space-between;margin-bottom:4px}}
.bar-name{{font-size:14px;color:#4E5968;font-weight:500}}
.bar-val{{font-size:14px;font-weight:700;font-variant-numeric:tabular-nums}}
.bar-pct{{font-size:11px;color:#8B95A1;margin-left:4px;font-weight:500}}
.bar-track{{height:8px;background:#F2F3F6;border-radius:4px;overflow:hidden}}
.bar-fill{{height:100%;background:linear-gradient(90deg,#3182F6,#5BA0FA);border-radius:4px;transition:width .6s ease}}
.bar-delta{{font-size:12px;font-weight:600;margin-top:2px;text-align:right;font-variant-numeric:tabular-nums}}
.list-row{{display:flex;align-items:center;padding:12px 0;border-bottom:1px solid #F2F3F6;gap:8px}}
.list-row:last-child{{border:none}}
.list-name{{flex:1;font-size:14px;color:#4E5968;font-weight:500}}
.list-val{{font-size:14px;font-weight:700;font-variant-numeric:tabular-nums}}
.list-delta{{font-size:12px;font-weight:600;min-width:60px;text-align:right;font-variant-numeric:tabular-nums}}
.sec-div{{font-size:12px;color:#8B95A1;font-weight:600;padding:16px 0 8px;margin-top:4px;border-top:1px solid #F2F3F6;display:flex;justify-content:space-between}}
.metric-card{{padding:18px 0;border-bottom:1px solid #F2F3F6}}
.metric-card:last-child{{border:none}}
.metric-top{{display:flex;align-items:center;gap:6px;margin-bottom:6px}}
.metric-emoji{{font-size:16px}}
.metric-label{{font-size:14px;color:#6B7684;font-weight:600}}
.metric-spark{{margin-left:auto;display:flex;align-items:center}}
.metric-num-row{{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}}
.metric-num{{font-size:26px;font-weight:800;letter-spacing:-1px;font-variant-numeric:tabular-nums}}
.spark-warn{{font-size:11px;font-weight:700;color:#F04452;background:#FEE2E2;padding:2px 8px;border-radius:6px;letter-spacing:-0.2px}}

/* PNL FLOW (카드 2 매출→직접→간접→영업이익 구조) */
.pnl-flow{{display:flex;flex-direction:column;gap:6px;margin-top:8px}}
.pnl-row{{padding:14px 16px;border-radius:14px;background:#F8F9FA}}
.pnl-row.pnl-sales{{background:#EBF3FE}}
.pnl-row.pnl-net{{background:#DBEAFE;border:1.5px solid #93C5FD}}
.pnl-row.pnl-cost{{background:#F8F9FA}}
.pnl-row.pnl-profit{{background:#F0FDF4;border:1.5px solid #BBF7D0}}
.pnl-equals-net{{color:#1E40AF !important}}
.pnl-row-top{{display:flex;align-items:center;gap:6px;margin-bottom:4px}}
.pnl-emoji{{font-size:14px}}
.pnl-label{{font-size:13px;font-weight:700;color:#191F28;letter-spacing:-0.2px}}
.pnl-sub{{font-size:10px;color:#8B95A1;font-weight:500;margin-left:auto}}
.pnl-num{{font-size:22px;font-weight:800;letter-spacing:-0.6px;font-variant-numeric:tabular-nums;color:#191F28}}
.pnl-profit-num{{font-size:26px;color:#0A7B3E}}
.pnl-minus,.pnl-equals{{
  text-align:center;font-size:18px;font-weight:700;color:#B0B8C1;
  padding:2px 0;line-height:1;
}}
.pnl-equals{{color:#0A7B3E}}
.chip-row{{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap}}
.chip{{font-size:12px;font-weight:700;padding:4px 10px;border-radius:8px;font-variant-numeric:tabular-nums;letter-spacing:-0.3px}}
.proj-item{{background:#F8F9FA;border-radius:16px;padding:18px 20px;margin-bottom:10px}}
.proj-item:last-child{{margin:0}}
.proj-header{{display:flex;align-items:center;gap:8px;margin-bottom:14px}}
.proj-status{{font-size:10px}}
.proj-title{{font-size:15px;font-weight:700;letter-spacing:-0.3px}}
.proj-spark-row{{display:flex;align-items:center;gap:8px;margin-top:12px;padding-top:12px;border-top:1px solid #E5E8EB;justify-content:flex-end}}
.proj-spark-minmax{{font-size:9px;color:#B0B8C1;font-weight:500;font-variant-numeric:tabular-nums;letter-spacing:-0.1px}}
.proj-spark-label{{font-size:11px;color:#8B95A1;font-weight:600}}
.proj-spark-svg{{display:flex;align-items:center}}
/* BIG CHART (카드2 상단) */
.big-chart{{background:#F8F9FA;border-radius:14px;padding:12px;margin-bottom:18px}}
.proj-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px 16px}}
.pm-label{{font-size:11px;color:#8B95A1;font-weight:600;display:block}}
.pm-val{{font-size:15px;font-weight:700;font-variant-numeric:tabular-nums;letter-spacing:-0.3px}}

/* 카드3 새 레이아웃 */
.proj-main-row{{
  display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;
  padding:14px 0;border-top:1px solid #F2F3F6;border-bottom:1px solid #F2F3F6;
}}
.proj-main-cell{{display:flex;flex-direction:column;gap:4px}}
.pm-val-lg{{font-size:17px;font-weight:800;font-variant-numeric:tabular-nums;letter-spacing:-0.4px;color:#191F28}}
.proj-sub-row{{
  display:flex;gap:14px;padding:10px 0 4px;
  font-size:11px;color:#6B7684;font-weight:500;
}}
.proj-sub-item{{font-variant-numeric:tabular-nums}}
.alloc-row{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}}
.alloc-box{{border-radius:16px;padding:20px;text-align:center}}
.alloc-box .al{{font-size:12px;font-weight:600;margin-bottom:6px}}
.alloc-box .av{{font-size:24px;font-weight:800;letter-spacing:-0.8px}}
.alloc-box .as{{font-size:12px;margin-top:6px;font-weight:500}}
.ab-direct{{background:#EBF3FE;color:#1B64DA}}
.ab-indirect{{background:#E8FAF0;color:#0A7B3E}}
.alloc-detail{{padding:14px 0;border-bottom:1px solid #F2F3F6;display:flex;justify-content:space-between;align-items:center}}
.alloc-detail:last-child{{border:none}}
.ad-left{{display:flex;flex-direction:column;gap:2px}}
.ad-label{{font-size:14px;color:#6B7684;font-weight:500}}
.ad-sub{{font-size:11px;color:#B0B8C1;font-variant-numeric:tabular-nums;font-weight:500}}
.ad-val{{font-size:18px;font-weight:700;font-variant-numeric:tabular-nums}}
.other-box{{margin-top:16px;background:#FFF8E7;border-radius:14px;padding:16px 18px}}
.other-title{{font-size:12px;font-weight:700;color:#92651A;margin-bottom:12px;letter-spacing:-0.2px}}
.other-row{{padding:8px 0;border-bottom:1px solid #F5E9C7}}
.other-row:last-child{{border:none;padding-bottom:0}}
.other-row:first-of-type{{padding-top:0}}
.other-top{{display:flex;justify-content:space-between;align-items:baseline}}
.other-wbs{{font-size:13px;font-weight:700;color:#3D2C0A;letter-spacing:-0.2px}}
.other-val{{font-size:13px;font-weight:700;color:#3D2C0A;font-variant-numeric:tabular-nums}}
.other-acct{{font-size:11px;color:#92651A;margin-top:2px;font-weight:500}}
/* MKT BOX */
.mkt-box{{margin-top:16px;background:#FCE7F3;border-radius:14px;padding:16px 18px}}
.mkt-title{{font-size:12px;font-weight:700;color:#9D174D;margin-bottom:14px;letter-spacing:-0.2px}}
.mkt-row{{padding:8px 0;border-bottom:1px solid #FBCFE8}}
.mkt-row:last-child{{border:none;padding-bottom:0}}
.mkt-row:first-of-type{{padding-top:0}}
.mkt-top{{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px}}
.mkt-wbs{{font-size:13px;font-weight:700;color:#831843;letter-spacing:-0.2px}}
.mkt-cc{{display:inline-block;font-size:10px;color:#BE185D;background:#FCE7F3;padding:1px 6px;border-radius:4px;font-weight:600;margin-left:6px;border:0.5px solid #FBCFE8}}
.mkt-val{{font-size:13px;font-weight:700;color:#831843;font-variant-numeric:tabular-nums}}
.mkt-pct{{font-size:11px;color:#BE185D;margin-left:4px;font-weight:600}}
.mkt-bar{{height:6px;background:#FBCFE8;border-radius:3px;overflow:hidden}}
.mkt-fill{{height:100%;background:linear-gradient(90deg,#EC4899,#F472B6);border-radius:3px}}
.ap-item{{background:#F8F9FA;border-radius:14px;padding:14px 16px;margin-bottom:8px}}
.ap-item:last-child{{margin:0}}
.ap-name{{font-size:14px;font-weight:700;margin-bottom:10px;letter-spacing:-0.3px}}
.ap-grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.ap-cell{{display:flex;flex-direction:column;gap:2px}}
.ap-lb{{font-size:11px;color:#8B95A1;font-weight:600}}
.ap-vl{{font-size:14px;font-weight:700;font-variant-numeric:tabular-nums;letter-spacing:-0.2px}}
.ap-dl{{font-size:11px;font-weight:600;font-variant-numeric:tabular-nums;margin-top:1px}}
.empty{{padding:40px 20px;text-align:center;color:#8B95A1;font-size:13px}}
/* BREAKDOWN CARD */
.bd-block{{background:#F8F9FA;border-radius:16px;padding:18px 20px;margin-bottom:12px}}
.bd-block:last-child{{margin:0}}
.bd-header{{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:12px}}
.bd-title{{font-size:15px;font-weight:700;letter-spacing:-0.3px;color:#191F28}}
.bd-total{{font-size:18px;font-weight:800;letter-spacing:-0.5px;color:#191F28;font-variant-numeric:tabular-nums}}
.bd-summary{{display:flex;align-items:center;background:#fff;border-radius:12px;padding:14px 16px;margin-bottom:14px;gap:14px}}
.bd-sum-item{{flex:1;display:flex;flex-direction:column;gap:4px}}
.bd-sum-divider{{width:1px;height:32px;background:#E5E8EB}}
.bd-sum-lb{{font-size:11px;color:#8B95A1;font-weight:600}}
.bd-sum-vl{{font-size:20px;font-weight:800;letter-spacing:-0.5px;font-variant-numeric:tabular-nums;display:flex;align-items:baseline;gap:4px}}
.bd-sum-pct{{font-size:11px;color:#8B95A1;font-weight:600}}
.bd-stack{{display:flex;height:10px;border-radius:5px;overflow:hidden;margin-bottom:14px;background:#E5E8EB}}
.bd-seg{{height:100%;transition:opacity .2s}}
.bd-seg:hover{{opacity:.8}}
.bd-cats{{display:flex;flex-direction:column;gap:8px}}
.bd-cat-row{{display:flex;align-items:center;gap:8px;font-size:13px}}
.bd-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.bd-cat-name{{color:#4E5968;font-weight:600;flex-shrink:0}}
.bd-inside{{font-size:10px;color:#8B5CF6;background:#F3F0FF;padding:1px 6px;border-radius:4px;font-weight:600}}
.bd-cat-mm{{margin-left:auto;font-weight:700;color:#191F28;font-variant-numeric:tabular-nums}}
.bd-cat-pct{{font-size:11px;color:#8B95A1;min-width:30px;text-align:right;font-variant-numeric:tabular-nums;font-weight:600}}
.nav{{padding:10px 0 24px;display:flex;justify-content:center;align-items:center;gap:14px;background:#fff;flex-shrink:0}}
.nav-btn{{width:44px;height:44px;border-radius:50%;border:none;background:#F2F3F6;font-size:20px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#4E5968;font-weight:500;transition:all .15s}}
.nav-btn:active{{transform:scale(.93);background:#E5E8EB}}
.nav-btn:disabled{{opacity:.25;cursor:default}}
.nav-btn:disabled:active{{transform:none;background:#F2F3F6}}
.dots{{display:flex;gap:6px}}
.dot{{width:6px;height:6px;border-radius:3px;background:#D1D6DB;transition:all .3s ease}}
.dot.on{{background:#191F28;width:18px}}
</style></head><body>
<div class="wrap">
  <div class="hdr">
    <div class="hdr-badge">📊 {MONTH_LABEL}</div>
    <div class="hdr-title">{DIVISION}</div>
    <div class="hdr-sub">{SUB} · {c2['hc'] if c2['hc'] else '-'}MM</div>
  </div>
  <div class="carousel" id="carousel">
    <div class="track" id="track">
{card_pnl1}
      <!-- CARD 2: 총지출 -->
      <div class="slide">
        <div class="card">
          <div class="card-label">{card1_label}</div>
          <div class="card-title">이번달 우리 팀 지출액</div>
          <div class="hero">
            <div class="hero-label">총 지출액 (마케팅비 포함)</div>
            <div class="hero-num">{fmt(c2['gt']['c'])}</div>
            <div class="hero-delta" style="color:{ncol(c2['gt']['d'])}">전월 대비 {fmtd(c2['gt']['d'])}</div>
            {('<div class="hero-breakdown"></div>') if False else ''}
          </div>
          <div style="font-size:13px;color:#8B95A1;font-weight:600;margin-bottom:12px">부서 직접 비용 · {fmt(c2['dt']['c'])}</div>
          {dept_rows}
          {proj_section}
          {mkt_section}
          {other_section}
        </div>
      </div>
{card_pnl2}
{card_breakdown}
    </div>
  </div>
  <div class="nav">
    <button class="nav-btn" id="prev" disabled>&#8249;</button>
    <div class="dots" id="dots"></div>
    <button class="nav-btn" id="next">&#8250;</button>
  </div>
</div>
<script>
const T=document.getElementById('track'),S=T.children,N=S.length;
let I=0;
const D=document.getElementById('dots'),P=document.getElementById('prev'),X=document.getElementById('next');
for(let i=0;i<N;i++){{let d=document.createElement('div');d.className='dot'+(i===0?' on':'');d.onclick=()=>go(i);D.appendChild(d)}}
function go(i){{
  I=Math.max(0,Math.min(N-1,i));
  T.style.transform='translateX(-'+I*100+'%)';
  P.disabled=I===0;X.disabled=I===N-1;
  document.querySelectorAll('.dot').forEach((d,j)=>d.className='dot'+(j===I?' on':''));
  S[I].scrollTop=0;
}}
P.onclick=()=>go(I-1);X.onclick=()=>go(I+1);
let sx=0;
document.getElementById('carousel').addEventListener('touchstart',e=>sx=e.touches[0].clientX);
document.getElementById('carousel').addEventListener('touchend',e=>{{
  const dx=e.changedTouches[0].clientX-sx;
  if(Math.abs(dx)>50){{if(dx<0)go(I+1);else go(I-1)}}
}});
if(N===1){{X.style.display='none';P.style.display='none'}}
</script>
</body></html>"""


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    divs = load_division_config()
    
    print(f"📊 {len(divs)}개 팀 브리프 생성 시작 (퍼블리싱 산하)\n")
    
    index_links = []
    for div in divs:
        print(f"▶ {div['name']} ({div['parent_sil']})")
        try:
            # Card 2: 팀 버전 = 배부결과 직접비 기반 (코스트센터 + WBS)
            c2 = extract_card2_team(div['name'], div['projects'])
            # 원장 기반 '기타' 상세는 그대로 재사용 (프로젝트 리스트 기준)
            c2['other_detail'] = analyze_other_from_ledger(div['projects'])
            # 마케팅비 상세(WBS 기준) - 해당 팀 프로젝트에 귀속된 전체 마케팅비
            if div['projects']:
                c2['mkt_detail'] = analyze_marketing_by_projects(div['projects'])
            else:
                c2['mkt_detail'] = None
            c2['extra_mkt'] = 0  # 팀 버전은 배부결과에서 이미 WBS 타부서분 포함
            
            if div['projects']:
                c3, c4 = extract_card3_4(div['projects'])
                if c2.get('mkt_detail') and c2['mkt_detail'].get('items'):
                    mkt_by_wbs = {it['wbs']: it['val'] for it in c2['mkt_detail']['items']}
                    for p in c4:
                        p['mkt'] = mkt_by_wbs.get(p['n'], 0)
                else:
                    for p in c4:
                        p['mkt'] = 0
                c6 = extract_card6(
                    div['projects'],
                    card1_total=c2['gt']['c'],
                    sales_net=c3['매출(Net)']['a'],
                    op_income=c3['영업이익']['a'],
                )
                print(f"   ✓ 손익: 영업이익 {fmt(c3['영업이익']['a'])} / 프로젝트 {len(c4)}개 / 직접비 {fmt(c6['cd'])}")
            else:
                c3, c4, c6 = None, None, None
                print(f"   ✓ 손익 카드 제외 (프로젝트 없음)")
            
            c5 = extract_workload(div.get('teams'), filter_projects=div.get('projects'))
            if c5:
                print(f"   ✓ MM: 당월 {c5['curr']:.1f} (전월 {c5['prev']:.1f}, {c5['delta']:+.1f})")
            
            print(f"   ✓ 팀비용: {fmt(c2['gt']['c'])} (CC {fmt(c2['dt']['c'])} + WBS {fmt(c2['pt']['c'])})")
            
            html = gen_html(div, c2, c3, c4, c6, c5)
            fname = f"brief_{div['name']}_{CURR_MONTH}.html"
            path = os.path.join(OUTPUT_DIR, fname)
            with open(path,'w',encoding='utf-8') as f:
                f.write(html)
            print(f"   ✅ {fname}\n")
            index_links.append({'name':div['name'],'sub':div['sub'],'file':fname,'total':c2['gt']['c']})
        except Exception as e:
            print(f"   ❌ 실패: {e}\n")
            import traceback; traceback.print_exc()
    
    # 인덱스 페이지 생성
    links_html = ""
    for l in index_links:
        links_html += f"""
        <a href="{l['file']}" class="idx-card">
          <div class="idx-name">{l['name']}</div>
          <div class="idx-sub">{l['sub']}</div>
          <div class="idx-total">{fmt(l['total'])}</div>
        </a>"""
    
    index_html = f"""<!DOCTYPE html><html lang="ko"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>월간 팀 숫자 브리프 · {MONTH_LABEL}</title>
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Pretendard',sans-serif;background:#F2F3F6;color:#191F28;padding:24px 16px;min-height:100vh}}
.idx-wrap{{max-width:430px;margin:0 auto}}
.idx-hdr{{padding:8px 8px 24px}}
.idx-badge{{display:inline-block;font-size:11px;font-weight:600;color:#3182F6;background:#EBF3FE;padding:4px 10px;border-radius:100px;margin-bottom:8px}}
.idx-title{{font-size:24px;font-weight:800;letter-spacing:-0.6px}}
.idx-desc{{font-size:13px;color:#8B95A1;margin-top:4px}}
.idx-card{{display:block;background:#fff;border-radius:16px;padding:20px 22px;margin-bottom:10px;text-decoration:none;color:inherit;transition:transform .15s}}
.idx-card:active{{transform:scale(.98)}}
.idx-name{{font-size:17px;font-weight:700;letter-spacing:-0.3px}}
.idx-sub{{font-size:12px;color:#8B95A1;margin-top:2px;font-weight:500}}
.idx-total{{font-size:22px;font-weight:800;letter-spacing:-0.8px;margin-top:10px;color:#3182F6;font-variant-numeric:tabular-nums}}
</style></head><body>
<div class="idx-wrap">
  <div class="idx-hdr">
    <div class="idx-badge">📊 {MONTH_LABEL} · 팀 기준</div>
    <div class="idx-title">월간 팀 숫자 브리프</div>
    <div class="idx-desc">팀을 선택하면 상세 카드가 열립니다</div>
  </div>
  {links_html}
</div></body></html>"""
    
    with open(os.path.join(OUTPUT_DIR, "index.html"),'w',encoding='utf-8') as f:
        f.write(index_html)
    print(f"📑 index.html 생성 완료")
    print(f"\n✅ 전체 완료!")
