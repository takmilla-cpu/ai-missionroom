"""
generate_dashboard.py
엑셀 파일(data/0507데이터.xlsx)을 읽어 data/dashboard.json 생성
"""
import pandas as pd
from datetime import date, datetime
import numpy as np, json, sys
from collections import defaultdict, Counter

TODAY = date.today()
CUTOFF_WEEKS = 3  # 3주 이상 미진행 = 지연
NEW_HIRE_CUTOFF = pd.Timestamp('2026-02-25')
EXCEL_PATH = 'data/latest.xlsx'
OUTPUT_PATH = 'data/dashboard.json'

# ── 휴직 처리
df_leave = pd.read_excel(EXCEL_PATH, sheet_name='휴직현황')
df_leave = df_leave.dropna(subset=['성명'])
def is_on_leave(row):
    start = row['휴직시작일']; end = row['휴직종료일']
    if isinstance(start, str) or pd.isna(start): return False
    start = start.date() if hasattr(start,'date') else start
    end = date(2099,12,31) if pd.isna(end) else (end.date() if hasattr(end,'date') else end)
    return start <= TODAY <= end
df_leave['휴직중'] = df_leave.apply(is_on_leave, axis=1)
on_leave_names = set(df_leave[df_leave['휴직중']]['성명'].str.strip())

# ── 전사 / 수료현황 데이터
df_all = pd.read_excel(EXCEL_PATH, sheet_name='전사 데이터')
df_status = pd.read_excel(EXCEL_PATH, sheet_name='수료 현황 데이터')
df_all['이메일_lower'] = df_all['사내메일'].str.lower().str.strip()
df_status['이메일_lower'] = df_status['이메일'].str.lower().str.strip()
df_all['휴직중'] = df_all['성명'].str.strip().isin(on_leave_names)
df_target = df_all[(df_all['STATUS']=='대상자') & (~df_all['휴직중'])].copy()
df_target['is_new'] = df_target['최초입사일'] >= NEW_HIRE_CUTOFF
df_target_m = df_target.rename(columns={'부서':'조직부서'})

# 관리자 정보 (전체 데이터 기준)
mgr_info = {}
for _, r in df_all.iterrows():
    name = str(r['성명']).strip()
    bz = str(r['사업부']).strip() if pd.notna(r['사업부']) else ''
    gr = str(r['직책']).strip() if pd.notna(r['직책']) else ''
    if name and bz:
        mgr_info[name] = {'gr': gr, 'bz': bz}

df = df_status.merge(
    df_target_m[['이메일_lower','성명','사업부','본부','조직부서','직무','직책','고용형태','최초입사일','is_new','관리자']],
    on='이메일_lower', how='inner'
).copy()

def parse_d(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return None
    if hasattr(val, 'date'): return val.date() if callable(val.date) else val
    s = str(val).strip()
    if s in ['nan','NaT','None','','-']: return None
    for fmt in ['%Y. %m. %d','%Y.%m.%d','%Y-%m-%d']:
        try: return datetime.strptime(s, fmt).date()
        except: pass
    return None

CUTOFF = date(TODAY.year, TODAY.month, TODAY.day) - pd.Timedelta(weeks=CUTOFF_WEEKS)
CUTOFF = CUTOFF.date() if hasattr(CUTOFF, 'date') else CUTOFF

df['login_date'] = df['가입일'].apply(parse_d)
df['logged'] = df['login_date'].notna()
df['cw'] = sum((df[f'{i}주차'].astype(str).str.strip()=='완료') for i in range(1,11))
df['login_days'] = df['login_date'].apply(lambda d: (TODAY - d).days if d else 0)
for i in range(1,11): df[f'w{i}'] = df[f'{i}주차 완료일'].apply(parse_d)
def get_last(row):
    ds = [row[f'w{i}'] for i in range(1,11) if row[f'w{i}']]
    return max(ds) if ds else None
df['last'] = df.apply(get_last, axis=1)

def rst(row):
    if not row['logged']: return 'nolog'
    if row['cw'] == 10: return 'done'
    if row['login_days'] >= 21 and row['cw'] == 0: return 'abandon'
    last = row['last']
    if last is None: return 'prog'
    return 'prog' if last >= CUTOFF else 'delay'
df['rst'] = df.apply(rst, axis=1)
not_in = df_target_m[~df_target_m['이메일_lower'].isin(df_status['이메일_lower'])].copy()
not_in['rst'] = 'nolog'

def cl(v):
    s = str(v).strip() if v is not None else ''
    return '' if s in ['-','nan','NaN','NaT','None'] else s

ST_LABEL = {'done':'완료','prog':'진행중','delay':'3주↑ 지연','abandon':'방치','nolog':'미로그인'}

# ── RAW 데이터
all_rows = []
for _, r in df.iterrows():
    all_rows.append({'n':cl(r['이름']),'e':r['이메일'],'bz':cl(r['사업부']),'hb':cl(r['본부']),'dp':cl(r['조직부서']),'ro':cl(r['직무']),'gr':cl(r['직책']),'ht':cl(r['고용형태']),'hd':str(r['최초입사일'])[:10],'c':int(r['cw']),'lcd':str(r['last']) if r['last'] else '','ld':str(r['login_date']) if r['login_date'] else '','ldays':int(r['login_days']) if r['login_date'] else 0,'st':r['rst']})
for _, r in not_in.iterrows():
    all_rows.append({'n':cl(r['성명']),'e':cl(r['사내메일']),'bz':cl(r['사업부']),'hb':cl(r['본부']),'dp':cl(r['조직부서']),'ro':cl(r['직무']),'gr':cl(r['직책']),'ht':cl(r['고용형태']),'hd':str(r['최초입사일'])[:10],'c':0,'lcd':'','ld':'','ldays':0,'st':'nolog'})

cnt = Counter(p['st'] for p in all_rows)

# ── INIT_DATA (전체현황 탭 요약)
from collections import defaultdict
biz_map = defaultdict(lambda:{'prog':0,'done':0,'delay':0,'nolog':0,'abandon':0,'total':0})
for p in all_rows:
    bz = p['bz'] or '(미분류)'
    biz_map[bz][p['st']] += 1
    biz_map[bz]['total'] += 1
biz_list = [{'bz':k,'done':v['done'],'prog':v['prog'],'delay':v['delay'],'abandon':v['abandon'],'nolog':v['nolog'],'total':v['total']} for k,v in sorted(biz_map.items(), key=lambda x:-x[1]['total'])]

wk_cnt = []
for w in range(1, 11):
    c = sum(1 for p in all_rows if p['c'] >= w)
    pct = round(c / len(all_rows) * 100, 1) if all_rows else 0
    wk_cnt.append({'w': w, 'c': c, 'p': pct})

init_data = {
    'summary':{'total':len(all_rows),'completed':cnt['done'],'in_progress':cnt['prog'],'no_login':cnt['nolog']},
    'biz': biz_list,
    'weeks': wk_cnt,
    'delayed': [p for p in all_rows if p['st']=='delay'],
    'newb': [p for p in all_rows if p['st'] in ('done','prog','delay','abandon','nolog') and p['hd'] >= '2026-02-25'],
    'nolog': [p for p in all_rows if p['st']=='nolog'],
}

# ── INSIGHTS
team_map = defaultdict(list)
for p in all_rows:
    team_map[p['dp'] or '(미분류)'].append(p)

nl, dl, ab = {}, {}, {}
ad, ot = [], []
for t, ms in team_map.items():
    sc = Counter(p['st'] for p in ms)
    if sc['nolog'] > 0: nl[t] = {'count':sc['nolog'],'total':len(ms),'members':[{'name':p['n'],'grade':p['gr'],'cw':p['c'],'st':'nolog','stl':'미로그인'} for p in ms if p['st']=='nolog']}
    if sc['delay'] > 0: dl[t] = {'count':sc['delay'],'total':len(ms),'members':[{'name':p['n'],'grade':p['gr'],'cw':p['c'],'st':'delay','stl':'3주↑ 지연'} for p in ms if p['st']=='delay']}
    if sc['abandon'] > 0 and t != '(미분류)': ab[t] = {'count':sc['abandon'],'total':len(ms),'members':[{'name':p['n'],'grade':p['gr'],'cw':p['c'],'st':'abandon','stl':'방치'} for p in ms if p['st']=='abandon']}
    if sc['done'] == len(ms): ad.append({'team':t,'total':len(ms),'members':[{'name':p['n'],'grade':p['gr'],'cw':p['c'],'st':'done','stl':'완료'} for p in ms]})
    if sc.get('delay',0)+sc.get('nolog',0)+sc.get('abandon',0) == 0 and len(ms)>0: ot.append({'team':t,'total':len(ms),'members':[{'name':p['n'],'grade':p['gr'],'cw':p['c'],'st':p['st'],'stl':ST_LABEL[p['st']]} for p in ms]})

insights = {
    'all_done': ad,
    'on_track': ot,
    'top_nologin': sorted([dict(team=k,**v) for k,v in nl.items()], key=lambda x:-x['count'])[:5],
    'top_delayed': sorted([dict(team=k,**v) for k,v in dl.items()], key=lambda x:-x['count'])[:5],
    'top_abandon': sorted([dict(team=k,**v) for k,v in ab.items()], key=lambda x:-x['count'])[:5],
}

# ── 관리자 탭 데이터
mgr_map2 = defaultdict(lambda: defaultdict(list))
for p in all_rows:
    mgr = cl(df[df['이메일_lower']==p['e']]['관리자'].values[0]) if p['e'] in df['이메일_lower'].values else ''
    if not mgr:
        row_ni = not_in[not_in['이메일_lower']==p['e']]
        if len(row_ni): mgr = cl(row_ni.iloc[0].get('관리자',''))
    if not mgr: continue
    dept = p['dp'] or p['bz'] or '(미분류)'
    mgr_map2[mgr][dept].append({'n':p['n'],'gr':p['gr'],'c':p['c'],'st':p['st'],'stl':ST_LABEL[p['st']]})

managers = []
for mgr, depts in mgr_map2.items():
    all_m = [m for ms in depts.values() for m in ms]
    c2 = Counter(m['st'] for m in all_m)
    total = len(all_m)
    rate = round((c2['delay']+c2['abandon'])/total*100) if total else 0
    info = mgr_info.get(mgr,{'gr':'','bz':''})
    dl2 = [{'dept':d,'total':len(ms),'done':Counter(m['st'] for m in ms)['done'],'prog':Counter(m['st'] for m in ms)['prog'],'delay':Counter(m['st'] for m in ms)['delay'],'abandon':Counter(m['st'] for m in ms)['abandon'],'nolog':Counter(m['st'] for m in ms)['nolog'],'members':ms} for d,ms in depts.items()]
    dl2.sort(key=lambda x:-(x['delay']+x['abandon']))
    managers.append({'name':mgr,'gr':info['gr'],'bz':info['bz'],'total':total,'done':c2['done'],'prog':c2['prog'],'delay':c2['delay'],'abandon':c2['abandon'],'nolog':c2['nolog'],'rate':rate,'depts':dl2})
managers.sort(key=lambda x:-x['rate'])

biz_map3 = defaultdict(list)
for m in managers:
    biz_map3[m['bz'] or '(미분류)'].append(m)
biz_mgr = []
for bz, mlist in biz_map3.items():
    total = sum(m['total'] for m in mlist)
    delay = sum(m['delay'] for m in mlist)
    abandon = sum(m['abandon'] for m in mlist)
    rate = round((delay+abandon)/total*100) if total else 0
    biz_mgr.append({'bz':bz,'mgr_cnt':len(mlist),'total':total,'done':sum(m['done'] for m in mlist),'prog':sum(m['prog'] for m in mlist),'delay':delay,'abandon':abandon,'nolog':sum(m['nolog'] for m in mlist),'rate':rate,'managers':sorted(mlist,key=lambda x:-x['rate'])})
biz_mgr.sort(key=lambda x:-x['rate'])

# ── 최종 JSON 저장
result = {
    'generated': str(TODAY),
    'raw': all_rows,
    'init_data': init_data,
    'insights': insights,
    'biz_mgr': biz_mgr,
}
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)

print(f"✅ dashboard.json 생성 완료")
print(f"   기준일: {TODAY}")
print(f"   전체: {len(all_rows)}명 / 완료: {cnt['done']} / 지연: {cnt['delay']} / 방치: {cnt['abandon']}")
