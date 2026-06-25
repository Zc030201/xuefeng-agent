#!/usr/bin/env python3
"""雪峰Agent — 单文件服务器：HTML UI + API + 数据库查询"""
import os, re, json, sqlite3, gzip, shutil, urllib.request, urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from http.server import ThreadingHTTPServer
import sys
try: sys.stdout.reconfigure(encoding="utf-8")
except: pass

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, 'admission_clean.db')
GZ_PATH = os.path.join(HERE, 'admission_clean.db.gz')
if not os.path.exists(DB_PATH) and os.path.exists(GZ_PATH):
    with gzip.open(GZ_PATH, 'rb') as gz:
        with open(DB_PATH, 'wb') as f:
            shutil.copyfileobj(gz, f)

HAS_DB = os.path.exists(DB_PATH)
USER_XLSX = os.path.join(HERE, '自定义数据.xlsx')
USER_DATA = []
def clean_num(v):
    if v is None: return None
    s = str(v).strip()
    if not s: return None
    s = s.replace(',','').replace('，','').replace(' ','').replace('分','').replace('位','')
    try: return int(float(s))
    except: return None
def load_user_data():
    if not os.path.exists(USER_XLSX): return
    try: import openpyxl
    except ImportError: return
    global USER_DATA; USER_DATA = []
    if not os.path.exists(USER_XLSX): return
    try:
        wb = openpyxl.load_workbook(USER_XLSX, data_only=True)
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            school = str(row[0]).strip()
            if len(school) < 2 or school in ['学校名称','院校名称']: continue
            note = str(row[8]).strip() if row[8] else ''
            if '示例' in note or '不参与排序' in note: continue
            major = str(row[1]).strip() if row[1] else ''
            cat = str(row[2]).strip() if row[2] else ''
            if '物理' in cat: category = '物理类'
            elif '历史' in cat: category = '历史类'
            elif '综合' in cat: category = '综合'
            else: category = cat
            prov = str(row[3]).strip() if row[3] else ''
            prov = prov.replace('省','').replace('市','').strip()
            s24 = clean_num(row[4]); r24 = clean_num(row[5])
            s25 = clean_num(row[6]); r25 = clean_num(row[7])
            if s24 and s24 < 100 and r24 and r24 > 300: s24, r24 = r24, s24
            if s25 and s25 < 100 and r25 and r25 > 300: s25, r25 = r25, s25
            if s24 and r24: USER_DATA.append({'school':school,'major':major,'year':2024,'category':category,'score':s24,'rank':r24,'province':prov})
            if s25 and r25: USER_DATA.append({'school':school,'major':major,'year':2025,'category':category,'score':s25,'rank':r25,'province':prov})
        wb.close()
    except Exception as e: print(f'[user data] {e}')
load_user_data()

PROVINCES = ['北京','天津','上海','重庆','河北','山西','辽宁','吉林','黑龙江','江苏','浙江','安徽',
             '福建','江西','山东','河南','湖北','湖南','广东','广西','海南','四川','贵州','云南',
             '西藏','陕西','甘肃','青海','宁夏','新疆','内蒙古']

def query_db(province=None, school=None, major=None, limit=50):
    if not HAS_DB: return None
    conn = sqlite3.connect(DB_PATH)
    conds, params = [], []
    if province: conds.append("province LIKE ?"); params.append(f"%{province}%")
    if school: conds.append("school LIKE ?"); params.append(f"%{school}%")
    if major: conds.append("major LIKE ?"); params.append(f"%{major}%")
    if not conds: conn.close(); return None
    sql = f"SELECT province,year,school_name,major_name,score,rank FROM admission WHERE {' AND '.join(conds)} AND rank>100 ORDER BY year DESC,rank ASC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [{'province':r[0],'year':r[1],'school_name':r[2],'major_name':r[3],'score':r[4],'rank':r[5]} for r in rows]

def web_search(query, n=5):
    results = []
    OFFICIAL = {'北京':'bjeea.cn','天津':'zhaokao.net','河北':'hebeea.edu.cn','山西':'sxkszx.cn','内蒙古':'nm.zsks.cn','辽宁':'lnzsks.com','吉林':'jleea.edu.cn','黑龙江':'lzk.hl.cn','上海':'shmeea.edu.cn','江苏':'jseea.cn','浙江':'zjzs.net','安徽':'ahzsks.cn','福建':'eeafj.cn','江西':'jxeea.cn','山东':'sdzk.cn','河南':'haeea.cn','湖北':'hbea.edu.cn','湖南':'hneeb.cn','广东':'eeagd.edu.cn','广西':'gxeea.cn','海南':'hainanu.edu.cn','重庆':'cqksy.cn','四川':'sceea.cn','贵州':'zsksy.guizhou.gov.cn','云南':'ynzs.cn','西藏':'zsks.edu.xizang.gov.cn','陕西':'sneea.cn','甘肃':'ganseea.cn','青海':'qhjyks.com','宁夏':'nxjyks.cn','新疆':'xjzk.gov.cn'}
    sites = ['gaokao.chsi.com.cn','eol.cn']
    for prov, dom in OFFICIAL.items():
        if prov in query: sites.insert(0, dom); break
    for site in sites[:3]:
        if len(results) >= n: break
        try:
            sq = f'{query} site:{site}'
            url = 'https://www.baidu.com/s?wd=' + urllib.parse.quote(sq)
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=8) as resp:
                html = resp.read().decode('utf-8', errors='ignore')
            for pat in [r'class="c-abstract"[^>]*>(.*?)</span>', r'<span class="content-right_[^"]*">(.*?)</span>']:
                import re as _re
                for s in _re.findall(pat, html):
                    clean = _re.sub(r'<[^>]+>', '', s).strip()
                    if len(clean) > 20 and clean not in results:
                        results.append(f'[{site}] {clean[:300]}')
        except: continue
    if not results: results.append('未搜到结果')
    return results[:n]#!/usr/bin/env python3
"""雪峰Agent — 单文件服务器：HTML UI + API + 数据库查询"""
import os, re, json, sqlite3, gzip, shutil, urllib.request, urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from http.server import ThreadingHTTPServer
import sys
try: sys.stdout.reconfigure(encoding="utf-8")
except: pass

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, 'admission_clean.db')
GZ_PATH = os.path.join(HERE, 'admission_clean.db.gz')
if not os.path.exists(DB_PATH) and os.path.exists(GZ_PATH):
    with gzip.open(GZ_PATH, 'rb') as gz:
        with open(DB_PATH, 'wb') as f:
            shutil.copyfileobj(gz, f)

HAS_DB = os.path.exists(DB_PATH)
USER_XLSX = os.path.join(HERE, '自定义数据.xlsx')
USER_DATA = []
def clean_num(v):
    if v is None: return None
    s = str(v).strip()
    if not s: return None
    s = s.replace(',','').replace('，','').replace(' ','').replace('分','').replace('位','')
    try: return int(float(s))
    except: return None
def load_user_data():
    if not os.path.exists(USER_XLSX): return
    try: import openpyxl
    except ImportError: return
    global USER_DATA; USER_DATA = []
    if not os.path.exists(USER_XLSX): return
    try:
        wb = openpyxl.load_workbook(USER_XLSX, data_only=True)
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            school = str(row[0]).strip()
            if len(school) < 2 or school in ['学校名称','院校名称']: continue
            note = str(row[8]).strip() if row[8] else ''
            if '示例' in note or '不参与排序' in note: continue
            major = str(row[1]).strip() if row[1] else ''
            cat = str(row[2]).strip() if row[2] else ''
            if '物理' in cat: category = '物理类'
            elif '历史' in cat: category = '历史类'
            elif '综合' in cat: category = '综合'
            else: category = cat
            prov = str(row[3]).strip() if row[3] else ''
            prov = prov.replace('省','').replace('市','').strip()
            s24 = clean_num(row[4]); r24 = clean_num(row[5])
            s25 = clean_num(row[6]); r25 = clean_num(row[7])
            if s24 and s24 < 100 and r24 and r24 > 300: s24, r24 = r24, s24
            if s25 and s25 < 100 and r25 and r25 > 300: s25, r25 = r25, s25
            if s24 and r24: USER_DATA.append({'school':school,'major':major,'year':2024,'category':category,'score':s24,'rank':r24,'province':prov})
            if s25 and r25: USER_DATA.append({'school':school,'major':major,'year':2025,'category':category,'score':s25,'rank':r25,'province':prov})
        wb.close()
    except Exception as e: print(f'[user data] {e}')
load_user_data()

PROVINCES = ['北京','天津','上海','重庆','河北','山西','辽宁','吉林','黑龙江','江苏','浙江','安徽',
             '福建','江西','山东','河南','湖北','湖南','广东','广西','海南','四川','贵州','云南',
             '西藏','陕西','甘肃','青海','宁夏','新疆','内蒙古']

def query_db(province=None, school=None, major=None, limit=50):
    if not HAS_DB: return None
    conn = sqlite3.connect(DB_PATH)
    conds, params = [], []
    if province: conds.append("province LIKE ?"); params.append(f"%{province}%")
    if school: conds.append("school LIKE ?"); params.append(f"%{school}%")
    if major: conds.append("major LIKE ?"); params.append(f"%{major}%")
    if not conds: conn.close(); return None
    sql = f"SELECT province,year,school_name,major_name,score,rank FROM admission WHERE {' AND '.join(conds)} AND rank>100 ORDER BY year DESC,rank ASC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [{'province':r[0],'year':r[1],'school_name':r[2],'major_name':r[3],'score':r[4],'rank':r[5]} for r in rows]

def web_search(query, n=5):
    """百度搜索兜底 — 当 Tavily 不可用时使用"""
    results = []
    try:
        url = "https://www.baidu.com/s?wd=" + urllib.parse.quote(query)
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        with urllib.request.urlopen(req, timeout=8) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        # 提取搜索结果摘要
        snippets = re.findall(r'<span class="content-right_[^"]*">(.*?)</span>', html)
        for s in snippets[:n]:
            clean = re.sub(r'<[^>]+>', '', s).strip()
            if len(clean) > 20:
                results.append(clean[:300])
        if not results:
            # 备选：匹配任意摘要片段
            fallback = re.findall(r'class="c-abstract"[^>]*>(.*?)</span>', html)
            for s in fallback[:n]:
                clean = re.sub(r'<[^>]+>', '', s).strip()
                if len(clean) > 20:
                    results.append(clean[:300])
        if not results:
            results.append("百度搜索未返回可用结果，建议注册 Tavily Key（tavily.com 免费）以获得更精准的AI搜索。")
    except Exception as e:
        results.append(f"搜索暂不可用（{e}）。建议注册 Tavily Key（tavily.com 免费）以获得更精准的AI搜索。")
    return results if results else ["搜索无结果。请在前端API设置中填入Tavily Key以启用联网搜索（tavily.com免费注册）。"]

class Handler(BaseHTTPRequestHandler):
    def _send(self, data, code=200):
        self.send_response(code)
        self.send_header('Content-Type','application/json;charset=utf-8')
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Cache-Control','no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma','no-cache')
        self.send_header('Expires','0')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Methods','GET,OPTIONS')
        self.send_header('Access-Control-Allow-Headers','*')
        self.end_headers()

    def do_GET(self):
        if self.path == '/ping':
            return self._send({'ok':True,'db':HAS_DB})
        if self.path.startswith('/query'):
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            rows = query_db(qs.get('province',[''])[0], qs.get('school',[''])[0], qs.get('major',[''])[0])
            return self._send({'db':rows,'count':len(rows) if rows else 0})
        if self.path.startswith('/recommend'):
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            prov = qs.get('province',[''])[0]
            major = qs.get('major',[''])[0]
            keyword = qs.get('keyword',[''])[0]
            try: rank = int(qs.get('rank',['0'])[0])
            except: rank = 0
            try: score = int(qs.get('score',['0'])[0])
            except: score = 0
            print(f"[RECOMMEND] prov={prov} rank={rank} score={score} kw={keyword[:30] if keyword else 'none'}")
            if prov:
                has_ud = any(u.get('province','') in prov for u in USER_DATA if u.get('score'))
                if has_ud:
                    um_raw = {}
                    for u in USER_DATA:
                        if u.get('province','') not in prov: continue
                        if not u.get('score') or not u.get('rank'): continue
                        k = u['school'] + '|' + (u.get('major',''))
                        if k not in um_raw: um_raw[k] = {'school':u['school'],'major':u.get('major',''),'scores':[],'ranks':[],'years':[]}
                        um_raw[k]['scores'].append(u['score']); um_raw[k]['ranks'].append(u['rank']); um_raw[k]['years'].append(u['year'])
                    um_all = []
                    for k, v in um_raw.items():
                        r = v['ranks']; s = v['scores']; avg_sc = int(sum(s)/len(s)); avg_rk = int(sum(r)/len(r))
                        yr = v['major']
                        yr = v["major"]
                        if len(r)>=2:
                            if v["years"][0]==2024: yr += " [24:"+str(s[0])+"分/"+str(r[0])+"位 25:"+str(s[1])+"分/"+str(r[1])+"位]"
                            else: yr += " [24:"+str(s[1])+"分/"+str(r[1])+"位 25:"+str(s[0])+"分/"+str(r[0])+"位]"
                        elif len(r)==1: yr += " ["+str(v["years"][0])+":"+str(s[0])+"分/"+str(r[0])+"位]"
                        um_all.append({"school":v["school"],"major":yr,"score":avg_sc,"rank":avg_rk,"year":"综合","source":"user"})
                    um_all.sort(key=lambda x: x["rank"])
                    n = len(um_all)
                    ch = um_all[:n//3] if n>=3 else um_all
                    wn = um_all[n//3:2*n//3] if n>=3 else []
                    ba = um_all[2*n//3:] if n>=3 else []
                    ch.insert(0, {'school':'【死命令】','major':'只准推荐下面学校·不准推荐表外任何学校·补充建议提其他校须标注网络搜索仅供参考','score':0,'rank':0,'year':'','source':'system'})
                    return self._send({'rank':rank,'score':score,'chong':ch,'wen':wn,'bao':ba,'user_data':um_all,'source':'custom_only'})
            if prov and (rank > 0 or score > 0):
                conn = sqlite3.connect(DB_PATH)
                base = "province LIKE ? AND (score>0 OR rank>0)"
                bp = [f'%{prov}%']
                if major: base += " AND major_name LIKE ?"; bp.append(f'%{major}%')
                if keyword:
                    kws = keyword.split(',')
                    kw_conds = []
                    for kw in kws:
                        kw_conds.append("(major_name LIKE ? OR school_name LIKE ?)")
                        bp.append(f'%{kw}%'); bp.append(f'%{kw}%')
                    base += " AND (" + " OR ".join(kw_conds) + ")"

                chong = []; wen = []; bao = []

                # Try rank-based first, fall back to score-based
                if rank > 0:
                    chong = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} AND rank>0 AND rank<? AND rank>=? ORDER BY rank ASC LIMIT 50",
                        bp+[rank, max(1,int(rank*0.90))]).fetchall()]
                    wen = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 50",
                        bp+[rank, int(rank*1.3)]).fetchall()]
                    bao = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} AND rank>0 AND rank>? AND rank<=? ORDER BY rank ASC LIMIT 50",
                        bp+[int(rank*1.3), int(rank*1.6)]).fetchall()]

                # If no results with keyword, retry without keyword (broader search)
                if not (chong or wen or bao) and keyword:
                    chong = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE province LIKE ? AND rank>0 AND rank<? AND rank>=? ORDER BY rank ASC LIMIT 50",
                        [f'%{prov}%', rank, max(1,int(rank*0.90))]).fetchall()]
                    wen = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE province LIKE ? AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 50",
                        [f'%{prov}%', rank, int(rank*1.3)]).fetchall()]
                    bao = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE province LIKE ? AND rank>0 AND rank>? AND rank<=? ORDER BY rank ASC LIMIT 50",
                        [f'%{prov}%', int(rank*1.3), int(rank*1.6)]).fetchall()]

                # If rank query returned nothing, try score-based
                if not (chong or wen or bao) and score > 0:
                    # First try with keyword
                    chong = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} AND score>? AND score<=? ORDER BY score DESC LIMIT 80",
                        bp+[score, score+25]).fetchall()]
                    wen = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} AND score>=? AND score<=? ORDER BY score ASC LIMIT 50",
                        bp+[score-25, score+25]).fetchall()]
                    bao = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                        conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} AND score>=? AND score<? ORDER BY score ASC LIMIT 50",
                        bp+[score-50, score-25]).fetchall()]
                    # If keyword filtered everything, retry without keyword
                    if not (chong or wen or bao):
                        base2 = "province LIKE ? AND (score>0 OR rank>0)"
                        bp2 = [f'%{prov}%']
                        chong = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                            conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base2} AND score>? AND score<=? ORDER BY score DESC LIMIT 80",
                            bp2+[score, score+25]).fetchall()]
                        wen = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                            conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base2} AND score>=? AND score<=? ORDER BY score ASC LIMIT 50",
                            bp2+[score-25, score+25]).fetchall()]
                        bao = [{'school':r[0],'major':r[1],'score':r[2],'rank':r[3],'year':r[4]} for r in
                            conn.execute(f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base2} AND score>=? AND score<? ORDER BY score ASC LIMIT 50",
                            bp2+[score-50, score-25]).fetchall()]
                conn.close()
                return self._send({'rank':rank,'score':score,'chong':chong,'wen':wen,'bao':bao})
            return self._send({'error':'need province and rank or score'},400)
        if self.path == '/userdata':
            return self._send({'data':USER_DATA,'count':len(USER_DATA)})
        if self.path == '/reload_userdata':
            load_user_data()
            return self._send({'ok':True,'count':len(USER_DATA)})
        if self.path.startswith('/search'):
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            q = qs.get('q',[''])[0]
            if q: return self._send({'results':web_search(q)})
            return self._send({'results':[]})

        # Serve image files
        for img in ['img_suit.png','img_scifi.png']:
            if self.path == '/'+img:
                ip = os.path.join(HERE, img)
                if os.path.exists(ip):
                    self.send_response(200)
                    self.send_header('Content-Type','image/png')
                    self.send_header('Cache-Control','max-age=3600')
                    self.end_headers()
                    with open(ip,'rb') as f: self.wfile.write(f.read())
                    return

        # Serve the main UI page
        self.send_response(200)
        self.send_header('Content-Type','text/html;charset=utf-8')
        self.send_header('Cache-Control','no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma','no-cache')
        self.send_header('Expires','0')
        self.end_headers()
        self.wfile.write(HTML_PAGE.encode('utf-8'))

    def log_message(self, format, *args):
        msg = format%args if args else format
        if '/recommend' in msg or '/query' in msg or '/ping' in msg or '/search' in msg:
            print(f"[REQ] {msg}")


# ========== HTML 页面（从 index.html 加载）==========
with open(os.path.join(HERE, 'index.html'), 'r', encoding='utf-8') as _f:
    HTML_PAGE = _f.read()

def main():
    port = 8765
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    print(f'雪峰Agent: http://127.0.0.1:{port}/')
    print(f'数据库: {"已加载" if HAS_DB else "未找到"}')
    try: server.serve_forever()
    except KeyboardInterrupt: server.shutdown(); print('\n已停止')

if __name__ == '__main__': main()
