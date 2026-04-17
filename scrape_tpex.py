"""
興櫃處置預警系統 — Playwright 自動爬蟲更新
用法: python scrape_tpex.py
每次執行會從 TPEx 爬取最近30天的注意/處置資料並更新 HTML
"""
import json, re, os
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(SCRIPT_DIR, '興櫃處置預警系統.html')
IND_FILE = os.path.join(SCRIPT_DIR, 'industry_map.json')

def roc(dt):
    return f"{dt.year-1911}/{dt.month}/{dt.day}"

def west(dt):
    return f"{dt.year}/{dt.month:02d}/{dt.day:02d}"

def set_dates_and_query(page, start_dt, end_dt):
    """設定日期並查詢，回傳表格資料"""
    page.evaluate(f'''() => {{
        const s = document.querySelector('input[name="startDate"]');
        const e = document.querySelector('input[name="endDate"]');
        jQuery(s).data('value', '{west(start_dt)}');
        jQuery(e).data('value', '{west(end_dt)}');
        s.removeAttribute('readonly'); e.removeAttribute('readonly');
        s.value = '{roc(start_dt)}'; e.value = '{roc(end_dt)}';
        document.querySelectorAll('form')[2].querySelectorAll('input[name="type"]').forEach(r => {{ if (r.value === 'all') r.checked = true; }});
    }}''')
    page.click('button:has-text("查詢")')
    page.wait_for_timeout(4000)

    return page.evaluate(r'''() => {
        const table = document.querySelector('table');
        if (!table) return [];
        let data = [];
        table.querySelectorAll('tr').forEach((r, i) => {
            if (i === 0) return;
            const c = r.querySelectorAll('td');
            if (c.length < 6) return;
            const code = c[1].textContent.trim();
            if (!/^\d{4}$/.test(code)) return;
            const info = c[4].textContent;
            let clauses = [];
            if (info.includes('第一款')) clauses.push('1');
            if (info.includes('第二款')) clauses.push('2');
            if (info.includes('第三款')) clauses.push('3');
            if (info.includes('第四款')) clauses.push('4');
            data.push({
                code, name: c[2].textContent.trim(),
                date: c[5].textContent.trim(), clauses,
                price: parseFloat(c[6]?.textContent?.trim()) || 0
            });
        });
        return data;
    }''')

def scrape_all(page, url, start_dt, end_dt):
    """分批爬取日期範圍內所有資料"""
    page.goto(url, wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(3000)

    all_records = []
    seen = set()
    d = start_dt
    while d <= end_dt:
        d_end = min(d + timedelta(days=1), end_dt)
        batch = set_dates_and_query(page, d, d_end)
        for r in batch:
            key = f"{r['code']}_{r['date']}"
            if key not in seen:
                seen.add(key)
                all_records.append(r)
        n = len(batch)
        print(f"    {roc(d)}~{roc(d_end)}: {n} 筆")
        d = d_end + timedelta(days=1)
        while d.weekday() >= 5 and d <= end_dt:
            d += timedelta(days=1)

    return all_records

def scrape_disposal(page, start_dt, end_dt):
    """爬取處置資料"""
    page.goto("https://www.tpex.org.tw/zh-tw/announce/market/esb-disposal.html",
              wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(3000)

    page.evaluate(f'''() => {{
        const s = document.querySelector('input[name="startDate"]');
        const e = document.querySelector('input[name="endDate"]');
        jQuery(s).data('value', '{west(start_dt)}');
        jQuery(e).data('value', '{west(end_dt)}');
        s.removeAttribute('readonly'); e.removeAttribute('readonly');
        s.value = '{roc(start_dt)}'; e.value = '{roc(end_dt)}';
        document.querySelectorAll('form')[2].querySelectorAll('input[name="type"]').forEach(r => {{ if (r.value === 'all') r.checked = true; }});
    }}''')
    page.click('button:has-text("查詢")')
    page.wait_for_timeout(4000)

    return page.evaluate(r'''() => {
        const table = document.querySelector('table');
        if (!table) return [];
        let data = [];
        table.querySelectorAll('tr').forEach((r, i) => {
            if (i === 0) return;
            const c = r.querySelectorAll('td');
            if (c.length < 6 || !c[2].textContent.trim()) return;
            const code = c[2].textContent.trim();
            if (!/^\d{4}$/.test(code)) return;
            data.push({
                code, name: c[3].textContent.trim(),
                date: c[1].textContent.trim(),
                period: c[5].textContent.trim(),
                reason: c[6]?.textContent?.trim() || ''
            });
        });
        return data;
    }''')

def merge_into_html(attn, disp):
    """合併資料進 HTML"""
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        html = f.read()
    m = re.search(r'const DATA = (\{.*?\});', html, re.DOTALL)
    if not m:
        print("找不到 DATA"); return
    data = json.loads(m.group(1))

    # 載入產業對照表
    ind_map = {}
    if os.path.exists(IND_FILE):
        with open(IND_FILE, encoding='utf-8') as f:
            ind_map = json.load(f)

    # 合併注意紀錄
    new_tl = 0
    new_stocks = 0
    added_codes = set()
    for rec in attn:
        code_dot = rec['code'] + '.0'
        mmdd = rec['date'][-5:]
        has_23 = any(c in rec['clauses'] for c in ['2', '3'])
        has_234 = any(c in rec['clauses'] for c in ['2', '3', '4'])
        tl_entry = {'d': mmdd, 'c': rec['clauses'], 'a': has_23, 'b': has_234}

        stock = next((s for s in data['stocks'] if s['code'] == code_dot), None)
        if stock:
            if not any(t['d'] == mmdd for t in stock['tl']):
                stock['tl'].append(tl_entry)
                stock['tl'].sort(key=lambda x: x['d'])
                stock['n30'] = len(stock['tl'])
                stock['latest'] = f"2026/{mmdd}"
                stock['conds'] = rec['clauses']
                stock['price'] = rec['price']
                new_tl += 1
        elif code_dot not in added_codes:
            data['stocks'].append({
                'code': code_dot, 'name': rec['name'],
                'ind': ind_map.get(rec['code'], '未分類'),
                'price': rec['price'], 'latest': f"2026/{mmdd}",
                'conds': rec['clauses'], 'rA': 0, 'rB': 0, 'rem': 3,
                'tl': [tl_entry], 'n30': 1, 'disp': False
            })
            added_codes.add(code_dot)
            new_stocks += 1

    # 合併處置紀錄
    for rec in disp:
        code_dot = rec['code'] + '.0'
        d = f"2026/{rec['date'][-5:]}"
        if not any(x['code'] == code_dot and x['date'] == d for x in data['recent_disposals']):
            data['recent_disposals'].append({
                'code': code_dot, 'name': rec['name'],
                'date': d, 'reason': rec['reason']
            })

    data['date'] = datetime.now().strftime('%Y/%m/%d')
    data['total_stocks'] = len(data['stocks'])

    data_json = json.dumps(data, ensure_ascii=False)
    html = re.sub(r'const DATA = \{.*?\};', f'const DATA = {data_json};', html, count=1, flags=re.DOTALL)
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  新增 {new_tl} 筆時間線, {new_stocks} 檔新股票")
    print(f"  總計 {len(data['stocks'])} 檔")

def main():
    today = datetime.now()
    start = today - timedelta(days=30)

    print("啟動 Playwright 瀏覽器...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        print("爬取注意股票...")
        attn = scrape_all(page,
            "https://www.tpex.org.tw/zh-tw/announce/market/esb-attention.html",
            start, today)
        print(f"  共 {len(attn)} 筆注意紀錄")

        print("爬取處置股票...")
        disp = scrape_disposal(page, start, today)
        print(f"  共 {len(disp)} 筆處置紀錄")

        browser.close()

    if attn or disp:
        print("更新 HTML...")
        merge_into_html(attn, disp)
        print(f"\n=== 更新完成 ({today.strftime('%Y/%m/%d')}) ===")
    else:
        print("沒有新資料")

if __name__ == '__main__':
    main()
