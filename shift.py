import datetime
import jpholiday
import calendar
import csv
import os
import random
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 1. データ読み込み
# ==========================================

def load_staff_master(file_path):
    staff_list = []
    if not os.path.exists(file_path): return staff_list
    with open(file_path, mode='r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            staff_list.append({
                "id": int(row['id']),
                "name": str(row['name']).strip(),
                "can_kitchen": True if str(row['can_kitchen']) == '1' else False,
                "grade": int(row['grade'])
            })
    return staff_list

def load_hope_data_from_sheets(spreadsheet_id, year):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds_json = os.environ.get("GCP_CREDENTIALS")
    try:
        if creds_json:
            import json
            creds = service_account.Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
        else:
            creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='フォームの回答 1!A:C').execute()
        rows = result.get('values', [])
        
        hope_dict = {}
        if not rows: return hope_dict
        for row in rows[1:]:
            if len(row) < 3: continue
            name = row[1].strip()
            dates = row[2].replace(';', ',').split(',')
            if name not in hope_dict: hope_dict[name] = []
            for d_text in dates:
                nums = re.findall(r'\d+', d_text)
                if len(nums) >= 2:
                    try: hope_dict[name].append(datetime.date(year, int(nums[0]), int(nums[1])))
                    except: continue
        return hope_dict
    except: return {}

# ==========================================
# 2. ロジック（繁忙日判定含む）
# ==========================================

def generate_calendar():
    year, month = 2026, 3
    calendar_list = []
    for d in range(1, 32):
        try:
            target_date = datetime.date(year, month, d)
            # 翌日が休日か、当日が金土なら繁忙日
            tomorrow = target_date + datetime.timedelta(days=1)
            is_busy = target_date.weekday() in [4, 5] or jpholiday.is_holiday(tomorrow)
            calendar_list.append({
                "date": target_date,
                "weekday": ["月","火","水","木","金","土","日"][target_date.weekday()],
                "is_busy": is_busy,
                "req_k": 3 if is_busy else 2,
                "req_h": 2
            })
        except: break
    return year, month, calendar_list

def assign_shift(calendar_list, staff_members, hope_data):
    final_shift = {}
    total_count = {s['name']: 0 for s in staff_members}
    for day in calendar_list:
        target_date = day['date']
        available = [s for s in staff_members if target_date in hope_data.get(s['name'], [])]
        selected_today = []
        def pick(candidates, num):
            picked = []
            for _ in range(num):
                if not candidates: break
                candidates.sort(key=lambda x: (total_count[x['name']], 1 if x['grade'] in [p['grade'] for p in picked] else 0, random.random()))
                p = candidates.pop(0)
                picked.append(p)
                total_count[p['name']] += 1
            return picked
        k = pick([s for s in available if s['can_kitchen']], day['req_k'])
        selected_today.extend(k)
        h = pick([s for s in available if s['id'] not in [x['id'] for x in selected_today]], day['req_h'])
        selected_today.extend(h)
        final_shift[target_date] = selected_today
    return final_shift

# ==========================================
# 3. 出力（アクセント復活デザイン）
# ==========================================

def export_to_html(assigned_data, days, year, month, staff_members, hope_data):
    html_filename = f"shift_{year}_{month}.html"
    all_names = [s['name'] for s in staff_members]
    
    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {{ font-family: sans-serif; background: #f4f7f6; padding: 10px; }}
        .wrap {{ overflow-x: auto; background: white; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        table {{ width: 100%; border-collapse: collapse; min-width: 600px; }}
        th, td {{ border: 1px solid #eee; padding: 10px 5px; text-align: center; }}
        th {{ background: #007bff; color: white; position: sticky; top: 0; font-size: 13px; }}
        .work {{ background: #e7f3ff; color: #007bff; font-weight: bold; }}
        .can {{ color: #ccc; }}
        .busy-row {{ background: #fff4f4 !important; }}
        .hi {{ background: #ffeb3b !important; color: black !important; font-weight: bold; border: 2px solid #fbc02d; }}
        .btn {{ cursor: pointer; padding: 7px 14px; border-radius: 20px; border: 1px solid #007bff; color: #007bff; display: inline-block; margin: 3px; background: white; font-size: 13px; transition: 0.2s; }}
        .btn:hover {{ background: #007bff; color: white; }}
    </style></head><body>
    <h2 style="text-align:center; color: #333;">📅 {year}年{month}月 シフト表</h2>
    <div style="text-align:center; margin-bottom:15px;">"""
    
    for n in all_names: html += f'<span class="btn" onclick="highlight(\'{n}\')">{n}</span> '
    html += '</div><div class="wrap"><table><tr><th>日付</th>'
    for n in all_names: html += f'<th>{n}</th>'
    html += '</tr>'
    
    for d in days:
        # 繁忙日の行にクラスを適用
        cl = 'class="busy-row"' if d['is_busy'] else ""
        html += f'<tr {cl}><td>{d["date"].strftime("%m/%d")}({d["weekday"]})</td>'
        assigned = [m['name'] for m in assigned_data[d['date']]]
        for n in all_names:
            if n in assigned: html += f'<td class="work" data-n="{n}">出勤</td>'
            elif d['date'] in hope_data.get(n, []): html += f'<td class="can" data-n="{n}">◯</td>'
            else: html += f'<td data-n="{n}">-</td>'
        html += '</tr>'
    
    html += """</table></div><script>
    function highlight(n) {
        document.querySelectorAll('td').forEach(td => td.classList.remove('hi'));
        document.querySelectorAll('td[data-n="'+n+'"]').forEach(td => {
            if(td.innerText !== '-') td.classList.add('hi');
        });
    }
    </script></body></html>"""
    
    with open(html_filename, "w", encoding="utf-8") as f: f.write(html)

if __name__ == "__main__":
    sid = os.environ.get("SPREADSHEET_ID", "13xykI-3nJH91uWUbdvP-xDMGi2zjizKcJZEbDJoPAA4")
    staff = load_staff_master('staff_master.csv')
    y, m, days = generate_calendar()
    
    # ターゲットとなるファイル名
    target_html = f"shift_{y}_{m}.html"
    
    # 【重要】実行前に、リポジトリ内の古いHTMLをOSレベルで削除する
    if os.path.exists(target_html):
        os.remove(target_html)
        print(f"🗑️ 古いファイル {target_html} をリセットしました。")

    hope = load_hope_data_from_sheets(sid, y)
    
    if staff:
        assigned = assign_shift(days, staff, hope)
        # ここで新しいHTMLが「ゼロから」作られます
        export_to_html(assigned, days, y, m, staff, hope)
        print("✅ 現在のスプレッドシートの内容のみで再作成しました。")
