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
# 1. データ読み込みセクション
# ==========================================

def load_staff_master(file_path):
    staff_list = []
    if not os.path.exists(file_path): return staff_list
    try:
        with open(file_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                staff_list.append({
                    "id": int(row['id']),
                    "name": str(row['name']).strip(),
                    "can_kitchen": True if str(row['can_kitchen']) == '1' else False,
                    "rank": str(row['rank']).strip(),
                    "grade": int(row['grade'])
                })
        return staff_list
    except Exception as e:
        print(f"名簿読み込みエラー: {e}")
        return staff_list

def load_hope_data_from_sheets(spreadsheet_id, year):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds_json = os.environ.get("GCP_CREDENTIALS")
    
    try:
        if creds_json:
            import json
            info = json.loads(creds_json)
            creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        else:
            creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
            
        service = build('sheets', 'v4', credentials=creds)
        # A:D列（タイムスタンプ, 名前, 希望日, 時間指定メモ）を取得
        range_name = 'フォームの回答 1!A:D' 
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        rows = result.get('values', [])

        # 常に最新の回答だけで作り直すため、ここで辞書を初期化
        hope_dict = {} # { "名前": { 日付: "メモ" } }
        
        if not rows: return hope_dict

        # 同じ名前の人が複数回答した場合、下の行（最新）で上書きする仕組み
        for row in rows[1:]:
            if len(row) < 3: continue
            name = row[1].strip()
            dates_text = row[2].replace(';', ',')
            memo_text = row[3] if len(row) >= 4 else "" # D列のメモ
            
            if name not in hope_dict: hope_dict[name] = {}
            
            # 1. チェックボックスの日付を処理
            raw_dates = dates_text.split(',')
            for rd in raw_dates:
                nums = re.findall(r'\d+', rd)
                if len(nums) >= 2:
                    try:
                        d = datetime.date(year, int(nums[0]), int(nums[1]))
                        hope_dict[name][d] = "" # まずはメモなしで登録
                    except: continue
            
            # 2. D列のメモから時間を抽出して紐付け
            if memo_text:
                # メモ内の日付（例: 3/5）を探して、その日のメモとして保存
                memo_items = re.split(r'[,、\s\n]+', memo_text)
                for item in memo_items:
                    m_nums = re.findall(r'\d+', item)
                    if len(m_nums) >= 2:
                        try:
                            md = datetime.date(year, int(m_nums[0]), int(m_nums[1]))
                            if md in hope_dict[name]:
                                # 日付部分を除いた残りをメモとする
                                note = re.sub(r'\d+/\d+', '', item).strip()
                                hope_dict[name][md] = note
                        except: continue
        return hope_dict
    except Exception as e:
        print(f"APIエラー: {e}")
        return {}

# ==========================================
# 2. ロジックセクション
# ==========================================

def generate_calendar():
    year, month = 2026, 3
    calendar_list = []
    for d in range(1, 32):
        try:
            target_date = datetime.date(year, month, d)
            tomorrow = target_date + datetime.timedelta(days=1)
            is_busy = target_date.weekday() in [4, 5] or jpholiday.is_holiday(tomorrow)
            calendar_list.append({
                "date": target_date,
                "weekday": ["月","火","水","木","金","土","日"][target_date.weekday()],
                "is_busy": is_busy,
                "req_kitchen": 3 if is_busy else 2,
                "req_hall": 2
            })
        except: break
    return year, month, calendar_list

def assign_shift(calendar_list, staff_members, hope_data):
    final_shift = {}
    total_attendance = {s['name']: 0 for s in staff_members}
    
    for day in calendar_list:
        target_date = day['date']
        # 希望を出しているスタッフ（hope_dataにある人）
        available = [s for s in staff_members if target_date in hope_data.get(s['name'], {})]
        
        selected_today = []

        def pick_staff(candidates, num_needed):
            picked = []
            for _ in range(num_needed):
                if not candidates: break
                # 出勤数が少ない順 ＞ 学年分散 ＞ ランダム
                candidates.sort(key=lambda x: (
                    total_attendance[x['name']], 
                    1 if x['grade'] in [p['grade'] for p in picked] else 0,
                    random.random()
                ))
                p = candidates.pop(0)
                picked.append(p)
                total_attendance[p['name']] += 1
            return picked

        k_candidates = [s for s in available if s['can_kitchen']]
        selected_k = pick_staff(k_candidates, day['req_kitchen'])
        selected_today.extend(selected_k)

        already_ids = [s['id'] for s in selected_today]
        h_candidates = [s for s in available if s['id'] not in already_ids]
        selected_h = pick_staff(h_candidates, day['req_hall'])
        selected_today.extend(selected_h)

        final_shift[target_date] = selected_today
    return final_shift

# ==========================================
# 3. 出力セクション (HTML)
# ==========================================

def export_to_html(assigned_data, days, year, month, staff_members, hope_data):
    html_filename = f"shift_{year}_{month}.html"
    all_staff_names = [s['name'] for s in staff_members]
    
    html_start = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{ font-family: sans-serif; background: #f8f9fa; margin: 0; padding: 10px; font-size: 12px; }}
            .wrapper {{ max-width: 100%; overflow-x: auto; }}
            table {{ width: 100%; border-collapse: collapse; background: white; }}
            th, td {{ border: 1px solid #dee2e6; padding: 5px; text-align: center; min-width: 60px; }}
            th {{ background: #007bff; color: white; position: sticky; top: 0; }}
            .work {{ background: #e7f3ff; color: #007bff; font-weight: bold; }}
            .memo {{ font-size: 10px; color: #ff5722; display: block; }}
            .can-work {{ color: #ccc; }}
            .highlight-member {{ background: #ffeb3b !important; color: #000 !important; }}
            .staff-btn {{ cursor: pointer; padding: 4px 8px; border-radius: 12px; border: 1px solid #007bff; color: #007bff; display: inline-block; margin: 2px; }}
        </style>
    </head>
    <body>
        <h3 style="text-align:center;">📅 {year}年{month}月 シフト表</h3>
        <div style="text-align:center; margin-bottom:10px;">
    """
    
    buttons = "".join([f'<span class="staff-btn" onclick="hi(\'{n}\')">{n}</span> ' for n in all_staff_names])
    table_header = '<div class="wrapper"><table><tr><th>日付</th>' + "".join([f'<th>{n}</th>' for n in all_staff_names]) + '</tr>'
    
    table_rows = ""
    for d in days:
        assigned_names = [m['name'] for m in assigned_data[d['date']]]
        table_rows += f"<tr><td>{d['date'].strftime('%m/%d')}({d['weekday']})</td>"
        
        for name in all_staff_names:
            memo = hope_data.get(name, {}).get(d['date'], "")
            memo_html = f'<span class="memo">{memo}</span>' if memo else ""
            
            if name in assigned_names:
                table_rows += f'<td class="work" data-name="{name}">出勤{memo_html}</td>'
            elif d['date'] in hope_data.get(name, {}):
                table_rows += f'<td class="can-work" data-name="{name}">◯{memo_html}</td>'
            else:
                table_rows += f'<td data-name="{name}">-</td>'
        table_rows += "</tr>"

    html_end = """</table></div>
        <script>
            function hi(n) {
                document.querySelectorAll('td').forEach(td => td.classList.remove('highlight-member'));
                document.querySelectorAll('td[data-name="' + n + '"]').forEach(td => {
                    if(td.innerText !== '-') td.classList.add('highlight-member');
                });
            }
        </script>
    </body></html>
    """
    with open(html_filename, "w", encoding="utf-8") as f:
        f.write(html_start + buttons + table_header + table_rows + html_end)

if __name__ == "__main__":
    sid = os.environ.get("SPREADSHEET_ID", "13xykI-3nJH91uWUbdvP-xDMGi2zjizKcJZEbDJoPAA4")
    staff = load_staff_master('staff_master.csv')
    y, m, days = generate_calendar()
    hope = load_hope_data_from_sheets(sid, y)
    if staff:
        assigned = assign_shift(days, staff, hope)
        export_to_html(assigned, days, y, m, staff, hope)
        print("✅ 更新完了（最新回答のみ反映）")
