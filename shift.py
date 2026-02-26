import datetime
import jpholiday
import calendar
import csv
import os
import random
import re
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 1. データ読み込みセクション
# ==========================================

def load_staff_master(file_path):
    staff_list = []
    if not os.path.exists(file_path):
        return staff_list
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
    # (既存のSheets API読み込み処理と同じため省略してもOKですが、環境に合わせてsecretsから読み込む形にします)
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
        range_name = 'フォームの回答 1!A:C' 
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        rows = result.get('values', [])

        hope_dict = {}
        if not rows: return hope_dict

        for row in rows[1:]:
            if len(row) < 3: continue
            name = row[1].strip()
            raw_text = row[2].replace(';', ',')
            raw_dates = raw_text.split(',')
            if name not in hope_dict: hope_dict[name] = []
            for rd in raw_dates:
                nums = re.findall(r'\d+', rd)
                if len(nums) >= 2:
                    try: hope_dict[name].append(datetime.date(year, int(nums[0]), int(nums[1])))
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
        target_date = datetime.date(year, month, d)
        tomorrow = target_date + datetime.timedelta(days=1)
        is_busy = target_date.weekday() in [4, 5] or jpholiday.is_holiday(tomorrow)
        calendar_list.append({
            "date": target_date,
            "weekday": ["月","火","水","木","金","土","日"][target_date.weekday()],
            "is_busy": is_busy,
            "req_kitchen": 3 if is_busy else 2,
            "req_hall": 2, # 繁忙日(3+2=5), 通常日(2+2=4)
            "holiday": jpholiday.is_holiday_name(target_date)
        })
    return year, month, calendar_list

def assign_shift(calendar_list, staff_members, hope_data):
    final_shift = {}
    attendance_count = {s['name']: 0 for s in staff_members}
    
    for day in calendar_list:
        target_date = day['date']
        # 週の始まり(月曜)に出勤数をリセット気味に調整（週2回確保のため）
        if target_date.weekday() == 0:
            for name in attendance_count: attendance_count[name] = 0

        # その日の希望者
        available = [s for s in staff_members if target_date in hope_data.get(s['name'], [])]
        
        selected_today = []

        # 割り当て用内部関数（学年分散と週2回優先）
        def pick_staff(candidates, num_needed):
            picked = []
            for _ in range(num_needed):
                if not candidates: break
                # スコア付け: 出勤数が少ないほど、かつ現在の選択メンバーと学年が違うほどスコアが高い
                candidates.sort(key=lambda x: (
                    attendance_count[x['name']], 
                    1 if x['grade'] in [p['grade'] for p in picked] else 0,
                    random.random()
                ))
                p = candidates.pop(0)
                picked.append(p)
                attendance_count[p['name']] += 1
            return picked

        # 1. キッチン選出
        k_candidates = [s for s in available if s['can_kitchen']]
        selected_k = pick_staff(k_candidates, day['req_kitchen'])
        selected_today.extend(selected_k)

        # 2. ホール選出
        already_ids = [s['id'] for s in selected_today]
        h_candidates = [s for s in available if s['id'] not in already_ids]
        selected_h = pick_staff(h_candidates, day['req_hall'])
        selected_today.extend(selected_h)

        final_shift[target_date] = selected_today
        
    return final_shift

# ==========================================
# 3. 出力セクション (HTML)
# ==========================================

def export_to_html(assigned_data, days, year, month, staff_members):
    html_filename = f"shift_{year}_{month}.html"
    all_staff_names = [s['name'] for s in staff_members]
    
    html_start = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{ font-family: sans-serif; background: #f8f9fa; margin: 0; padding: 10px; font-size: 14px; }}
            .wrapper {{ max-width: 100%; overflow-x: auto; }}
            h2 {{ text-align: center; color: #333; }}
            table {{ width: 100%; border-collapse: collapse; background: white; }}
            th, td {{ border: 1px solid #dee2e6; padding: 8px; text-align: center; }}
            th {{ background: #007bff; color: white; position: sticky; top: 0; }}
            .work {{ background: #e7f3ff; color: #007bff; font-weight: bold; }}
            .busy-row {{ background: #fff4f4 !important; }}
            .highlight-member {{ background: #ffeb3b !important; color: #000 !important; }}
            .staff-name-btn {{ cursor: pointer; padding: 5px 10px; border-radius: 15px; background: white; border: 1px solid #007bff; color: #007bff; display: inline-block; margin: 2px; }}
        </style>
    </head>
    <body>
        <h2>📅 {year}年{month}月 シフト表</h2>
        <div style="text-align:center; margin-bottom:10px;">
    """
    
    buttons = "".join([f'<span class="staff-name-btn" onclick="highlight(\'{n}\')">{n}</span> ' for n in all_staff_names])
    table_header = '<div class="wrapper"><table><tr><th>日付</th>' + "".join([f'<th>{n}</th>' for n in all_staff_names]) + '</tr>'
    
    table_rows = ""
    for d in days:
        assigned_names = [m['name'] for m in assigned_data[d['date']]]
        row_cl = 'class="busy-row"' if d['is_busy'] else ""
        table_rows += f"<tr {row_cl}><td>{d['date'].strftime('%m/%d')}({d['weekday']})</td>"
        for name in all_staff_names:
            if name in assigned_names:
                table_rows += f'<td class="work" data-name="{name}">出勤</td>'
            else:
                table_rows += f'<td data-name="{name}">-</td>'
        table_rows += "</tr>"

    html_end = """</table></div>
        <script>
            function highlight(name) {
                document.querySelectorAll('td').forEach(td => td.classList.remove('highlight-member'));
                document.querySelectorAll('td[data-name="' + name + '"]').forEach(td => {
                    if(td.innerText === '出勤') td.classList.add('highlight-member');
                });
            }
        </script>
    </body></html>
    """
    with open(html_filename, "w", encoding="utf-8") as f:
        f.write(html_start + buttons + table_header + table_rows + html_end)

# ==========================================
# 4. メイン実行部
# ==========================================

if __name__ == "__main__":
    SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "13xykI-3nJH91uWUbdvP-xDMGi2zjizKcJZEbDJoPAA4")
    STAFF_CSV = 'staff_master.csv'
    
    staff_members = load_staff_master(STAFF_CSV)
    year, month, days = generate_calendar()
    hope_data = load_hope_data_from_sheets(SPREADSHEET_ID, year)

    if staff_members and hope_data:
        assigned_data = assign_shift(days, staff_members, hope_data)
        export_to_html(assigned_data, days, year, month, staff_members)
        print("✅ シフト更新完了")
