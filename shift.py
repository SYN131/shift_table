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
    """スタッフ名簿CSVを読み込む"""
    staff_list = []
    if not os.path.exists(file_path):
        print(f"【警告】名簿が見つかりません: {file_path}")
        return staff_list
    try:
        with open(file_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                staff_list.append({
                    "id": int(row['id']),
                    "name": str(row['name']).strip(),
                    "can_kitchen": True if str(row['can_kitchen']) == '1' else False,
                    "rank": "新人" if str(row['rank']) == 'begi' else "一般"
                })
        return staff_list
    except Exception as e:
        print(f"名簿読み込みエラー: {e}")
        return staff_list

def load_hope_data_from_sheets(spreadsheet_id, year):
    """Google Sheets APIから直接回答を取得する"""
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    SERVICE_ACCOUNT_FILE = 'credentials.json'

    try:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            print(f"【エラー】{SERVICE_ACCOUNT_FILE} が見つかりません。")
            return None

        creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        # 「フォームの回答 1」シートのA列(タイムスタンプ), B列(氏名), C列(希望日)を想定
        range_name = 'フォームの回答 1!A:C' 
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        rows = result.get('values', [])

        if not rows:
            print('⚠️ スプレッドシートにデータが見つかりません。')
            return None

        hope_dict = {}
        for row in rows[1:]: # ヘッダーをスキップ
            if len(row) < 3: continue
            name = row[1].strip()
            raw_text = row[2].replace(';', ',') # セミコロン区切り対策
            raw_dates = raw_text.split(',')

            if name not in hope_dict:
                hope_dict[name] = []

            for rd in raw_dates:
                nums = re.findall(r'\d+', rd)
                if len(nums) >= 2:
                    m_val, d_val = int(nums[0]), int(nums[1])
                    try:
                        hope_dict[name].append(datetime.date(year, m_val, d_val))
                    except: continue
        
        print(f"✅ Googleスプレッドシートから {len(hope_dict)} 名分の希望を取得しました。")
        return hope_dict
    except Exception as e:
        print(f"⚠️ Sheets API接続エラー: {e}")
        return None

# ==========================================
# 2. ロジック・カレンダー生成セクション
# ==========================================

def generate_calendar():
    """2026年3月のカレンダー枠を生成"""
    year, month = 2026, 3
    calendar_list = []
    # 3月は31日まで
    for d in range(1, 32):
        target_date = datetime.date(year, month, d)
        tomorrow = target_date + datetime.timedelta(days=1)
        # 金土、または翌日が祝日なら繁忙日
        is_busy = target_date.weekday() in [4, 5] or jpholiday.is_holiday(tomorrow)
        calendar_list.append({
            "date": target_date,
            "weekday": ["月","火","水","木","金","土","日"][target_date.weekday()],
            "is_busy": is_busy,
            "req_staff": 5 if is_busy else 4,
            "req_kitchen": 3 if is_busy else 2,
            "holiday": jpholiday.is_holiday_name(target_date)
        })
    return year, month, calendar_list

def assign_shift(calendar_list, staff_members, hope_data):
    """自動割り当て実行"""
    final_shift = {}
    for day in calendar_list:
        target_date = day['date']
        # その日の希望者
        available_staff = [s for s in staff_members if target_date in hope_data.get(s['name'], [])]

        # キッチン担当
        k_candidates = [s for s in available_staff if s['can_kitchen']]
        selected_k = random.sample(k_candidates, min(len(k_candidates), day['req_kitchen']))
        
        # ホール担当
        already_in = [s['name'] for s in selected_k]
        h_candidates = [s for s in available_staff if s['name'] not in already_in]
        needed_h = day['req_staff'] - len(selected_k)
        selected_h = random.sample(h_candidates, min(len(h_candidates), max(0, needed_h)))
        
        final_shift[target_date] = selected_k + selected_h
    return final_shift

# ==========================================
# 3. 出力セクション (Excel / HTML)
# ==========================================

def export_to_excel(assigned_data, days, year, month):
    """Excelファイル出力"""
    output_list = []
    for d in days:
        date_obj = d['date']
        members = assigned_data[date_obj]
        member_names = ", ".join([m['name'] for m in members]) if members else "(希望者なし)"
        shortage = d['req_staff'] - len(members)
        output_list.append({
            "日付": date_obj, "曜日": d['weekday'], "状態": "繁忙日" if d['is_busy'] else "通常",
            "祝日": d['holiday'], "出勤メンバー": member_names,
            "必要人数": d['req_staff'], "現在の人数": len(members), "不足人数": max(0, shortage)
        })
    
    df = pd.DataFrame(output_list)
    filename = f"shift_{year}_{month}.xlsx"
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='シフト表')
            ws = writer.sheets['シフト表']
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['E'].width = 45
        print(f"✅ Excelファイル '{filename}' を作成しました。")
    except PermissionError:
        print(f"❌ エラー: {filename} を閉じてから再実行してください。")

def export_to_html(assigned_data, days, year, month, staff_members):
    """スマホ用ハイライト機能付きHTML出力"""
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
            .wrapper {{ max-width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }}
            h2 {{ text-align: center; color: #333; margin-top: 20px; }}
            table {{ width: 100%; border-collapse: collapse; background: white; min-width: 600px; }}
            th, td {{ border: 1px solid #dee2e6; padding: 10px; text-align: center; white-space: nowrap; }}
            th {{ background: #007bff; color: white; position: sticky; top: 0; z-index: 2; }}
            .sticky-col {{ position: sticky; left: 0; background: #f8f9fa; z-index: 3; border-right: 2px solid #dee2e6; }}
            th.sticky-col {{ z-index: 4; background: #0056b3; }}
            tr:nth-child(even) {{ background: #f2f2f2; }}
            .work {{ background: #e7f3ff; color: #007bff; font-weight: bold; }}
            .rest {{ color: #ccc; font-size: 0.8em; }}
            .busy-row {{ background: #fff4f4 !important; }}
            .sun {{ color: #d9534f; }} .sat {{ color: #007bff; }}
            .holiday-name {{ font-size: 10px; display: block; color: #d9534f; }}
            .highlight-member {{ background: #ffeb3b !important; color: #000 !important; border: 2px solid #fbc02d !important; }}
            .staff-name-btn {{ 
                cursor: pointer; padding: 8px 12px; border-radius: 20px; background: white; 
                display: inline-block; margin: 4px; border: 1px solid #007bff; color: #007bff;
                font-size: 13px; transition: 0.2s;
            }}
            .instructions {{ text-align: center; margin-bottom: 15px; color: #666; font-size: 12px; }}
        </style>
        <title>{year}年{month}月シフト表</title>
    </head>
    <body>
        <h2>📅 {year}年{month}月 シフト一覧表</h2>
        <p class="instructions">👇 名前をタップで自分の出勤日をハイライト！</p>
        <div style="text-align:center; margin-bottom:15px;">
    """
    
    buttons = "".join([f'<span class="staff-name-btn" onclick="highlight(\'{n}\')">{n}</span> ' for n in all_staff_names])
    table_header = '<div class="wrapper"><table id="shiftTable"><tr><th class="sticky-col">日付</th>' + "".join([f'<th>{n}</th>' for n in all_staff_names]) + '</tr>'
    
    table_rows = ""
    for d in days:
        assigned_names = [m['name'] for m in assigned_data[d['date']]]
        day_cl = "sun" if d['weekday'] == "日" else ("sat" if d['weekday'] == "土" else "")
        row_st = 'class="busy-row"' if d['is_busy'] else ""
        h_txt = f'<span class="holiday-name">{d["holiday"]}</span>' if d['holiday'] else ""
        
        table_rows += f"<tr {row_st}><td class='sticky-col {day_cl}'><b>{d['date'].strftime('%m/%d')}</b>({d['weekday']}){h_txt}</td>"
        for name in all_staff_names:
            if name in assigned_names:
                table_rows += f'<td class="work" data-name="{name}">出勤</td>'
            else:
                table_rows += f'<td class="rest" data-name="{name}">-</td>'
        table_rows += "</tr>"

    html_end = """
                </table>
            </div>
        <script>
            function highlight(name) {
                document.querySelectorAll('td').forEach(td => td.classList.remove('highlight-member'));
                document.querySelectorAll('td[data-name="' + name + '"]').forEach(td => {
                    if(td.innerText === '出勤') {
                        td.classList.add('highlight-member');
                    }
                });
            }
        </script>
    </body>
    </html>
    """
    
    with open(html_filename, "w", encoding="utf-8") as f:
        f.write(html_start + buttons + table_header + table_rows + html_end)
    print(f"✅ 多機能HTML '{html_filename}' を作成しました。")

# ==========================================
# 4. メイン実行部
# ==========================================

if __name__ == "__main__":
    print("--- シフト作成システム 起動 ---")
    
    # 【重要設定】スプレッドシートのIDをここに貼り付け
    SPREADSHEET_ID = "13xykI-3nJH91uWUbdvP-xDMGi2zjizKcJZEbDJoPAA4"
    STAFF_CSV = 'staff_master.csv'
    
    # 処理開始
    staff_members = load_staff_master(STAFF_CSV)
    year, month, days = generate_calendar()
    hope_data = load_hope_data_from_sheets(SPREADSHEET_ID, year)

    if not staff_members:
        print("名簿データがないため終了します。")
    elif hope_data is None:
        print("スプレッドシートの取得に失敗しました。")
    else:
        # 自動計算
        assigned_data = assign_shift(days, staff_members, hope_data)

        # 出力
        export_to_excel(assigned_data, days, year, month)
        export_to_html(assigned_data, days, year, month, staff_members)

    print("\n--- 全ての処理が完了しました ---")