import os
import json
import base64
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import google.generativeai as genai
import gspread
from google.oauth2.credentials import Credentials
from dotenv import load_dotenv
from PIL import Image
import io

load_dotenv()

app = Flask(__name__, static_folder='.')
CORS(app)

# --- Helpers ---
def get_gspread_client(access_token):
    try:
        if not access_token:
            return None
        creds = Credentials(token=access_token)
        return gspread.authorize(creds)
    except Exception as e:
        print(f"Auth Error: {e}")
        return None

# --- Google Sheets Logic (History Only) ---
def get_accounting_history(spreadsheet_id, access_token):
    print("Fetching accounting history (Semantic Reference) from sheets...")
    try:
        client = get_gspread_client(access_token)
        if not client:
             return []

        sh = client.open_by_key(spreadsheet_id)
        sheet = sh.worksheet("仕訳明細")
        data = sheet.get_all_values()
        
        # 直近の履歴を取得（ヘッダーを除く最後から150件程度）
        history = []
        if len(data) > 1:
            recent_rows = data[1:][-150:]
            for row in recent_rows:
                if len(row) >= 6:
                    history.append({
                        "counterparty": row[4].strip(),
                        "memo": row[5].strip(),
                        "account": row[1].strip()
                    })
        return history
    except Exception as e:
        print(f"Error fetching history: {e}")
        return []

import csv
import io

# --- CSV Logic ---
def analyze_csv(csv_bytes, history=[]):
    print("Analyzing CSV statement with semantic memory...")
    try:
        text = csv_bytes.decode('shift_jis', errors='replace') 
        if '確定日' not in text and '利用日' not in text and ',' not in text:
             text = csv_bytes.decode('utf-8', errors='replace')

        f = io.StringIO(text)
        reader = csv.reader(f)
        rows = list(reader)
        
        csv_text = "\n".join([",".join(row) for row in rows[:50]])
        
        model = genai.GenerativeModel('gemini-2.0-flash-exp')
        
        history_str = "\n".join([f"- {h['counterparty']} ({h['memo']}) => {h['account']}" for h in history]) if history else "なし"
        
        prompt = f"""
        あなたは優秀な会計士です。明細データ（CSV）から仕訳を作成してください。
        このCSVは主に「クレジットカード利用明細」または「銀行入出金」です。

        ルール:
        1. **クレジットカード明細の場合**、貸方勘定科目は原則として「**未払金**」を使用してください。
        2. **銀行口座のCSVの場合**、貸方または借方は「普通預金」などが適切です。
        3. 過去の履歴を参考に、最適な勘定科目を選んでください。

        過去の仕訳履歴（参考）：
        {history_str}

        JSON形式（配列）で出力：
        [
          {{
            "date": "YYYY-MM-DD",
            "debit_account": "借方勘定科目",
            "credit_account": "貸方勘定科目",
            "amount": 数値,
            "counterparty": "取引先名",
            "memo": "詳細・摘要"
          }}
        ]
        
        明細データ:
        {csv_text}
        """
        
        response = model.generate_content(prompt)
        content = response.text.strip()
        if "```json" in content:
            content = content.split("```json")[-1].split("```")[0].strip()
        return json.loads(content)
    except Exception as e:
        print(f"Error in analyze_csv: {e}")
        return []

# --- Google Sheets Logic (Duplicate Check) ---
def get_existing_entries(spreadsheet_id, access_token):
    print("Fetching existing entries for duplicate check...")
    try:
        client = get_gspread_client(access_token)
        if not client:
             return set()

        sh = client.open_by_key(spreadsheet_id)
        
        # Check both Auto and Manual sheets
        existing = set()
        
        for sheet_name in ["仕訳明細", "仕訳明細（手入力）"]:
            try:
                sheet = sh.worksheet(sheet_name)
                data = sheet.get_all_values()
                if len(data) > 1:
                    for row in data[1:]:
                        if len(row) >= 5:
                            date = row[0].strip()
                            amount = str(row[3]).strip()
                            counterparty = row[4].strip()
                            existing.add(f"{date}_{amount}_{counterparty}")
            except:
                pass # Sheet might not exist yet
                
        return existing
    except Exception as e:
        print(f"Error fetching existing entries: {e}")
        return set()

# --- AI Logic ---
def analyze_document(file_bytes, mime_type, history=[]):
    print(f"Analyzing {mime_type} with semantic memory...")
    models_to_try = ['gemini-2.0-flash-exp', 'gemini-1.5-flash']
    
    history_str = json.dumps(history, ensure_ascii=False, indent=2) if history else "なし"

    for model_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name)
            
            prompt = f"""
            あなたは日本の税務・会計士です。
            渡された画像の「領収書（レシート）」または「請求書」から仕訳を作成してください。
            
            **重要：決済方法の判定**
            画像内の支払情報を確認し、貸方勘定科目を以下のように決定してください：
            - **クレジットカード払い、カード利用、VISA/JCB/Master等の記載がある場合** → 「**未払金**」
            - **電子マネー（PayPay, Suica等）、後払い決済の場合** → 「**未払金**」
            - 現金、Cash、または支払方法の記載がない場合 → 「現金」（または「小口現金」）
            - 銀行振込の請求書の場合 → 「買掛金」または「未払金」

            履歴（優先）:
            {history_str}

            JSON形式（配列）で出力してください。他の説明は不要です。
            [
              {{
                "date": "YYYY-MM-DD",
                "debit_account": "借方勘定科目",
                "credit_account": "貸方勘定科目（未払金/現金/買掛金など）",
                "amount": 数値,
                "counterparty": "取引先名",
                "memo": "詳細・摘要"
              }}
            ]
            """
            
            if mime_type.startswith('image/'):
                content_part = Image.open(io.BytesIO(file_bytes))
            else:
                content_part = {
                    "mime_type": mime_type,
                    "data": file_bytes
                }
                
            response = model.generate_content([prompt, content_part])
            
            content = response.text.strip()
            if "```json" in content:
                content = content.split("```json")[-1].split("```")[0].strip()
            elif "```" in content:
                content = content.split("```")[-1].split("```")[0].strip()
            
            start_idx = content.find("[")
            end_idx = content.rfind("]")
            if start_idx != -1 and end_idx != -1:
                content = content[start_idx:end_idx+1]
                
            return json.loads(content)
        except Exception as e:
            print(f"Error with {model_name}: {e}")
            continue
    return []

# --- Google Sheets Logic (Save Only) ---
def save_to_sheets(data, spreadsheet_id, access_token):
    print(f"Saving {len(data)} items to sheets...")
        
    try:
        client = get_gspread_client(access_token)
        if not client:
             print("Credentials not found or invalid token")
             return False

        sh = client.open_by_key(spreadsheet_id)
        
        # --- Sheet 1: 仕訳明細 (Auto) ---
        try:
            sheet1 = sh.worksheet("仕訳明細")
        except:
            sheet1 = sh.add_worksheet(title="仕訳明細", rows="1000", cols="6")
            
        headers = ["日にち", "借方", "貸方", "金額", "取引先", "摘要（内容）"]
        existing_values = sheet1.get_all_values()
        
        if not existing_values:
            sheet1.append_row(headers)
            sheet1.freeze(rows=1)
        elif existing_values[0] != headers:
            if not any(existing_values[0]):
                 sheet1.update('A1', [headers])
                 sheet1.freeze(rows=1)

        # --- Sheet 1.5: 仕訳明細（手入力） ---
        try:
            sheet_manual = sh.worksheet("仕訳明細（手入力）")
        except:
            sheet_manual = sh.add_worksheet(title="仕訳明細（手入力）", rows="1000", cols="6")
            sheet_manual.append_row(headers)
            sheet_manual.freeze(rows=1)

        # --- Sheet 2: 損益計算書 (P/L) ---
        try:
            sheet_pl = sh.worksheet("損益計算書")
        except gspread.exceptions.WorksheetNotFound:
            sheet_pl = sh.add_worksheet(title="損益計算書", rows="100", cols="10")
        
        # P/L Layout Reconstruction
        sheet_pl.clear()
        sheet_pl.update('A1', [["損益計算書 (P/L) - 月次推移なし・全体集計"]])
        sheet_pl.update('A3', [["【借方（費用・資産増加）】", "金額", "", "【貸方（収益・負債増加）】", "金額"]])
        
        # Formula for Debits (Expenses) - Aggregating from BOTH sheets
        debit_formula = "=QUERY({'仕訳明細'!A2:F; '仕訳明細（手入力）'!A2:F}, \"select Col2, sum(Col4) where Col2 is not null group by Col2 label sum(Col4) ''\", 0)"
        sheet_pl.update_acell('A4', debit_formula)
        
        # Formula for Credits (Revenue) - Aggregating from BOTH sheets
        credit_formula = "=QUERY({'仕訳明細'!A2:F; '仕訳明細（手入力）'!A2:F}, \"select Col3, sum(Col4) where Col3 is not null group by Col3 label sum(Col4) ''\", 0)"
        sheet_pl.update_acell('D4', credit_formula)
        
        sheet_pl.freeze(rows=3)
        
        # --- Sheet 3: 貸借対照表 (B/S) ---
        try:
            sheet_bs = sh.worksheet("貸借対照表")
        except:
            sheet_bs = sh.add_worksheet(title="貸借対照表", rows="100", cols="6")
            sheet_bs.update('A1', [["簡易貸借対照表 (B/S)"]])
            sheet_bs.update('A3', [["勘定科目", "借方合計", "貸方合計", "残高 (資産は正/負債は負)"]])
            
            # Simple aggregation for B/S reference
            sheet_bs.update('A4', [["※P/Lと同様に、仕訳明細から集計を行います。詳細なB/S作成には期首残高が必要です。"]])

        # データの書き込み (仕訳明細へ)
        rows = []
        for item in data:
            rows.append([
                str(item.get('date', '')),
                str(item.get('debit_account', '')),
                str(item.get('credit_account', '')),
                item.get('amount', 0),
                str(item.get('counterparty', '')),
                str(item.get('memo', ''))
            ])
        
        sheet1.append_rows(rows, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        print(f"Detailed Error saving to sheets: {str(e)}")
        return False

# --- Routes ---
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory('.', path)

@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    if 'files' not in request.files:
        return jsonify({"error": "No files uploaded"}), 400
    
    # Extract Config from Request
    api_key = request.form.get('gemini_api_key')
    spreadsheet_id = request.form.get('spreadsheet_id')
    access_token = request.form.get('access_token')

    if not api_key or not spreadsheet_id or not access_token:
        return jsonify({"error": "Missing configuration or authentication"}), 401

    # Configure Gemini per request
    genai.configure(api_key=api_key)
    
    history = get_accounting_history(spreadsheet_id, access_token)
    existing_entries = get_existing_entries(spreadsheet_id, access_token)
    
    files = request.files.getlist('files')
    all_results = []
    
    for file in files:
        filename = file.filename.lower()
        mime_type = file.content_type
        file_bytes = file.read()
        
        if filename.endswith('.csv'):
            results = analyze_csv(file_bytes, history)
        else:
            results = analyze_document(file_bytes, mime_type, history)
            
        for item in results:
            key = f"{item.get('date', '')}_{item.get('amount', '')}_{item.get('counterparty', '')}"
            if key in existing_entries:
                item['is_duplicate'] = True
            else:
                item['is_duplicate'] = False
            
        all_results.extend(results)
    
    return jsonify(all_results)

@app.route('/api/save', methods=['POST'])
def api_save():
    req_data = request.json
    
    data = req_data.get('data', [])
    api_key = req_data.get('gemini_api_key') # Not needed for save but consistent
    spreadsheet_id = req_data.get('spreadsheet_id')
    access_token = req_data.get('access_token')

    if not spreadsheet_id or not access_token:
        return jsonify({"error": "Missing configuration or authentication"}), 401
    
    success = save_to_sheets(data, spreadsheet_id, access_token)
    if success:
        return jsonify({"message": "Success"})
    else:
        return jsonify({"error": "Failed to save to sheets"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)

