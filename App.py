import re, os, pandas as pd
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

app = Flask(__name__)

# โหลดไฟล์ Excel ครั้งเดียวตอนเปิดโปรแกรม
# หาไฟล์ที่ขึ้นต้นด้วย All_MAT อัตโนมัติ ไม่ว่านามสกุลจะเป็นอะไร
import glob, os
excel_files = glob.glob("All_MAT*")
if not excel_files:
    print("❌ หาไฟล์ All_MAT ไม่เจอ! เช็คว่าวางไฟล์ Excel ไว้ในโฟลเดอร์เดียวกับ app.py หรือยัง")
    raise SystemExit(1)
excel_path = excel_files[0]
print(f"✅ เจอไฟล์ Excel: {excel_path}")
df = pd.read_excel(excel_path, sheet_name="item", dtype=str).fillna("")
df['CODE_SEARCH'] = df['Item Code'].str.upper().str.replace(r'\s+', '', regex=True)

# === ใส่ของตัวเองตรงนี้ ===
import os
CHANNEL_ACCESS_TOKEN = os.environ.get("CHANNEL_ACCESS_TOKEN", "J6DT3zSxZwwH7isgQqHmBp8RbSf70FNwI7C8JeQ6mc6oDuTK8s8CJGPh5DE3fjPymTM62PuW2P0bHD4u5FJowKitHOy5IQFJ884E5nAIpgvOCtMDfNyHJGYuqCOM8IawRrKwi1wlcOPwcCp1BZvaoQdB04t89/1O/w1cDnyilFU=")
CHANNEL_SECRET = os.environ.get("CHANNEL_SECRET", "5e9789299f9ab74f6752e2a03e5907b7")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

def clean_mat(text):
    # ตัดแค่คำว่า MAT อะไร / คืออะไร ออก แต่เก็บชื่อของไว้ครบๆ เช่น IMM - 7x50 2-PT 100GE CFP
    t = text.strip()
    t = re.sub(r'\bMAT\b', '', t, flags=re.IGNORECASE)
    t = re.sub(r'อะไร|คืออะไร|คือของอะไร|คือ', '', t)
    t = re.sub(r'\?', '', t)
    return t.strip()

def normalize(s):
    # ตัดช่องว่าง ขีด _ - . ออกให้หมด เพื่อให้ 4850G1 = R4850G1 = MR4850G1 เจอเหมือนกัน
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

df['NAME_SEARCH'] = df['Item Name'].apply(normalize)
df['CODE_SEARCH_NORM'] = df['CODE_SEARCH'].apply(normalize)

def search_mat(code_raw):
    code = normalize(code_raw)
    if not code:
        return None, 'not_found'
    # 1. เจอ Item Code ตรงๆ
    exact = df[df['CODE_SEARCH_NORM'] == code]
    if not exact.empty:
        return exact.iloc[0], 'exact'
    # 2. เจอในชื่อของแบบตรงๆ (4850G1)
    if len(code) >= 4:
        in_name = df[df['NAME_SEARCH'].str.contains(code, na=False)]
        if not in_name.empty:
            in_name = in_name.copy()
            in_name['LEN'] = in_name['Item Name'].str.len()
            in_name = in_name.sort_values('LEN')
            return in_name.head(5), 'in_name'
    # 3. ค้นแบบแยกคำ (FAN MDB 48V 0.60A) - ให้พิมพ์ผิดนิดหน่อยก็เจอ
    # แยกคำที่พิมพ์มาเป็นชิ้นๆ แล้วให้คะแนนว่าชื่อของชิ้นไหนมีคำพวกนี้เยอะสุด
    tokens = [t for t in re.split(r'[^A-Z0-9]+', code_raw.upper()) if len(t) >= 2]
    # ตัดคำทั่วไปทิ้ง
    stop = {'MAT', 'FAN', 'FOR', 'THE', 'AND', 'CABINET'}
    # FAN เก็บไว้เพราะเป็นคำสำคัญ แต่ถ้ามีคำอื่นด้วยจะให้คะแนนรวม
    if len(tokens) >= 2:
        def score_row(name_norm):
            s = 0
            for tok in tokens:
                tok_norm = normalize(tok)
                if tok_norm and tok_norm in name_norm:
                    s += len(tok_norm)
                elif len(tok_norm) >= 3 and tok_norm[:3] in name_norm:
                    # เจอแค่ 3 ตัวแรก เช่น MDB vs NMB ยังไม่ให้คะแนนเต็ม แต่ให้คะแนนนิดหน่อยถ้าใกล้เคียง
                    s += 1
            return s
        df['_SCORE'] = df['NAME_SEARCH'].apply(score_row)
        # ต้องเจออย่างน้อย 2 คำ หรือคะแนนรวม >= 6
        best = df[df['_SCORE'] >= 6].sort_values('_SCORE', ascending=False).head(5)
        if not best.empty:
            # ถ้าคำที่พิมพ์มาตรงกับของชิ้นเดียวจริงๆ ให้ถือว่าเจอ
            return best, 'in_name'
    # 4. เจอ Item Code ขึ้นต้นเหมือนกัน
    like = df[df['CODE_SEARCH_NORM'].str.startswith(code)]
    if not like.empty:
        return like.head(10), 'like'
    return None, 'not_found' 
    # 1. เจอ Item Code ตรงๆ
    exact = df[df['CODE_SEARCH_NORM'] == code]
    if not exact.empty:
        return exact.iloc[0], 'exact'
    # 2. เจอในชื่อของ (Item Name) - อันนี้คือเคส 4850G1, IMM 7x50
    # ต้องยาว >= 4 ตัวอักษร ไม่งั้นเจอเยอะเกิน
    if len(code) >= 4:
        in_name = df[df['NAME_SEARCH'].str.contains(code, na=False)]
        if not in_name.empty:
            # เอาตัวที่ชื่อสั้นที่สุดขึ้นก่อน เพราะตรงสุด
            in_name = in_name.copy()
            in_name['LEN'] = in_name['Item Name'].str.len()
            in_name = in_name.sort_values('LEN')
            return in_name.head(5), 'in_name'
    # 3. เจอ Item Code ขึ้นต้นเหมือนกัน
    like = df[df['CODE_SEARCH_NORM'].str.startswith(code)]
    if not like.empty:
        return like.head(10), 'like'
    return None, 'not_found'

def format_item(row):
    return (
        f"MAT: {row['Item Code']}\n"
        f"ชื่อ: {row['Item Name']}\n"
        f"Model: {row['Model']} | Vendor: {row['Vendor']}\n"
        f"Domain: {row['Domain']} | Unit: {row['Unit']}"
    )

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = event.message.text
    code = clean_mat(user_text)
    if not code:
        reply = "พิมพ์รหัส MAT มาเลยครับ เช่น 4850G1 หรือ 4850G1 MAT อะไร"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
        return
    result, mode = search_mat(code)
    if mode == 'exact':
        reply = "เจอแล้วครับ ✅\n\n" + format_item(result)
    elif mode == 'in_name':
        lines = []
        for i, (_, r) in enumerate(result.iterrows()):
            lines.append(f"{i+1}. MAT: {r['Item Code']}\n   ชื่อ: {r['Item Name'][:70]}\n   Model: {r['Model']} | Vendor: {r['Vendor']}")
        reply = f"เจอในชื่อของครับ ✅ มี {len(result)} รายการที่ชื่อมีคำว่า '{code}':\n\n" + "\n\n".join(lines) + "\n\nถ้าไม่ใช่ตัวที่หา พิมพ์รหัส Item Code เต็มๆ มาอีกรอบได้ครับ"
    elif mode == 'like':
        lines = [f"{i+1}. {r['Item Code']} - {r['Item Name'][:50]}" for i, (_, r) in enumerate(result.iterrows())]
        reply = f"ไม่เจอ {code} ตรงๆ แต่เจอใกล้เคียง {len(result)} รายการครับ:\n\n" + "\n".join(lines) + "\n\nพิมพ์รหัสเต็มๆ มาอีกรอบได้ไหมครับ?"
    else:
        reply = "เช็คหน่อยว่า:" + chr(10) + "1. พิมพ์ชื่อครบไหม?" + chr(10) + "2. เป็นของใหม่ที่ยังไม่เข้าระบบหรือเปล่า?"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
