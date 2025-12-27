# -*- coding: utf-8 -*-
"""
Created on Thu Dec 25 19:25:58 2025

@author: 88690
"""
from flask import Flask, render_template, jsonify
import pandas as pd
import os
import json
from datetime import datetime, timedelta

app = Flask(__name__)

EXCEL_FILE = 'schedule.xlsx'

def get_excel_data():
    if not os.path.exists(EXCEL_FILE): return []
    try:
        df = pd.read_excel(EXCEL_FILE).fillna('')
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date'])
        
        color_map = {'聖務': '#8e44ad', '講經說法': '#3498db', '科儀活動': '#9b59b6'}
        
        events = []
        for _, row in df.iterrows():
            name = str(row['Name']).strip()
            shift_type = str(row['Shift']).strip()
            note = str(row['Note']).strip() or "無備註"
            
            display_title = name if name != "" else shift_type
            if display_title == "": display_title = "未命名活動"
            
            bg_color = color_map.get(shift_type, '#e74c3c')
            class_name = 'fc-event-neon' if shift_type == '科儀活動' else ''
            
            events.append({
                'title': display_title,
                'start': row['Date'].strftime('%Y-%m-%d'),
                'color': bg_color,
                'className': class_name,
                'textColor': '#ffffff',
                'extendedProps': { 'name': name, 'shift': shift_type, 'note': note }
            })
        return events
    except Exception as e:
        print(f"Excel 錯誤: {e}"); return []

@app.route('/')
def index():
    marquee_messages = []
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE).fillna('')
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            now = datetime.now()
            rituals = df[df['Shift'] == '科儀活動']
            for _, row in rituals.iterrows():
                if pd.isnull(row['Date']): continue
                delta = (row['Date'] - now).days
                roc_year = row['Date'].year - 1911
                date_str = f"民國 {roc_year} 年 {row['Date'].month} 月 {row['Date'].day} 日"
                note = str(row['Note'])
                highlight_note = f"<span style='color: #8d4b3d; font-weight: 800; border-bottom: 2px solid #e9ecef;'>{note}</span>"
                addr = "地點 : 台南市歸仁區仁愛五街34號"

                if 0 <= delta < 15:
                    msg = f"🏮 台南道場誠摯邀請十方大眾共襄盛舉🏮於 {date_str} {highlight_note} {addr}"
                    marquee_messages.append(msg)
                elif 15 <= delta < 45:
                    msg = f"🙏 歡迎蒞臨🙏於 {date_str} 參加 {highlight_note} | {addr}"
                    marquee_messages.append(msg)
                elif 45 <= delta < 90:
                    msg = f"✨ 即將到來✨ {date_str} {highlight_note}"
                    marquee_messages.append(msg)
        except: pass
    final_marquee = "　　　✦　　　".join(marquee_messages) if marquee_messages else "歡迎蒞臨 聖蓮宮 台南道場"
    return render_template('index.html', marquee_text=final_marquee)

@app.route('/api/events')
def events():
    return json.dumps(get_excel_data(), ensure_ascii=False), 200, {'Content-Type': 'application/json'}

if __name__ == '__main__':
    # Render 會自動分配 PORT 環境變數
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host='0.0.0.0', port=port)