# -*- coding: utf-8 -*-
"""
Created on Thu Dec 25 19:25:58 2025

@author: 88690
"""

from flask import Flask, render_template, jsonify
import pandas as pd
import os
import json

app = Flask(__name__)

EXCEL_FILE = 'schedule.xlsx'

def get_excel_data():
    if not os.path.exists(EXCEL_FILE):
        return []
    
    try:
        # 讀取 Excel 並將所有 NaN 替換為空字串
        df = pd.read_excel(EXCEL_FILE).fillna('')
        
        # 轉換日期格式
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date'])
        
        # 需求：定義不同 Shift 的顏色
        color_map = {
            '聖務': '#8e44ad', # 紫色 (Purple)
            '講經說法': '#3498db'  # 藍色 (Blue)
        }
        # 預設顏色 (若不屬於以上兩者)
        default_color = '#e74c3c' # 紅色 (Red)
        
        events = []
        for _, row in df.iterrows():
            # 處理 Name 與 Shift 的空值與字串轉換
            name = str(row['Name']).strip() if str(row['Name']).strip() != '' else " "
            shift_type = str(row['Shift']).strip() if str(row['Shift']).strip() != '' else " "
            note = str(row['Note']).strip() if str(row['Note']).strip() != '' else "無備註"
            
            # 根據 Shift 內容決定顏色
            bg_color = color_map.get(shift_type, default_color)
            
            events.append({
                'title': name,
                'start': row['Date'].strftime('%Y-%m-%d'),
                'color': bg_color,
                'textColor': '#ffffff',
                'extendedProps': {
                    'shift': shift_type,
                    'note': note
                }
            })
        return events
    except Exception as e:
        print(f"Excel 讀取錯誤: {e}")
        return []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/events')
def events():
    data = get_excel_data()
    return json.dumps(data, ensure_ascii=False), 200, {'Content-Type': 'application/json'}

if __name__ == '__main__':
    # host='0.0.0.0' 方便讓區域網路內的其他電腦連線
    app.run(debug=True, host='0.0.0.0', port=5000)