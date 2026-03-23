#!/usr/bin/env python3
"""
Flask 簡化版應用：當 Streamlit 不可用時的備用方案
輕量級、依賴少、相容性好
"""
import os
import sys
import webbrowser
import threading
from pathlib import Path
from werkzeug.utils import secure_filename
from flask import Flask, render_template_string, request, jsonify, send_file
import pandas as pd
import io

# 匯入核心分析函數
sys.path.insert(0, str(Path(__file__).parent))
from functions import parse_transcript, extract_text_from_uploaded_file

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 上傳限制
app.config['UPLOAD_FOLDER'] = Path(__file__).parent / 'temp_uploads'
app.config['UPLOAD_FOLDER'].mkdir(exist_ok=True)

# 簡單的 HTML 模板
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>立法委員發言統計系統 - Flask 簡化版</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 800px;
            width: 100%;
            padding: 40px;
        }
        h1 {
            color: #333;
            margin-bottom: 10px;
            font-size: 28px;
        }
        .subtitle {
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .warning {
            background: #fff3cd;
            border: 1px solid #ffc107;
            color: #856404;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 13px;
        }
        .section {
            margin-bottom: 25px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            color: #333;
            font-weight: 500;
        }
        input[type="file"], select {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 6px;
            font-size: 14px;
            font-family: inherit;
        }
        input[type="file"]:focus, select:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        button {
            background: #667eea;
            color: white;
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            width: 100%;
            transition: background 0.3s;
        }
        button:hover {
            background: #764ba2;
        }
        button:disabled {
            background: #ccc;
            cursor: not-allowed;
        }
        .progress {
            display: none;
            margin-top: 15px;
            text-align: center;
        }
        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .result {
            display: none;
            margin-top: 20px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 6px;
        }
        .result h3 {
            color: #28a745;
            margin-bottom: 10px;
        }
        .error {
            display: none;
            color: #dc3545;
            margin-top: 15px;
            padding: 12px;
            background: #f8d7da;
            border-radius: 6px;
        }
        .template-btn {
            background: #6c757d;
            margin-bottom: 10px;
        }
        .template-btn:hover {
            background: #5a6268;
        }
        .footer {
            margin-top: 30px;
            text-align: center;
            font-size: 12px;
            color: #999;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🏛️ 立法委員發言統計系統</h1>
        <p class="subtitle">Flask 簡化版 - 當 Streamlit 不可用時使用</p>
        
        <div class="warning">
            ℹ️ 這是輕量級備用版本。完整功能請使用 Streamlit 網頁版。
        </div>
        
        <form id="analysisForm">
            <div class="section">
                <label>📋 下載範本名單</label>
                <button type="button" class="template-btn" onclick="downloadTemplate()">
                    下載範本（Excel）
                </button>
            </div>

            <div class="section">
                <label for="labelFile">委員名單 (Excel)</label>
                <input type="file" id="labelFile" name="label_file" accept=".xlsx,.xls" required>
            </div>

            <div class="section">
                <label for="transcriptFile">逐字稿 (txt/doc/docx)</label>
                <input type="file" id="transcriptFile" name="transcript_file" accept=".txt,.doc,.docx" required>
            </div>

            <button type="submit">🚀 開始分析</button>
        </form>

        <div class="progress" id="progress">
            <div class="spinner"></div>
            <p>正在分析...</p>
        </div>

        <div class="error" id="error">
            分析失敗
        </div>

        <div class="result" id="result">
            <h3>✅ 分析完成</h3>
            <p id="resultText"></p>
            <button onclick="downloadResult()" style="margin-top: 15px;">
                📥 下載 Excel 報表
            </button>
        </div>

        <div class="footer">
            <p>版本：Flask 1.0 | Python 3.8+</p>
        </div>
    </div>

    <script>
        let resultData = null;

        function downloadTemplate() {
            window.location.href = '/download-template';
        }

        document.getElementById('analysisForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const labelFile = document.getElementById('labelFile').files[0];
            const transcriptFile = document.getElementById('transcriptFile').files[0];
            
            if (!labelFile || !transcriptFile) {
                alert('請選擇所有檔案');
                return;
            }
            
            const formData = new FormData();
            formData.append('label_file', labelFile);
            formData.append('transcript_file', transcriptFile);
            
            document.getElementById('progress').style.display = 'block';
            document.getElementById('error').style.display = 'none';
            document.getElementById('result').style.display = 'none';
            
            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });
                
                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.error || '分析失敗');
                }
                
                resultData = await response.json();
                
                document.getElementById('resultText').innerHTML = `
                    會議名稱：${resultData.meeting_info.name}<br>
                    會議日期：${resultData.meeting_info.date}<br>
                    主席：${resultData.meeting_info.chairman}<br>
                    <br>
                    分析完成！請下載 Excel 報表查看詳細統計。
                `;
                
                document.getElementById('progress').style.display = 'none';
                document.getElementById('result').style.display = 'block';
            } catch (err) {
                document.getElementById('error').textContent = '錯誤：' + err.message;
                document.getElementById('error').style.display = 'block';
                document.getElementById('progress').style.display = 'none';
            }
        });

        function downloadResult() {
            if (!resultData) return;
            
            const link = document.createElement('a');
            link.href = '/download-result';
            link.download = 'legislator_stats.xlsx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    </script>
</body>
</html>
"""

# 全局變數存儲結果
analysis_result = None


@app.route('/')
def index():
    """主頁"""
    return render_template_string(HTML_TEMPLATE)


@app.route('/download-template')
def download_template():
    """下載範本名單"""
    template_df = pd.DataFrame({
        '委員會': ['內政', '內政', '經濟', '財政'],
        '黨籍': ['民進黨', '國民黨', '民眾黨', '無黨籍'],
        '姓名': ['張宏陸', '牛煦庭', '張啓楷', '陳超明']
    })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, sheet_name='委員名單', index=False)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='委員名單_範本.xlsx'
    )


@app.route('/analyze', methods=['POST'])
def analyze():
    """分析逐字稿"""
    global analysis_result
    
    try:
        # 取得上傳的檔案
        label_file = request.files.get('label_file')
        transcript_file = request.files.get('transcript_file')
        
        if not label_file or not transcript_file:
            return jsonify({'error': '須上傳名單和逐字稿'}), 400
        
        # 讀取名單
        df_config = pd.read_excel(label_file)
        
        # 欄位標準化
        rename_map = {}
        if '黨籍' in df_config.columns and '政黨' not in df_config.columns:
            rename_map['黨籍'] = '政黨'
        if rename_map:
            df_config = df_config.rename(columns=rename_map)
        
        required_columns = {'姓名', '政黨'}
        if not required_columns.issubset(df_config.columns):
            return jsonify({'error': "名單格式需包含『姓名』與『政黨』欄位"}), 400
        
        # 清理資料
        df_config = df_config.dropna(subset=['姓名', '政黨']).copy()
        df_config['姓名'] = df_config['姓名'].astype(str).str.strip()
        df_config['政黨'] = df_config['政黨'].astype(str).str.strip()
        
        # 讀取逐字稿
        transcript_bytes = transcript_file.read()
        text, convert_note = extract_text_from_uploaded_file(transcript_file.filename, transcript_bytes)
        
        # 解析
        labels = dict(zip(df_config['姓名'], df_config['政黨']))
        info, stats = parse_transcript(text, labels)
        
        # 生成結果 DataFrame
        row_data = {
            ("基本資訊", "日期", ""): info['date'],
            ("基本資訊", "會議名稱", ""): info['name'],
            ("基本資訊", "主席", ""): info['chairman']
        }
        for name, party in labels.items():
            row_data[(party, name, "次數")] = stats.get(name, {}).get('count', 0)
            row_data[(party, name, "字數")] = stats.get(name, {}).get('words', 0)
        
        # 創建 DataFrame
        df = pd.DataFrame([row_data])
        df.columns = pd.MultiIndex.from_tuples(df.columns)
        
        # 重新排序
        name_index_map = {name: i for i, name in enumerate(labels.keys())}
        
        def list_based_sort_key(col_tuple):
            category, name, metric = col_tuple
            if category == "基本資訊":
                basic_info_order = {"日期": 0, "會議名稱": 1, "主席": 2}
                info_rank = basic_info_order.get(name, 999)
                return (-1, info_rank, str(metric))
            idx = name_index_map.get(name, 9999)
            m_rank = 0 if metric == "次數" else 1
            return (0, idx, m_rank)
        
        sorted_cols = sorted(df.columns, key=list_based_sort_key)
        df = df[sorted_cols]
        
        # 存儲結果用於下載
        analysis_result = df
        
        return jsonify({
            'success': True,
            'meeting_info': info,
            'stats': {name: stats[name]['count'] for name in labels}
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download-result')
def download_result():
    """下載分析結果"""
    if analysis_result is None:
        return jsonify({'error': '無可用的分析結果'}), 400
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        analysis_result.to_excel(writer, sheet_name='統計結果')
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='legislator_stats.xlsx'
    )


def open_browser():
    """延遲打開瀏覽器"""
    import time
    time.sleep(2)
    webbrowser.open('http://127.0.0.1:5000')


if __name__ == '__main__':
    print("=" * 60)
    print("🏛️  立法委員會議發言統計系統 - Flask 簡化版")
    print("=" * 60)
    print("\n🚀 正在啟動...") 
    print("   網址：http://127.0.0.1:5000")
    print("\n💡 按 Ctrl+C 停止程式")
    print("=" * 60)
    
    # 在後台執行緒打開瀏覽器
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    # 啟動 Flask
    app.run(debug=False, port=5000)
