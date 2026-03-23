# 委員實質法案審查發言次/字數函式庫
import re
import io
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path

from charset_normalizer import from_bytes
from docx import Document
from config import user_config

party_order, info_order, metric_order = user_config()


def decode_text_bytes(file_bytes):
    """
    盡可能自動偵測文字編碼，提升 Windows 上的相容性。
    """
    if not file_bytes:
        return ""

    fallback_encodings = [
        'utf-8-sig',
        'utf-8',
        'cp950',
        'big5',
        'utf-16',
        'utf-16le',
        'utf-16be',
        'gb18030'
    ]

    best_match = from_bytes(file_bytes).best()
    if best_match:
        try:
            return str(best_match)
        except Exception:
            pass

    for encoding in fallback_encodings:
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue

    return file_bytes.decode('utf-8', errors='replace')


def extract_text_from_docx_bytes(file_bytes):
    """
    將 docx 內容擷取成純文字。
    """
    document = Document(io.BytesIO(file_bytes))
    return "\n".join(paragraph.text for paragraph in document.paragraphs)


def convert_doc_to_docx_bytes(file_name, file_bytes):
    """
    自動將 .doc 轉成 .docx，優先使用：
    1) Windows + Word COM (win32com)
    2) LibreOffice soffice (macOS/Linux/Windows)
    
    on macOS: 檢查常見路徑，包括 /Applications/LibreOffice.app
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        source_path = Path(temp_dir) / file_name
        target_path = source_path.with_suffix('.docx')
        source_path.write_bytes(file_bytes)

        # 方案 1: Windows + Word COM
        if platform.system().lower() == 'windows':
            try:
                import win32com.client

                word_app = win32com.client.Dispatch('Word.Application')
                word_app.Visible = False
                document = word_app.Documents.Open(str(source_path.resolve()))
                document.SaveAs(str(target_path.resolve()), FileFormat=16)
                document.Close(False)
                word_app.Quit()
                return target_path.read_bytes()
            except Exception:
                pass

        # 方案 2: LibreOffice soffice (所有平台)
        soffice_candidates = [
            shutil.which('soffice'),  # PATH 中的 soffice
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',  # macOS 標準路徑
            '/usr/bin/soffice',  # Linux 標準路徑
            '/usr/local/bin/soffice',  # Homebrew (Intel Mac)
            '/opt/homebrew/bin/soffice',  # Homebrew (Apple Silicon)
        ]
        
        for soffice_path in soffice_candidates:
            if soffice_path and Path(soffice_path).exists():
                try:
                    subprocess.run(
                        [
                            soffice_path,
                            '--headless',
                            '--convert-to',
                            'docx',
                            '--outdir',
                            temp_dir,
                            str(source_path)
                        ],
                        check=True,
                        capture_output=True,
                        text=True,
                        timeout=30
                    )
                    if target_path.exists():
                        return target_path.read_bytes()
                except Exception:
                    pass

    # 根據平台顯示不同的錯誤訊息
    system = platform.system().lower()
    if system == 'windows':
        error_msg = (
            "無法自動將 .doc 轉為 .docx。\n"
            "Windows 請確認已安裝 Microsoft Word；或改上傳 .docx / .txt。"
        )
    elif system == 'darwin':  # macOS
        error_msg = (
            "無法自動將 .doc 轉為 .docx。\n"
            "macOS 請安裝 LibreOffice：\n"
            "  brew install libreoffice\n"
            "或改上傳 .docx / .txt。"
        )
    else:  # Linux
        error_msg = (
            "無法自動將 .doc 轉為 .docx。\n"
            "Linux 請安裝 LibreOffice：\n"
            "  sudo apt install libreoffice (Debian/Ubuntu)\n"
            "  sudo yum install libreoffice (CentOS/RHEL)\n"
            "或改上傳 .docx / .txt。"
        )
    
    raise ValueError(error_msg)


def extract_text_from_uploaded_file(file_name, file_bytes):
    """
    支援 txt/doc/docx：
    - txt: 自動編碼偵測後轉文字
    - docx: 直接抽取純文字
    - doc: 先轉 docx 再抽取純文字
    """
    suffix = Path(file_name).suffix.lower()

    if suffix == '.txt':
        return decode_text_bytes(file_bytes), None

    if suffix == '.docx':
        return extract_text_from_docx_bytes(file_bytes), "已自動由 docx 轉為純文字"

    if suffix == '.doc':
        converted_docx_bytes = convert_doc_to_docx_bytes(file_name, file_bytes)
        return extract_text_from_docx_bytes(converted_docx_bytes), "已自動由 doc → docx → 純文字"

    raise ValueError(f"不支援的檔案格式：{suffix}")
        
def parse_transcript(text, labels):
    """
    解析逐字稿，回傳會議資訊與發言統計
    """
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    lines = text.split('\n')
    
    # 初始化變數
    stats = {name: {'count': 0, 'words': 0, 'logs': []} for name in labels.keys()}
    meeting_info = {
        'date': '',
        'name': '',
        'chairman': ''
    }
    
    current_speaker = None
    
    # --- 步驟 A: 解析檔頭資訊 (Header Parsing) ---
    # 這裡使用正則表達式 (Regex) 來抓取特定格式
    
    # 1. 抓取第一行作為會議名稱
    if len(lines) > 0:
        meeting_info['name'] = lines[0].strip()
        
    # 2. 抓取日期 (格式：中華民國XXX年XX月XX日)
    date_match = re.search(r'中華民國(\d+)年(\d+)月(\d+)日', text)
    if date_match:
        # 格式化為 114.12.17
        meeting_info['date'] = f"{date_match.group(1)}.{date_match.group(2)}.{date_match.group(3)}"
        
    # 3. 抓取主席 (格式：主　　席　王委員美惠)
    # 邏輯：找到「主席」那一行，並移除「委員」二字來對應標準姓名
    chairman_match = re.search(r'主\s*席\s*(.*)', text)
    if chairman_match:
        raw_chairman = chairman_match.group(1).strip()
        # 移除 "委員" 二字以便對應 (例如 "王委員美惠" -> "王美惠")
        clean_chairman = raw_chairman.replace("委員", "").replace(" ", "")
        meeting_info['chairman'] = clean_chairman

    # --- 步驟 B: 逐行解析發言 (Body Parsing) ---
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # [修正 1]：擴大 Regex 範圍
        # 允許：中文字 (\u4e00-\u9fa5)、英文字母 (a-zA-Z)、空白 (\s)、點 (.)
        # 這樣就能抓到 "鄭天財Sra Kacaw委員" 或 "王委員美惠"
        speaker_match = re.match(r'^([\u4e00-\u9fa5a-zA-Z\s\.]+)\s*[：:]', line)
        
        if speaker_match:
            raw_name = speaker_match.group(1)
            
            # 判斷是誰在說話
            speaker_identity = None
            
            if "主席" in raw_name:
                speaker_identity = meeting_info['chairman']
            else:
                # [修正 2]：更聰明的名字清洗邏輯
                # 1. 先移除 "委員"
                temp_name = raw_name.replace("委員", "").strip()
                
                # 2. 嘗試比對（完全匹配，針對 Sra Kacaw 這種有空白的名字）
                if temp_name in labels:
                    speaker_identity = temp_name
                # 3. 嘗試比對（移除所有空白，針對 "黃 捷" 這種名字）
                elif temp_name.replace(" ", "") in labels:
                    speaker_identity = temp_name.replace(" ", "")
                # 4. 針對 "黃委員捷" 變成 "黃捷" 的狀況 (已在步驟1處理，但雙重保險)
            
            # 如果是我們關注的對象，切換 current_speaker
            if speaker_identity and speaker_identity in stats:
                current_speaker = speaker_identity
                stats[current_speaker]['count'] += 1
                
                # 計算這一行剩下的文字長度 (扣除名字和冒號)
                content = line[len(speaker_match.group(0)):]
                stats[current_speaker]['words'] += len(content)

                stats[current_speaker]['logs'].append(f"**[{stats[current_speaker]['count']}]** {content.strip()}")
            else:
                # 如果是官員或不在名單的人說話，將 current_speaker 設為 None，不計入
                current_speaker = None
                
        else:
            # 如果這一行沒有冒號，表示是上一位講者的延續發言
            if current_speaker:
                stats[current_speaker]['words'] += len(line)
                stats[current_speaker]['logs'][-1] += f"\n{line}"
    return meeting_info, stats

def custom_sort_key(col_tuple):
    """
    這個函數會幫每一個欄位打分數，分數越小排越前面。
    col_tuple 結構: (黨派/分類, 姓名/欄位名, 數據類型)
    """
    category, name, metric = col_tuple
    
    # [權重 1] 第一層：黨派順序
    try:
        rank1 = party_order.index(category)
    except ValueError:
        rank1 = 99  # 如果有沒定義到的黨派，排到最後面

    # [權重 2] 第二層：
    # 如果是基本資訊，依照 info_order 排序
    # 如果是委員，我們希望依照「你的 Label 字典順序」或「姓名筆畫」
    # 這裡我們簡單點，讓同一黨派內的委員依照姓名排序
    if category == "基本資訊":
        try:
            rank2 = info_order.index(name)
        except ValueError:
            rank2 = 99
        # 為了確保基本資訊永遠在該群組最前，給個小權重
        rank3 = 0 
    else:
        # 委員姓名排序：這裡使用 Python 預設字串排序，
        # 如果你希望特定委員排前面，也可以像 party_order 一樣建一個 list
        rank2 = name 
        
        # [權重 3] 第三層：次數 vs 字數
        try:
            rank3 = metric_order.index(metric)
        except ValueError:
            rank3 = 99

    # 回傳一個 Tuple 讓 Python 進行多重排序
    # Python 排序邏輯：先比 rank1，一樣則比 rank2，再一樣比 rank3
    return (rank1, rank2, rank3)