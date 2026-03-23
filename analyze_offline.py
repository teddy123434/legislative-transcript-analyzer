#!/usr/bin/env python3
"""
命令行離線分析工具：純 Python 實現，無需任何網頁框架
適用於所有環境，穩定性最高
"""
import sys
import argparse
from pathlib import Path
import pandas as pd

# 匯入核心分析函數
from functions import parse_transcript, extract_text_from_uploaded_file


def prompt_file(prompt_text, accept_extensions=None):
    """交互式選擇檔案"""
    while True:
        file_path = input(f"\n{prompt_text}\n> ").strip()
        
        if not file_path:
            print("❌ 檔案路徑不能為空")
            continue
        
        p = Path(file_path).expanduser()
        
        if not p.exists():
            print(f"❌ 檔案不存在：{p}")
            continue
        
        if not p.is_file():
            print(f"❌ 路徑不是檔案：{p}")
            continue
        
        if accept_extensions:
            if p.suffix.lower() not in accept_extensions:
                exts = ", ".join(accept_extensions)
                print(f"❌ 檔案格式不支援，需要 {exts}")
                continue
        
        return p


def interactive_mode():
    """交互式模式"""
    print("\n" + "=" * 60)
    print("🏛️  立法委員發言統計系統 - 命令行工具（交互模式）")
    print("=" * 60)
    
    # 選擇委員名單
    label_file = prompt_file(
        "📋 請選擇委員名單檔案（Excel）\n例如：/Users/teddy/委員名單.xlsx",
        accept_extensions=['.xlsx', '.xls']
    )
    
    # 選擇逐字稿
    transcript_file = prompt_file(
        "📄 請選擇逐字稿檔案（txt/doc/docx）\n例如：/Users/teddy/逐字稿.txt",
        accept_extensions=['.txt', '.doc', '.docx']
    )
    
    # 選擇輸出位置
    while True:
        output_path = input("\n💾 請輸入輸出檔案路徑（預設：./legislator_stats.xlsx）\n> ").strip()
        if not output_path:
            output_path = "./legislator_stats.xlsx"
        
        output_path = Path(output_path).expanduser()
        
        # 檢查父資料夾是否存在
        if not output_path.parent.exists():
            print(f"❌ 輸出資料夾不存在：{output_path.parent}")
            continue
        
        break
    
    # 開始分析
    analyze(label_file, transcript_file, output_path)


def analyze(label_file, transcript_file, output_file=None):
    """執行分析"""
    if output_file is None:
        output_file = Path("legislator_stats.xlsx")
    else:
        output_file = Path(output_file)
    
    try:
        print("\n" + "=" * 60)
        print("📊 開始分析...")
        print("=" * 60)
        
        # 讀取委員名單
        print(f"\n📋 讀取名單：{label_file}")
        df_config = pd.read_excel(label_file)
        
        # 欄位標準化
        rename_map = {}
        if '黨籍' in df_config.columns and '政黨' not in df_config.columns:
            rename_map['黨籍'] = '政黨'
        if rename_map:
            df_config = df_config.rename(columns=rename_map)
        
        required_columns = {'姓名', '政黨'}
        if not required_columns.issubset(df_config.columns):
            raise ValueError("名單格式需包含『姓名』與『政黨』欄位")
        
        # 清理資料
        df_config = df_config.dropna(subset=['姓名', '政黨']).copy()
        df_config['姓名'] = df_config['姓名'].astype(str).str.strip()
        df_config['政黨'] = df_config['政黨'].astype(str).str.strip()
        
        labels = dict(zip(df_config['姓名'], df_config['政黨']))
        print(f"   ✅ 已載入 {len(labels)} 位委員")
        
        # 讀取逐字稿
        print(f"\n📄 讀取逐字稿：{transcript_file}")
        with open(transcript_file, 'rb') as f:
            transcript_bytes = f.read()
        
        text, convert_note = extract_text_from_uploaded_file(
            transcript_file.name,
            transcript_bytes
        )
        if convert_note:
            print(f"   🔄 {convert_note}")
        print(f"   ✅ 已讀取 {len(text)} 個字符")
        
        # 解析逐字稿
        print("\n🔍 分析發言內容...")
        info, stats = parse_transcript(text, labels)
        
        print(f"   會議名稱：{info['name']}")
        print(f"   會議日期：{info['date']}")
        print(f"   主席：{info['chairman']}")
        
        # 統計有發言的委員
        active_count = sum(1 for s in stats.values() if s['count'] > 0)
        print(f"   有發言的委員：{active_count} 位")
        
        # 生成結果 DataFrame
        print("\n📊 生成統計表格...")
        row_data = {
            ("基本資訊", "日期", ""): info['date'],
            ("基本資訊", "會議名稱", ""): info['name'],
            ("基本資訊", "主席", ""): info['chairman']
        }
        for name, party in labels.items():
            row_data[(party, name, "次數")] = stats.get(name, {}).get('count', 0)
            row_data[(party, name, "字數")] = stats.get(name, {}).get('words', 0)
        
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
        
        # 匯出 Excel
        print(f"\n💾 匯出結果：{output_file}")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='統計結果')
        
        print(f"   ✅ 已保存")
        
        # 顯示前 5 位高發言者
        print("\n🏆 發言次數前 5 名：")
        top_speakers = sorted(
            [(name, stats[name]['count']) for name in labels if stats[name]['count'] > 0],
            key=lambda x: x[1],
            reverse=True
        )[:5]
        
        for i, (name, count) in enumerate(top_speakers, 1):
            party = labels[name]
            print(f"   {i}. {name} ({party})：{count} 次")
        
        print("\n" + "=" * 60)
        print("✅ 分析完成！")
        print("=" * 60)
        
        return True
    
    except Exception as e:
        print(f"\n❌ 分析失敗：{e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函數"""
    parser = argparse.ArgumentParser(
        description="立法委員發言統計系統 - 命令行工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用範例：
  1. 交互模式（推薦）：
     python analyze_offline.py
  
  2. 命令行參數模式：
     python analyze_offline.py --label 委員名單.xlsx --transcript 逐字稿.txt --output 結果.xlsx
"""
    )
    
    parser.add_argument(
        '--label',
        type=str,
        help='委員名單檔案路徑 (Excel format)'
    )
    parser.add_argument(
        '--transcript',
        type=str,
        help='逐字稿檔案路徑 (txt/doc/docx format)'
    )
    parser.add_argument(
        '--output',
        type=str,
        default='legislator_stats.xlsx',
        help='輸出 Excel 檔案路徑 (預設：legislator_stats.xlsx)'
    )
    
    args = parser.parse_args()
    
    # 如果提供了所有三個參數，執行分析
    if args.label and args.transcript:
        label_file = Path(args.label)
        transcript_file = Path(args.transcript)
        
        if not label_file.exists():
            print(f"❌ 檔案不存在：{label_file}")
            sys.exit(1)
        
        if not transcript_file.exists():
            print(f"❌ 檔案不存在：{transcript_file}")
            sys.exit(1)
        
        success = analyze(label_file, transcript_file, args.output)
        sys.exit(0 if success else 1)
    
    # 否則進入交互模式
    interactive_mode()


if __name__ == '__main__':
    main()
