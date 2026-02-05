#!/usr/bin/env python3
"""
互動式 CCW 逐字稿分析工具
"""

import sys
from pathlib import Path
from ccw_analyzer import CCWTranscriptAnalyzer


def interactive_menu():
    """互動式選單"""
    analyzer = CCWTranscriptAnalyzer()
    speaker_list = None
    
    while True:
        print("\n" + "="*50)
        print("公督盟委員會逐字稿分析工具")
        print("="*50)
        if speaker_list:
            print(f"📋 已指定委員: {', '.join(sorted(speaker_list))}")
        else:
            print("📋 未指定委員 (將抓取所有發言人)")
        print("-"*50)
        print("1. 設定委員名單")
        print("2. 清除委員限制 (抓取所有)")
        print("3. 讀取 Word 檔案 (.docx)")
        print("4. 讀取文字檔案 (.txt)")
        print("5. 貼上逐字稿文本")
        print("6. 查看所有委員統計")
        print("7. 查看特定委員統計")
        print("8. 輸出到 Excel")
        print("9. 結束程式")
        print("="*50)
        
        choice = input("請選擇操作 (1-9): ").strip()
        
        if choice == "1":
            speakers_input = input("請輸入委員名單，用逗號分隔 (例: 主席,陳委員培瑜,王委員義川): ").strip()
            if speakers_input:
                speaker_list = [s.strip() for s in speakers_input.split(',')]
                analyzer = CCWTranscriptAnalyzer(speaker_list)
                print(f"✓ 已設定委員名單: {', '.join(speaker_list)}")
            else:
                print("✗ 未輸入任何委員")
        
        elif choice == "2":
            speaker_list = None
            analyzer = CCWTranscriptAnalyzer()
            print("✓ 已清除委員限制，將抓取所有發言人")
        
        elif choice == "3":
            file_path = input("請輸入 Word 檔案路徑 (.docx): ").strip()
            try:
                analyzer.load_from_word(file_path)
                print(f"✓ 成功讀取檔案，找到 {len(analyzer.get_all_speakers())} 位發言人")
            except FileNotFoundError:
                print(f"✗ 找不到檔案: {file_path}")
            except Exception as e:
                print(f"✗ 讀取出錯: {e}")
        
        elif choice == "4":
            file_path = input("請輸入文字檔案路徑 (.txt): ").strip()
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                analyzer.parse_transcript(text)
                print(f"✓ 成功讀取檔案，找到 {len(analyzer.get_all_speakers())} 位發言人")
            except FileNotFoundError:
                print(f"✗ 找不到檔案: {file_path}")
            except Exception as e:
                print(f"✗ 讀取出錯: {e}")
        
        elif choice == "5":
            print("請輸入逐字稿文本 (完成後輸入 'END' 並按 Enter):")
            lines = []
            while True:
                line = input()
                if line.strip().upper() == "END":
                    break
                lines.append(line)
            text = "\n".join(lines)
            analyzer.parse_transcript(text)
            print(f"✓ 成功解析，找到 {len(analyzer.get_all_speakers())} 位發言人")
        
        elif choice == "6":
            if not analyzer.get_all_speakers():
                print("✗ 尚未載入任何逐字稿")
            else:
                analyzer.print_summary()
        
        elif choice == "7":
            if not analyzer.get_all_speakers():
                print("✗ 尚未載入任何逐字稿")
            else:
                print("\n現有發言人:")
                for speaker in sorted(analyzer.get_all_speakers()):
                    print(f"  - {speaker}")
                speaker_name = input("\n請輸入委員名字: ").strip()
                analyzer.print_summary(speaker_name)
        
        elif choice == "8":
            if not analyzer.get_all_speakers():
                print("✗ 尚未載入任何逐字稿")
            else:
                output_path = input("請輸入輸出檔案路徑 (預設: ccw_analysis.xlsx): ").strip()
                if not output_path:
                    output_path = "ccw_analysis.xlsx"
                try:
                    analyzer.export_to_excel(output_path)
                except Exception as e:
                    print(f"✗ 輸出失敗: {e}")
        
        elif choice == "9":
            print("謝謝使用，再見！")
            break
        
        else:
            print("✗ 無效的選擇，請重試")


if __name__ == "__main__":
    # 如果命令行有參數，直接處理檔案
    if len(sys.argv) > 1:
        analyzer = CCWTranscriptAnalyzer()
        file_path = sys.argv[1]
        
        try:
            # 檢測檔案類型
            if file_path.lower().endswith('.docx'):
                analyzer.load_from_word(file_path)
            else:
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                analyzer.parse_transcript(text)
            
            analyzer.print_summary()
            
            output_file = sys.argv[2] if len(sys.argv) > 2 else "ccw_analysis.xlsx"
            analyzer.export_to_excel(output_file)
            print(f"\n✓ 已輸出統計結果到: {output_file}")
        except FileNotFoundError:
            print(f"✗ 找不到文件: {file_path}")
            sys.exit(1)
        except Exception as e:
            print(f"✗ 錯誤: {e}")
            sys.exit(1)
    else:
        # 互動式模式
        interactive_menu()
