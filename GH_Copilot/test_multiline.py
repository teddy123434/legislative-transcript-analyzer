#!/usr/bin/env python3
"""測試多行續接功能"""

from ccw_analyzer import CCWTranscriptAnalyzer

# 測試範例：主席的發言有三行
sample_text = """主席：提案條文都已經宣讀完畢，請問各位委員，我們是不是不用再進行大體討論，直接進行逐條討論，有無異議？（無）好。
再來我們進行逐條討論協商。
報告委員會，修正動議第1案、第2案、第3案、第4案，因為委員已經有提案了，所以撤案。
我們現在進行協商，謝謝。
黃委員捷：謝謝。針對第七條，我覺得行政院版是OK的，只是我覺得在執行上可能會……因為其實原本賴瑞隆的版本裡面有寫到一些更細緻的部分。
這是第二句續接。
主席：請李柏毅委員。
李委員柏毅：我的提案跟行政院版本是一樣的。"""

analyzer = CCWTranscriptAnalyzer()
analyzer.parse_transcript(sample_text)

print("【所有發言人統計】")
print("=" * 60)
for speaker in analyzer.get_all_speakers():
    data = analyzer.speaker_data[speaker]
    print(f"\n{speaker}:")
    print(f"  發言次數: {data['count']}")
    print(f"  總字數: {data['total_chars']}")
    print(f"  發言內容:")
    for i, content in enumerate(data['content'], 1):
        # 顯示前100字，展示是否有續接
        preview = content[:100] + "..." if len(content) > 100 else content
        print(f"    {i}. {preview}")
        if "\n" in content:
            print(f"       (包含 {content.count(chr(10))} 個續行)")
