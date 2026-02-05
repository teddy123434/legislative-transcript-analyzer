def user_config():
    # 1. 定義你想要的「黨派順序」 (這是最外層的分類)
    # 程式會依照這個列表的順序由左至右排列
    party_order = ["基本資訊", "民進黨", "國民黨", "民眾黨", "無黨籍"]

    # 2. 定義「基本資訊」內部的順序
    info_order = ["日期", "會議名稱", "主席"]

    # 3. 定義數據欄位的順序
    metric_order = ["次數", "字數"]
    
    return party_order, info_order, metric_order