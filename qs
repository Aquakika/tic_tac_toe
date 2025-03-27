import pandas as pd
import re

def normalize_name(name):
    """只保留中文字符作为品名关键词"""
    if pd.isna(name):
        return ''
    name = str(name)
    name = re.sub(r'[^\u4e00-\u9fff]', '', name)  # 只保留中文
    return name

# 读取 A 库存和 B 文件
a_file = "QS.xlsx"
b_file = "1111.xlsx"

df_a = pd.read_excel(a_file)
df_b = pd.read_excel(b_file)

# 规范化品名（只提取中文）
df_a["品名_norm"] = df_a["品名"].apply(normalize_name)
df_b["品名_norm"] = df_b["品名"].apply(normalize_name)

# 构建 A 库存品名到型号的映射
a_stock_dict = {}
for _, row in df_a.iterrows():
    name = row["品名_norm"]
    model = row["型号"]
    if name in a_stock_dict:
        a_stock_dict[name].append(model)
    else:
        a_stock_dict[name] = [model]

# 进行匹配并分配型号
matched_models = []
model_count = {}

for _, row in df_b.iterrows():
    name = row["品名_norm"]
    if name in a_stock_dict:
        models = a_stock_dict[name]
        if name not in model_count:
            model_count[name] = 0
        
        if model_count[name] < len(models):
            matched_models.append(models[model_count[name]])
            model_count[name] += 1
        else:
            matched_models.append("")  # 型号不足，留空
    else:
        matched_models.append("")  # 无匹配，留空

# 添加匹配结果到原 B 文件中
df_b["匹配型号"] = matched_models

# 保存结果为新文件
df_b.drop(columns=["品名_norm"], inplace=True)
df_b.to_excel("B_file_with_model.xlsx", index=False)

print("匹配完成，结果已保存为 'B_file_with_model.xlsx'")
