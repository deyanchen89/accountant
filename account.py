import pandas as pd
import re

# 读取Excel文件中的SAP和HQ两个sheet
df_sap = pd.read_excel("data.xlsx", sheet_name="SAP", dtype=str)
df_hq = pd.read_excel("data.xlsx", sheet_name="HQ", dtype=str)

# 确保前两列正确命名（可根据实际列名调整）
df_sap.columns = ['Col1', 'Col2'] + list(df_sap.columns[2:])
df_hq.columns = ['Col1', 'Col2'] + list(df_hq.columns[2:])

# 将HQ的第一列转为字符串并去除空白
hq_dict = dict(zip(df_hq['Col1'].astype(str).str.strip(), df_hq['Col2'].astype(float)))

# 存储需要写入结果文件的SAP第一列值
results = []

def get_transformed_ids(sap_id):
    # 检查是否包含 - 或 /
    if '-' in sap_id or '/' in sap_id:
        parts = re.split(r'[-/]', sap_id)
        base = parts[0]
        prefix = parts[0][:-len(parts[1])]
        results = [base]
        for l in parts[1:]:
            transformed = prefix + l
            results.append(transformed)
        
    return results

# 遍历SAP数据
for idx, row in df_sap.iterrows():
    sap_id = str(row['Col1']).strip()
    try:
        sap_amount = abs(float(row['Col2']))
    except:
        continue  # 如果无法转换为数字，跳过

    candidate_ids = get_transformed_ids(sap_id)

    matched_values = [abs(hq_dict.get(cid, 0)) for cid in candidate_ids if cid in hq_dict]

    if not matched_values:
        # HQ中完全没有匹配ID，添加到结果
        results.append(sap_id)
    elif len(candidate_ids) == 1:
        # 直接匹配：判断金额是否相等
        if sap_amount != matched_values[0]:
            results.append(sap_id)
    else:
        # 多个转换后匹配：判断金额之和
        if sap_amount != sum(matched_values):
            results.append(sap_id)

# 将结果写入result.xlsx
pd.DataFrame(results, columns=['SAP Col1']).to_excel("result.xlsx", index=False)
