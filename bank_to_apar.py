import pandas as pd
import os

# 文件名定义
bank_file = 'Aivres Bank Record_2025.05.23.xlsx'
ap_files = {
    '2025.03': 'EW_AP_2025.03.xlsx',
    '2025.04': 'EW_AP_2025.04.xlsx'
}
ar_files = {
    '2025.03': 'EW_AR_2025.03.xlsx',
    '2025.04': 'EW_AR_2025.04.xlsx'
}

def get_abs_float(value):
    """尝试将值转换为绝对值浮点数，如果失败则返回None"""
    try:
        return abs(float(value))
    except:
        return None

def compare_column(bank_df, ref_series, col_name, result_file):
    correct_rows = []
    incorrect_rows = []

    # 处理参考列中的数据，转为绝对值浮点型集合
    ref_values = set()
    for val in ref_series.dropna():
        abs_val = get_abs_float(val)
        if abs_val is not None:
            ref_values.add(abs_val)

    for _, row in bank_df.iterrows():
        # 跳过Date列为空的行
        if pd.isna(row.get('Date')):
            continue

        value = row.get(col_name)
        # 跳过空或非数值的值
        val_float = get_abs_float(value)
        if val_float is None:
            continue

        if val_float in ref_values:
            correct_rows.append(row)
        else:
            incorrect_rows.append(row)

    # 写入结果到Excel文件
    correct_df = pd.DataFrame(correct_rows)
    incorrect_df = pd.DataFrame(incorrect_rows)

    with pd.ExcelWriter(result_file, engine='openpyxl', mode='w') as writer:
        correct_df.to_excel(writer, sheet_name='correct', index=False)
        incorrect_df.to_excel(writer, sheet_name='incorrect', index=False)

def compare_and_write_results(month):
    # 读取 Bank Record 中对应月份的 Sheet
    bank_df = pd.read_excel(bank_file, sheet_name=month, engine='openpyxl')

    # -------------------
    # Debit 对比 AP G列
    # -------------------
    ap_df = pd.read_excel(ap_files[month], sheet_name='record', engine='openpyxl')
    ap_column = ap_df.iloc[:, 6]  # G列是第7列，索引为6

    compare_column(
        bank_df=bank_df,
        ref_series=ap_column,
        col_name='Debit',
        result_file=f"EW_AP_{month}_result.xlsx"
    )

    # -------------------
    # Credit 对比 AR H列
    # -------------------
    ar_df = pd.read_excel(ar_files[month], sheet_name='record', engine='openpyxl')
    ar_column = ar_df.iloc[:, 7]  # H列是第8列，索引为7

    compare_column(
        bank_df=bank_df,
        ref_series=ar_column,
        col_name='Credit',
        result_file=f"EW_AR_{month}_result.xlsx"
    )

# 执行每个月的比对
for month in ['2025.03', '2025.04']:
    compare_and_write_results(month)

print("所有对比完成，结果已保存到对应 *_result.xlsx 文件中。")
