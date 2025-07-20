import os
import pandas as pd

def load_bank_data(bank_file, sheet_name, value_column):
    df = pd.read_excel(bank_file, sheet_name=sheet_name, engine="openpyxl")
    df = df[pd.notnull(df["Date"]) & pd.notnull(df[value_column])]
    df = df.copy()
    df["value"] = pd.to_numeric(df[value_column], errors='coerce').abs()
    df = df[pd.notnull(df["value"])]  # 过滤非数值
    return df

def load_ap_ar_data(file, col_index):
    df = pd.read_excel(file, sheet_name="record", engine="openpyxl")
    df["value"] = pd.to_numeric(df.iloc[:, col_index], errors='coerce').abs()
    df = df[pd.notnull(df["value"])]  # 过滤非数值
    return df

def compare_and_export(bank_df, ap_ar_df, out_dir, out_prefix, bank_name, ap_ar_name):
    bank_values = set(bank_df["value"])
    ap_ar_values = set(ap_ar_df["value"])

    in_both = bank_values & ap_ar_values
    only_in_bank = bank_values - ap_ar_values
    only_in_ap_ar = ap_ar_values - bank_values

    # A -> AP/AR
    in_both_bank = bank_df[bank_df["value"].isin(in_both)]
    only_in_bank_df = bank_df[bank_df["value"].isin(only_in_bank)]

    # AP/AR -> A
    in_both_ap_ar = ap_ar_df[ap_ar_df["value"].isin(in_both)]
    only_in_ap_ar_df = ap_ar_df[ap_ar_df["value"].isin(only_in_ap_ar)]

    os.makedirs(out_dir, exist_ok=True)

    writer1 = pd.ExcelWriter(os.path.join(out_dir, f"{out_prefix}_{bank_name}{ap_ar_name}_result.xlsx"), engine="openpyxl")
    in_both_bank.to_excel(writer1, sheet_name=f"In{bank_name}And{ap_ar_name}", index=False)
    only_in_bank_df.to_excel(writer1, sheet_name=f"In{bank_name}only", index=False)
    writer1.close()

    writer2 = pd.ExcelWriter(os.path.join(out_dir, f"{out_prefix}_{ap_ar_name}{bank_name}_result.xlsx"), engine="openpyxl")
    in_both_ap_ar.to_excel(writer2, sheet_name=f"In{ap_ar_name}And{bank_name}", index=False)
    only_in_ap_ar_df.to_excel(writer2, sheet_name=f"In{ap_ar_name}only", index=False)
    writer2.close()

def process_month(month_str):
    bank_file = "Aivres Bank Record_2025.05.23.xlsx"
    ap_file = f"EW_AP_{month_str}.xlsx"
    ar_file = f"EW_AR_{month_str}.xlsx"
    output_dir = f"{month_str.replace('.', '')}"

    print(f"Processing month {month_str}...")

    # --- AP vs Bank Debit ---
    bank_debit_df = load_bank_data(bank_file, sheet_name=month_str, value_column="Debit")
    ap_df = load_ap_ar_data(ap_file, col_index=6)  # 第6列索引为5
    compare_and_export(bank_debit_df, ap_df, output_dir, month_str.replace('.', ''), "Bank", "AP")

    # --- AR vs Bank Credit ---
    bank_credit_df = load_bank_data(bank_file, sheet_name=month_str, value_column="Credit")
    ar_df = load_ap_ar_data(ar_file, col_index=7)  # 第7列索引为6
    compare_and_export(bank_credit_df, ar_df, output_dir, month_str.replace('.', ''), "Bank", "AR")

# 主处理流程
if __name__ == "__main__":
    process_month("2025.03")
    process_month("2025.04")
