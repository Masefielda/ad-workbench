import pandas as pd
import numpy as np

BASE = r"D:\Desktop"
files = {
    "手机号": f"{BASE}\\线索评分表-手机号.xlsx",
    "设备号": f"{BASE}\\线索评分表-设备号.xlsx",
}

def load_score_df(path):
    # 表头在第2行（index=1），跳过第1行
    df = pd.read_excel(path, sheet_name="下单日期", header=1)
    df.columns = ["下单日期", "评分", "低价课订单数", "高价课订单数", "转率", "产值", "评分占比"]
    # 日期向下填充
    df["下单日期"] = df["下单日期"].ffill()
    # 排除汇总行
    df = df[~df["下单日期"].astype(str).str.contains("汇总", na=False)]
    # 评分必须是1-10的数字
    df = df[pd.to_numeric(df["评分"], errors="coerce").between(1, 10)]
    df["评分"] = df["评分"].astype(int)
    for col in ["低价课订单数", "高价课订单数", "产值"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    # 产值列是客单价，总产值 = 产值 × 低价课订单数
    df["总产值"] = df["产值"] * df["低价课订单数"]
    return df

def agg_by_score(df, label):
    g = df.groupby("评分").agg(
        低价课订单数=("低价课订单数", "sum"),
        高价课订单数=("高价课订单数", "sum"),
        总产值=("总产值", "sum"),
    ).reset_index()
    g["转化率"] = np.where(g["低价课订单数"] > 0, g["高价课订单数"] / g["低价课订单数"], 0)
    g["ARPU"] = np.where(g["低价课订单数"] > 0, g["总产值"] / g["低价课订单数"], 0)
    g.columns = ["评分", f"{label}_低价课订单数", f"{label}_高价课订单数", f"{label}_总产值", f"{label}_转化率", f"{label}_ARPU"]
    return g

# 加载并汇总
phone_df = load_score_df(files["手机号"])
dev_df = load_score_df(files["设备号"])

phone_agg = agg_by_score(phone_df, "手机号")
dev_agg = agg_by_score(dev_df, "设备号")

# 合并
result = pd.merge(phone_agg, dev_agg, on="评分", how="outer").sort_values("评分").reset_index(drop=True)

# 合计行
total_row = {"评分": "合计"}
for col in result.columns:
    if col == "评分":
        continue
    if "转化率" in col or "ARPU" in col:
        low_col = col.replace("转化率", "低价课订单数").replace("ARPU", "低价课订单数")
        high_col = col.replace("转化率", "高价课订单数").replace("ARPU", "总产值")
        if low_col in result.columns and high_col in result.columns:
            low_sum = result[low_col].sum()
            high_sum = result[high_col].sum()
            total_row[col] = high_sum / low_sum if low_sum > 0 else 0
        else:
            total_row[col] = 0
    else:
        total_row[col] = result[col].sum()

result = pd.concat([result, pd.DataFrame([total_row])], ignore_index=True)

# 输出
out_path = f"{BASE}\\线索评分表-汇总.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="线索评分", index=False)

print(f"已输出: {out_path}")
print(f"行数: {len(result)}")
print(result.to_string(index=False))
