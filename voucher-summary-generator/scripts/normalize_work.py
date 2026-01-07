#!/usr/bin/env python3
"""
将“工作表”明细规范化：项目在列0，名称在列2，自动检测金额列，生成收入类型并做房费透视检查。

用法：
    python3 normalize_work.py -i "2025年8月总台.xlsx" -s "工作表" -o "2025年8月_work_long.csv" -t 351260

输出：
    - CSV 长表：项目、名称、金额、收入类型、source_file、source_row
    - 控制台：收入类型计数、房费透视（H/L/T/R/S/Z）、房费总计和调整S（若提供转账金额）
"""

import argparse
import re
from pathlib import Path
import pandas as pd


def detect_amount_col(df: pd.DataFrame) -> int:
    """选择数值型占比最高的列作为金额列。"""
    best_col = None
    best_count = -1
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        count = s.notna().sum()
        if count > best_count:
            best_count = count
            best_col = c
    return best_col


def derive_income_type(name: str):
    if not isinstance(name, str):
        return None
    m = re.search(r"[A-Za-z]", name)
    return m.group(0).upper() if m else None


def clean_name(name: str):
    if not isinstance(name, str):
        return name
    name = name.strip()
    if "_" in name:
        name = name.split("_")[0].strip()
    return name


def normalize_work(input_file: Path, sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(input_file, sheet_name=sheet, header=None)
    raw.columns = range(raw.shape[1])

    # 如果首行包含“项目”，视为表头，移除
    start_row = 1 if raw.iloc[0].astype(str).str.contains("项目").any() else 0
    data = raw.iloc[start_row:].copy()

    # 按固定列号命名
    data = data.rename(columns={0: "项目", 2: "名称"})

    # 检测金额列
    amt_col = detect_amount_col(data)
    data = data.rename(columns={amt_col: "金额"})

    # 清洗
    data["名称"] = data["名称"].apply(clean_name)
    data["金额"] = pd.to_numeric(data["金额"], errors="coerce")
    data = data.dropna(subset=["金额"])
    # 半日租归并房费
    data.loc[data["项目"] == "半日租", "项目"] = "房费"
    # 收入类型
    data["收入类型"] = data["名称"].apply(derive_income_type)

    data = data[["项目", "名称", "金额", "收入类型"]].copy()
    data["source_file"] = input_file.name
    data["source_row"] = data.index + start_row
    return data


def main():
    parser = argparse.ArgumentParser(description="规范化工作表并生成收入类型")
    parser.add_argument("-i", "--input", required=True, help="输入Excel文件")
    parser.add_argument("-s", "--sheet", default="工作表", help="工作表名，默认 工作表")
    parser.add_argument("-o", "--output", required=True, help="输出CSV文件")
    parser.add_argument("-t", "--transfer", type=float, default=None, help="转账金额（可选，用于调整S计算）")
    args = parser.parse_args()

    inp = Path(args.input)
    df = normalize_work(inp, args.sheet)
    df.to_csv(args.output, index=False)

    print(f"[info] 已保存长表: {args.output}, 行数 {len(df)}")
    print("[info] 收入类型分布:")
    print(df["收入类型"].value_counts().to_string())

    pivot = df.pivot_table(index="收入类型", columns="项目", values="金额", aggfunc="sum", fill_value=0)
    room_total = pivot.get("房费", pd.Series(dtype=float)).sum()
    h = pivot.get("房费", pd.Series(dtype=float)).get("H", 0)
    l = pivot.get("房费", pd.Series(dtype=float)).get("L", 0)
    t = pivot.get("房费", pd.Series(dtype=float)).get("T", 0)
    print("[info] 房费透视：")
    print(pivot.get("房费", pd.Series(dtype=float)))
    if args.transfer is not None:
        adjust_s = room_total - args.transfer - h - l - t
        print(f"[info] 房费总计={room_total:.2f} 转账={args.transfer} H={h} L={l} T={t} 调整S={adjust_s:.2f}")


if __name__ == "__main__":
    main()
