#!/usr/bin/env python3
"""
按收入类型过滤（默认 H），从工作表明细生成动态列透视（名称 × 项目），用于价税分离等场景。

用法示例：
    python3 generate_tax_sep.py -i "2025年8月总台.xlsx" -s "工作表" -o "2025年8月_H_价税分离.xlsx"
"""

import argparse
import re
from pathlib import Path
import pandas as pd


def detect_amount_col(df: pd.DataFrame) -> int:
    best_col, best_count = None, -1
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        cnt = s.notna().sum()
        if cnt > best_count:
            best_count, best_col = cnt, c
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


def load_work(input_file: Path, sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(input_file, sheet_name=sheet, header=None)
    raw.columns = range(raw.shape[1])

    # 定位表头行
    header_idx = None
    col_project = col_name = None
    for idx, row in raw.iterrows():
        if row.astype(str).str.contains("项目").any() or row.astype(str).str.contains("名称").any():
            header_idx = idx
            if row.astype(str).str.contains("项目").any():
                col_project = row[row.astype(str).str.contains("项目")].index.min()
            if row.astype(str).str.contains("名称").any():
                col_name = row[row.astype(str).str.contains("名称")].index.min()
            break
    start_row = header_idx + 1 if header_idx is not None else 0
    col_project = 0 if col_project is None else col_project
    col_name = 2 if col_name is None else col_name

    data = raw.iloc[start_row:].copy()
    data = data.rename(columns={col_project: "项目", col_name: "名称"})

    amt_col = detect_amount_col(data)
    if amt_col is None:
        raise ValueError("未检测到金额列")
    data = data.rename(columns={amt_col: "金额"})

    data["名称"] = data["名称"].apply(clean_name)
    data["金额"] = pd.to_numeric(data["金额"], errors="coerce")
    data = data.dropna(subset=["金额"])
    data.loc[data["项目"] == "半日租", "项目"] = "房费"
    data["收入类型"] = data["名称"].apply(derive_income_type)
    return data


def main():
    ap = argparse.ArgumentParser(description="生成价税分离透视（按收入类型过滤）")
    ap.add_argument("-i", "--input", required=True, help="输入工作表所在的 Excel 文件")
    ap.add_argument("-s", "--sheet", default="工作表", help="工作表名，默认 工作表")
    ap.add_argument("-o", "--output", help="输出 Excel 文件路径")
    ap.add_argument("-t", "--type", default="H", help="收入类型过滤值，默认 H")
    args = ap.parse_args()

    inp = Path(args.input)
    out = Path(args.output) if args.output else inp.with_name(f"{inp.stem}_{args.type}_tax.xlsx")

    df = load_work(inp, args.sheet)
    df = df[df["收入类型"].astype(str).str.upper() == args.type.upper()]

    pivot = df.pivot_table(index="名称", columns="项目", values="金额", aggfunc="sum", fill_value=0)
    pivot["总计"] = pivot.sum(axis=1)
    pivot.loc["合计"] = pivot.sum()

    pivot.to_excel(out)
    print(f"[info] 已生成透视: {out}, 行数 {len(pivot)}")


if __name__ == "__main__":
    main()
