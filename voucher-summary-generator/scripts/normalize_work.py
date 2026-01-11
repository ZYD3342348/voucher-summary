#!/usr/bin/env python3
"""
将“工作表”明细规范化：自动定位表头（项目/名称/金额），生成收入类型并做房费透视检查。

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
    # 规则：取遇到的第一个英文字母作为收入类型
    s = re.sub(r"\s+", "", name)
    m = re.search(r"[A-Za-z]", s)
    return m.group(0).upper() if m else None


def clean_name(name: str):
    if not isinstance(name, str):
        return name
    # 对齐 Excel「数据分列」：按 "_" 分割，只保留第一段
    s = re.sub(r"\s+", "", name)
    if "_" in s:
        s = s.split("_")[0].strip()
    return s


def _locate_header(raw: pd.DataFrame):
    """
    在 sheet 内扫描表头行，优先找同时包含“项目”和“名称”的行。
    返回：(header_idx, col_project, col_name, col_amount)
    """
    header_idx = None
    col_project = col_name = col_amount = None
    for idx, row in raw.iterrows():
        row_s = row.astype(str)
        has_project = row_s.str.contains("项目").any()
        has_name = row_s.str.contains("名称").any()
        if has_project and has_name:
            header_idx = idx
            col_project = row[row_s.str.contains("项目")].index.min()
            col_name = row[row_s.str.contains("名称")].index.min()
            if row_s.str.contains("金额").any():
                col_amount = row[row_s.str.contains("金额")].index.min()
            break
    return header_idx, col_project, col_name, col_amount


def normalize_work(input_file: Path, sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(input_file, sheet_name=sheet, header=None)
    raw.columns = range(raw.shape[1])

    header_idx, col_project, col_name, col_amount = _locate_header(raw)
    start_row = (header_idx + 1) if header_idx is not None else 0
    data = raw.iloc[start_row:].copy()

    # 若未命中表头，退回历史经验：项目列0，名称列2
    if col_project is None:
        col_project = 0
    if col_name is None:
        col_name = 2

    if col_project not in data.columns:
        raise ValueError(f"未找到项目列（推断列={col_project}，请检查“{sheet}”表结构）")
    if col_name not in data.columns:
        raise ValueError(f"未找到名称列（推断列={col_name}，请检查“{sheet}”表结构）")

    # 命名
    data = data.rename(columns={col_project: "项目", col_name: "名称"})

    # 金额列：优先用表头命中列，否则在“非项目/名称列”中自动检测
    if col_amount is not None and col_amount in data.columns:
        amt_col = col_amount
    else:
        candidates = [c for c in data.columns if c not in {"项目", "名称"}]
        amt_col = detect_amount_col(data[candidates]) if candidates else detect_amount_col(data)
        if amt_col is None:
            raise ValueError("未检测到金额列")
    data = data.rename(columns={amt_col: "金额"})

    # 先提取收入类型，再按 "_" 清洗名称（对齐人工：先取首字母，再分列）
    data["收入类型"] = data["名称"].apply(derive_income_type)
    data["名称"] = data["名称"].apply(clean_name)
    data["金额"] = pd.to_numeric(data["金额"], errors="coerce")
    data = data.dropna(subset=["金额"])
    # 半日租归并房费
    data.loc[data["项目"] == "半日租", "项目"] = "房费"

    data = data[["项目", "名称", "金额", "收入类型"]].copy()
    data["source_file"] = input_file.name
    # data.index 是原始 raw 的行号（0-based），直接作为定位依据
    data["source_row"] = data.index
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
    df.to_csv(args.output, index=False, encoding="utf-8-sig")

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
