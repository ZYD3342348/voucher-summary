#!/usr/bin/env python3
"""
将“总数”表规范为长表（科目代码、名称、借方、贷方），方便后续程序化取数。

用法：
    python3 normalize_total.py -i "2025年10月总台测试.xlsx" -s "总数" -o "2025年10月总数_long.csv"

规则：
- 读取指定工作表，清除全空行/列。
- 只保留包含名称或借方/贷方金额的行，跳过“合计”行。
- 输出列：code（科目代码，可为空）、name、debit、credit、source_file、source_row。
- 若需要转账金额，可在长表中按 name == "转账" 筛选，优先取贷方（credit），否则借方。
"""

import argparse
from pathlib import Path
import pandas as pd


def normalize_total(input_file: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(input_file, sheet_name=sheet, header=None)
    # 丢弃全空行/列
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = range(df.shape[1])

    # 期待的列位置：0 科目代码, 1 名称, 2 借方, 3 贷方（若不存在，尽量兼容）
    col_code = 0 if 0 in df.columns else None
    col_name = 1 if 1 in df.columns else None
    col_debit = 2 if 2 in df.columns else None
    col_credit = 3 if 3 in df.columns else None

    records = []
    for idx, row in df.iterrows():
        name = row[col_name] if col_name is not None else None
        debit = row[col_debit] if col_debit is not None else None
        credit = row[col_credit] if col_credit is not None else None
        code = row[col_code] if col_code is not None else None

        # 跳过合计行
        if isinstance(name, str) and "合计" in name:
            continue

        # 过滤全空行
        if pd.isna(name) and pd.isna(debit) and pd.isna(credit):
            continue

        records.append(
            {
                "code": code,
                "name": name if not (isinstance(name, float) and pd.isna(name)) else None,
                "debit": debit if not pd.isna(debit) else None,
                "credit": credit if not pd.isna(credit) else None,
                "source_file": input_file.name,
                "source_row": int(idx),
            }
        )

    return pd.DataFrame(records)


def main():
    parser = argparse.ArgumentParser(description="规范化总数表为长表")
    parser.add_argument("-i", "--input", required=True, help="输入Excel文件")
    parser.add_argument("-s", "--sheet", default="总数", help="工作表名，默认 总数")
    parser.add_argument("-o", "--output", required=True, help="输出CSV文件路径")
    args = parser.parse_args()

    inp = Path(args.input)
    df = normalize_total(inp, args.sheet)
    df.to_csv(args.output, index=False)
    # 输出转账提示
    tx_rows = df[df["name"] == "转账"]
    if not tx_rows.empty:
        tx_row = tx_rows.iloc[0]
        tx_amt = tx_row["credit"] if pd.notna(tx_row["credit"]) else tx_row["debit"]
        print(f"[info] 转账行: row {tx_row['source_row']} 金额={tx_amt}")
    print(f"[info] 已保存: {args.output}, 行数 {len(df)}")


if __name__ == "__main__":
    main()
