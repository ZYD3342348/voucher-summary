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


def _clean_cell(v):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, str):
        v = v.strip()
        return v if v != "" else None
    return v


def _locate_header(df: pd.DataFrame):
    """
    根据“名称”行定位表头；借方/贷方可在该行或前几行找到。
    返回 (header_row_idx, col_code, col_name, col_debit, col_credit)
    """
    header_idx = None
    col_name = col_debit = col_credit = col_code = None

    for idx, row in df.iterrows():
        if row.astype(str).str.contains("名称").any():
            header_idx = idx
            col_name = row[row.astype(str).str.contains("名称")].index.min()
            break

    if header_idx is None:
        # 尝试仅借方/贷方行
        for idx, row in df.iterrows():
            if row.astype(str).str.contains("借方").any() or row.astype(str).str.contains("贷方").any():
                header_idx = idx
                col_debit = row[row.astype(str).str.contains("借方")].index.min() if row.astype(str).str.contains("借方").any() else None
                col_credit = row[row.astype(str).str.contains("贷方")].index.min() if row.astype(str).str.contains("贷方").any() else None
                # 经验规则：名称在借方左侧2列，代码在名称左侧1列
                col_name = col_debit - 2 if col_debit is not None and col_debit >= 2 else None
                col_code = col_name - 1 if col_name is not None and col_name >= 1 else None
                return header_idx, col_code, col_name, col_debit, col_credit
        return None, None, None, None, None

    # 在表头行以及之前的行里找借方/贷方标识
    search_region = df.loc[:header_idx]
    for _, r in search_region.iterrows():
        if col_debit is None and r.astype(str).str.contains("借方").any():
            col_debit = r[r.astype(str).str.contains("借方")].index.min()
        if col_credit is None and r.astype(str).str.contains("贷方").any():
            col_credit = r[r.astype(str).str.contains("贷方")].index.min()

    # 代码列：在名称列左侧的最近一列
    if col_name is not None:
        left_cols = [c for c in df.columns if c < col_name]
        if left_cols:
            col_code = max(left_cols)

    return header_idx, col_code, col_name, col_debit, col_credit


def normalize_total(input_file: Path, sheet: str) -> pd.DataFrame:
    df_raw = pd.read_excel(input_file, sheet_name=sheet, header=None)
    df = df_raw.dropna(how="all").dropna(axis=1, how="all")
    header_idx, col_code, col_name, col_debit, col_credit = _locate_header(df)
    if header_idx is None:
        raise ValueError("未找到包含 名称/借方/贷方 的表头行")

    data = df.loc[header_idx + 1 :].copy()
    records = []
    for idx, row in data.iterrows():
        name = row[col_name] if col_name is not None and col_name in row else None
        debit = row[col_debit] if col_debit is not None and col_debit in row else None
        credit = row[col_credit] if col_credit is not None and col_credit in row else None
        code = row[col_code] if col_code is not None and col_code in row else None

        name = _clean_cell(name)
        code = _clean_cell(code)

        # 跳过合计行
        if isinstance(name, str) and "合计" in name:
            continue

        # 跳过“总计/合计”等空白名称的尾部汇总行（常见：code/name 为空，但借贷方给出总额）
        if name is None and code is None:
            continue

        # 过滤全空行
        if (name is None or name == "") and pd.isna(debit) and pd.isna(credit):
            continue

        records.append(
            {
                "code": code,
                "name": name,
                "debit": float(debit) if isinstance(debit, (int, float)) and not pd.isna(debit) else None,
                "credit": float(credit) if isinstance(credit, (int, float)) and not pd.isna(credit) else None,
                "source_file": input_file.name,
                "source_row": int(idx),
            }
        )

    return pd.DataFrame(records)


def summarize(df: pd.DataFrame):
    """按名称分组，计算转账、汇总及倒推校验。"""
    def credit_sum(names):
        mask = df['name'].astype(str).isin(names)
        return df.loc[mask, 'credit'].fillna(0).sum()

    debit_total = df['debit'].fillna(0).sum()

    transfer = None
    tx_rows = df[df["name"] == "转账"]
    if not tx_rows.empty:
        tx_row = tx_rows.iloc[0]
        transfer = tx_row["credit"] if pd.notna(tx_row["credit"]) else tx_row["debit"]

    internal_cost = 0
    ic_rows = df[df["name"] == "转内部成本"]
    if not ic_rows.empty:
        ic_row = ic_rows.iloc[0]
        internal_cost = ic_row["credit"] if pd.notna(ic_row["credit"]) else ic_row["debit"]

    bank = credit_sum(['银行转账', '银行转帐', '银行', '银行汇总', 'AR支票预收'])
    wechat = credit_sum(['微信', '微信支付', '微信汇总'])
    cash = credit_sum(['现金结账', '现金', '现金汇总', 'AR现金预收'])
    lkl = credit_sum(['拉卡拉', '拉卡拉预收', '银联POS预收', '拉卡拉汇总'])
    fiscal = credit_sum(['财政', '财政汇总'])

    voucher_credit = debit_total - (transfer or 0) - internal_cost
    pending = voucher_credit - bank - wechat - cash - lkl - fiscal
    voucher_debit = bank + wechat + cash + lkl + fiscal + pending

    return {
        "debit_total": debit_total,
        "transfer": transfer,
        "internal_cost": internal_cost,
        "bank": bank,
        "wechat": wechat,
        "cash": cash,
        "lkl": lkl,
        "fiscal": fiscal,
        "voucher_credit": voucher_credit,
        "pending": pending,
        "voucher_debit": voucher_debit,
    }


def main():
    parser = argparse.ArgumentParser(description="规范化总数表为长表")
    parser.add_argument("-i", "--input", required=True, help="输入Excel文件")
    parser.add_argument("-s", "--sheet", default="总数", help="工作表名，默认 总数")
    parser.add_argument("-o", "--output", required=True, help="输出CSV文件路径")
    args = parser.parse_args()

    inp = Path(args.input)
    df = normalize_total(inp, args.sheet)
    df.to_csv(args.output, index=False, encoding="utf-8-sig")
    # 输出转账提示
    tx_rows = df[df["name"] == "转账"]
    if not tx_rows.empty:
        tx_row = tx_rows.iloc[0]
        tx_amt = tx_row["credit"] if pd.notna(tx_row["credit"]) else tx_row["debit"]
        print(f"[info] 转账行: row {tx_row['source_row']} 金额={tx_amt}")
    summary = summarize(df)
    print("[summary] 借方合计={debit_total:.2f} 转账={transfer} 转内部成本={internal_cost}".format(**summary))
    print("[summary] 银行={bank:.2f} 微信={wechat:.2f} 现金={cash:.2f} 拉卡拉={lkl:.2f} 财政={fiscal:.2f}".format(**summary))
    print("[summary] 凭证贷方={voucher_credit:.2f} 应挂账金额={pending:.2f} 凭证借方={voucher_debit:.2f}".format(**summary))
    print(f"[info] 已保存: {args.output}, 行数 {len(df)}")


if __name__ == "__main__":
    main()
