#!/usr/bin/env python3
"""
批量跑测试数据并生成汇总输出（用于验证公式链与转账提取逻辑）。

提取转账金额规则：
1) 在“总数”工作表中找到首个包含“转账”字样的行。
2) 取该行的数值单元格里最大的数（过滤掉明显是科目编号的小数字，例如 <1000 的数字）。
   这样避免把整行求和造成金额偏大。
"""

import sys
from pathlib import Path
import pandas as pd

# 将上级目录加入路径以导入 generate_summary
sys.path.append(str(Path(__file__).resolve().parent))
from generate_summary import generate_summary  # type: ignore


def detect_transfer_amount(path: Path) -> float:
    """
    从总数表提取转账金额：
    1) 先找单元格内容恰好等于“转账”的行（不含括号/后缀），取首行。
    2) 在该行中，从“转账”所在列向右，取首个数值且>=1000；若不存在，再取该行中数值>=1000的最小值（避免把合计/倒推等更大的数值误取）。
    3) 若不存在恰好“转账”行，再退回含“转账”关键字的行，取首行并同样的取值策略。
    """
    xls = pd.ExcelFile(path)
    if '总数' not in xls.sheet_names:
        return 0.0
    df = pd.read_excel(path, sheet_name='总数', header=None)
    exact_mask = df.applymap(lambda x: isinstance(x, str) and x.strip() == '转账')
    if exact_mask.any().any():
        target_idx = df[exact_mask.any(axis=1)].index.min()
    else:
        fuzzy_mask = df.apply(lambda r: r.astype(str).str.contains('转账', na=False)).any(axis=1)
        if not fuzzy_mask.any():
            return 0.0
        target_idx = df[fuzzy_mask].index.min()

    row = df.loc[target_idx]
    # 尝试按“转账”所在列向右取首个金额
    row_series = row.copy()
    transfer_cols = [i for i, v in row_series.items() if isinstance(v, str) and v.strip() == '转账']
    if transfer_cols:
        start_col = transfer_cols[0] + 1
        for col in range(start_col, len(row_series)):
            v = row_series.iloc[col]
            if isinstance(v, (int, float)) and not pd.isna(v) and abs(float(v)) >= 1000:
                return float(v)
    # 退回：取该行数值>=1000的最小值，避免取到合计/倒推的更大数
    nums = [
        float(x)
        for x in row_series
        if isinstance(x, (int, float)) and not pd.isna(x) and abs(float(x)) >= 1000
    ]
    return min(nums) if nums else 0.0


def choose_income_sheet(path: Path):
    """优先选择 '收入类型表'，否则 '收入类型'，否则 None."""
    xls = pd.ExcelFile(path)
    if '收入类型表' in xls.sheet_names:
        return '收入类型表'
    if '收入类型' in xls.sheet_names:
        return '收入类型'
    return None


def main():
    test_dir = Path(__file__).resolve().parent / '测试数据'
    files = [
        Path('2025年10月总台测试.xlsx'),
        Path('2025年11月总台收入.xlsx'),
        Path('2025年12月总台收入.xlsx'),
        Path('2025年8月总台.xlsx'),
        Path('2025年9月总台收入.xlsx'),
    ]
    results = []
    for f in files:
        path = test_dir / f
        if not path.exists():
            results.append((f.name, 'missing', '文件不存在'))
            continue
        sheet = choose_income_sheet(path)
        if not sheet:
            results.append((f.name, 'skip', '缺少收入类型表/收入类型工作表'))
            continue
        transfer = detect_transfer_amount(path)
        out_file = path.with_name(path.stem + '_汇总输出.xlsx')
        try:
            generate_summary(str(path), str(out_file), transfer_amount=transfer, sheet_name=sheet)
            results.append((f.name, 'ok', transfer, out_file.name))
        except Exception as e:
            results.append((f.name, 'error', str(e)))

    print("\n运行结果:")
    for r in results:
        print(r)


if __name__ == '__main__':
    main()
