#!/usr/bin/env python3
"""
简易 Web 报表（Streamlit）：
- 上传工作簿，自动长表化“工作表”，名称清洗、收入类型提取
- 过滤收入类型（默认 H），生成名称×项目透视
- 生成可编辑调整表（不计税/5%/6% 分桶），即时计算税分拆摘要（总计 + 按项目）
- 支持下载多 Sheet Excel（透视、调整表、税分拆摘要、项目税分拆）

运行：
    streamlit run app.py
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# 本地导入
from scripts.normalize_work import normalize_work  # type: ignore


def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    pivot = df.pivot_table(index="名称", columns="项目", values="金额", aggfunc="sum", fill_value=0)
    pivot["总计"] = pivot.sum(axis=1)
    pivot.loc["合计"] = pivot.sum()
    return pivot.round(2)


def build_adjust_table(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["名称", "项目"], as_index=False)["金额"].sum()
    adj = g.copy()
    adj["不计税收入"] = 0.0
    adj["计税收入-5%"] = 0.0
    adj["计税收入-6%"] = adj["金额"].astype(float)
    adj["备注"] = ""
    return adj


def summarize_tax(adj: pd.DataFrame) -> pd.DataFrame:
    rows = []
    notax = adj["不计税收入"].sum()
    tax5 = adj["计税收入-5%"].sum()
    tax6 = adj["计税收入-6%"].sum()

    def net_tax(amount, rate):
        if amount <= 0:
            return 0.0, 0.0
        net = amount / (1 + rate)
        tax = amount - net
        return net, tax

    net5, taxamt5 = net_tax(tax5, 0.05)
    net6, taxamt6 = net_tax(tax6, 0.06)

    rows.append(["不计税收入", round(notax, 2), "", ""])
    rows.append(["计税收入-5%", round(tax5, 2), round(net5, 2), round(taxamt5, 2)])
    rows.append(["计税收入-6%", round(tax6, 2), round(net6, 2), round(taxamt6, 2)])
    rows.append(["合计", round(notax + tax5 + tax6, 2), round(net5 + net6, 2), round(taxamt5 + taxamt6, 2)])
    return pd.DataFrame(rows, columns=["类别", "含税收入", "不含税收入", "税额"]).round(2)


def summarize_tax_by_project(adj: pd.DataFrame) -> pd.DataFrame:
    proj_rows = []
    for proj, grp in adj.groupby("项目"):
        notax = grp["不计税收入"].sum()
        tax5 = grp["计税收入-5%"].sum()
        tax6 = grp["计税收入-6%"].sum()

        def net_tax(amount, rate):
            if amount <= 0:
                return 0.0, 0.0
            net = amount / (1 + rate)
            tax = amount - net
            return net, tax

        net5, taxamt5 = net_tax(tax5, 0.05)
        net6, taxamt6 = net_tax(tax6, 0.06)

        proj_rows.append(
            {
                "项目": proj,
                "不计税收入": round(notax, 2),
                "计税收入-5%": round(tax5, 2),
                "计税收入-6%": round(tax6, 2),
                "含税收入合计": round(notax + tax5 + tax6, 2),
                "不含税收入": round(net5 + net6, 2),
                "税额": round(taxamt5 + taxamt6, 2),
            }
        )

    df_proj = pd.DataFrame(proj_rows)
    if not df_proj.empty:
        total_row = {
            "项目": "合计",
            "不计税收入": df_proj["不计税收入"].sum(),
            "计税收入-5%": df_proj["计税收入-5%"].sum(),
            "计税收入-6%": df_proj["计税收入-6%"].sum(),
            "含税收入合计": df_proj["含税收入合计"].sum(),
            "不含税收入": df_proj["不含税收入"].sum(),
            "税额": df_proj["税额"].sum(),
        }
        df_proj = pd.concat([df_proj, pd.DataFrame([total_row])], ignore_index=True)
    return df_proj.round(2)


def main():
    st.set_page_config(page_title="简易报表", layout="wide")
    st.title("简易报表（工作表 → 透视/调整表/税分拆）")

    uploaded = st.file_uploader("上传工作簿（含“工作表”）", type=["xlsx", "xls"])
    income_type = st.selectbox("收入类型过滤", ["H", "L", "R", "S", "T", "Z"], index=0)
    sheet_name = st.text_input("工作表名称", value="工作表")

    if uploaded:
        # 保存临时文件供 normalize_work 使用
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.read())
            tmp_path = Path(tmp.name)

        df = normalize_work(tmp_path, sheet_name)
        df = df[df["收入类型"].astype(str).str.upper() == income_type.upper()]

        st.subheader("透视（名称 × 项目）")
        pivot = build_pivot(df)
        st.dataframe(pivot)

        st.subheader("调整表（可编辑，不计税/5%/6% 分桶）")
        adj_default = build_adjust_table(df)
        edited = st.data_editor(adj_default, num_rows="dynamic")
        st.caption("提示：分桶之和应≤金额；未修改的默认全计税6%，可手工调整。")

        st.subheader("税分拆摘要（总计）")
        summary = summarize_tax(edited)
        st.dataframe(summary)

        st.subheader("项目维度税分拆")
        summary_proj = summarize_tax_by_project(edited)
        st.dataframe(summary_proj)

        # 下载 Excel
        with io.BytesIO() as buffer:
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                pivot.reset_index().rename(columns={"index": "名称"}).to_excel(writer, sheet_name="透视", index=False)
                edited.to_excel(writer, sheet_name="调整表", index=False)
                summary.to_excel(writer, sheet_name="税分拆摘要", index=False)
                summary_proj.to_excel(writer, sheet_name="项目税分拆", index=False)
            st.download_button("下载 Excel", data=buffer.getvalue(), file_name="报表输出.xlsx", mime="application/vnd.ms-excel")

        # 清理临时文件
        tmp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
