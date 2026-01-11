#!/usr/bin/env python3
"""
简易 Web 报表（Streamlit）：
- 上传单个工作簿（需含"工作表"和"总数"），自动完成：
  - 工作表：规范化/过滤收入类型 → 透视 → 分配表（可编辑）→ 税分拆
  - 总数：规范化为长表 → 定位"转账"贷方金额 → 倒推校验指标
- 全量校验通过才允许导出（校验失败会阻断导出）
- 导出单个工作簿，包含所有结果 Sheet

运行：
    streamlit run app.py
    # 如果本机 streamlit 命令的 shebang 损坏，可用：
    python3 -m streamlit run app.py
"""

import io
import sys
import tempfile
from pathlib import Path
import hashlib
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

# 必须是脚本中第一条 Streamlit 命令（且只能调用一次）
st.set_page_config(page_title="总台收入工作台", layout="wide")

# 确保 scripts 可导入
BASE = Path(__file__).resolve().parent
PARENT = BASE.parent
for p in (BASE, PARENT):
    if str(p) not in sys.path:
        sys.path.append(str(p))

from scripts.normalize_work import normalize_work  # type: ignore
from scripts.normalize_total import normalize_total  # type: ignore


def load_custom_css():
    """加载自定义CSS样式"""
    css_path = BASE / "static" / "custom_styles.css"
    if not css_path.exists():
        return
    try:
        custom_css = css_path.read_text(encoding="utf-8")
    except Exception as e:
        st.warning(f"加载自定义样式失败: {e}")
        return
    st.markdown(
        f"""<style>
{custom_css}
</style>""",
        unsafe_allow_html=True,
    )



def stable_file_id(uploaded: "st.runtime.uploaded_file_manager.UploadedFile") -> str:
    # 用内容 hash 做 session key，避免同名文件冲突
    data = uploaded.getvalue()
    return hashlib.sha1(data).hexdigest()


def to_number(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0).astype(float)


ALLOCATION_COLS = ("不计税分配", "计税分配-5%", "计税分配-6%")

DEFAULT_TOTAL_MATCH = {
    "银行": ["银行转账", "银行转帐", "银行", "银行汇总", "AR支票预收"],
    "微信": ["微信", "微信支付", "微信汇总"],
    "现金": ["现金结账", "现金", "现金汇总", "AR现金预收"],
    "拉卡拉": ["拉卡拉", "拉卡拉预收", "银联POS预收", "拉卡拉汇总"],
    "财政": ["财政", "财政汇总"],
}


def get_sheet_names(xlsx_path: Path) -> list[str]:
    try:
        with pd.ExcelFile(xlsx_path) as xf:
            return list(xf.sheet_names)
    except Exception:
        return []


def pick_default_sheet(sheet_names: list[str], preferred: str) -> Optional[str]:
    if preferred in sheet_names:
        return preferred
    for s in sheet_names:
        if preferred in s:
            return s
    return sheet_names[0] if sheet_names else None


def build_validation_tables(alloc: pd.DataFrame) -> dict[str, pd.DataFrame]:
    df = alloc.copy()
    for col in ("金额", *ALLOCATION_COLS):
        if col in df.columns:
            df[col] = to_number(df[col])
        else:
            df[col] = 0.0

    df["分配合计"] = df[list(ALLOCATION_COLS)].sum(axis=1)
    df["差额(分配-金额)"] = (df["分配合计"] - df["金额"]).round(2)

    def _agg(group_col: str) -> pd.DataFrame:
        g = (
            df.groupby(group_col, as_index=False)[["金额", *ALLOCATION_COLS, "分配合计"]]
            .sum()
            .sort_values(by="分配合计", ascending=False)
        )
        g["差额(分配-金额)"] = (g["分配合计"] - g["金额"]).round(2)
        return g

    by_name = _agg("名称") if "名称" in df.columns else pd.DataFrame()
    by_project = _agg("项目") if "项目" in df.columns else pd.DataFrame()
    total = pd.DataFrame(
        [
            {
                "名称/项目": "合计",
                "金额": df["金额"].sum(),
                **{c: df[c].sum() for c in ALLOCATION_COLS},
                "分配合计": df["分配合计"].sum(),
                "差额(分配-金额)": round(df["分配合计"].sum() - df["金额"].sum(), 2),
            }
        ]
    )
    return {"行级": df, "按名称": by_name, "按项目": by_project, "总计": total}


def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    pivot = df.pivot_table(index="名称", columns="项目", values="金额", aggfunc="sum", fill_value=0)
    pivot["总计"] = pivot.sum(axis=1)
    pivot.loc["合计"] = pivot.sum()
    return pivot.round(2)


def build_allocation_table(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["名称", "项目"], as_index=False)["金额"].sum()
    alloc = g.copy()
    alloc["不计税分配"] = 0.0
    alloc["计税分配-5%"] = 0.0
    alloc["计税分配-6%"] = alloc["金额"].astype(float)
    alloc["备注"] = ""
    return alloc


def summarize_tax(alloc: pd.DataFrame) -> pd.DataFrame:
    rows = []
    notax = to_number(alloc.get("不计税分配", pd.Series(dtype=float))).sum()
    tax5 = to_number(alloc.get("计税分配-5%", pd.Series(dtype=float))).sum()
    tax6 = to_number(alloc.get("计税分配-6%", pd.Series(dtype=float))).sum()

    def net_tax(amount, rate):
        if amount <= 0:
            return 0.0, 0.0
        net = amount / (1 + rate)
        tax = amount - net
        return net, tax

    net5, taxamt5 = net_tax(tax5, 0.05)
    net6, taxamt6 = net_tax(tax6, 0.06)

    rows.append(["不计税分配", round(notax, 2), "", ""])
    rows.append(["计税分配-5%", round(tax5, 2), round(net5, 2), round(taxamt5, 2)])
    rows.append(["计税分配-6%", round(tax6, 2), round(net6, 2), round(taxamt6, 2)])
    rows.append(["合计", round(notax + tax5 + tax6, 2), round(net5 + net6, 2), round(taxamt5 + taxamt6, 2)])
    return pd.DataFrame(rows, columns=["类别", "含税收入", "不含税收入", "税额"]).round(2)


def summarize_tax_by_project(alloc: pd.DataFrame) -> pd.DataFrame:
    proj_rows = []
    if "项目" not in alloc.columns:
        return pd.DataFrame()

    for proj, grp in alloc.groupby("项目"):
        notax = to_number(grp.get("不计税分配", pd.Series(dtype=float))).sum()
        tax5 = to_number(grp.get("计税分配-5%", pd.Series(dtype=float))).sum()
        tax6 = to_number(grp.get("计税分配-6%", pd.Series(dtype=float))).sum()

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
                "不计税分配": round(notax, 2),
                "计税分配-5%": round(tax5, 2),
                "计税分配-6%": round(tax6, 2),
                "含税收入合计": round(notax + tax5 + tax6, 2),
                "不含税收入": round(net5 + net6, 2),
                "税额": round(taxamt5 + taxamt6, 2),
            }
        )

    df_proj = pd.DataFrame(proj_rows)
    if not df_proj.empty:
        total_row = {
            "项目": "合计",
            "不计税分配": df_proj["不计税分配"].sum(),
            "计税分配-5%": df_proj["计税分配-5%"].sum(),
            "计税分配-6%": df_proj["计税分配-6%"].sum(),
            "含税收入合计": df_proj["含税收入合计"].sum(),
            "不含税收入": df_proj["不含税收入"].sum(),
            "税额": df_proj["税额"].sum(),
        }
        df_proj = pd.concat([df_proj, pd.DataFrame([total_row])], ignore_index=True)
    return df_proj.round(2)


def parse_name_list(text: str) -> list[str]:
    if not text:
        return []
    # 支持换行/逗号/顿号/分号分隔
    parts: list[str] = []
    for line in str(text).splitlines():
        for p in line.replace("，", ",").replace("、", ",").replace(";", ",").split(","):
            p = p.strip()
            if p:
                parts.append(p)
    # 去重保序
    seen = set()
    out = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def sum_credit_by_names(total_long: pd.DataFrame, names: list[str]) -> tuple[float, pd.DataFrame]:
    if total_long.empty or not names or "name" not in total_long.columns:
        return 0.0, pd.DataFrame()
    df = total_long.copy()
    df["name_norm"] = df["name"].astype(str).str.strip()
    hits = df[df["name_norm"].isin(names)].copy()
    amt = float(to_number(hits.get("credit", pd.Series(dtype=float))).sum())
    return amt, hits.drop(columns=["name_norm"], errors="ignore")


def extract_transfer_credit(total_long: pd.DataFrame) -> Tuple[Optional[float], pd.DataFrame]:
    if total_long.empty or "name" not in total_long.columns:
        return None, pd.DataFrame()
    df = total_long.copy()
    df["name_norm"] = df["name"].astype(str).str.strip()
    hits = df[df["name_norm"] == "转账"].copy()
    if hits.empty:
        return None, pd.DataFrame()
    credit = to_number(hits.get("credit", pd.Series(dtype=float))).sum()
    return float(credit), hits.drop(columns=["name_norm"], errors="ignore")


def build_total_summary(total_long: pd.DataFrame, transfer_credit: float, match_map: dict[str, list[str]]) -> tuple[dict, dict[str, pd.DataFrame]]:
    # 按需求：转账必须取贷方（credit），并用它回算倒推指标
    internal_cost = 0.0
    if "name" in total_long.columns:
        ic = total_long[total_long["name"].astype(str).str.strip() == "转内部成本"]
        internal_cost = float(to_number(ic.get("credit", pd.Series(dtype=float))).sum())

    debit_total = float(to_number(total_long.get("debit", pd.Series(dtype=float))).sum())

    hit_tables: dict[str, pd.DataFrame] = {}
    bank, hit_tables["银行"] = sum_credit_by_names(total_long, match_map.get("银行", []))
    wechat, hit_tables["微信"] = sum_credit_by_names(total_long, match_map.get("微信", []))
    cash, hit_tables["现金"] = sum_credit_by_names(total_long, match_map.get("现金", []))
    lkl, hit_tables["拉卡拉"] = sum_credit_by_names(total_long, match_map.get("拉卡拉", []))
    fiscal, hit_tables["财政"] = sum_credit_by_names(total_long, match_map.get("财政", []))

    voucher_credit = debit_total - transfer_credit - internal_cost
    pending = voucher_credit - bank - wechat - cash - lkl - fiscal
    voucher_debit = bank + wechat + cash + lkl + fiscal + pending

    return ({
        "借方合计": round(debit_total, 2),
        "转账(贷方)": round(transfer_credit, 2),
        "转内部成本": round(internal_cost, 2),
        "银行": round(bank, 2),
        "微信": round(wechat, 2),
        "现金": round(cash, 2),
        "拉卡拉": round(lkl, 2),
        "财政": round(fiscal, 2),
        "凭证贷方": round(voucher_credit, 2),
        "应挂账金额": round(pending, 2),
        "凭证借方": round(voucher_debit, 2),
    }, hit_tables)


def main():
    # 加载自定义CSS样式（在 set_page_config 之后）
    load_custom_css()
    
    # 自定义标题区域
    st.markdown("""
        <div class="main-container">
            <h1 class="main-title">总台收入工作台</h1>
            <p class="sub-title">中式浪漫 · 财务报表处理系统</p>
        </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader('上传 Excel（需含“工作表”和“总数”）', type=["xlsx", "xls"])

    if not uploaded:
        st.markdown("""
            <div class="info-card">
                <h3>📊 使用说明</h3>
                <p>上传Excel文件后，您可以：</p>
                <ul>
                    <li>选择工作表/总数表、收入类型（单选/多选）</li>
                    <li>编辑"分配表"，进行灵活的税费分配</li>
                    <li>实时校验数据平衡性</li>
                    <li>导出包含所有报表的汇总工作簿</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
        return

    file_id = stable_file_id(uploaded)

    # 保存临时文件供 scripts 使用
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded.getvalue())
        tmp_path = Path(tmp.name)

    sheet_names = get_sheet_names(tmp_path)
    default_work = pick_default_sheet(sheet_names, "工作表")
    default_total = pick_default_sheet(sheet_names, "总数")

    with st.sidebar:
        st.markdown("""
            <div class="info-card" style="margin-bottom: 1rem;">
                <h4>⚙️ 参数设置</h4>
            </div>
        """, unsafe_allow_html=True)
        
        if sheet_names:
            work_sheet = st.selectbox("工作表 Sheet", sheet_names, index=sheet_names.index(default_work) if default_work in sheet_names else 0)
            total_sheet = st.selectbox("总数 Sheet", sheet_names, index=sheet_names.index(default_total) if default_total in sheet_names else 0)
        else:
            work_sheet = st.text_input("工作表 Sheet", value="工作表")
            total_sheet = st.text_input("总数 Sheet", value="总数")

        mode = st.radio("收入类型选择", ["单选", "多选"], horizontal=True, index=0)
        all_types = ["H", "L", "R", "S", "T", "Z"]
        if mode == "单选":
            types = [st.selectbox("收入类型", all_types, index=0)]
        else:
            types = st.multiselect("收入类型（可多选）", all_types, default=["H"])
        tol = st.number_input("校验容差", min_value=0.0, value=0.01, step=0.01, format="%.2f")

        with st.expander("总数命中名单（默认 + 可追加）", expanded=False):
            st.caption('说明：下面只填写“追加项”，默认命中规则始终保留。多个名称可用换行/逗号分隔。')
            extra_bank = st.text_area("银行：追加名称", value="", height=80)
            extra_wechat = st.text_area("微信：追加名称", value="", height=80)
            extra_cash = st.text_area("现金：追加名称", value="", height=80)
            extra_lkl = st.text_area("拉卡拉：追加名称", value="", height=80)
            extra_fiscal = st.text_area("财政：追加名称", value="", height=80)

        extra_map = {
            "银行": parse_name_list(extra_bank),
            "微信": parse_name_list(extra_wechat),
            "现金": parse_name_list(extra_cash),
            "拉卡拉": parse_name_list(extra_lkl),
            "财政": parse_name_list(extra_fiscal),
        }
        match_map = {
            k: sorted(set(DEFAULT_TOTAL_MATCH.get(k, []) + extra_map.get(k, [])))
            for k in DEFAULT_TOTAL_MATCH.keys()
        }

    selected_types = sorted({str(t).upper() for t in types if str(t).strip()})
    state_prefix = f"{file_id}:{work_sheet}:{total_sheet}:{','.join(selected_types)}"

    errors: list[str] = []
    work_long = pd.DataFrame()
    total_long = pd.DataFrame()

    if not selected_types:
        errors.append("未选择收入类型（至少选择 1 个）。")

    try:
        work_long = normalize_work(tmp_path, work_sheet)
    except Exception as e:
        errors.append(f"读取/规范化工作表失败：{e}")

    try:
        total_long = normalize_total(tmp_path, total_sheet)
    except Exception as e:
        errors.append(f"读取/规范化总数表失败：{e}")

    # 工作表过滤
    work_filtered = pd.DataFrame()
    if not work_long.empty and selected_types:
        if "收入类型" not in work_long.columns:
            errors.append('工作表缺少“收入类型”列（脚本应自动生成，若为 None 请检查名称列）。')
        else:
            work_filtered = work_long[work_long["收入类型"].astype(str).str.upper().isin(selected_types)].copy()
            if work_filtered.empty:
                errors.append("工作表按所选收入类型过滤后为空（请检查收入类型选择或原始数据）。")

    # 总数表定位转账（贷方）
    transfer_credit = None
    transfer_hits = pd.DataFrame()
    total_summary: dict | None = None
    total_hit_tables: dict[str, pd.DataFrame] = {}
    if not total_long.empty:
        transfer_credit, transfer_hits = extract_transfer_credit(total_long)
        if transfer_credit is None:
            errors.append('总数表未定位到 name=“转账”的行，无法获取转账金额（贷方）。')
        elif abs(transfer_credit) <= tol:
            errors.append(f'总数表已定位到“转账”，但贷方合计为 {transfer_credit:.2f}（视为无效）。')
        else:
            total_summary, total_hit_tables = build_total_summary(total_long, float(transfer_credit), match_map)

    tab_report, tab_total, tab_taxpivot, tab_export = st.tabs(["📊 工作表报表", "🔍 总数校验", "💰 价税透视", "📤 导出"])

    alloc_state_key = f"alloc:{state_prefix}"
    alloc = pd.DataFrame()
    pivot = pd.DataFrame()
    validations: dict[str, pd.DataFrame] = {}
    summary = pd.DataFrame()
    summary_proj = pd.DataFrame()

    with tab_report:
        st.markdown("""
            <div class="info-card">
                <h4>📊 工作表数据处理</h4>
                <p>包含透视分析、分配管理、税费计算等核心功能</p>
            </div>
        """, unsafe_allow_html=True)
        
        if work_filtered.empty:
            st.error("工作表数据不可用：请先解决侧边栏参数或上传文件问题。")
        else:
            st.markdown(f"""
                <div style="background: rgba(255,255,255,0.8); padding: 1rem; border-radius: 8px; margin: 1rem 0; border-left: 4px solid #BF9E6B;">
                    <strong>📈 数据概览</strong><br>
                    收入类型：{', '.join(selected_types)} | 明细行数：{len(work_filtered):,} 条
                </div>
            """, unsafe_allow_html=True)

            st.markdown("##### 🔄 透视分析（名称 × 项目）")
            pivot = build_pivot(work_filtered)
            st.dataframe(pivot, use_container_width=True)

            st.markdown("##### ⚙️ 分配管理（可编辑：不计税 / 5% / 6%）")
            alloc_default = build_allocation_table(work_filtered)

            if alloc_state_key not in st.session_state:
                st.session_state[alloc_state_key] = alloc_default

            alloc = st.data_editor(
                st.session_state[alloc_state_key],
                key=f"editor:{state_prefix}",
                num_rows="dynamic",
                use_container_width=True,
            )
            st.session_state[alloc_state_key] = alloc

            c1, c2 = st.columns([1, 3])
            with c1:
                if st.button("🔄 重新计算", type="primary"):
                    st.rerun()
            with c2:
                st.caption('编辑后按回车或点出单元格，再点“重新计算”，即可按当前分配重新生成校验/税分拆/导出。')

            st.markdown("##### ✅ 校验分析（分配平衡检查）")
            validations = build_validation_tables(alloc)
            row_v = validations["行级"]
            bad = row_v[row_v["差额(分配-金额)"].abs() > tol].copy()

            total_amount = float(to_number(row_v["金额"]).sum()) if not row_v.empty else 0.0
            total_alloc = float(to_number(row_v["分配合计"]).sum()) if not row_v.empty else 0.0
            total_diff = round(total_alloc - total_amount, 2)

            m1, m2, m3 = st.columns(3)
            m1.metric("💰 金额总计", f"{total_amount:,.2f}")
            m2.metric("⚖️ 分配合计", f"{total_alloc:,.2f}")
            m3.metric("📊 总差额", f"{total_diff:,.2f}")

            if not bad.empty:
                errors.append(f'分配不平衡：{len(bad)} 行“分配合计”与“金额”不一致（容差 ±{tol}）。')

            if bad.empty:
                st.success("✅ 通过：行级分配已平衡（允许负数金额/负数分配）。")
            else:
                st.error("❌ 未通过：存在不平衡分配（将阻断导出）。")
                show_cols = ["名称", "项目", "金额", *ALLOCATION_COLS, "分配合计", "差额(分配-金额)", "备注"]
                show_cols = [c for c in show_cols if c in bad.columns]
                if not bad.empty:
                    bad = bad.sort_values(by="差额(分配-金额)", key=lambda s: s.abs(), ascending=False)
                    st.dataframe(bad[show_cols], use_container_width=True, height=260)

            st.markdown("##### 📋 汇总校验")
            if not validations["按名称"].empty:
                st.dataframe(validations["按名称"], use_container_width=True, height=200)
                if (validations["按名称"]["差额(分配-金额)"].abs() > tol).any():
                    errors.append("按名称汇总存在不平衡差额（将阻断导出）。")
            if not validations["按项目"].empty:
                st.dataframe(validations["按项目"], use_container_width=True, height=200)
                if (validations["按项目"]["差额(分配-金额)"].abs() > tol).any():
                    errors.append("按项目汇总存在不平衡差额（将阻断导出）。")
            st.dataframe(validations["总计"], use_container_width=True)
            if abs(float(validations["总计"].iloc[0]["差额(分配-金额)"])) > tol:
                errors.append("总计存在不平衡差额（将阻断导出）。")

            st.markdown("##### 💰 税分拆摘要（总计）")
            summary = summarize_tax(alloc)
            st.dataframe(summary, use_container_width=True)

            st.markdown("##### 📊 项目维度税分拆")
            summary_proj = summarize_tax_by_project(alloc)
            if not summary_proj.empty:
                st.dataframe(summary_proj, use_container_width=True)

    with tab_total:
        st.markdown("""
            <div class="info-card">
                <h4>🔍 总数表分析</h4>
                <p>包含长表转换、转账定位、倒推校验等财务核心功能</p>
            </div>
        """, unsafe_allow_html=True)
        
        if total_long.empty:
            st.error("总数表数据不可用：请检查总数 Sheet 名称是否正确。")
        else:
            if transfer_credit is not None:
                st.info(f"💰 转账（贷方合计）：{transfer_credit:,.2f}")
            if transfer_hits.empty:
                st.error('未找到“转账”定位行（将阻断导出）。')
            else:
                with st.expander("📋 转账定位明细（name=转账）", expanded=False):
                    st.dataframe(transfer_hits, use_container_width=True)

            if total_summary is not None:
                st.markdown("##### 📊 倒推校验指标")
                metrics_df = pd.DataFrame(
                    [{"指标": k, "金额": v} for k, v in total_summary.items()]
                )
                st.dataframe(metrics_df, use_container_width=True, height=360)

                with st.expander("📋 命中明细（按渠道）", expanded=False):
                    for k in ["银行", "微信", "现金", "拉卡拉", "财政"]:
                        names = match_map.get(k, [])
                        st.markdown(f"**{k}**（命中名称：{', '.join(names) if names else '无'}）")
                        hits = total_hit_tables.get(k, pd.DataFrame())
                        if hits is None or hits.empty:
                            st.caption("未命中任何行。")
                        else:
                            st.dataframe(hits[["source_row", "code", "name", "debit", "credit"]], use_container_width=True, height=160)

            with st.expander("📄 总数长表预览", expanded=False):
                st.dataframe(total_long, use_container_width=True, height=320)

    with tab_taxpivot:
        st.markdown("""
            <div class="info-card">
                <h4>💰 价税透视分析</h4>
                <p>动态列透视，按所选收入类型合并，用于价税分离和对账场景</p>
            </div>
        """, unsafe_allow_html=True)
        
        if work_filtered.empty:
            st.error("工作表数据不可用。")
        else:
            tax_pivot = build_pivot(work_filtered)
            st.dataframe(tax_pivot, use_container_width=True)
            st.markdown("""
                <div style="background: rgba(255,255,255,0.8); padding: 1rem; border-radius: 8px; margin: 1rem 0;">
                    <strong>📝 说明：</strong>此透视为动态列（名称×项目），用于后续价税分离/对账场景；当前与"透视"一致（按多选类型合并）。
                </div>
            """, unsafe_allow_html=True)

    with tab_export:
        st.markdown("""
            <div class="info-card">
                <h4>📤 数据导出</h4>
                <p>生成包含所有分析结果的Excel工作簿</p>
            </div>
        """, unsafe_allow_html=True)
        
        # 允许用户选择要导出的 Sheet（仍然会先做完整校验；校验失败会阻断导出）
        available_sheets = [
            "工作表_long",
            "透视",
            "分配表",
            "校验-行",
            "校验-名称",
            "校验-项目",
            "校验-总计",
            "税分拆",
            "总数_long",
            "总数校验",
            "转账定位",
        ]
        # 按实际数据可用性过滤（避免出现空/不存在的 sheet 选项）
        available_sheets = [s for s in available_sheets if s != "总数校验"] + (
            ["总数校验"] if total_summary is not None else []
        )
        export_key = f"export_sheets:{state_prefix}"
        selected_sheets = st.multiselect(
            "选择要导出的工作表（Sheet）",
            options=available_sheets,
            default=available_sheets,
            key=export_key,
        )

        if not selected_sheets:
            errors.append("未选择任何要导出的工作表（Sheet）。")

        can_export = len(errors) == 0 and not work_filtered.empty and not total_long.empty and bool(selected_sheets)
        if can_export:
            st.success("✅ 所有校验通过：可以下载汇总工作簿。")
        else:
            st.error("❌ 存在问题：已阻断导出。请先修正以下项：")
            for msg in sorted(set(errors)):
                st.markdown(f"""
                    <div style="background: rgba(214, 139, 179, 0.1); padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px; border-left: 4px solid #D68BB3;">
                        • {msg}
                    </div>
                """, unsafe_allow_html=True)

        # 生成并下载 Excel（仅在能导出时生成，避免浪费）
        buffer = None
        if can_export:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                # 工作表
                if "工作表_long" in selected_sheets:
                    work_filtered.to_excel(writer, sheet_name="工作表_long", index=False)
                if "透视" in selected_sheets:
                    pivot.reset_index().rename(columns={"index": "名称"}).to_excel(writer, sheet_name="透视", index=False)
                if "分配表" in selected_sheets:
                    alloc.to_excel(writer, sheet_name="分配表", index=False)
                if "校验-行" in selected_sheets:
                    validations["行级"].to_excel(writer, sheet_name="校验-行", index=False)
                if "校验-名称" in selected_sheets:
                    validations["按名称"].to_excel(writer, sheet_name="校验-名称", index=False)
                if "校验-项目" in selected_sheets:
                    validations["按项目"].to_excel(writer, sheet_name="校验-项目", index=False)
                if "校验-总计" in selected_sheets:
                    validations["总计"].to_excel(writer, sheet_name="校验-总计", index=False)
                if "税分拆" in selected_sheets:
                    # 同一个工作表中放置"税分拆摘要"和"项目维度税分拆"，各自保留表头，便于阅读
                    sheet = "税分拆"
                    summary.to_excel(writer, sheet_name=sheet, index=False, startrow=1)
                    ws = writer.sheets.get(sheet)
                    if ws is not None:
                        ws.write(0, 0, "税分拆摘要")

                    if summary_proj is not None and not summary_proj.empty:
                        title_row = 1 + len(summary) + 2
                        data_row = title_row + 1
                        if ws is not None:
                            ws.write(title_row, 0, "项目维度税分拆")
                        summary_proj.to_excel(writer, sheet_name=sheet, index=False, startrow=data_row)

                # 总数
                if "总数_long" in selected_sheets:
                    total_long.to_excel(writer, sheet_name="总数_long", index=False)
                if "总数校验" in selected_sheets and total_summary is not None:
                    pd.DataFrame([total_summary]).to_excel(writer, sheet_name="总数校验", index=False)
                if "转账定位" in selected_sheets:
                    transfer_hits.to_excel(writer, sheet_name="转账定位", index=False)

            buffer.seek(0)

        st.markdown("""
            <div style="text-align: center; margin-top: 2rem;">
        """, unsafe_allow_html=True)
        
        st.download_button(
            "📥 下载汇总工作簿",
            data=buffer.getvalue() if buffer is not None else b"",
            file_name="报表输出.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=not can_export,
        )
        
        st.markdown("""
            </div>
        """, unsafe_allow_html=True)

    # 清理临时文件（不影响导出：导出已在内存）
    tmp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
