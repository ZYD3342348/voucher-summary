# 项目总结：总台收入工作台（Excel → 透视 / 分配 / 税分拆 / 总数校验）

本项目把“总台”原始 Excel 中的 **工作表明细** 与 **总数表** 整合为一套可复用的处理链路：既支持命令行批处理，也提供本机 Streamlit Web 工作台（可编辑分配、自动校验、导出单一工作簿）。

---

## 1. 目标与范围

- 输入：单个 Excel（默认含 `工作表` + `总数`，也支持在 Web 侧边栏选择实际 Sheet）。
- 输出：单个 Excel 工作簿（多 Sheet，且可在导出时勾选要包含哪些 Sheet）。
- 场景：本机单人使用，不做登录/多租户/云部署。

---

## 2. 功能清单

### 2.1 Web 工作台（Streamlit）

路径：`voucher-summary-generator/app.py`

- **工作表报表**
  - 工作表明细规范化（项目/名称/金额自动定位）
  - 收入类型筛选（单选/多选合并）
  - 透视：`名称 × 项目`（动态列）
  - 分配表：对每个 `名称 × 项目` 的金额进行 `不计税/5%/6%` 人工分配
  - 税分拆：自动计算不含税与税额汇总（总计 + 项目维度）
  - 校验：行级 + 按名称汇总 + 按项目汇总 + 总计（不通过则阻断导出）
- **总数校验**
  - 总数表规范化为长表（`code/name/debit/credit/source_row`）
  - 自动定位“转账”（按贷方合计）
  - 倒推校验指标（凭证借贷方、应挂账等）
  - **命中名单：默认 + 可追加**（银行/微信/现金/拉卡拉/财政），并展示命中明细便于追溯
- **价税透视**
  - 按所选收入类型合并的 `名称 × 项目` 动态列透视（作为价税分离/对账底稿）
- **导出**
  - 只导出一个工作簿 `报表输出.xlsx`
  - 导出前全量校验，任意错误会阻断下载
  - 允许勾选要导出的工作表（Sheet）

### 2.2 命令行脚本

- `generate_summary.py`（根目录）：从 `收入类型表` 生成单表“汇总/凭证链”输出（用于公式链验证、调整 S 等）。
- `voucher-summary-generator/scripts/normalize_work.py`：将 `工作表` 规范化为长表（并打印收入类型分布与房费透视）。
- `voucher-summary-generator/scripts/normalize_total.py`：将 `总数` 规范化为长表，并输出倒推校验指标（CLI 版）。
- `voucher-summary-generator/scripts/report_engine.py`：从 `工作表` 生成（透视/调整表/税分拆摘要/项目税分拆）多 Sheet 报表（CLI 版雏形）。
- `voucher-summary-generator/scripts/generate_tax_sep.py`：生成按收入类型过滤的 `名称 × 项目` 动态列透视（CLI 版）。

---

## 3. 技术栈与选择理由

- **Python 3.9+**
  - 本机可直接运行，生态成熟，适合财务 Excel 自动化。
- **pandas**
  - 强项是表格清洗、透视、分组汇总；Excel I/O 稳定且开发效率高。
- **openpyxl**
  - 用于需要写入公式/样式/单表布局的场景（例如根目录 `generate_summary.py` 的公式链输出）。
- **xlsxwriter（通过 pandas.ExcelWriter）**
  - 用于 Web/报表导出多 Sheet；写入速度快，兼容性好。
- **Streamlit**
  - 低门槛 Web 工作台：支持上传文件、可编辑表格、即时计算、下载按钮；无需单独前端工程。

---

## 4. 数据处理规则（关键口径）

### 4.1 名称清洗（对齐 Excel「数据分列」）

在 `工作表` 中，名称列遵循你在 Excel 的操作习惯：

1) **先提取收入类型**：从“清洗前名称”中，找到遇到的**第一个英文字母**，作为收入类型（大写）。
2) **再清洗名称**：清除所有空白字符，然后按 `_` 分列，只保留第一段（等价 `split('_')[0]`）。

实现：`voucher-summary-generator/scripts/normalize_work.py`

### 4.2 总数校验（倒推凭证借贷方）

核心逻辑（Web 版）：

- `借方合计`：总数长表中所有借方求和。
- `转账(贷方)`：`name == "转账"` 的贷方合计（按贷方取数）。
- `转内部成本`：`name == "转内部成本"` 的贷方合计。
- `凭证贷方`：`借方合计 - 转账(贷方) - 转内部成本`
- `资金类合计`：银行/微信/现金/拉卡拉/财政（按命中名单在贷方求和）
- `应挂账金额`：`凭证贷方 - 资金类合计`
- `凭证借方`：`资金类合计 + 应挂账金额`（因此必然等于凭证贷方，用于配平）

实现：`voucher-summary-generator/app.py`（`build_total_summary`）

---

## 5. 使用方法

### 5.1 启动 Web

```bash
cd /Users/zengyuntan/python/凭证处理/2025年12月/voucher-summary-generator
streamlit run app.py
# 若本机 streamlit 命令报 bad interpreter，可用：
python3 -m streamlit run app.py
```

操作流程：

1) 上传单个 Excel（含 `工作表` 和 `总数`）。
2) 选择收入类型（单选/多选）。
3) 在「分配表」里修改分配。
4) 查看「校验」是否通过（不通过会阻断导出）。
5) 在「导出」勾选要导出的 Sheet，下载 `报表输出.xlsx`。

### 5.2 运行命令行

汇总/凭证链：

```bash
cd /Users/zengyuntan/python/凭证处理/2025年12月
python3 generate_summary.py
```

工作表/总数规范化与报表（示例）：

```bash
python3 voucher-summary-generator/scripts/normalize_total.py -i "测试数据/2025年8月总台.xlsx" -s "总数" -o "2025年8月_总数_long.csv"
python3 voucher-summary-generator/scripts/report_engine.py -i "测试数据/2025年8月总台.xlsx" -w "工作表" -o "2025年8月_报表.xlsx" -t H
python3 voucher-summary-generator/scripts/generate_tax_sep.py -i "测试数据/2025年8月总台.xlsx" -s "工作表" -o "2025年8月_H_价税分离.xlsx" -t H
```

### 5.3 Windows 运行与打包建议（无 Python 环境的电脑）

本项目是 Python 方案：在“没有 Python 环境”的 Windows 电脑上，要么安装 Python，要么做打包/便携发布。推荐先保证“可运行”，再做打包。

**在 Windows 上先跑通（推荐命令）**

1) 安装 Python 3.9+（建议 3.10/3.11），并勾选 “Add Python to PATH”  
2) 在仓库根目录执行依赖安装：

```powershell
python -m pip install -U pip
python -m pip install streamlit pandas openpyxl xlsxwriter
```

3) 启动 Web（避免 `streamlit.exe` 路径问题，建议用模块方式）：

```powershell
python -m streamlit run voucher-summary-generator\\app.py
```

**方式 A：便携版（推荐先做，最省心）**

思路：把 `venv` 连同项目一起打包成 zip，解压即用（不追求单 exe）。

```powershell
python -m venv .venv
.\\.venv\\Scripts\\pip install -U pip
.\\.venv\\Scripts\\pip install streamlit pandas openpyxl xlsxwriter
```

然后写一个 `run.bat`（示例）：

```bat
@echo off
call .venv\\Scripts\\activate
python -m streamlit run voucher-summary-generator\\app.py
pause
```

**方式 B：打包成 exe（体积更大，坑更多）**

可以使用 PyInstaller/Nuitka 等工具，但 Streamlit + pandas 组合通常体积较大（常见为 100MB～数百 MB），并且需要处理依赖/动态导入问题。建议在方式 A 稳定后再做方式 B。

**注意：字体/样式**

前端样式默认**不依赖联网字体**（离线可用）。如希望使用 Google Fonts 的更美观字体，可在 `voucher-summary-generator/static/custom_styles.css` 里取消 `@import` 的注释；离线环境会自动使用系统字体（功能不受影响，仅观感略有变化）。

---

## 6. 过程踩坑与修复（经验沉淀）

- **Streamlit 必须用 `streamlit run` 启动**
  - 直接 `python app.py` 会出现 `ScriptRunContext` 等警告，页面交互也会异常。
- **streamlit 命令可能因 shebang 失效**
  - 本机如果出现 `bad interpreter`，用 `python3 -m streamlit run app.py` 兜底。
- **不同月份“工作表”的列位置可能不同**
  - 早期按固定列号（项目0/名称2）会在某些月份（如 10 月）直接 KeyError。
  - 现已改为“扫描表头定位 + 自动识别金额列”。
- **“总数”表尾部的汇总行会导致重复计入**
  - 部分文件在表尾存在 `code/name 为空` 但借/贷给出总计的汇总行；若当作明细，会把借方合计翻倍。
  - 现已在总数规范化时过滤该类汇总行。
- **Web 编辑表格会 rerun，需保留编辑态**
  - 通过 `st.session_state` 保存分配表，避免每次刷新回到默认值。
- **负数金额是合法业务场景**
  - 校验口径改为仅校验“分配合计是否等于金额”，不再因负数阻断导出。
- **部分 UI 无法彻底中文化**
  - Streamlit 自带右上角菜单无法完全中文化，已通过页面按钮与提示做“可用性兜底”。

---

## 7. 如何修改与维护（推荐路径）

### 7.1 修改总数命中规则（默认 + 追加）

- 默认命中名单：`voucher-summary-generator/app.py` 中 `DEFAULT_TOTAL_MATCH`
- Web 追加命中：侧边栏「总数命中名单（默认 + 可追加）」
- 排查命中：总数校验 Tab → 「命中明细（按渠道）」展开查看命中的具体行

### 7.2 修改名称清洗 / 收入类型提取

统一入口：

- `voucher-summary-generator/scripts/normalize_work.py`：Web 与大多数 CLI 都复用这里的逻辑
- 若要保持一致，请同步更新：
  - `voucher-summary-generator/scripts/generate_tax_sep.py`
  - 根目录 `generate_summary.py`

关键函数：

- `derive_income_type()`：收入类型提取
- `clean_name()`：名称清洗（当前对齐 Excel「数据分列」）

### 7.3 修改税率/分配列

位置：`voucher-summary-generator/app.py`

- 分配列：`ALLOCATION_COLS`
- 税分拆计算：`summarize_tax()` / `summarize_tax_by_project()`

如需新增税率（例如 3%），建议：

1) 扩展 `ALLOCATION_COLS`
2) 在 `summarize_tax()` 里补对应 `net_tax()` 计算
3) 在校验里把新列纳入“分配合计”

### 7.4 修改导出内容

位置：`voucher-summary-generator/app.py` 的「导出」tab

- `available_sheets`：控制可选导出项
- 写入逻辑：`pd.ExcelWriter(...).to_excel(...)`

---

## 8. 验证与回归

建议回归命令（会跑一组测试数据输出）：

```bash
cd /Users/zengyuntan/python/凭证处理/2025年12月
python3 run_tests.py
```
