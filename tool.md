# 润滑油产品匹配查询工具

## 项目背景
作为润滑油产品经理，每天需要处理大量销售询问的产品匹配工作。本工具旨在让销售人员实现自主查询，减少产品经理的重复性工作。

## 工具形式
Excel (.xlsx) 文件为根数据，通过运行 `build_all.py` 更新，通过 html 文件实现跨表自动查询。

## 核心文件

| 文件 | 路径 | 说明 |
|------|------|------|
| **生成工具** | `build_all.py` | 统一管线：Excel→JSON→HTML 生成 |
| **HTML 工具** | `lubricant_product_matching_tool.html` | 自包含单页 HTML，嵌入 JSON 数据，浏览器直接打开 |
| **合并数据** | `product_data_merged.json` | 管线输出的 JSON |
| **根数据源** | `竞品产品对照表_已填充.xlsx` | 用户编辑的 Excel，含润美映射，管线首选读取 |
| **原始表** | `竞品产品对照表.xlsx` | 原始竞品表（无润美列）的副本 |
| **映射库** | `../Lubricant Product Matching Inquiry/product_data.json` | 润美映射库（106 条映射） |
| **填充脚本** | `fill_excel_runmei.py` | 用映射数据回填 Excel 的润美列 |
| **最终清洗** | `final_clean.py` | 去噪声 + 合并产品族 |
| **记忆文件** | `chauxngongju.md` | 本文件 |

## Excel 列映射说明

_已填充.xlsx 的列 → 内部字段：
- `品牌` → brand
- `类型` → category
- `名称` → product_cn
- `英文名称` → product_en
- `粘度等级` → viscosity
- `润美产品` → runmei
- `特性` → source

原始表略有不同：`产品类型` → category, `竞品产品名称` → product_cn, `竞品英文名称` → product_en, `ISO VG / 粘度等级` → viscosity, `润美对应产品（请填写）` → runmei, `数据来源` → source

## 数据管线流程（build_all.py）

1. **读数据源**：`_已填充.xlsx` → `竞品产品对照表.xlsx` → `competitor_final.json`（依次尝试）
2. **载入映射**：从 `product_data.json` 读 106 条映射，建三个索引（精确名、产品族、品牌-品类）
3. **匹配润美**（4 步 Python 策略）：
   - 精确名称匹配
   - 去品牌前缀后匹配
   - 产品族 + 粘度交叉匹配
   - 品牌 + 品类兜底匹配
4. **去重**：按产品名去重，优先保留有润美映射的条目
5. **品牌标准化**：`壳牌/Shell` → `壳牌`
6. **生成**：`product_data_merged.json` + `lubricant_product_matching_tool.html`

## HTML 工具功能

### 三个面板
1. **竞品匹配** — 输入类别/品牌/型号/粘度/等级，自动评分匹配润美产品
2. **分类浏览** — 按产品类型 + 品牌浏览，带文本过滤
3. **快速搜索** — 全局全文搜索

### JS 评分算法（matchScore）
- **品类**：+15 精确匹配，+8 相关匹配
- **品牌**：+15 匹配；-15 不匹配惩罚（用户明确选品牌时）
- **型号**：token 比例制（≥80% = +20，≥50% = +12，≥25% = +6）
- **粘度**：+10 精确，+5 相邻等级（遵循 ISO VG 序列）
- **等级代号**：+8/每项（HM/HV/CKC/CKD/TSA/PAO/SHC 等）
- **特性关键词**：+2/每项
- **润美存在**：+5；泛搜索（无型号/粘度/等级）额外 +5
- **噪声惩罚**：长名 -3，含标点 -3
- **匹配等级阈值**：≥30 强烈推荐，≥18 推荐匹配，≥10 可能匹配

### 生成注意事项
- Python f-string 中 JS 的 `{`/`}` 需双写为 `{{`/`}}`
- 复制按钮用 `data-runmei` 属性 + `escAttr()` 避免 onclick 引号转义问题
- 品牌下拉选项随品类动态过滤
- localStorage 保存上次搜索状态

### JS 关键函数
- `tokenize(s)` — 分词（支持驼峰 + 数字分隔）
- `tokenMatch(inputTokens, productTokens)` — token 比例匹配
- `extractGrade(text)` — 提取等级代号（正则匹配 8 类模式）
- `extractVG(text)` — 提取 ISO VG 粘度值
- `normalizeBrand(b)` — 品牌标准化
- `matchScore(input, product)` — 综合评分
- `getMatchLevel(score)` — 评分→等级转换
- `searchCompetitor()` — 竞品匹配搜索入口
- `copyRunmei(btn)` — 从 `data-runmei` 属性读取并复制
- `quickPreset(cat, brand, model, vis)` — 预设快捷按钮
- `esc(s)` / `escAttr(s)` — HTML 内容/属性值转义

## 当前数据（上次构建 2026-05-19）
- 产品总数：208
- 润美映射：116
- 品牌：8 个（壳牌、美孚、雪佛龙、长城、克鲁勃、嘉实多、福斯、昆仑）
- 品类：10 个（液压油、齿轮油、空压机油、涡轮机油、轴承油、循环油、压缩机油、润滑脂等）
- HTML 大小：101.3 KB

## 已知要点
- `build_all.py` 中品类映射（如 空压机油/压缩机油 → 空压机油）需在 CAT_MAP 字典维护
- JS 端 CAT_MAP 与 Python 端保持一致
- 快速搜索使用独立评分逻辑（非 matchScore），可单独优化
