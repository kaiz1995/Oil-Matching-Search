# 润美润滑油产品匹配查询工具 — 开发仓库

## 项目概览

竞品润滑油与润美（Runmei）产品智能匹配查询工具。数据维护在 Excel，通过 Python 管线自动构建 JSON + 离线 HTML，推送到 GitHub Pages 供手机 App 更新。

## 目录结构

```
润滑产品查询工具/
├── build_all.py                     # 📌 统一构建管线（核心）
├── RunmeiMatching-1.1.2.html       # 网页工具产物（离线单页，内嵌数据）
├── data.json                        # 数据产物（供 App 拉取）
├── product_data_merged.json         # 合并后数据（中间产物）
│
├── 竞品产品对照表.xlsx              # Excel 数据源（原始）
├── 竞品产品对照表_已填充.xlsx        # Excel 数据源（已填润美映射，首选读取）
│
├── build_matching_data.py           # 辅助：从 product_data.json 生成匹配映射
├── build_html_tool.py               # 辅助：生成 HTML（build_all 子模块）
├── fill_excel_runmei.py             # 辅助：用映射库回填 Excel 的润美列
├── fill_params_from_pdf.py          # 辅助：从 PDF 说明书提取参数填 Excel
├── extract_competitor_data.py       # 辅助：提取竞品对照表数据
├── clean_and_export.py              # 辅助：清洗 + 导出
├── final_clean.py                   # 辅助：去噪声 + 合并产品族
├── split_html_data.py               # 辅助：将内嵌 HTML 拆分为外置 data.json
│
├── competitor_final.json            # 竞品目录 JSON（Excel 不可用时 fallback）
├── fallback.json                    # Android 离线数据副本
│
├── logo.png / logo2.png / logo3.png # 网页 Logo
├── jingpinziliao/                   # 竞品技术资料（PDF）
├── screenshots/                     # 截图
│
└── 竞品产品对照表_已填充.xlsx        # 当前最新数据源
```

## 数据管线

```
Excel 编辑（用户操作）
    │ 填写竞品 + 润美映射 + 润美产品目录
    ▼
build_all.py（Python 管线）
    │ 1. 读取 Excel（_已填充.xlsx）
    │ 2. 载入 product_data.json 映射库
    │ 3. 智能匹配润美（名称→族→品牌品类）
    │ 4. 去重 + 品牌标准化
    │ 5. 嵌入 Logo + 更新 HTML
    │ 6. 复制到 GitHub Pages 仓库
    │ 7. git commit + push
    │ 8. 复制 Android assets
    ▼
Oil-Matching-Search/data.json ──→ GitHub Pages CDN ──→ App 在线更新
Oil-Matching-Search/RunmeiMatching-X.X.X.apk ──→ App 下载安装
```

## Excel 数据源说明

### Sheet: `竞品产品对照表`（竞品数据）

| 列 | 字段 | 说明 |
|----|------|------|
| A | 品牌 | 竞品品牌（壳牌/美孚/长城等） |
| B | 类型 | 产品类型（液压油/齿轮油等） |
| C | 名称 | 竞品产品名称 |
| D | 英文名称 | 竞品英文名 |
| E | 粘度等级 | ISO VG 等级 |
| F | 润美产品 | 🖊️ **用户填写**的对应润美产品 |
| G | 特性 | 特性描述 / 数据来源 |

### Sheet: `润美产品查询`（润美产品目录，31 列）

| 列 | 字段 | 技术参数 |
|----|------|---------|
| A | 技术系列 | 系列分类 |
| B | 产品牌号 | 产品名 |
| C~G | 包装/行业/应用/特性/备注 | 基础信息 |
| I~T | kv40~weld_load | 12 项技术参数 |
| AE | fzg | ✅ **FZG 失效负荷等级** |

### 其他 Sheet

- **说明** — A1 单元格 = 版本注脚（显示在工具底部）
- **版本更新记录** — 版本号 / 更新说明 / 是否强制更新
- **发布页** — 下载地址名称 / URL

## 构建命令

```bash
# 更新数据 + 推送到 GitHub Pages
python build_all.py

# 构建 APK + 创建 GitHub Release
python build_all.py --release

# 仅构建，跳过 git 操作
python build_all.py --skip-git

# 构建（跳过 APK）
python build_all.py --release --skip-apk
```

## 数据统计（最新）

| 指标 | 数值 |
|------|------|
| 产品总数 | 326 |
| 已匹配润美 | 183（56.1%） |
| 品牌数 | 15 |
| 品类数 | 13 |
| 润美产品目录 | 96 |

## 润美匹配策略（Python 端 4 步）

1. **精确名称匹配** — 归一化后精确匹配竞品名
2. **去品牌前缀匹配** — 去掉"壳牌""美孚"等前缀再匹配
3. **产品族匹配** — 同系列不同粘度交叉匹配
4. **品牌 + 品类兜底** — 同品牌同品类按粘度交叉

**Excel 手填值优先**：F 列（润美产品）有值时，跳过自动匹配。

## 发布流程

```bash
# 1. 编辑 Excel，填写润美映射 + 润美产品目录
# 2. 在"版本更新记录"sheet 添加新版本行
# 3. 运行完整发布
python build_all.py --release

# 4. 自动完成：
#    - data.json → Oil-Matching-Search 仓库
#    - APK → Oil-Matching-Search 仓库
#    - git commit + push + tag
#    - GitHub Release 创建
#    - Android assets 更新
```

## 部署仓库

- **GitHub Pages（CDN）**: `kaiz1995/Oil-Matching-Search`
- **在线数据**: `https://kaiz1995.github.io/Oil-Matching-Search/data.json`
- **APK 下载**: `https://kaiz1995.github.io/Oil-Matching-Search/RunmeiMatching.apk`

## 技术参数列映射

| 键名 | 中文名 | Excel 列 |
|------|--------|---------|
| kv40 | 运动粘度40℃(mm²/s) | I |
| kv100 | 运动粘度100℃(mm²/s) | J |
| vi | 粘度指数 | K |
| pour_point | 倾点(℃) | L |
| flash_point | 闪点(℃) | M |
| copper_corrosion | 铜片腐蚀(100℃,3h) | N |
| rust_a | 液相锈蚀A法 | O |
| rust_b | 液相锈蚀B法 | P |
| cleanliness_nas | 清洁度(NAS1638) | Q |
| cleanliness_iso | 清洁度(ISO 4406) | R |
| wear_scar | 磨斑直径(mm) | S |
| weld_load | 烧结负荷(N) | T |
| fzg | FZG失效负荷等级 | AE |

## Android 项目

- 路径：`C:\Users\张大脸小太阳\AndroidStudioProjects\RunmeiMatching\`
- Kotlin + WebView 封装，离线数据支持
- 在线更新检测 + DownloadManager 下载安装
- JS Bridge（`AndroidBridge`）提供剪贴板、Toast、版本检测
- 内置 `assets/index.html` + `assets/fallback.json`
