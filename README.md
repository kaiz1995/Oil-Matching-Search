# 润美润滑油产品匹配工具 (Runmei Oil Product Matching Tool)

## 项目简介

竞品润滑油与润美产品智能匹配查询工具。帮助销售、技术人员快速将壳牌、美孚、雪佛龙等品牌润滑油产品对应到润美（Runmei）等效产品，支持产品查找、参数对比、粘度计算等功能。

## 功能特性

### 智能搜索
- **跨字段自由组合搜索** — 支持品牌、型号、粘度等任意组合，如搜"壳牌S2 68"自动匹配品牌+型号+粘度
- **模糊匹配** — 分段分词 AND 逻辑，输入碎片信息也能找到对应产品
- **搜索结果高亮** — 匹配关键词在结果中高亮显示
- **相关性排序** — 按匹配度评分，最相关结果排最前

### 产品对照
- **竞品→润美映射** — 224 款竞品产品对照，171 款已匹配润美对应型号（72.7% 匹配率）
- **多维筛选** — 按品牌（10个）、品类（9类）、技术系列筛选
- **产品详情** — 技术参数（粘度、闪点、倾点等）、应用行业、包装规格
- **一键复制** — 润美对应产品名一键复制到剪贴板

### 润美产品目录
- 91 款润美产品完整技术参数
- 按技术系列浏览筛选
- 详细参数对比表格

### 粘度计算工具
- ASTM D2270 标准粘度指数计算
- 基础油调合配比计算
- 40℃/100℃ 运动粘度换算

### Android App
- 原生 Android WebView 封装，离线数据支持
- 在线检查更新，DownloadManager 下载 + 自动安装
- 内置 fallback 数据，无网络也可查询
- 国内网络优化：APK 通过 GitHub Pages 分发

## 数据来源

数据维护在 Excel 文件 `竞品产品对照表.xlsx` 中，包含：

| 品牌 | 覆盖品类 |
|------|---------|
| 壳牌/Shell | 液压油、齿轮油、涡轮机油、空压机油、轴承油等 |
| 美孚/Mobil | 液压油、齿轮油、涡轮机油等 |
| 雪佛龙/Chevron | 液压油、齿轮油等 |
| 长城/Sinopec | 液压油、齿轮油、空压机油等 |
| 克鲁勃/Klüber | 特种润滑油脂 |
| 其他 | 嘉实多、福斯、安索、昆仑等 |

## 技术架构

### 数据管线

```
Excel 数据源
    ↓
build_all.py (Python 数据合并+匹配)
    ↓
data.json ───→ GitHub Pages (CDN分发)
    │
    ├──→ HTML 网页工具 (在线版)
    │
    └──→ Android App (离线版)
```

### 技术栈

| 层 | 技术 |
|----|------|
| 数据源 | Excel (openpyxl) |
| 后端管线 | Python 3 (JSON 处理、模糊匹配) |
| 网页前端 | 原生 JavaScript (无框架)、CSS 变量 |
| Android | Kotlin + WebView + DownloadManager |
| 分发 | GitHub Pages（数据 + APK） |

### 项目结构

```
oil-pm/
├── build_all.py                 # 统一构建管线
├── lubricant_product_matching_tool.html  # 网页工具 (核心)
├── data.json                    # 构建产物 (产品数据)
├── 竞品产品对照表.xlsx           # Excel 数据源
├── 竞品产品对照表_已填充.xlsx     # 填充后的数据源
├── competitor_final.json        # JSON 数据源 (备选)
├── product_data_merged.json     # 合并后 JSON
└── RunmeiMatching.apk           # Android APK

Oil-Matching-Search/            # GitHub Pages 仓库
├── data.json                   # 在线数据 (CDN)
├── RunmeiMatching.apk          # APK 下载
└── README.md                   # 本文件

RunmeiMatching/                 # Android 项目
├── app/src/main/
│   ├── java/.../MainActivity.kt  # WebView + JS Bridge
│   ├── assets/index.html         # 内置网页
│   └── AndroidManifest.xml
└── app/build.gradle.kts
```

## 使用方式

### 网页版
直接访问 GitHub Pages：`https://kaiz1995.github.io/Oil-Matching-Search/`（需要自行部署 HTML 文件）

或本地打开 `lubricant_product_matching_tool.html`（内置数据，离线可用）。

### Android App
下载 `RunmeiMatching.apk` 安装，支持在线更新和离线查询。

### 搜索技巧
- 简单搜：`46` → 列出所有粘度 46 的产品
- 组合搜：`壳牌得力士 46` → 品牌+型号+粘度
- 缩写搜：`S2 68` → 系列+粘度
- 品类搜：`齿轮油 320` → 品类+粘度

## 开发指南

### 更新数据
1. 编辑 `竞品产品对照表.xlsx`
2. 运行 `python build_all.py` → 自动生成 data.json 和 HTML
3. 推送到 GitHub Pages

### 发布新版 APK
```bash
python build_all.py --release
```
会自动构建 APK、创建 GitHub Release、更新 data.json。

### 环境要求
- Python 3.8+
- openpyxl, Pillow
- JDK 17+ (构建 APK)
- Android SDK (构建 APK)

## 版本历史

| 版本 | 更新内容 |
|------|---------|
| 1.0.3 | 修复复制按钮不可用的问题 |
| 1.0.2 | 搜索匹配逻辑优化、支持跨字段自由组合搜索、搜索结果高亮 |
| 1.0.1 | 匹配逻辑优化、支持多维度搜索 |
| 1.0.0 | 初始版本，竞品-润美产品对照查询 |

## License

Internal use.
