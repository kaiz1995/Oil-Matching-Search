#!/usr/bin/env python3
"""
Unified build pipeline: Excel → JSON → HTML
1. Reads 竞品产品对照表.xlsx as PRIMARY source (user edits this file)
2. Reads product_data.json (detailed Runmei mappings, as supplement)
3. Falls back to competitor_final.json if Excel unavailable
4. Smart matching to fill Runmei where possible
5. Outputs: product_data_merged.json + lubricant_product_matching_tool.html
"""
import json, re, sys, os
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r'D:\HuaweiMoveData\Users\张大脸小太阳\Documents\cc workspace\oil-pm'
MAPPED = r'D:\HuaweiMoveData\Users\张大脸小太阳\Documents\cc workspace\Lubricant Product Matching Inquiry\product_data.json'
CATALOG = f'{BASE_DIR}/competitor_final.json'
EXCEL_PRIMARY = f'{BASE_DIR}/竞品产品对照表.xlsx'
EXCEL_FILLED = f'{BASE_DIR}/竞品产品对照表_已填充.xlsx'

# ============================================================
# STEP 1: Load data sources
# ============================================================

with open(MAPPED, 'r', encoding='utf-8') as f:
    mapped_data = json.load(f)
mappings = [m for m in mapped_data['products'] if m.get('competitorProduct')]
print(f"Mapped products: {len(mappings)}")

# Try reading from Excel first (user's primary editing surface)
def read_catalog_from_excel(excel_path):
    """Read product catalog from Excel file. Returns list of dicts or None."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
        headers = [str(ws.cell(1, c).value or '') for c in range(1, ws.max_column + 1)]
        # Map headers to field names
        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if '品牌' in h:
                col_map['brand'] = i
            elif '类型' in h or '产品类型' in h:
                col_map['category'] = i
            elif '竞品产品名称' in h or '产品名称' in h or h == '名称':
                col_map['product_cn'] = i
            elif '英文' in h:
                col_map['product_en'] = i
            elif '粘度' in h or 'vg' in h.lower():
                col_map['viscosity'] = i
            elif '润美' in h:
                col_map['runmei'] = i
            elif '来源' in h or '数据来源' in h or '特性' in h:
                col_map['source'] = i

        products = []
        for row in range(2, ws.max_row + 1):
            def cell(col_name):
                idx = col_map.get(col_name)
                if idx is None:
                    return ''
                val = ws.cell(row, idx + 1).value
                return str(val).strip() if val else ''

            cn = cell('product_cn')
            brand = cell('brand')
            if not cn or not brand:
                continue  # Skip empty rows

            # Skip header-like rows
            if cn in ('竞品产品名称', '产品名称', '品牌'):
                continue

            products.append({
                'brand': brand,
                'category': cell('category'),
                'product_cn': cn,
                'product_en': cell('product_en'),
                'viscosity': cell('viscosity'),
                'runmei': cell('runmei'),
                'source': cell('source'),
            })
        wb.close()
        return products if products else None
    except Exception as e:
        print(f"  [WARN] Cannot read Excel: {e}")
        return None

# Try filled Excel first (user's primary editing surface), then original, then JSON
catalog = None
for excel_path in [EXCEL_FILLED, EXCEL_PRIMARY]:
    if os.path.exists(excel_path):
        catalog = read_catalog_from_excel(excel_path)
        if catalog:
            print(f"Catalog from Excel: {len(catalog)} products ({os.path.basename(excel_path)})")
            break

if not catalog:
    # Fallback to JSON catalog
    with open(CATALOG, 'r', encoding='utf-8') as f:
        catalog = json.load(f)
    # Filter noise entries
    NOISE_WORDS = [
        '服务流程', '服务体系的组成', '全程支持', 'LubeAssist',
        'EfficiencySupport', 'Energy克鲁勃能效', 'Maintain克鲁勃维护',
        'Monitor克鲁勃监控', 'Renew克鲁勃更新', 'OilSuite',
        '技术支持', '咨询服务', '培训服务',
    ]
    catalog = [p for p in catalog if not any(k in p.get('product_cn', '') for k in NOISE_WORDS)]
    print(f"Catalog from JSON (fallback): {len(catalog)} products")

# ============================================================
# STEP 2: Normalization & matching utilities
# ============================================================

def normalize(s):
    """Normalize for comparison"""
    s = s.lower().strip()
    s = re.sub(r'[™®\s]+', ' ', s)
    s = re.sub(r'[^\w\s一-鿿]', '', s)
    return s.strip()

def extract_viscosities(name):
    return set(int(x) for x in re.findall(r'\b(\d{2,3})\b', name))

def extract_family(name):
    """Extract product family name (brand + series, without viscosity)"""
    name = re.sub(r'\d{2,3}(?:\s*/\s*\d{2,3})*', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def brand_match(b1, b2):
    """Check if two brand strings refer to the same brand"""
    b1, b2 = normalize(b1), normalize(b2)
    for b in ['壳牌', 'shell', '美孚', 'mobil', '雪佛龙', 'chevron',
              '长城', 'sinopec', '克鲁勃', 'kluber', 'klüber',
              '嘉实多', 'castrol', '福斯', 'fuchs', '安索', 'amsoil',
              '昆仑', 'kunlun']:
        if b in b1 and b in b2:
            return True
    return b1 == b2

# Category normalization
CAT_MAP = {
    '空压机油/压缩机油': '空压机油',
    '汽轮机油/涡轮机油': '涡轮机油',
    '轴承油/循环油': '轴承油',
}

def cat_match(c1, c2):
    c1 = CAT_MAP.get(c1, c1)
    c2 = CAT_MAP.get(c2, c2)
    return c1 == c2

# ============================================================
# STEP 3: Build lookup indexes from mapped data
# ============================================================

name_index = {}       # normalized name → mapping
family_index = {}     # (family, category) → [mappings]
brand_cat_index = {}  # (brand, category) → [mappings]

for m in mappings:
    if not m.get('competitorProduct'):
        continue
    norm = normalize(m['competitorProduct'])
    family = normalize(extract_family(m['competitorProduct']))
    cat = m['category']
    brand = m['competitorBrand']
    runmei = m.get('runmeiProduct', '')

    name_index[norm] = m
    if family and runmei:
        key = (family, cat)
        if key not in family_index:
            family_index[key] = []
        family_index[key].append(m)

    bck = (brand, cat)
    if bck not in brand_cat_index:
        brand_cat_index[bck] = []
    brand_cat_index[bck].append(m)

print(f"Name index: {len(name_index)} entries")
print(f"Family index: {len(family_index)} families")
print(f"Brand-cat index: {len(brand_cat_index)} groups")

# ============================================================
# STEP 4: Smart matching function
# ============================================================

def find_runmei(catalog_entry):
    """Find best Runmei match for a competitor product from catalog"""
    cn_name = catalog_entry['product_cn']
    brand = catalog_entry['brand']
    cat = catalog_entry['category']

    norm = normalize(cn_name)

    # Step 1: Exact name match
    if norm in name_index:
        m = name_index[norm]
        return m.get('runmeiProduct', ''), m

    # Step 2: Name match without brand prefix
    for prefix in ['壳牌', '美孚', '雪佛龙', '长城', '克鲁勃', '长城牌',
                   'Shell', 'Mobil', 'Chevron', 'Klüber', 'Kluber']:
        if prefix in cn_name:
            bare = normalize(cn_name.replace(prefix, ''))
            if bare in name_index:
                m = name_index[bare]
                return m.get('runmeiProduct', ''), m
            break

    # Step 3: Family-based match (same product family, different viscosity)
    cat_brand = brand.split('/')[0].strip()  # "壳牌/Shell" → "壳牌"
    cn_family = normalize(extract_family(cn_name))

    for (family, fcat), entries in family_index.items():
        family_parts = set(family.split())
        cn_parts = set(cn_family.split())
        common = family_parts & cn_parts
        # Allow family match across related categories
        if len(common) >= 2 and cat in (fcat, '液压油', '齿轮油', '空压机油', '涡轮机油', '汽轮机油'):
            for m in entries:
                if not m.get('runmeiProduct'):
                    continue
                cat_vis = set(extract_viscosities(cn_name + ' ' + catalog_entry.get('viscosity', '')))
                m_vis = set(extract_viscosities(m['competitorProduct']))
                if not cat_vis or not m_vis or (cat_vis & m_vis):
                    return m.get('runmeiProduct', ''), m
            # Fallback: return first mapping
            for m in entries:
                if m.get('runmeiProduct'):
                    return m.get('runmeiProduct', ''), m

    # Step 4: Brand + category based fallback
    bck = (cat_brand, cat)
    if bck in brand_cat_index:
        entries = brand_cat_index[bck]
        cat_vis = extract_viscosities(catalog_entry.get('viscosity', '') + ' ' + cn_name)
        best = None
        for m in entries:
            if not m.get('runmeiProduct'):
                continue
            m_vis = extract_viscosities(m.get('competitorProduct', '') + ' ' + m.get('competitorProductEN', ''))
            if cat_vis and m_vis and (set(cat_vis) & set(m_vis)):
                return m.get('runmeiProduct', ''), m
            if not best:
                best = m
        if best:
            return best.get('runmeiProduct', ''), best

    return '', None

# ============================================================
# STEP 5: Process all catalog entries
# ============================================================

merged = []
matched_count = 0
user_filled_count = 0

for entry in catalog:
    # Check if user already filled Runmei in Excel - that takes priority
    user_runmei = entry.get('runmei', '').strip()
    if user_runmei:
        runmei = user_runmei
        mapping = None
        user_filled_count += 1
        # Try to find mapping for specs
        _, mapping = find_runmei(entry)
    else:
        runmei, mapping = find_runmei(entry)
        if runmei:
            matched_count += 1

    merged.append({
        'brand': entry['brand'],
        'category': entry['category'],
        'product_cn': entry['product_cn'],
        'product_en': entry.get('product_en', ''),
        'viscosity': entry.get('viscosity', ''),
        'source': entry.get('source', ''),
        'runmei': runmei,
        'application': mapping.get('application', '') if mapping else '',
        'compSpecs': mapping.get('compSpecs', {}) if mapping else {},
        'runmeiSpecs': mapping.get('runmeiSpecs', {}) if mapping else {},
    })

total_runmei = matched_count + user_filled_count
print(f"User-filled: {user_filled_count}, Auto-matched: {matched_count}, Total Runmei: {total_runmei}/{len(merged)} ({100*total_runmei/len(merged):.1f}%)")

# Add mapped products not in catalog
existing_names = set(normalize(m['product_cn']) for m in merged)
for m in mappings:
    if not m.get('competitorProduct'):
        continue
    norm = normalize(m['competitorProduct'])
    if norm not in existing_names:
        existing_names.add(norm)
        merged.append({
            'brand': m['competitorBrand'],
            'category': m['category'],
            'product_cn': m['competitorProduct'],
            'product_en': m.get('competitorProductEN', ''),
            'viscosity': '',
            'source': m.get('sourceSheet', ''),
            'runmei': m.get('runmeiProduct', ''),
            'application': m.get('application', ''),
            'compSpecs': m.get('compSpecs', {}),
            'runmeiSpecs': m.get('runmeiSpecs', {}),
        })

# Deduplicate by product_cn, keeping entries with runmei
deduped = {}
for m in merged:
    key = normalize(m['product_cn'])
    if key not in deduped or (m['runmei'] and not deduped[key]['runmei']):
        deduped[key] = m
merged = list(deduped.values())
print(f"After dedup: {len(merged)} products")

# ============================================================
# STEP 6: Save product_data_merged.json
# ============================================================

OUTPUT_JSON = f'{BASE_DIR}/product_data_merged.json'
with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(merged, f, ensure_ascii=False, indent=2)
print(f"JSON saved: {OUTPUT_JSON}")

# ============================================================
# STEP 6b: Save data.json (standalone, with version, for Android app)
# ============================================================

from datetime import date
today_str = date.today().isoformat()
data_export = {
    'version': f'{today_str}-v1',
    'updatedAt': today_str,
    'totalProducts': len(merged),
    'products': merged,
}
OUTPUT_DATA_JSON = f'{BASE_DIR}/data.json'
with open(OUTPUT_DATA_JSON, 'w', encoding='utf-8') as f:
    json.dump(data_export, f, ensure_ascii=False, indent=2)
print(f"Data JSON saved: {OUTPUT_DATA_JSON} ({len(merged)} products)")

# ============================================================
# STEP 7: Generate HTML tool (without embedded data, async loading)
# ============================================================

products = merged
# Normalize brand display names (prefer short Chinese-only form)
BRAND_PREFER = {
    '壳牌/Shell': '壳牌', '壳牌': '壳牌',
    '美孚/Mobil': '美孚', '美孚': '美孚',
    '雪佛龙/Chevron': '雪佛龙', '雪佛龙': '雪佛龙',
    '长城/Sinopec': '长城', '长城': '长城',
    '克鲁勃/Klüber': '克鲁勃', '克鲁勃': '克鲁勃',
}
for p in products:
    if p['brand'] in BRAND_PREFER:
        p['brand'] = BRAND_PREFER[p['brand']]

brands = sorted(set(p['brand'] for p in products))
categories = sorted(set(p['category'] for p in products))

# We no longer embed data_json in HTML - it loads from remote/cache
brands_json = json.dumps(brands, ensure_ascii=False)
categories_json = json.dumps(categories, ensure_ascii=False)
cat_map_json = json.dumps(CAT_MAP, ensure_ascii=False)

# JavaScript code for asynchronous data loading
# Defined outside f-string to avoid {} escaping issues
load_data_js = '''
async function loadData() {
  if (DATA_LOADING) return;
  DATA_LOADING = true;
  var dataUrl = 'https://zk55806334-lang.github.io/Oil-Matching-Search/data.json';

  try {
    var resp = await fetch(dataUrl + '?_=' + Date.now());
    var json = await resp.json();
    if (json.products && json.products.length > 0) {
      DATA = json.products;
      DATA_VERSION = json.version || '';
      localStorage.setItem('oil_pm_data', JSON.stringify(json));
      DATA_LOADING = false;
      return;
    }
  } catch(e) {}

  try {
    var cached = localStorage.getItem('oil_pm_data');
    if (cached) {
      var json = JSON.parse(cached);
      if (json.products && json.products.length > 0) {
        DATA = json.products;
        DATA_VERSION = json.version || '';
        DATA_LOADING = false;
        return;
      }
    }
  } catch(e) {}

  try {
    if (window.AndroidBridge) {
      var fb = AndroidBridge.getFallbackData();
      if (fb) {
        var json = JSON.parse(fb);
        if (json.products && json.products.length > 0) {
          DATA = json.products;
          DATA_VERSION = json.version || '';
        }
      }
    }
  } catch(e) {}

  if (DATA.length === 0) {
    var el = document.getElementById('toast');
    if (el) { el.textContent = '\\u6570\\u636e\\u52a0\\u8f7d\\u5931\\u8d25\\uff0c\\u8bf7\\u68c0\\u67e5\\u7f51\\u7edc'; el.classList.add('show'); }
  }
  DATA_LOADING = false;
}

async function checkForUpdates() {
  if (!DATA_VERSION) return;
  try {
    var resp = await fetch('https://zk55806334-lang.github.io/Oil-Matching-Search/data.json?_=' + Date.now());
    var json = await resp.json();
    if (json.version && json.version !== DATA_VERSION && json.products) {
      DATA = json.products;
      DATA_VERSION = json.version;
      localStorage.setItem('oil_pm_data', JSON.stringify(json));
      var el = document.getElementById('toast');
      if (el) { el.textContent = '\\u6570\\u636e\\u5df2\\u66f4\\u65b0\\u81f3 ' + DATA_VERSION; el.classList.add('show');
        setTimeout(function() { el.classList.remove('show'); }, 3000); }
    }
  } catch(e) {}
}
'''

HTML = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>润美润滑油产品匹配工具</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif; background: #f0f2f5; color: #1a1a1a; min-height: 100vh; }}
.header {{ background: linear-gradient(135deg, #1a3c5e 0%, #2a5a8e 100%); color: #fff; padding: 14px 24px; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 8px rgba(0,0,0,.12); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }}
.header h1 {{ font-size: 18px; font-weight: 600; }}
.header-stats {{ font-size: 12px; opacity: .85; display: flex; gap: 16px; flex-wrap: wrap; }}
.header-stats span {{ white-space: nowrap; }}
.container {{ max-width: 920px; margin: 0 auto; padding: 16px 16px 40px; }}
.tabs {{ display: flex; gap: 0; margin-bottom: 16px; background: #fff; border-radius: 10px; padding: 4px; box-shadow: 0 1px 3px rgba(0,0,0,.06); }}
.tab {{ flex: 1; padding: 10px 8px; text-align: center; font-size: 13px; cursor: pointer; border: none; background: none; color: #666; border-radius: 8px; transition: all .2s; }}
.tab.active {{ background: #1a3c5e; color: #fff; font-weight: 500; }}
.tab:hover:not(.active) {{ background: #f0f2f5; }}
.panel {{ display: none; }}
.panel.active {{ display: block; }}
.card {{ background: #fff; border-radius: 10px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,.06); }}
.card-title {{ font-size: 14px; font-weight: 600; margin-bottom: 12px; color: #1a3c5e; display: flex; align-items: center; gap: 6px; }}
.form-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 10px; }}
@media (max-width: 600px) {{ .form-row {{ grid-template-columns: 1fr; }} }}
.form-group {{ display: flex; flex-direction: column; gap: 3px; }}
.form-group label {{ font-size: 12px; color: #888; font-weight: 500; }}
.form-group input, .form-group select, .form-group textarea {{
  font-size: 14px; border-radius: 8px; border: 1px solid #d0d5dd; padding: 9px 12px;
  background: #fff; color: #1a1a1a; width: 100%; font-family: inherit;
  transition: border-color .15s;
}}
.form-group input:focus, .form-group select:focus, .form-group textarea:focus {{
  border-color: #1a3c5e; outline: none; box-shadow: 0 0 0 3px rgba(26,60,94,.1);
}}
.form-group textarea {{ resize: vertical; min-height: 56px; }}
.hint {{ font-size: 11px; color: #aaa; margin-top: 2px; }}
.btn-row {{ display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }}
.search-btn {{
  padding: 10px 28px; background: #1a3c5e; color: #fff;
  border: none; border-radius: 8px; font-size: 14px; font-weight: 500; cursor: pointer;
  display: inline-flex; align-items: center; gap: 6px; transition: all .15s;
}}
.search-btn:hover {{ background: #2a5a8e; }}
.btn-outline {{ padding: 8px 14px; background: #fff; color: #1a3c5e; border: 1px solid #1a3c5e; border-radius: 8px; font-size: 13px; cursor: pointer; transition: all .15s; }}
.btn-outline:hover {{ background: #f0f5fa; }}
.btn-clear {{ padding: 8px 14px; background: #fff; color: #888; border: 1px solid #d0d5dd; border-radius: 8px; font-size: 13px; cursor: pointer; }}
.btn-clear:hover {{ background: #f5f5f5; }}
.preset-bar {{ display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; }}
.preset-chip {{ font-size: 12px; padding: 4px 10px; border-radius: 14px; background: #f0f2f5; color: #555; border: none; cursor: pointer; transition: all .15s; white-space: nowrap; }}
.preset-chip:hover {{ background: #d4e6f1; color: #1a3c5e; }}
.result-summary {{ display: flex; gap: 12px; flex-wrap: wrap; margin: 12px 0 8px; padding: 10px 14px; background: #fff; border-radius: 8px; border: 1px solid #e8e8e8; font-size: 13px; color: #888; align-items: center; }}
.result-summary strong {{ color: #1a3c5e; }}
.summary-divider {{ width: 1px; height: 16px; background: #ddd; }}
.result-card {{
  background: #fff; border-radius: 10px; padding: 14px; margin-bottom: 8px;
  box-shadow: 0 1px 3px rgba(0,0,0,.06); cursor: pointer; transition: all .15s;
  border: 1px solid #e8e8e8; position: relative;
}}
.result-card:hover {{ border-color: #1a3c5e; box-shadow: 0 2px 8px rgba(26,60,94,.1); }}
.result-card.best {{ border-left: 4px solid #27ae60; }}
.result-card.good {{ border-left: 4px solid #2980b9; }}
.result-card.fair {{ border-left: 4px solid #f39c12; }}
.result-card.none {{ border-left: 4px solid #bbb; }}
.card-top {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px; gap: 8px; }}
.card-name {{ font-size: 14px; font-weight: 600; word-break: break-all; }}
.card-meta {{ font-size: 12px; color: #888; margin-top: 2px; }}
.score-badge {{ font-size: 11px; padding: 3px 10px; border-radius: 12px; white-space: nowrap; font-weight: 500; flex-shrink: 0; }}
.badge-high {{ background: #d5f5e3; color: #1a7e40; }}
.badge-mid {{ background: #d4e6f1; color: #1a5276; }}
.badge-low {{ background: #fdebd0; color: #935116; }}
.badge-info {{ background: #f0f0f0; color: #888; }}
.runmei-box {{ margin-top: 8px; padding: 10px 14px; background: #f0fdf4; border-radius: 8px; border: 1px solid #bbf7d0; display: flex; justify-content: space-between; align-items: center; gap: 8px; }}
.runmei-info {{ flex: 1; }}
.runmei-label {{ font-size: 11px; color: #15803d; font-weight: 600; }}
.runmei-name {{ font-size: 15px; font-weight: 600; color: #166534; margin-top: 2px; }}
.runmei-app {{ font-size: 12px; color: #666; margin-top: 2px; }}
.copy-btn {{ font-size: 11px; padding: 4px 10px; background: #fff; border: 1px solid #bbf7d0; border-radius: 6px; color: #15803d; cursor: pointer; white-space: nowrap; flex-shrink: 0; }}
.copy-btn:hover {{ background: #d5f5e3; }}
.copy-btn.copied {{ background: #bbf7d0; }}
.no-runmei {{ margin-top: 6px; padding: 6px 10px; background: #fffbeb; border-radius: 6px; border: 1px solid #fde68a; font-size: 12px; color: #92400e; }}
.tags {{ display: flex; gap: 4px; flex-wrap: wrap; margin-top: 4px; }}
.tag {{ font-size: 11px; padding: 2px 8px; border-radius: 4px; background: #f0f2f5; color: #666; }}
.tag.green {{ background: #d5f5e3; color: #1a7e40; }}
.empty-state {{ text-align: center; padding: 48px 20px; color: #aaa; font-size: 14px; }}
.empty-state .example {{ font-size: 12px; color: #ccc; margin-top: 8px; line-height: 1.8; }}
.result-expand {{ display: none; margin-top: 10px; padding-top: 10px; border-top: 1px solid #eee; font-size: 13px; color: #666; line-height: 1.8; }}
.result-card.expanded .result-expand {{ display: block; }}
.free-search {{ position: relative; }}
.free-search input {{ width: 100%; padding: 12px 16px; border: 1px solid #d0d5dd; border-radius: 10px; font-size: 15px; font-family: inherit; }}
.free-search input:focus {{ border-color: #1a3c5e; outline: none; box-shadow: 0 0 0 3px rgba(26,60,94,.1); }}
.stats-bar {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 10px; }}
.stat-item {{ font-size: 12px; color: #888; background: #fff; padding: 6px 12px; border-radius: 20px; border: 1px solid #e8e8e8; }}
.stat-item strong {{ color: #1a3c5e; }}
.quick-results {{ max-height: 520px; overflow-y: auto; }}
.browse-filter {{ margin-bottom: 10px; }}
.browse-filter input {{ width: 100%; padding: 8px 12px; border: 1px solid #d0d5dd; border-radius: 8px; font-size: 13px; }}
.toast {{ position: fixed; top: 80px; left: 50%; transform: translateX(-50%); background: #1a3c5e; color: #fff; padding: 8px 20px; border-radius: 20px; font-size: 13px; z-index: 999; opacity: 0; transition: opacity .3s; pointer-events: none; }}
.toast.show {{ opacity: 1; }}
</style>
</head>
<body>

<div class="header">
  <div><h1>润美润滑油产品匹配工具</h1></div>
  <div class="header-stats">
    <span>竞品: {len(products)} 款</span>
    <span>润美: {sum(1 for p in products if p["runmei"])} 款</span>
    <span>品类: {len(categories)}</span>
    <span>品牌: {len(brands)}</span>
  </div>
</div>

<div class="toast" id="toast"></div>

<div class="container">
  <div class="tabs">
    <button class="tab active" onclick="switchTab('competitor')">竞品匹配</button>
    <button class="tab" onclick="switchTab('browse')">分类浏览</button>
    <button class="tab" onclick="switchTab('quick')">快速搜索</button>
  </div>

  <!-- ============ PANEL 1: COMPETITOR MATCHING ============ -->
  <div class="panel active" id="panel-competitor">
    <div class="card">
      <div class="card-title">输入竞品信息，自动匹配润美对应产品</div>
      <div class="preset-bar">
        <span style="font-size:11px;color:#aaa;padding:4px 0;">快捷:</span>
        <button class="preset-chip" onclick="quickPreset('液压油','壳牌','Tellus S2 M 46','VG 46')">壳牌 Tellus 液压油</button>
        <button class="preset-chip" onclick="quickPreset('液压油','美孚','DTE 10 Excel 32','VG 32')">美孚 DTE 液压油</button>
        <button class="preset-chip" onclick="quickPreset('齿轮油','壳牌','Omala S2 G 320','VG 320')">壳牌 Omala 齿轮油</button>
        <button class="preset-chip" onclick="quickPreset('齿轮油','美孚','SHC 600','VG 320')">美孚 SHC 齿轮油</button>
        <button class="preset-chip" onclick="quickPreset('空压机油','壳牌','Corena S3 R 46','VG 46')">壳牌 Corena 空压机油</button>
        <button class="preset-chip" onclick="quickPreset('涡轮机油','壳牌','Turbo T 46','VG 46')">壳牌 Turbo 涡轮机油</button>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>产品类型</label>
          <select id="cp-category" onchange="onCategoryChange()">
            <option value="">全部类型（选填可提高精度）</option>
            {''.join(f'<option value="{c}">{c}</option>' for c in categories)}
          </select>
        </div>
        <div class="form-group">
          <label>竞品品牌</label>
          <select id="cp-brand">
            <option value="">全部品牌（选填可提高精度）</option>
            {''.join(f'<option value="{b}">{b}</option>' for b in brands)}
          </select>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>竞品型号 / 名称</label>
          <input type="text" id="cp-model" placeholder="如: Tellus S2 M 46、得力士 S4 VX、DTE 10 Excel、Omala..." autocomplete="off" list="model-suggestions" />
          <datalist id="model-suggestions"></datalist>
          <span class="hint">可输入完整型号或部分关键词，中英文均可。输入后点下拉建议快速选择</span>
        </div>
        <div class="form-group">
          <label>ISO 粘度等级</label>
          <select id="cp-viscosity">
            <option value="">不限粘度</option>
            <option>VG 2</option><option>VG 5</option><option>VG 7</option><option>VG 10</option>
            <option>VG 15</option><option>VG 22</option><option>VG 32</option>
            <option>VG 46</option><option>VG 68</option><option>VG 100</option>
            <option>VG 150</option><option>VG 220</option><option>VG 320</option>
            <option>VG 460</option><option>VG 680</option><option>VG 1000</option><option>VG 1500</option>
          </select>
        </div>
      </div>
      <div class="form-group" style="margin-bottom:10px;">
        <label>技术等级 / 性能要求（选填）</label>
        <textarea id="cp-specs" placeholder="如: 抗磨 HM、高粘指 HV、极压 CKC、全合成 PAO、GTL、食品级、可生物降解、难燃、低温、风电..."></textarea>
        <span class="hint">输入等级代号或特性关键词，系统自动提取匹配。Ctrl+Enter 快速搜索</span>
      </div>
      <div class="btn-row">
        <button class="search-btn" onclick="searchCompetitor()">匹配润美产品</button>
        <button class="btn-clear" onclick="clearForm()">清空重填</button>
      </div>
    </div>
    <div id="cp-results"></div>
  </div>

  <!-- ============ PANEL 2: BROWSE ============ -->
  <div class="panel" id="panel-browse">
    <div class="card">
      <div class="card-title">按产品类型浏览竞品对照表</div>
      <div class="form-row">
        <div class="form-group">
          <label>产品类型</label>
          <select id="browse-cat" onchange="browseCategory()">
            <option value="">请选择类型</option>
            {''.join(f'<option value="{c}">{c}</option>' for c in categories)}
          </select>
        </div>
        <div class="form-group">
          <label>品牌筛选</label>
          <select id="browse-brand" onchange="browseCategory()">
            <option value="">全部品牌</option>
            {''.join(f'<option value="{b}">{b}</option>' for b in brands)}
          </select>
        </div>
      </div>
      <div class="browse-filter">
        <input type="text" id="browse-filter-text" placeholder="在当前结果中搜索..." oninput="browseCategory()" />
      </div>
    </div>
    <div id="browse-results"></div>
  </div>

  <!-- ============ PANEL 3: QUICK SEARCH ============ -->
  <div class="panel" id="panel-quick">
    <div class="card">
      <div class="card-title">全局快速搜索（搜索竞品名、品牌、润美产品、粘度、等级...）</div>
      <div class="free-search">
        <input type="text" id="quick-search" placeholder="试试输入: Tellus、得力士、PAO、风电、HM、CKC..." autocomplete="off" oninput="quickSearch()" />
      </div>
      <div class="stats-bar" style="margin-top:10px;" id="quick-stats"></div>
    </div>
    <div class="quick-results" id="quick-results"></div>
  </div>
</div>

<script>
var DATA = [];
var DATA_VERSION = '';
var DATA_LOADING = false;
var BRANDS = {brands_json};
var CATEGORIES = {categories_json};
var CAT_MAP = {cat_map_json};

{load_data_js}
// ============== MATCHING ENGINE ==============

function tokenize(s) {{
  if (!s) return [];
  var raw = s.toLowerCase().split(/[\\s/\\-\\.\\(\\)\\+]+/).filter(Boolean);
  var out = [];
  raw.forEach(function(t) {{
    var parts = t.split(/(?<=[a-z])(?=[A-Z0-9])|(?<=[A-Z])(?=[A-Z][a-z])|(?<=[0-9])(?=[A-Za-z])|(?<=[A-Za-z])(?=[0-9])/);
    parts.forEach(function(p) {{ out.push(p); }});
  }});
  return out.filter(function(t) {{ return t.length >= 1; }});
}}

function extractGrade(text) {{
  var grades = [];
  var patterns = [
    /\b(HM|HV|HLP?|HVLPD|HS)\b/gi,
    /\b(CKC|CKD|CKE|CKT|CKP)\b/gi,
    /\b(TSA|TSE|TGB|TGSB)\b/gi,
    /\b(VDL|DAC|DAH|DAG)\b/gi,
    /\b(FD|FC)\b/gi,
    /\b(S1|S2|S3|S4|S5)\b/gi,
    /\b(PAO|GTL|PAG|SHC|EP|MP|XMP|OG)\b/gi,
  ];
  patterns.forEach(function(pat) {{
    var m = text.match(pat);
    if (m) m.forEach(function(g) {{ grades.push(g.toUpperCase()); }});
  }});
  return grades;
}}

function extractVG(text) {{
  var m = text.match(/VG\\s*(\\d{{2,4}})/i);
  if (m) return parseInt(m[1]);
  m = text.match(/\\b(\\d{{2,3}})\\s*(?:号|#|cSt)?\\b/);
  if (m) {{ var v = parseInt(m[1]); if (v >= 2 && v <= 1500) return v; }}
  return null;
}}

function normalizeBrand(b) {{
  var map = {{ '壳牌/Shell':'壳牌','壳牌':'壳牌','美孚/Mobil':'美孚','美孚':'美孚','雪佛龙/Chevron':'雪佛龙','雪佛龙':'雪佛龙','长城/Sinopec':'长城','长城':'长城','克鲁勃/Klüber':'克鲁勃','克鲁勃':'克鲁勃','嘉实多':'嘉实多','福斯':'福斯','安索':'安索','昆仑':'昆仑' }};
  return map[b] || b;
}}

function getProductTokens(p) {{
  return tokenize(p.product_cn + ' ' + p.product_en + ' ' + p.category + ' ' + (p.application || ''));
}}

function getProductGrade(p) {{
  return extractGrade(p.product_cn + ' ' + p.product_en + ' ' + (p.application || ''));
}}

function getProductVG(p) {{
  return extractVG(p.viscosity || '') || extractVG(p.product_cn);
}}

function tokenMatch(inputTokens, productTokens) {{
  if (!inputTokens || inputTokens.length === 0) return {{ count: 0, ratio: 0 }};
  var matched = 0;
  inputTokens.forEach(function(t) {{
    var tl = t.toLowerCase();
    for (var i = 0; i < productTokens.length; i++) {{
      var pl = productTokens[i].toLowerCase();
      // Numeric tokens: exact match only (avoid "1000" matching "100")
      if (/^\\d+$/.test(tl) && /^\\d+$/.test(pl)) {{
        if (tl === pl) {{ matched++; break; }}
      }}
      // Non-numeric: substring match
      else if (pl.includes(tl) || tl.includes(pl)) {{ matched++; break; }}
    }}
  }});
  return {{ count: matched, ratio: inputTokens.length > 0 ? matched / inputTokens.length : 0 }};
}}

function matchScore(input, product) {{
  var score = 0, reasons = [];

  if (input.category && product.category) {{
    var icat = CAT_MAP[input.category] || input.category;
    var pcat = CAT_MAP[product.category] || product.category;
    if (icat === pcat) {{ score += 15; reasons.push('品类匹配'); }}
    else if (icat && pcat && (icat.includes(pcat) || pcat.includes(icat))) {{ score += 8; reasons.push('品类相关'); }}
  }}

  if (input.brand && product.brand) {{
    var ib = normalizeBrand(input.brand), pb = normalizeBrand(product.brand);
    if (ib && pb && ib === pb) {{ score += 15; reasons.push('品牌匹配'); }}
    else {{ score -= 15; reasons.push('品牌不匹配'); }}
  }}

  if (input.modelTokens && input.modelTokens.length > 0) {{
    var pTokens = getProductTokens(product);
    var mr = tokenMatch(input.modelTokens, pTokens);
    if (mr.ratio >= 0.8) {{ score += 20; reasons.push('型号高度匹配(' + mr.count + '/' + input.modelTokens.length + '词)'); }}
    else if (mr.ratio >= 0.5) {{ score += 12; reasons.push('型号部分匹配(' + mr.count + '/' + input.modelTokens.length + '词)'); }}
    else if (mr.ratio >= 0.25) {{ score += 6; reasons.push('型号弱匹配(' + mr.count + '/' + input.modelTokens.length + '词)'); }}
    else if (mr.count > 0) {{ score += 3; }}
  }}

  if (input.viscosity) {{
    var pvg = getProductVG(product);
    if (pvg && input.viscosity === pvg) {{ score += 10; reasons.push('粘度精确 VG' + pvg); }}
    else if (pvg) {{
      var vgGrades = [2,5,7,10,15,22,32,46,68,100,150,220,320,460,680,1000,1500];
      var ivgIdx = vgGrades.indexOf(input.viscosity), pvgIdx = vgGrades.indexOf(pvg);
      if (ivgIdx >= 0 && pvgIdx >= 0 && Math.abs(ivgIdx - pvgIdx) <= 1) {{ score += 5; reasons.push('粘度接近 VG' + pvg); }}
    }}
  }}

  if (input.grades && input.grades.length > 0) {{
    var pGrades = getProductGrade(product);
    input.grades.forEach(function(g) {{ if (pGrades.indexOf(g.toUpperCase()) >= 0) {{ score += 8; reasons.push('等级' + g.toUpperCase()); }} }});
  }}

  if (input.specKeywords && input.specKeywords.length > 0) {{
    var pText = (product.product_cn + ' ' + product.product_en + ' ' + (product.application||'') + ' ' + (product.runmei||'')).toLowerCase(), kwMatched = 0;
    input.specKeywords.forEach(function(kw) {{ if (pText.includes(kw.toLowerCase())) {{ kwMatched++; score += 2; }} }});
    if (kwMatched >= 2) reasons.push('特性' + kwMatched + '项');
  }}

  if (product.runmei) {{
    score += 5;
    if (!input.model && !input.viscosity && (!input.grades || input.grades.length === 0)) {{
      score += 5; reasons.push('推荐润美');
    }} else {{
      reasons.push('已匹配润美');
    }}
  }}

  var cn = product.product_cn || '';
  if (cn.length > 50) score -= 3;
  if (/[，。；：]/.test(cn) && cn.length > 20) score -= 3;

  return {{ score: score, reasons: reasons }};
}}

function getMatchLevel(score) {{
  if (score >= 30) return {{ cls:'best', label:'强烈推荐', badgeCls:'badge-high' }};
  if (score >= 18) return {{ cls:'good', label:'推荐匹配', badgeCls:'badge-mid' }};
  if (score >= 10) return {{ cls:'fair', label:'可能匹配', badgeCls:'badge-low' }};
  return {{ cls:'none', label:'参考', badgeCls:'badge-info' }};
}}

// ============== COMPETITOR SEARCH ==============

function parseInput() {{
  var category = document.getElementById('cp-category').value;
  var brand = document.getElementById('cp-brand').value;
  var model = document.getElementById('cp-model').value.trim().replace(/[™®]/g, '');
  var viscosityVal = document.getElementById('cp-viscosity').value;
  var specs = document.getElementById('cp-specs').value.trim();
  var viscosity = null;
  if (viscosityVal) {{ var vm = viscosityVal.match(/(\\d+)/); if (vm) viscosity = parseInt(vm[1]); }}
  var grades = extractGrade(model + ' ' + specs);
  var specKeywords = tokenize(specs);
  var noiseWords = ['的','型','油','级','等','及','与','或','和','且'];
  specKeywords = specKeywords.filter(function(k) {{ return noiseWords.indexOf(k) < 0 && k.length >= 2; }});
  // Auto-detect brand from model text when dropdown not used
  var detectedBrand = brand;
  if (!detectedBrand && model) {{
    var ml = model.toLowerCase();
    var brandMap = {{'壳牌':'壳牌','shell':'壳牌','美孚':'美孚','mobil':'美孚','雪佛龙':'雪佛龙','chevron':'雪佛龙','长城':'长城','sinopec':'长城','克鲁勃':'克鲁勃','kluber':'克鲁勃','klüber':'克鲁勃','嘉实多':'嘉实多','castrol':'嘉实多','福斯':'福斯','fuchs':'福斯'}};
    Object.keys(brandMap).forEach(function(k) {{ if (ml.indexOf(k) >= 0) detectedBrand = brandMap[k]; }});
  }}
  return {{ category:category, brand:brand, detectedBrand:detectedBrand, model:model, modelTokens:tokenize(model), viscosity:viscosity, grades:grades, specKeywords:specKeywords }};
}}

function searchCompetitor() {{
  var rawModel = document.getElementById('cp-model').value.trim();
  var rawSpecs = document.getElementById('cp-specs').value.trim();
  var query = rawModel || rawSpecs;
  var res = document.getElementById('cp-results');
  if (!query) {{
    res.innerHTML = '<div class="empty-state">请输入竞品名称或关键字后查询<div class="example">可在型号框或性能要求框中输入</div></div>';
    return;
  }}

  var q = query.toLowerCase();
  var category = document.getElementById('cp-category').value;
  var brand = document.getElementById('cp-brand').value;

  // Step 1: Precise name lookup (match product_cn in data)
  var matched = [];
  DATA.forEach(function(p) {{
    if (brand && normalizeBrand(p.brand) !== normalizeBrand(brand)) return;
    if (category) {{
      var icat = CAT_MAP[category] || category;
      var pcat = CAT_MAP[p.category] || p.category;
      if (icat !== pcat) return;
    }}
    var cn = p.product_cn.toLowerCase().trim();
    if (cn === q) matched.push({{ product: p, exact: true }});
    else if (cn.includes(q) || q.includes(cn)) matched.push({{ product: p, exact: false }});
  }});

  // Exact match first, then shorter names
  matched.sort(function(a, b) {{
    if (a.exact !== b.exact) return a.exact ? -1 : 1;
    return a.product.product_cn.length - b.product.product_cn.length;
  }});

  // Find first name match that has Runmei → show 1 result
  for (var i = 0; i < matched.length; i++) {{
    if (matched[i].product.runmei) {{
      var best = matched[i].product;
      var html = '<div class="result-summary">精准匹配: ' + esc(query) + '</div>';
      html += '<div class="result-card best" onclick="toggleCard(this)"><div class="card-top"><div>';
      html += '<div class="card-name">' + esc(best.product_cn) + '</div>';
      html += '<div class="card-meta">' + esc(best.brand) + ' · ' + esc(best.category);
      if (best.viscosity) html += ' · ' + esc(best.viscosity);
      html += '</div></div></div>';
      html += '<div class="runmei-box"><div class="runmei-info"><div class="runmei-label">润美对应产品</div>';
      html += '<div class="runmei-name">' + esc(best.runmei) + '</div>';
      if (best.application) html += '<div class="runmei-app">' + esc(best.application) + '</div></div>';
      html += '<button class="copy-btn" onclick="event.stopPropagation();copyRunmei(this)" data-runmei="' + escAttr(best.runmei) + '">复制</button></div>';
      html += '<div class="result-expand">';
      html += '<strong>英文名:</strong> ' + esc(best.product_en || '(无)') + '<br>';
      html += '<strong>数据来源:</strong> ' + esc(best.source || '') + '</div></div>';
      res.innerHTML = html;
      return;
    }}
  }}

  // Found name match but no Runmei
  if (matched.length > 0) {{
    res.innerHTML = '<div class="empty-state">已找到产品，但暂无润美映射<div class="example">请在Excel F列补充对应润美产品</div></div>';
    return;
  }}

  // Step 2: Fallback to fuzzy scoring (old algorithm) for broad search
  var input = parseInput();
  // If user typed in specs field only, use it as model text for token matching
  if (!rawModel && rawSpecs) {{
    input.model = rawSpecs;
    input.modelTokens = tokenize(rawSpecs);
    input.grades = extractGrade(rawSpecs);
  }}
  var scored = DATA.map(function(p) {{ var r = matchScore(input, p); return {{ product:p, score:r.score, reasons:r.reasons }}; }});
  if (input.model || input.brand || input.grades.length > 0) scored = scored.filter(function(s) {{ return s.score > 0; }});
  var activeBrand = input.brand || input.detectedBrand;
  if (activeBrand) {{
    var ab = normalizeBrand(activeBrand);
    scored = scored.filter(function(s) {{ return normalizeBrand(s.product.brand) === ab; }});
  }}
  scored.sort(function(a, b) {{ return b.score - a.score; }});
  scored = scored.filter(function(s) {{ return s.score >= 8; }});
  if (scored.length === 0) scored = scored.slice(0, 3);
  else scored = scored.slice(0, 5);

  if (scored.length === 0) {{
    res.innerHTML = '<div class="empty-state">未找到匹配产品<div class="example">尝试简化关键字<br>或在型号/性能框中输入不同关键词</div></div>';
    return;
  }}

  var hasRunmei = scored.filter(function(s) {{ return s.product.runmei; }}).length;
  var desc = [input.brand, query, input.viscosity ? 'VG ' + input.viscosity : '', input.category].filter(Boolean).join(' · ');
  var html = '<div class="result-summary"><strong>' + scored.length + '</strong> 条模糊匹配 · 润美 <strong style="color:#27ae60;">' + hasRunmei + '</strong> 条 · ' + esc(desc) + '</div>';
  scored.forEach(function(s) {{
    var p = s.product, level = getMatchLevel(s.score);
    html += '<div class="result-card ' + level.cls + '" onclick="toggleCard(this)"><div class="card-top"><div>';
    html += '<div class="card-name">' + esc(p.product_cn) + '</div>';
    html += '<div class="card-meta">' + esc(p.brand) + ' · ' + esc(p.category);
    if (p.viscosity) html += ' · ' + esc(p.viscosity);
    html += '</div></div>';
    html += '<span class="score-badge ' + level.badgeCls + '">' + s.score + '分</span></div>';
    if (p.runmei) {{
      html += '<div class="runmei-box"><div class="runmei-info"><div class="runmei-label">润美对应产品</div>';
      html += '<div class="runmei-name">' + esc(p.runmei) + '</div></div>';
      html += '<button class="copy-btn" onclick="event.stopPropagation();copyRunmei(this)" data-runmei="' + escAttr(p.runmei) + '">复制</button></div>';
    }} else {{
      html += '<div class="no-runmei">暂无润美映射</div>';
    }}
    html += '</div>';
  }});
  res.innerHTML = html;
}}

function toggleCard(el) {{ el.classList.toggle('expanded'); }}

function copyRunmei(btn) {{
  var text = btn.getAttribute('data-runmei');
  navigator.clipboard.writeText(text).then(function() {{
    btn.textContent = '已复制!'; btn.classList.add('copied');
    setTimeout(function() {{ btn.textContent = '复制'; btn.classList.remove('copied'); }}, 1500);
  }}).catch(function() {{
    showToast('复制失败，请手动选择');
  }});
}}

function showToast(msg) {{
  var t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  setTimeout(function() {{ t.classList.remove('show'); }}, 2000);
}}

function quickPreset(cat, brand, model, vis) {{
  document.getElementById('cp-category').value = cat;
  document.getElementById('cp-brand').value = brand;
  document.getElementById('cp-model').value = model;
  document.getElementById('cp-viscosity').value = vis;
  onCategoryChange();
  searchCompetitor();
}}

function clearForm() {{
  document.getElementById('cp-category').value = '';
  document.getElementById('cp-brand').value = '';
  document.getElementById('cp-model').value = '';
  document.getElementById('cp-viscosity').value = '';
  document.getElementById('cp-specs').value = '';
  document.getElementById('cp-results').innerHTML = '';
  onCategoryChange();
}}

function onCategoryChange() {{
  var cat = document.getElementById('cp-category').value;
  var brandSel = document.getElementById('cp-brand'), currentBrand = brandSel.value;
  brandSel.innerHTML = '<option value="">全部品牌（选填可提高精度）</option>';
  (cat ? [] : BRANDS).concat(cat ? (function(){{ var b=[]; DATA.forEach(function(p){{ if(p.category===cat && b.indexOf(p.brand)<0) b.push(p.brand); }}); return b; }})() : []).forEach(function(b) {{
    brandSel.innerHTML += '<option value="'+esc(b)+'"'+(b===currentBrand?' selected':'')+'>'+esc(b)+'</option>';
  }});

  // Update model suggestions datalist
  var dl = document.getElementById('model-suggestions');
  dl.innerHTML = '';
  var suggestions = cat ? DATA.filter(function(p){{ return p.category===cat; }}).slice(0,50) : DATA.slice(0,50);
  suggestions.forEach(function(p) {{
    dl.innerHTML += '<option value="'+esc(p.product_cn)+'">';
  }});
}}

// ============== BROWSE MODE ==============

function browseCategory() {{
  var cat = document.getElementById('browse-cat').value;
  var brand = document.getElementById('browse-brand').value;
  var filterText = (document.getElementById('browse-filter-text').value || '').trim().toLowerCase();
  var container = document.getElementById('browse-results');
  if (!cat) {{ container.innerHTML = '<div class="empty-state">请选择产品类型开始浏览<div class="example">可同时筛选品牌和输入关键词过滤</div></div>'; return; }}

  var filtered = DATA.filter(function(p) {{
    if (p.category !== cat) return false;
    if (brand && p.brand !== brand) return false;
    if (filterText) {{
      var txt = (p.product_cn + ' ' + p.product_en + ' ' + p.brand + ' ' + p.runmei).toLowerCase();
      if (txt.indexOf(filterText) < 0) return false;
    }}
    return true;
  }});

  var groups = {{}};
  filtered.forEach(function(p) {{
    if (!groups[p.brand]) groups[p.brand] = [];
    groups[p.brand].push(p);
  }});

  var html = '<div class="result-summary">共 <strong>' + filtered.length + '</strong> 个产品' + (brand ? ' · ' + esc(brand) : '') + (filterText ? ' · 过滤: ' + esc(filterText) : '') + '</div>';

  Object.keys(groups).sort().forEach(function(brandName) {{
    var items = groups[brandName], hasRunmei = items.filter(function(p){{return p.runmei;}}).length;
    html += '<div class="card"><div class="card-title">' + esc(brandName) + ' <span style="font-weight:400;font-size:12px;color:#888;">' + items.length + '条, ' + hasRunmei + '条已匹配润美</span></div>';
    items.forEach(function(p) {{
      html += '<div class="result-card" style="margin-bottom:6px;cursor:default;' + (p.runmei?'':'opacity:.75;') + '">';
      html += '<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">';
      html += '<div><div style="font-weight:600;">' + esc(p.product_cn) + '</div>';
      if (p.product_en) html += '<div style="font-size:12px;color:#888;">' + esc(p.product_en) + '</div>';
      html += '<div style="font-size:11px;color:#aaa;">' + esc(p.viscosity || '粘度未标注') + ' · ' + esc(p.source || '') + '</div></div>';
      if (p.runmei) {{
        html += '<div style="text-align:right;flex-shrink:0;">';
        html += '<div style="font-size:11px;color:#27ae60;font-weight:600;">润美对应</div>';
        html += '<div style="font-size:13px;font-weight:600;color:#166534;">' + esc(p.runmei) + '</div></div>';
      }}
      html += '</div></div>';
    }});
    html += '</div>';
  }});
  container.innerHTML = html;
}}

// ============== QUICK SEARCH ==============
var quickTimeout = null;
function quickSearch() {{ clearTimeout(quickTimeout); quickTimeout = setTimeout(doQuickSearch, 200); }}

function doQuickSearch() {{
  var q = document.getElementById('quick-search').value.trim();
  var container = document.getElementById('quick-results'), stats = document.getElementById('quick-stats');
  if (!q || q.length < 2) {{ container.innerHTML = '<div class="empty-state">输入至少2个字符开始搜索<div class="example">试试: Tellus, 得力士, PAO, 风电, CKC, 46, 壳牌液压油...</div></div>'; stats.innerHTML=''; return; }}

  var tokens = tokenize(q);
  var scored = DATA.map(function(p) {{
    var score = 0;
    var pText = (p.product_cn + ' ' + p.product_en + ' ' + p.brand + ' ' + p.category + ' ' + p.viscosity + ' ' + p.runmei + ' ' + (p.application||'')).toLowerCase();
    tokens.forEach(function(t) {{ if (pText.includes(t)) score += 5; try {{ if (new RegExp('\\\\b'+t.replace(/[.*+?^${{}}()|[\\]\\\\]/g,'\\\\$&'),'i').test(pText)) score += 3; }} catch(e){{}} }});
    if (p.runmei) score += 1;
    return {{ product:p, score:score }};
  }});
  scored = scored.filter(function(s){{return s.score>0;}}).sort(function(a,b){{return b.score-a.score;}}).slice(0,30);

  var hasRunmei = scored.filter(function(s){{return s.product.runmei;}}).length;
  stats.innerHTML = '<span class="stat-item">匹配 <strong>'+scored.length+'</strong> 条</span><span class="stat-item">润美 <strong>'+hasRunmei+'</strong> 条</span>';

  if (scored.length === 0) {{ container.innerHTML = '<div class="empty-state">未找到匹配</div>'; return; }}

  var html = '';
  scored.forEach(function(s) {{
    var p = s.product, level = getMatchLevel(s.score);
    html += '<div class="result-card ' + level.cls + '" onclick="toggleCard(this)">';
    html += '<div class="card-top"><div><div class="card-name">' + esc(p.product_cn) + '</div>';
    html += '<div class="card-meta">' + esc(p.brand) + ' · ' + esc(p.category);
    if (p.viscosity) html += ' · ' + esc(p.viscosity);
    html += '</div></div><span class="score-badge ' + level.badgeCls + '">' + s.score + '分</span></div>';
    if (p.runmei) {{
      html += '<div class="runmei-box"><div class="runmei-info"><div class="runmei-label">润美对应产品</div><div class="runmei-name">' + esc(p.runmei) + '</div></div>';
      html += '<button class="copy-btn" onclick="event.stopPropagation();copyRunmei(this)" data-runmei="' + escAttr(p.runmei) + '">复制</button></div>';
    }}
    html += '<div class="result-expand"><strong>英文名:</strong> ' + esc(p.product_en || '(无)') + '<br><strong>来源:</strong> ' + esc(p.source || '') + '</div></div>';
  }});
  container.innerHTML = html;
}}

// ============== UTILS ==============
function esc(s) {{ if (!s) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }}
function escAttr(s) {{ if (!s) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }}

function switchTab(tab) {{
  document.querySelectorAll('.tab').forEach(function(t,i){{ var tabs=['competitor','browse','quick']; t.classList.toggle('active',tabs[i]===tab); }});
  document.getElementById('panel-competitor').classList.toggle('active', tab==='competitor');
  document.getElementById('panel-browse').classList.toggle('active', tab==='browse');
  document.getElementById('panel-quick').classList.toggle('active', tab==='quick');
}}

document.addEventListener('DOMContentLoaded', function() {{
  // Load data then check for updates
  loadData().then(function() {{ checkForUpdates(); }});

  document.getElementById('cp-model').addEventListener('keydown', function(e) {{ if (e.key==='Enter') searchCompetitor(); }});
  document.getElementById('cp-specs').addEventListener('keydown', function(e) {{ if (e.ctrlKey && e.key==='Enter') searchCompetitor(); }});
  try {{
    var state = JSON.parse(localStorage.getItem('oil_pm_last'));
    if (state) {{
      if (state.category) document.getElementById('cp-category').value = state.category;
      if (state.brand) document.getElementById('cp-brand').value = state.brand;
      if (state.model) document.getElementById('cp-model').value = state.model;
      if (state.viscosity) document.getElementById('cp-viscosity').value = state.viscosity;
      if (state.specs) document.getElementById('cp-specs').value = state.specs;
    }}
  }} catch(e) {{}}
  ['cp-category','cp-brand','cp-model','cp-viscosity','cp-specs'].forEach(function(id) {{
    var el = document.getElementById(id);
    if (el) {{ el.addEventListener('change', saveLast); if (el.tagName==='INPUT'||el.tagName==='TEXTAREA') el.addEventListener('blur', saveLast); }}
  }});
}});

function saveLast() {{
  try {{ localStorage.setItem('oil_pm_last', JSON.stringify({{ category:document.getElementById('cp-category').value, brand:document.getElementById('cp-brand').value, model:document.getElementById('cp-model').value, viscosity:document.getElementById('cp-viscosity').value, specs:document.getElementById('cp-specs').value }})); }} catch(e) {{}}
}}
</script>
</body>
</html>'''

OUTPUT_HTML = f'{BASE_DIR}/lubricant_product_matching_tool.html'
with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"\nHTML tool saved: {OUTPUT_HTML}")
print(f"Size: {len(HTML.encode('utf-8')) / 1024:.1f} KB")
print(f"Products: {len(products)}")
print(f"With Runmei: {sum(1 for p in products if p['runmei'])}")
print(f"Brands: {len(brands)}")
print(f"Categories: {len(categories)}")
print("\nDone. Pipeline complete.")
