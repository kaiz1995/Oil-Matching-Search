"""Remove embedded DATA array from HTML, add async loading logic"""
import re

with open('lubricant_product_matching_tool.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Find the DATA array by locating script tag content
# Look for "var DATA = [" right after the <script> tag
script_start = html.find('<script>')
data_var = html.find('var DATA = [', script_start)
data_end = html.find('];', data_var)

if data_var == -1 or data_end == -1:
    print("ERROR: DATA array not found")
    # Debug: print positions
    print(f"script tag at: {script_start}")
    print(f"First 500 chars after script: {html[script_start:script_start+500]}")
    exit(1)

# The DATA array is: "var DATA = [" + content + "];"
# We need to include everything up to "];\n\nvar BRANDS"
# But "];" might appear inside the data too (e.g., in strings)
# So find the unique pattern after the array

after_data = html[data_end:]
# Look for the pattern that follows the DATA array
next_line_match = re.search(r'\]\)?;\s*\n\s*var (BRANDS|CATEGORIES)', after_data)
if next_line_match:
    actual_end = data_end + next_line_match.start() + 1  # +1 for the ]
    print(f"DATA array found: positions {data_var} to {actual_end}")
    old_block = html[data_var:actual_end]
else:
    # Just use up to "];" then check next char
    actual_end = data_end + 2  # include the ];
    # Read a bit after to confirm
    print(f"DATA array ends at pos {data_end}, next 50 chars: {repr(html[data_end:data_end+50])}")
    old_block = html[data_var:actual_end]

print(f"DATA block length: {len(old_block)} chars")

# New data loading code
new_block = '''var DATA = [];
var DATA_VERSION = '';
var DATA_LOADING = false;

async function loadData() {
  if (DATA_LOADING) return;
  DATA_LOADING = true;
  var dataUrl = 'https://zk55806334-lang.github.io/oil-pm-data/data.json';

  try {
    var resp = await fetch(dataUrl + '?_=' + Date.now());
    var json = await resp.json();
    if (json.products && json.products.length > 0) {
      DATA = json.products;
      DATA_VERSION = json.version || '';
      localStorage.setItem('oil_pm_data', JSON.stringify(json));
      console.log('Data loaded from remote, version:', DATA_VERSION);
      DATA_LOADING = false;
      return;
    }
  } catch(e) { console.log('Remote load failed'); }

  try {
    var cached = localStorage.getItem('oil_pm_data');
    if (cached) {
      var json = JSON.parse(cached);
      if (json.products && json.products.length > 0) {
        DATA = json.products;
        DATA_VERSION = json.version || '';
        console.log('Data loaded from cache, version:', DATA_VERSION);
        DATA_LOADING = false;
        return;
      }
    }
  } catch(e) {}

  try {
    if (window.AndroidBridge) {
      var fb = AndroidBridge.getFallbackData();
      if (fb) { var json = JSON.parse(fb);
        if (json.products && json.products.length > 0) {
          DATA = json.products;
          DATA_VERSION = json.version || '';
        }
      }
    }
  } catch(e) {}

  if (DATA.length === 0) {
    console.error('All data sources failed');
    var el = document.getElementById('toast');
    if (el) { el.textContent = '\\u6570\\u636e\\u52a0\\u8f7d\\u5931\\u8d25\\uff0c\\u8bf7\\u68c0\\u67e5\\u7f51\\u7edc'; el.classList.add('show'); }
  }
  DATA_LOADING = false;
}

async function checkForUpdates() {
  if (!DATA_VERSION) return;
  try {
    var resp = await fetch('https://zk55806334-lang.github.io/oil-pm-data/data.json?_=' + Date.now());
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

html = html.replace(old_block, new_block.strip())

# Now modify the DOMContentLoaded
old_dcl_marker = "document.addEventListener('DOMContentLoaded', function() {"
dcl_pos = html.find(old_dcl_marker)

if dcl_pos != -1:
    # Find the end of this DOMContentLoaded block
    # It ends with "});" followed by a newline and then "function saveLast"
    dcl_end_marker = "});\n\nfunction saveLast"
    dcl_end = html.find(dcl_end_marker, dcl_pos)

    if dcl_end != -1:
        old_dcl = html[dcl_pos:dcl_end + 3]  # include the });

        new_dcl = '''document.addEventListener('DOMContentLoaded', function() {
  loadData().then(function() { checkForUpdates(); });

  document.getElementById('cp-model').addEventListener('keydown', function(e) { if (e.key==='Enter') searchCompetitor(); });
  document.getElementById('cp-specs').addEventListener('keydown', function(e) { if (e.ctrlKey && e.key==='Enter') searchCompetitor(); });

  try {
    var state = JSON.parse(localStorage.getItem('oil_pm_last'));
    if (state) {
      if (state.category) document.getElementById('cp-category').value = state.category;
      if (state.brand) document.getElementById('cp-brand').value = state.brand;
      if (state.model) document.getElementById('cp-model').value = state.model;
      if (state.viscosity) document.getElementById('cp-viscosity').value = state.viscosity;
      if (state.specs) document.getElementById('cp-specs').value = state.specs;
    }
  } catch(e) {}

  ['cp-category','cp-brand','cp-model','cp-viscosity','cp-specs'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) { el.addEventListener('change', saveLast); if (el.tagName==='INPUT'||el.tagName==='TEXTAREA') el.addEventListener('blur', saveLast); }
  });
});'''

        html = html.replace(old_dcl, new_dcl)
        print("DOMContentLoaded replaced")
    else:
        print("WARNING: Could not find DOMContentLoaded end")
        print(f"  Context after dcl_pos: {html[dcl_pos:dcl_pos+200]}")
else:
    print("WARNING: DOMContentLoaded not found in HTML")

with open('lubricant_product_matching_tool.html', 'w', encoding='utf-8') as f:
    f.write(html)

new_size = len(html)
print(f"OK - HTML updated! New size: {new_size} bytes")
