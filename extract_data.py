"""Extract DATA array from HTML and save as standalone data.json"""
import re
import json

with open('lubricant_product_matching_tool.html', 'r', encoding='utf-8') as f:
    content = f.read()

match = re.search(r'var DATA = (\[.*?\]);', content, re.DOTALL)
if not match:
    print("ERROR: DATA array not found")
    exit(1)

data = json.loads(match.group(1))
output = {
    'version': '2026-05-19-v1',
    'updatedAt': '2026-05-19',
    'totalProducts': len(data),
    'products': data
}

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"OK - data.json created. Products: {len(data)}, Size: {len(json.dumps(output, ensure_ascii=False))} bytes")
