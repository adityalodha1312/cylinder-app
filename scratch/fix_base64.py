import re

with open('e:/Cylinder_MVP/scratch/logo_b64.txt', 'r', encoding='utf-8') as f:
    text = f.read().strip()
    # Extract just the base64 part
    if ',' in text:
        b64 = text.split(',')[1].strip()
    else:
        b64 = text

with open('e:/Cylinder_MVP/cylinder_full_script.gs', 'r', encoding='utf-8') as f:
    code = f.read()

code = re.sub(r'const logoBase64 = ".*?";', f'const logoBase64 = "{b64}";', code, flags=re.DOTALL)

with open('e:/Cylinder_MVP/cylinder_full_script.gs', 'w', encoding='utf-8') as f:
    f.write(code)

print("Fixed base64")
