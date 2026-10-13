import re, subprocess, os

def check_html(path):
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()
    
    scripts = re.findall(r'<script\b[^>]*>(.*?)</script>', html, re.DOTALL)
    print(f"Checking {path}: {len(scripts)} scripts found.")
    
    for idx, sc in enumerate(scripts):
        # Clean jinja tags
        clean = re.sub(r'\{%.*?%\}', '/* jinja */', sc)
        clean = re.sub(r'\{\{.*?\}\}', '"jinja_val"', clean)
        
        tmp = f"temp_check_{idx}.js"
        with open(tmp, 'w', encoding='utf-8') as tf:
            tf.write(clean)
        
        res = subprocess.run(['node', '--check', tmp], capture_output=True, text=True)
        if os.path.exists(tmp):
            os.remove(tmp)
            
        if res.returncode != 0:
            print(f"SYNTAX ERROR in script #{idx+1} of {path}:")
            print(res.stderr)
            return False
        else:
            print(f"Script #{idx+1} in {path} passed node --check syntax verification.")
    return True

if __name__ == '__main__':
    ok1 = check_html('cloud/templates/index.html')
    ok2 = check_html('templates/index.html')
    if ok1 and ok2:
        print("ALL JAVASCRIPT SYNTAX IN BOTH TEMPLATES IS 100% VALID!")
    else:
        exit(1)
