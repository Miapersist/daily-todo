import sys
sys.stdout = open(r"c:\Users\issuser\daily-todo\result.txt", "w", encoding="utf-8")
sys.stderr = sys.stdout

from zipfile import ZipFile
import re

docx_path = r"c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx"

print("Starting...")
try:
    zf = ZipFile(docx_path, 'r')
    names = zf.namelist()
    print(f"Files in zip: {names[:10]}")
    xml = zf.read('word/document.xml').decode('utf-8')
    zf.close()
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
    plain = ''.join(texts)
    print(f"Extracted {len(plain)} chars")
    # Write to separate file
    with open(r"c:\Users\issuser\daily-todo\test_output.txt", "w", encoding="utf-8") as f:
        f.write(plain)
    print("DONE")
except Exception as e:
    print(f"ERROR: {type(e).__name__}: {e}")
    import traceback
    traceback.print_exc()
