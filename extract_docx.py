"""Extract text from docx using zipfile and write to file."""
import sys, os, re
from zipfile import ZipFile

docx_path = r"c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx"
output_path = r"c:\Users\issuser\daily-todo\test_output.txt"

try:
    zf = ZipFile(docx_path, 'r')
    xml = zf.read('word/document.xml').decode('utf-8')
    zf.close()
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
    plain = ''.join(texts)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(plain)
    
    print(f'OK: {len(plain)} chars written to {output_path}')
except Exception as e:
    print(f'ERROR: {e}')
