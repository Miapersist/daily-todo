const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

const src = 'c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx';

try {
    const zip = new AdmZip(src);
    const xml = zip.readAsText('word/document.xml', 'utf8');
    
    // Extract all <w:t> texts
    const texts = [];
    const regex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let match;
    while ((match = regex.exec(xml)) !== null) {
        texts.push(match[1]);
    }
    
    // Join and write
    const plain = texts.join('');
    fs.writeFileSync('c:/Users/issuser/daily-todo/docx_text.txt', plain, 'utf8');
    console.log('Done. Chars:', plain.length);
} catch(e) {
    console.error('Error:', e.message);
}
