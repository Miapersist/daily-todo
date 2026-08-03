const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

const TEMPLATE = 'c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx';

const z = new AdmZip(TEMPLATE);
const xml = z.readAsText('word/document.xml', 'utf8');

// Write full XML for analysis
fs.writeFileSync('c:/Users/issuser/daily-todo/template_document.xml', xml, 'utf8');

// Extract structure info
// Find all paragraph styles
const pStyles = [...xml.matchAll(/<w:pStyle w:val="([^"]+)"/g)].map(m => m[1]);
console.log('Paragraph styles:', [...new Set(pStyles)]);

// Find all table-related info
const tblCount = (xml.match(/<w:tbl>/g) || []).length;
console.log('Tables:', tblCount);

// Find first 500 chars of each table
let tblIdx = 0;
const tblRegex = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
let tblMatch;
while ((tblMatch = tblRegex.exec(xml)) !== null) {
    console.log(`\n=== Table ${tblIdx} ===`);
    // Get all text content in this table
    const tblXml = tblMatch[0];
    const texts = [...tblXml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]);
    console.log('  Texts:', texts.slice(0, 20));
    tblIdx++;
}

// Find all sections (page layout)
const sectCount = (xml.match(/<w:sectPr/g) || []).length;
console.log('\nSections:', sectCount);

// Find the beginning of document body (first 2000 chars of body)
const bodyStart = xml.indexOf('<w:body>');
console.log('\nBody start (first 2000 chars):');
console.log(xml.substring(bodyStart, bodyStart + 2000));

console.log('\n\nTotal XML length:', xml.length);
