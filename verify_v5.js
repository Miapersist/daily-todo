const AdmZip = require('adm-zip');
const fs = require('fs');
const z = new AdmZip('c:/Users/issuser/daily-todo/市况联动_供应商报价_联动定价_UAT测试用例_v5.docx');
const xml = z.readAsText('word/document.xml', 'utf8');

const wts = [];
const regex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
let m;
while ((m = regex.exec(xml)) !== null) { if (m[1].trim()) wts.push(m[1]); }

const tcIds = wts.filter(t => t.match(/^(PRICE-LINK|INQ-|QUOTE-|LINK-|APPR-)/));
console.log('Test case IDs found:', tcIds.length);
tcIds.forEach(t => console.log('  ' + t));

const oldContent = wts.filter(t => t.includes('价格目录 -') || t.includes('工装价格库') || t.includes('CE 角色'));
console.log('\nOld content remaining:', oldContent.length);
if (oldContent.length > 0) oldContent.slice(0,5).forEach(t => console.log('  OLD:', t));
else console.log('  NONE - clean!');
