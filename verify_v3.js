const AdmZip = require('adm-zip');
const z = new AdmZip('c:/Users/issuser/daily-todo/市况联动_供应商报价_联动定价_UAT测试用例_v3.docx');
const xml = z.readAsText('word/document.xml', 'utf8');

// 提取所有非空 <w:t> 文本
const texts = [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).filter(t => t.trim());
console.log('Total non-empty <w:t> texts:', texts.length);
console.log('\nFirst 30 texts:');
texts.slice(0, 30).forEach((t, i) => console.log(`  [${i}] ${t}`));

// 找到测试用例相关文本
const tcTexts = texts.filter(t => t.includes('PRICE-LINK') || t.includes('INQ-') || t.includes('QUOTE-') || t.includes('LINK-') || t.includes('APPR-'));
console.log('\nTest case IDs found:', tcTexts.length);
tcTexts.forEach(t => console.log('  ' + t));

// 检查是否还有旧内容
const oldTexts = texts.filter(t => t.includes('价格目录 -') || t.includes('工装价格库 -') || t.includes('CE 角色') || t.includes('CE 映射'));
console.log('\nOld content remaining:', oldTexts.length);
if (oldTexts.length > 0) {
    oldTexts.slice(0, 5).forEach(t => console.log('  ' + t));
}
