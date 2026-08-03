const AdmZip = require('adm-zip');
const z = new AdmZip('c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx');
const xml = z.readAsText('word/document.xml', 'utf8');

// 找测试用例表格之间的内容
const tblRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
const tbls = [];
let m;
while ((m = tblRegex.exec(xml)) !== null) tbls.push({index: m.index, len: m[0].length});

// 看表2-表5之间的间隔内容
for (let i = 2; i < Math.min(6, tbls.length); i++) {
    const prevEnd = tbls[i-1].index + tbls[i-1].length;
    const gap = xml.substring(prevEnd, tbls[i].index);
    // 数段落
    const pCount = (gap.match(/<w:p[\s>]/g) || []).length;
    console.log(`Table ${i-1} -> Table ${i}: gap ${gap.length} chars, ${pCount} paragraphs`);
    // 看看gap里有什么
    const wts = [];
    const wtRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let wt;
    while ((wt = wtRegex.exec(gap)) !== null) wts.push(wt[1]);
    console.log('  Texts:', wts.filter(t=>t.trim()).slice(0,5));
}
