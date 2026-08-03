const AdmZip = require('adm-zip');
const z = new AdmZip('c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx');
const xml = z.readAsText('word/document.xml', 'utf8');

// 找到所有测试用例表格（表2+）
const tblRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
const tbls = [];
let m;
while ((m = tblRegex.exec(xml)) !== null) tbls.push({index: m.index, len: m[0].length});

// 看相邻两个测试用例表格之间
for (let i = 3; i < Math.min(6, tbls.length); i++) {
    const prevEnd = tbls[i-1].index + tbls[i-1].length;
    const gap = xml.substring(prevEnd, tbls[i].index);
    
    // 数段落
    const paras = [];
    const pRegex = /<w:p[\s>][\s\S]*?<\/w:p>/g;
    let pm;
    while ((pm = pRegex.exec(gap)) !== null) paras.push(pm[0]);
    
    console.log(`Table ${i-1} -> ${i}: ${paras.length} paragraphs`);
    paras.forEach((p, j) => {
        const wts = [];
        const wtRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
        let w;
        while ((w = wtRegex.exec(p)) !== null) wts.push(w[1]);
        const text = wts.join('').trim();
        if (text) console.log(`  P${j}: "${text.substring(0,60)}"`);
        else console.log(`  P${j}: [empty]`);
    });
}
