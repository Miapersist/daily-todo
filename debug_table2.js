const AdmZip = require('adm-zip');
const z = new AdmZip('c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx');
const xml = z.readAsText('word/document.xml', 'utf8');

// 找到第三个表格（第一个测试用例表）
const tblRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
let idx = 0;
let match;
while ((match = tblRegex.exec(xml)) !== null) {
    if (idx === 2) {
        console.log('=== Table 2 - Full rows ===');
        const tblXml = match[0];
        const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
        let trIdx = 0;
        let trMatch;
        while ((trMatch = trRegex.exec(tblXml)) !== null) {
            const tcRegex = /<w:tc[\s\S]*?<\/w:tc>/g;
            let cells = [];
            let tcMatch;
            while ((tcMatch = tcRegex.exec(trMatch[0])) !== null) {
                const wtRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
                let text = '';
                let wtMatch;
                while ((wtMatch = wtRegex.exec(tcMatch[0])) !== null) {
                    text += wtMatch[1];
                }
                cells.push(text);
            }
            console.log(`  TR${trIdx}: ${cells.length} cells -> [${cells.map(c => '"' + c.substring(0,40) + '"').join(', ')}]`);
            trIdx++;
        }
        break;
    }
    idx++;
}
