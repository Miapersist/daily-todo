const AdmZip = require('adm-zip');
const z = new AdmZip('c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx');
const xml = z.readAsText('word/document.xml', 'utf8');

// 找到第三个表格（第一个测试用例表）
const tblRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
let idx = 0;
let match;
while ((match = tblRegex.exec(xml)) !== null) {
    if (idx === 2) {
        console.log('=== Table 2 (first test case table) ===');
        const tblXml = match[0];
        
        // 找 <w:tr>
        const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
        let trIdx = 0;
        let trMatch;
        while ((trMatch = trRegex.exec(tblXml)) !== null) {
            console.log(`\n--- TR ${trIdx} ---`);
            const trXml = trMatch[0];
            
            // 找 <w:tc>
            const tcRegex = /<w:tc[\s\S]*?<\/w:tc>/g;
            let tcIdx = 0;
            let tcMatch;
            while ((tcMatch = tcRegex.exec(trXml)) !== null) {
                const tcXml = tcMatch[0];
                const wtRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
                let texts = '';
                let wtMatch;
                while ((wtMatch = wtRegex.exec(tcXml)) !== null) {
                    texts += wtMatch[1];
                }
                console.log(`  TC ${tcIdx}: "${texts}"`);
                tcIdx++;
            }
            trIdx++;
            if (trIdx >= 3) break;
        }
        break;
    }
    idx++;
}
