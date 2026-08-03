/**
 * 生成 UAT 测试用例 Word 文档 V3
 * 策略：保留原模板所有样式文件，读取 document.xml，
 * 精确识别每个 <w:tbl> 表格，按表格索引替换内容。
 * 表0-1：封面/数据准备（保留结构，替换文字）
 * 表2-23：测试用例（替换为新的测试数据）
 */
const AdmZip = require('adm-zip');
const fs = require('fs');

const TEMPLATE = 'c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx';
const OUTPUT = 'c:/Users/issuser/daily-todo/市况联动_供应商报价_联动定价_UAT测试用例_v3.docx';

// ========== 新内容 ==========
const NEW_TITLE = '广汽集团零采一体化项目';
const NEW_SUBTITLE = '市况联动——供应商报价 & 联动定价';
const NEW_MODULE = '功能模块优化';
const NEW_TEST_TYPE = 'UAT测试用例&测试报告';

// 前言
const NEW_INTRO = '验证市况联动板块优化 v4.1 中供应商报价流程和市况联动定价功能的正确性和完整性，确保各角色操作闭环通过。';
const NEW_METHOD = '黑盒测试，按用例逐步操作，验证预期结果与实际结果一致。';

// 数据准备
const DATA_PREP_ROWS = [
    ['1', '谷歌浏览器及用户账号', '登录 GPM 系统'],
    ['2', '市况材料主数据（铜、铝、PP等）', '已配置市况材料主数据'],
    ['3', '原材料行情数据', '已导入百川资讯/供应商来源行情'],
    ['4', '市况类型与材料关系', '已维护并生效'],
    ['5', '市况联动公式', '已配置材料单价要素+行情来源'],
    ['6', '零件认领', '已配置"是否联动调价/市况补偿"为"是"'],
    ['7', '零件价格联动维护', '已生效'],
    ['8', '测试用供应商账号', '可正常登录'],
];

// 测试用例 - 扁平化，每个用例一行
const ALL_TEST_CASES = [
    // 2.1 零件价格联动维护
    {id:'PRICE-LINK-001', name:'PP/钢板/铝板手工新增联动关系', pre:'1.已登录GPM系统\n2.配置管理→零件价格联动维护有权限', desc:'验证PP、钢板、铝板市况类型支持手工新增联动关系，字段名称使用新命名', scope:'配置管理→零件价格联动维护', steps:'1.进入"零件价格联动维护"页面\n2.点击"新增"\n3.选择PP/钢板/铝板市况类型\n4.填写材料编码、材料供应商等信息\n5.点击"保存"\n6.点击"生效"', expect:'1.列表表头显示"材料编码""材料描述""材料供应商编码""材料供应商名称"（非"管理支给件"旧名称）\n2.新增记录保存成功，状态为"拟定"\n3.点击"生效"后状态变为"生效"'},
    {id:'PRICE-LINK-002', name:'PP/钢板/铝板批量导入联动关系', pre:'1.同上\n2.已下载导入模板', desc:'验证PP、钢板、铝板市况类型支持Excel批量导入', scope:'配置管理→零件价格联动维护', steps:'1.点击"导入"\n2.下载导入模板\n3.按模板填写零件+PP/钢板/铝板市况类型数据\n4.上传模板', expect:'1.模板字段使用新命名"材料编码""材料描述"等\n2.校验通过的行导入成功\n3.失败行提示具体原因'},
    {id:'PRICE-LINK-003', name:'非PP/钢板/铝板市况类型导入失败', pre:'1.同上\n2.已准备含其他市况类型的导入文件', desc:'验证非PP/钢板/铝板市况类型数据导入时被拦截', scope:'配置管理→零件价格联动维护', steps:'1.在导入模板中填写非PP/钢板/铝板的市况类型（如铜）\n2.上传模板', expect:'导入失败，提示"该市况类型不支持导入，仅支持PP、钢板、铝板"'},
    {id:'PRICE-LINK-004', name:'管理支给件-总成件市况类型置灰', pre:'1.同上\n2.零件映射类型为"管理支给件-总成件"', desc:'验证总成件市况类型固定为空且置灰', scope:'配置管理→零件价格联动维护', steps:'1.查询映射类型为"管理支给件-总成件"的零件\n2.点击"新增"', expect:'市况类型字段置灰不可选择，显示为空'},
    {id:'PRICE-LINK-005', name:'零件认领生效后自动生成联动数据', pre:'1.同上\n2.零件认领中已配置非大宗料市况类型\n3.当前无联动数据', desc:'验证零件认领保存生效后，非大宗料联动关系自动生成', scope:'配置管理→零件认领→维护供应商', steps:'1.将零件"是否联动调价/市况补偿"设为"是"\n2.选择非大宗料市况类型\n3.保存并生效', expect:'1.零件价格联动维护页面自动生成一条联动数据\n2.材料供应商信息复制零件供应商信息\n3.状态为"生效"'},
    {id:'PRICE-LINK-006', name:'零件联动调价改为"否"后联动失效', pre:'1.同上\n2.已有联动数据', desc:'验证零件联动调价由"是"改为"否"后，联动数据不再参与计价', scope:'配置管理→零件认领', steps:'1.将零件"是否联动调价/市况补偿"由"是"改为"否"\n2.保存并生效', expect:'1.历史联动数据保留\n2.该零件不再进入市况联动定价单可联动范围'},
    
    // 2.2 询价单配置
    {id:'INQ-001', name:'询价单价格属性选择"联动定价"', pre:'1.已登录GPM系统（采购员账号）\n2.询价管理→询价单有权限', desc:'验证询价单价格属性新增"联动定价"选项', scope:'询价管理→询价单→新增/编辑', steps:'1.创建询价单\n2.选择供应商、零件\n3.在价格属性下拉中选择"联动定价"\n4.保存', expect:'1.价格属性下拉框包含"联动定价"选项\n2.选择"联动定价"后可正常保存'},
    {id:'INQ-002', name:'询价单配置"是否允许修改重量"', pre:'1.同上\n2.零件已配置市况联动', desc:'验证采购员可控制供应商是否允许修改投料重量和废料重量', scope:'询价管理→询价单→新增/编辑', steps:'1.创建询价单\n2.选择联动定价零件\n3."是否允许修改重量"选"否"\n4.发起询价', expect:'供应商报价单中投料重量和废料重量字段置灰不可编辑'},
    {id:'INQ-003', name:'询价单配置"是否允许修改市况材料单价"', pre:'1.同上', desc:'验证系统自动带价场景下，采购员可控制材料单价是否可修改', scope:'询价管理→询价单→新增/编辑', steps:'1.创建询价单\n2."是否允许修改市况材料单价"选"否"\n3.发起询价', expect:'供应商报价单中材料单价字段置灰不可编辑'},
    {id:'INQ-004', name:'询价单配置"是否允许修改其他信息"', pre:'1.同上', desc:'验证采购员可控制废料单价、合格率、加工费是否可修改', scope:'询价管理→询价单→新增/编辑', steps:'1.创建询价单\n2."是否允许修改其他信息"选"否"\n3.发起询价', expect:'供应商报价单中废料单价、合格率、加工费字段置灰不可编辑'},
    {id:'INQ-005', name:'历史报价默认值带入', pre:'1.同上\n2.该零件存在供应商已审核通过的历史报价', desc:'验证询价单发起时自动带入历史报价信息', scope:'询价管理→询价单→新增', steps:'1.选择有历史报价的零件和供应商\n2.创建询价单并发起', expect:'供应商报价单自动带入上一次通过报价的全部信息（材料单价、投料重量、合格率等）作为默认值'},
    
    // 2.3 供应商填写报价单
    {id:'QUOTE-001', name:'供应商查看待报价单据', pre:'1.已登录GPM系统（供应商账号）\n2.采购员已发起询价', desc:'验证供应商可查看待报价单据并进入报价', scope:'询价协同→报价单', steps:'1.进入报价单列表\n2.找到状态为"待报价"的单据\n3.点击"报价"', expect:'进入报价编辑页面，显示零件报价信息'},
    {id:'QUOTE-002', name:'供应商查看权限控制字段', pre:'1.同上\n2.采购员配置了不允许修改重量和单价', desc:'验证供应商页面中被禁止修改的字段置灰', scope:'询价协同→报价单→编辑', steps:'1.进入报价编辑页面\n2.查看材料单价字段\n3.查看投料重量字段', expect:'1.材料单价置灰不可编辑\n2.投料重量置灰不可编辑\n3.允许修改的字段（如加工费）为白色可编辑'},
    {id:'QUOTE-003', name:'供应商填写市况基准日期', pre:'1.同上\n2.采购员允许修改市况信息', desc:'验证供应商可选择市况基准日期区间', scope:'询价协同→报价单→编辑', steps:'1.点击市况基准开始日期\n2.选择日期\n3.点击市况基准结束日期\n4.选择日期', expect:'1.日期选择器正常弹出\n2.日期区间与联动频次匹配\n3.系统自动回填联动频次'},
    {id:'QUOTE-004', name:'供应商填写市况基准行情', pre:'1.同上', desc:'验证供应商可填写市况基准行情价格', scope:'询价协同→报价单→编辑', steps:'1.在市况基准行情字段输入价格\n2.保存', expect:'1.数值输入正常\n2.保存后数据不丢失'},
    {id:'QUOTE-005', name:'供应商提交报价', pre:'1.同上\n2.所有必填字段已填写', desc:'验证供应商可正常提交报价', scope:'询价协同→报价单→编辑', steps:'1.检查所有必填字段\n2.点击"提交"', expect:'1.提交成功\n2.报价单状态更新为"已报价"'},
    {id:'QUOTE-006', name:'提交时必填字段校验', pre:'1.同上\n2.存在未填写的必填字段', desc:'验证提交时校验必填字段', scope:'询价协同→报价单→编辑', steps:'1.不填写投料重量\n2.点击"提交"', expect:'提交失败，提示"请填写必填项"，必填字段标红'},
    
    // 2.4 市况联动定价单
    {id:'LINK-001', name:'新增市况联动定价单', pre:'1.已登录GPM系统（采购员账号）\n2.市况管理→联动定价→市况联动定价单有权限', desc:'验证采购员可创建市况联动定价单', scope:'市况管理→联动定价→市况联动定价单', steps:'1.点击"新增"\n2.选择供应商\n3.选择定价类型（正式价/暂估价）\n4.选择币种、税率\n5.保存', expect:'1.定价单创建成功\n2.状态为"拟定"\n3.自动带出该供应商名下全部零件'},
    {id:'LINK-002', name:'选择供应商后自动带出零件', pre:'1.同上', desc:'验证选择供应商后自动带出已报价和未报价零件', scope:'市况管理→联动定价→新增', steps:'1.选择供应商\n2.查看定价零件列表', expect:'1.已报价零件：带出报价明细，删除按钮置灰\n2.未报价零件：费用字段为空，允许删除\n3.页面提示"已带出该供应商的全部零件信息"'},
    {id:'LINK-003', name:'同步市况&计算', pre:'1.同上\n2.定价单中有零件', desc:'验证点击"同步市况&计算"后系统执行联动计算', scope:'市况管理→联动定价→编辑', steps:'1.点击"同步市况&计算"\n2.等待计算完成', expect:'1.系统按公式配置计算材料费和本次价格\n2.定价零件列表更新价格\n3.市况材料明细更新'},
    {id:'LINK-004', name:'行情来源切换', pre:'1.同上\n2.已同步市况计算', desc:'验证定价零件行情来源可切换', scope:'市况管理→联动定价→编辑', steps:'1.在定价零件行点击行情来源下拉\n2.切换为"供应商"', expect:'1.行情来源切换成功\n2.供应商行情来源时按供应商行情取值'},
    {id:'LINK-005', name:'导入年降及返利', pre:'1.同上\n2.已同步市况计算', desc:'验证年降比例和返利金额可导入', scope:'市况管理→联动定价→编辑', steps:'1.点击"导入"\n2.下载模板\n3.填写年降比例\n4.上传模板', expect:'1.年降比例导入成功\n2.年降金额自动计算（=上次价格×年降比例）\n3.返利金额导入成功'},
    {id:'LINK-006', name:'发起价格审批-校验通过', pre:'1.同上\n2.供应商、定价类型、币种、税率已填\n3.已报价零件已完成同步计算', desc:'验证满足条件时可发起价格审批', scope:'市况管理→联动定价→编辑', steps:'1.确认所有已报价零件已计算\n2.点击"发起价格审批"', expect:'1.系统创建拟定状态的价格审批单\n2.定价单状态变为"已生成价格审批单"\n3.价格审批单零件行数据正确'},
    {id:'LINK-007', name:'发起价格审批-校验失败', pre:'1.同上\n2.有零件未完成同步计算', desc:'验证未完成计算时不允许提交审批', scope:'市况管理→联动定价→编辑', steps:'1.不执行"同步市况&计算"\n2.直接点击"发起价格审批"', expect:'提交失败，提示"请完成市况同步及价格计算后再提交价格审批"'},
    
    // 2.5 价格审批与联动结果
    {id:'APPR-001', name:'审批通过后写入价格库', pre:'1.已登录GPM系统（审批人员账号）\n2.已生成联动定价审批单\n3.价格审批单处于审批流程中', desc:'验证审批通过后联动价格写入价格库', scope:'价格管理→价格审批单→审批', steps:'1.核对定价零件及价格明细\n2.审批通过', expect:'1.本次联动定价结果写入价格库\n2.零件价格目录更新为审批通过价格\n3.市况变动按规则统计'},
    {id:'APPR-002', name:'报价单终审后自动生成投料数据', pre:'1.同上\n2.报价单终审完成且已生成价格目录', desc:'验证报价终审后自动生成零件原材料投料重量数据', scope:'市况管理→市况基础信息→零件原材料投料重量维护', steps:'1.查询对应零件\n2.查看投料重量数据', expect:'1.自动生成新的生效数据\n2.原有生效数据变为失效\n3.生效时间为本次报价终审完成时间'},
    {id:'APPR-003', name:'报价单作废后投料数据作废', pre:'1.同上\n2.已有自动生成的投料数据', desc:'验证报价单作废时对应投料数据自动作废', scope:'询价管理→报价单→作废', steps:'1.作废已终审的报价单\n2.查看投料重量维护页面', expect:'对应批次投料数据自动作废'},
];

function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * 获取表格中所有行的 <w:t> 文本列表
 * 返回 [[row1_texts], [row2_texts], ...]
 */
function getTableWtCells(tblXml) {
    const rows = [];
    const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
    let trMatch;
    while ((trMatch = trRegex.exec(tblXml)) !== null) {
        const cells = [];
        const tcRegex = /<w:tc[\s\S]*?<\/w:tc>/g;
        let tcMatch;
        while ((tcMatch = tcRegex.exec(trMatch[0])) !== null) {
            const wtRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
            let cellText = '';
            let wtMatch;
            while ((wtMatch = wtRegex.exec(tcMatch[0])) !== null) {
                cellText += wtMatch[1];
            }
            cells.push(cellText);
        }
        rows.push(cells);
    }
    return rows;
}

/**
 * 生成一个新表格行的 XML（从原行复制样式，替换文本）
 */
function replaceRowTexts(rowXml, newTexts) {
    let result = '';
    let lastIdx = 0;
    const tcRegex = /(<w:tc[\s\S]*?<\/w:tc>)/g;
    let tcIdx = 0;
    let tcMatch;
    
    while ((tcMatch = tcRegex.exec(rowXml)) !== null) {
        result += rowXml.substring(lastIdx, tcMatch.index);
        
        let cellXml = tcMatch[1];
        if (tcIdx < newTexts.length) {
            // 替换这个单元格的 <w:t> 文本
            const wtRegex = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g;
            const wtMatches = [...cellXml.matchAll(wtRegex)];
            if (wtMatches.length > 0) {
                // 合并所有 wt 文本
                let newCellXml = cellXml;
                // 简单做法：替换第一个 <w:t> 为全部文本，清空后续 <w:t>
                const allWt = [...newCellXml.matchAll(wtRegex)];
                if (allWt.length > 0) {
                    const first = allWt[0];
                    newCellXml = newCellXml.substring(0, first.index) + 
                        first[1] + escapeXml(newTexts[tcIdx]) + first[3] +
                        newCellXml.substring(first.index + first[0].length);
                    // 清空后续 <w:t> 文本
                    for (let i = 1; i < allWt.length; i++) {
                        const wt = allWt[i];
                        const old = wt[1] + wt[2] + wt[3];
                        const repl = wt[1] + '' + wt[3];
                        newCellXml = newCellXml.replace(old, repl);
                    }
                }
                result += newCellXml;
            } else {
                result += cellXml;
            }
        } else {
            result += cellXml;
        }
        lastIdx = tcMatch.index + tcMatch[0].length;
        tcIdx++;
    }
    result += rowXml.substring(lastIdx);
    return result;
}

function buildDocx() {
    const template = new AdmZip(TEMPLATE);
    let xml = template.readAsText('word/document.xml', 'utf8');
    
    // ===== 第一步：替换封面/前言/目录中的文本 =====
    // 用正则找到所有 <w:t> 标签并按序替换关键文本
    const wtRegex = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g;
    const wtItems = [...xml.matchAll(wtRegex)];
    
    // 构建文本到新文本的映射（基于精确文本匹配）
    const textReplaceMap = new Map();
    
    // 封面
    textReplaceMap.set('广汽集团零采一体化项目', NEW_TITLE);
    textReplaceMap.set('市况联动相关', NEW_SUBTITLE);
    textReplaceMap.set('功能模块', NEW_MODULE);
    textReplaceMap.set('UAT测试用例&测试报告', NEW_TEST_TYPE);
    
    // 前言
    textReplaceMap.set('指导测试工作有序进行，使实施测试的有据可依；确保所实现的功能与客户预期的需求相符合。', NEW_INTRO);
    textReplaceMap.set('在测试过程中系统每个模块可以先调用通用测试用例（新增、查询、编辑、删除、导入、导出、分页）执行测试过程，然后根据功能测试用例执行业务流程测试。', NEW_METHOD);
    
    // 目录章节标题
    textReplaceMap.set('零件认领', '零件价格联动维护');
    textReplaceMap.set('新增"是否联动调价/市况补偿"字段', 'PP/钢板/铝板手工新增联动关系');
    textReplaceMap.set('市况物料维护', '询价单配置（采购员发起）');
    textReplaceMap.set('新增【市况物料维护】页面', '询价单价格属性选择"联动定价"');
    textReplaceMap.set('原材料行情维护', '供应商填写报价单');
    textReplaceMap.set('原材料行情维护移除工厂维度', '供应商查看待报价单据');
    textReplaceMap.set('询价单', '市况联动定价单');
    textReplaceMap.set('询价单新增控制字段', '新增市况联动定价单');
    textReplaceMap.set('报价单', '价格审批与联动结果');
    textReplaceMap.set('新增生管字段展示', '审批通过后写入价格库');
    textReplaceMap.set('零件价格联动维护', '');  // 原目录项移除
    textReplaceMap.set('零件价格联动维护优化', '');
    textReplaceMap.set('市况类型与材料关系维护', '');
    textReplaceMap.set('市况类型与材料关系维护', '');
    textReplaceMap.set('零件原材料投料重量维护', '');
    textReplaceMap.set('数据信息通过价格审批单审批通过后自动同步', '');
    textReplaceMap.set('单台原材料投料重量维护', '');
    textReplaceMap.set('数据信息通过外购成本月度报表月度锁定自动汇总生成', '');
    
    // 执行替换
    for (const item of wtItems) {
        const oldText = item[2];
        if (textReplaceMap.has(oldText)) {
            const newText = textReplaceMap.get(oldText);
            const oldFull = item[0];
            const newFull = item[1] + escapeXml(newText) + item[3];
            xml = xml.replace(oldFull, newFull);
        }
    }
    
    // ===== 第二步：替换表格内容 =====
    // 找到所有 <w:tbl>
    const tblRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
    let tblMatches = [...xml.matchAll(tblRegex)];
    
    console.log('Total tables:', tblMatches.length);
    
    // 表0：封面"测试人确认"
    // 表1：数据准备表格（3列：序号、提前准备事项、用途说明）
    // 表2-23：测试用例表格（10列）
    
    let newXml = '';
    let lastPos = 0;
    let tblIdx = 0;
    
    for (const tblMatch of tblMatches) {
        newXml += xml.substring(lastPos, tblMatch.index);
        let tblContent = tblMatch[0];
        
        if (tblIdx === 1) {
            // ===== 表1：数据准备表格 =====
            const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
            let trMatches = [...tblContent.matchAll(trRegex)];
            // 第一行是表头，保留
            // 从第二行开始替换为新的数据准备行
            if (trMatches.length >= 2) {
                const headerRow = trMatches[0][0];
                const sampleDataRow = trMatches[1][0]; // 用作样式模板
                
                let newTblContent = headerRow;
                for (let i = 0; i < DATA_PREP_ROWS.length; i++) {
                    newTblContent += replaceRowTexts(sampleDataRow, DATA_PREP_ROWS[i]);
                }
                
                // 替换整个表体
                const oldBodyStart = trMatches[1].index;
                const oldBodyEnd = trMatches[trMatches.length - 1].index + trMatches[trMatches.length - 1][0].length;
                tblContent = tblContent.substring(0, oldBodyStart) + newTblContent.substring(headerRow.length) + tblContent.substring(oldBodyEnd);
            }
        } else if (tblIdx >= 2) {
            // ===== 测试用例表格（表2+） =====
            const caseIdx = tblIdx - 2;
            if (caseIdx < ALL_TEST_CASES.length) {
                const tc = ALL_TEST_CASES[caseIdx];
                const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
                let trMatches = [...tblContent.matchAll(trRegex)];
                
                if (trMatches.length >= 2) {
                    // 第一行：表头（保留）
                    // 第二行：第一个数据行（用作样式模板）
                    const headerRow = trMatches[0][0];
                    const sampleRow = trMatches[1][0];
                    
                    // 构建新行数据
                    const newCells = [
                        tc.id,           // 用例编号
                        tc.name,         // 用例名称
                        tc.pre,          // 前置条件
                        tc.desc,         // 用例描述
                        tc.scope,        // 体现环节
                        tc.steps,        // 操作步骤
                        tc.expect,       // 预期结果
                        '',              // 图示
                        '',              // 测试结果
                    ];
                    
                    // 处理换行符：在 OOXML 中，\n 需要转为 <w:br/>
                    const processedCells = newCells.map(cell => cell.replace(/\n/g, '<w:br/>'));
                    
                    // 替换
                    let newRowXml = sampleRow;
                    const tcRegex = /(<w:tc[\s\S]*?<\/w:tc>)/g;
                    let tcMatches = [...newRowXml.matchAll(tcRegex)];
                    
                    for (let i = 0; i < Math.min(tcMatches.length, processedCells.length); i++) {
                        let cellXml = tcMatches[i][1];
                        
                        // 找到 <w:t> 标签
                        const wtRegex2 = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g;
                        const wts = [...cellXml.matchAll(wtRegex2)];
                        
                        if (wts.length > 0) {
                            // 替换第一个 <w:t> 文本
                            const first = wts[0];
                            let newCellXml = cellXml.substring(0, first.index) +
                                first[1] + escapeXml(processedCells[i]) + first[3] +
                                cellXml.substring(first.index + first[0].length);
                            
                            // 清空后续 <w:t> 文本
                            for (let j = 1; j < wts.length; j++) {
                                const wt = wts[j];
                                newCellXml = newCellXml.replace(wt[0], wt[1] + '' + wt[3]);
                            }
                            newRowXml = newRowXml.replace(cellXml, newCellXml);
                        }
                    }
                    
                    tblContent = headerRow + newRowXml;
                }
            }
        }
        
        newXml += tblContent;
        lastPos = tblMatch.index + tblMatch[0].length;
        tblIdx++;
    }
    newXml += xml.substring(lastPos);
    
    // ===== 写入输出文件 =====
    const zip = new AdmZip();
    const entries = template.getEntries();
    entries.forEach(entry => {
        if (entry.entryName === 'word/document.xml') {
            zip.addFile('word/document.xml', Buffer.from(newXml, 'utf8'));
        } else {
            zip.addFile(entry.entryName, entry.getData());
        }
    });
    
    zip.writeZip(OUTPUT);
    console.log('Generated:', OUTPUT);
    console.log('Total test cases:', ALL_TEST_CASES.length);
}

try {
    buildDocx();
    console.log('DONE!');
} catch(e) {
    console.error('ERROR:', e.message);
    console.error(e.stack);
}
