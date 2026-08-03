/**
 * 生成 UAT 测试用例 Word 文档 - 最终版
 * 
 * 原模板测试用例表格结构（纵向布局）：
 *   Row 0: 用例编号 | 值
 *   Row 1: 用例名称 | 值
 *   Row 2: 前置条件 | 值
 *   Row 3: 用例描述 | 值
 *   Row 4: 体现环节 | 操作步骤 | 预期结果 | 图示 | 测试结果 (5列表头)
 *   Row 5+: 步骤序号 | 步骤内容 | 预期结果 | 图示 | 测试结果
 *
 * 策略：保留模板完整 XML，替换封面文字、数据准备表格、测试用例表格内容。
 */
const AdmZip = require('adm-zip');

const TEMPLATE = 'c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx';
const OUTPUT = 'c:/Users/issuser/daily-todo/市况联动_供应商报价_联动定价_UAT测试用例_v4.docx';

// ========== 测试用例数据（27个） ==========
const ALL_TEST_CASES = [
    // 2.1 零件价格联动维护
    {id:'PRICE-LINK-001', name:'PP/钢板/铝板手工新增联动关系', pre:'1.已登录GPM系统\n2.配置管理→零件价格联动维护有权限', desc:'验证PP、钢板、铝板市况类型支持手工新增联动关系，字段名称使用新命名', scope:'配置管理→零件价格联动维护', steps:[['1','进入"零件价格联动维护"页面','成功打开零件价格联动维护菜单页面','',''],['2','点击"新增"，选择PP/钢板/铝板市况类型','弹出新增窗口，市况类型可选择','',''],['3','填写材料编码、材料供应商等信息，点击"保存"','保存成功，状态为"拟定"','',''],['4','点击"生效"','状态变为"生效"','','']]},
    {id:'PRICE-LINK-002', name:'PP/钢板/铝板批量导入联动关系', pre:'1.同上\n2.已下载导入模板', desc:'验证PP、钢板、铝板市况类型支持Excel批量导入', scope:'配置管理→零件价格联动维护', steps:[['1','点击"导入"，下载模板','模板字段使用新命名"材料编码""材料描述"等','',''],['2','按模板填写零件+PP/钢板/铝板市况类型数据','校验通过的行导入成功','',''],['3','查看导入失败行','失败行提示具体原因','','']]},
    {id:'PRICE-LINK-003', name:'非PP/钢板/铝板市况类型导入失败', pre:'1.同上\n2.已准备含其他市况类型的导入文件', desc:'验证非PP/钢板/铝板市况类型数据导入时被拦截', scope:'配置管理→零件价格联动维护', steps:[['1','在导入模板中填写非PP/钢板/铝板的市况类型（如铜）','提示"该市况类型不支持导入，仅支持PP、钢板、铝板"','',''],['2','上传模板','导入失败','','']]},
    {id:'PRICE-LINK-004', name:'管理支给件-总成件市况类型置灰', pre:'1.同上\n2.零件映射类型为"管理支给件-总成件"', desc:'验证总成件市况类型固定为空且置灰', scope:'配置管理→零件价格联动维护', steps:[['1','查询映射类型为"管理支给件-总成件"的零件','查询结果正常','',''],['2','点击"新增"','市况类型字段置灰不可选择，显示为空','','']]},
    {id:'PRICE-LINK-005', name:'零件认领生效后自动生成联动数据', pre:'1.同上\n2.零件认领中已配置非大宗料市况类型\n3.当前无联动数据', desc:'验证零件认领保存生效后，非大宗料联动关系自动生成', scope:'配置管理→零件认领→维护供应商', steps:[['1','将零件"是否联动调价/市况补偿"设为"是"','设置成功','',''],['2','选择非大宗料市况类型，保存并生效','零件价格联动维护页面自动生成一条联动数据','',''],['3','检查联动数据','材料供应商信息复制零件供应商信息，状态为"生效"','','']]},
    {id:'PRICE-LINK-006', name:'零件联动调价改为"否"后联动失效', pre:'1.同上\n2.已有联动数据', desc:'验证零件联动调价由"是"改为"否"后，联动数据不再参与计价', scope:'配置管理→零件认领', steps:[['1','将零件"是否联动调价/市况补偿"由"是"改为"否"','修改成功','',''],['2','保存并生效','历史联动数据保留，该零件不再进入市况联动定价单可联动范围','','']]},
    // 2.2 询价单配置
    {id:'INQ-001', name:'询价单价格属性选择"联动定价"', pre:'1.已登录GPM系统（采购员账号）\n2.询价管理→询价单有权限', desc:'验证询价单价格属性新增"联动定价"选项', scope:'询价管理→询价单→新增/编辑', steps:[['1','创建询价单，选择供应商、零件','询价单创建成功','',''],['2','在价格属性下拉中选择"联动定价"','下拉框包含"联动定价"选项','',''],['3','保存','选择"联动定价"后可正常保存','','']]},
    {id:'INQ-002', name:'询价单配置"是否允许修改重量"', pre:'1.同上\n2.零件已配置市况联动', desc:'验证采购员可控制供应商是否允许修改投料重量和废料重量', scope:'询价管理→询价单→新增/编辑', steps:[['1','创建询价单，选择联动定价零件','创建成功','',''],['2','"是否允许修改重量"选"否"，发起询价','供应商报价单中投料重量和废料重量字段置灰不可编辑','','']]},
    {id:'INQ-003', name:'询价单配置"是否允许修改市况材料单价"', pre:'1.同上', desc:'验证系统自动带价场景下，采购员可控制材料单价是否可修改', scope:'询价管理→询价单→新增/编辑', steps:[['1','创建询价单，"是否允许修改市况材料单价"选"否"','供应商报价单中材料单价字段置灰不可编辑','',''],['2','发起询价','供应商端确认字段状态','','']]},
    {id:'INQ-004', name:'询价单配置"是否允许修改其他信息"', pre:'1.同上', desc:'验证采购员可控制废料单价、合格率、加工费是否可修改', scope:'询价管理→询价单→新增/编辑', steps:[['1','创建询价单，"是否允许修改其他信息"选"否"','供应商报价单中废料单价、合格率、加工费字段置灰不可编辑','',''],['2','发起询价','供应商端确认字段状态','','']]},
    {id:'INQ-005', name:'历史报价默认值带入', pre:'1.同上\n2.该零件存在供应商已审核通过的历史报价', desc:'验证询价单发起时自动带入历史报价信息', scope:'询价管理→询价单→新增', steps:[['1','选择有历史报价的零件和供应商','零件和供应商选择成功','',''],['2','创建询价单并发起','供应商报价单自动带入上一次通过报价的全部信息（材料单价、投料重量、合格率等）','','']]},
    // 2.3 供应商填写报价单
    {id:'QUOTE-001', name:'供应商查看待报价单据', pre:'1.已登录GPM系统（供应商账号）\n2.采购员已发起询价', desc:'验证供应商可查看待报价单据并进入报价', scope:'询价协同→报价单', steps:[['1','进入报价单列表','显示待报价单据列表','',''],['2','找到状态为"待报价"的单据，点击"报价"','进入报价编辑页面，显示零件报价信息','','']]},
    {id:'QUOTE-002', name:'供应商查看权限控制字段', pre:'1.同上\n2.采购员配置了不允许修改重量和单价', desc:'验证供应商页面中被禁止修改的字段置灰', scope:'询价协同→报价单→编辑', steps:[['1','进入报价编辑页面','正常打开编辑页面','',''],['2','查看材料单价字段','材料单价置灰不可编辑','',''],['3','查看投料重量字段','投料重量置灰不可编辑，允许修改的字段为白色可编辑','','']]},
    {id:'QUOTE-003', name:'供应商填写市况基准日期', pre:'1.同上\n2.采购员允许修改市况信息', desc:'验证供应商可选择市况基准日期区间', scope:'询价协同→报价单→编辑', steps:[['1','点击市况基准开始日期，选择日期','日期选择器正常弹出','',''],['2','点击市况基准结束日期，选择日期','日期区间与联动频次匹配','',''],['3','查看联动频次','系统自动回填联动频次','','']]},
    {id:'QUOTE-004', name:'供应商填写市况基准行情', pre:'1.同上', desc:'验证供应商可填写市况基准行情价格', scope:'询价协同→报价单→编辑', steps:[['1','在市况基准行情字段输入价格','数值输入正常','',''],['2','点击保存','保存后数据不丢失','','']]},
    {id:'QUOTE-005', name:'供应商提交报价', pre:'1.同上\n2.所有必填字段已填写', desc:'验证供应商可正常提交报价', scope:'询价协同→报价单→编辑', steps:[['1','检查所有必填字段','所有必填字段已填写','',''],['2','点击"提交"','提交成功，报价单状态更新为"已报价"','','']]},
    {id:'QUOTE-006', name:'提交时必填字段校验', pre:'1.同上\n2.存在未填写的必填字段', desc:'验证提交时校验必填字段', scope:'询价协同→报价单→编辑', steps:[['1','不填写投料重量','投料重量为空','',''],['2','点击"提交"','提交失败，提示"请填写必填项"，必填字段标红','','']]},
    // 2.4 市况联动定价单
    {id:'LINK-001', name:'新增市况联动定价单', pre:'1.已登录GPM系统（采购员账号）\n2.市况管理→联动定价→市况联动定价单有权限', desc:'验证采购员可创建市况联动定价单', scope:'市况管理→联动定价→市况联动定价单', steps:[['1','点击"新增"','打开新增页面','',''],['2','选择供应商','供应商选择成功','',''],['3','选择定价类型（正式价/暂估价）、币种、税率','字段选择正常','',''],['4','点击"保存"','定价单创建成功，状态为"拟定"，自动带出该供应商名下全部零件','','']]},
    {id:'LINK-002', name:'选择供应商后自动带出零件', pre:'1.同上', desc:'验证选择供应商后自动带出已报价和未报价零件', scope:'市况管理→联动定价→新增', steps:[['1','选择供应商','供应商选择成功','',''],['2','查看定价零件列表','已报价零件带出报价明细，删除按钮置灰','',''],['3','查看未报价零件','未报价零件费用字段为空，允许删除','',''],['4','查看页面提示','提示"已带出该供应商的全部零件信息"','','']]},
    {id:'LINK-003', name:'同步市况&计算', pre:'1.同上\n2.定价单中有零件', desc:'验证点击"同步市况&计算"后系统执行联动计算', scope:'市况管理→联动定价→编辑', steps:[['1','点击"同步市况&计算"','系统开始计算','',''],['2','等待计算完成','系统按公式配置计算材料费和本次价格','',''],['3','检查定价零件列表和市况材料明细','定价零件列表更新价格，市况材料明细更新','','']]},
    {id:'LINK-004', name:'行情来源切换', pre:'1.同上\n2.已同步市况计算', desc:'验证定价零件行情来源可切换', scope:'市况管理→联动定价→编辑', steps:[['1','在定价零件行点击行情来源下拉','下拉框正常展开','',''],['2','切换为"供应商"','行情来源切换成功，按供应商行情取值','','']]},
    {id:'LINK-005', name:'导入年降及返利', pre:'1.同上\n2.已同步市况计算', desc:'验证年降比例和返利金额可导入', scope:'市况管理→联动定价→编辑', steps:[['1','点击"导入"，下载模板','模板下载成功','',''],['2','填写年降比例和返利金额','填写正常','',''],['3','上传模板','年降比例导入成功，年降金额自动计算（=上次价格×年降比例），返利金额导入成功','','']]},
    {id:'LINK-006', name:'发起价格审批-校验通过', pre:'1.同上\n2.供应商、定价类型、币种、税率已填\n3.已报价零件已完成同步计算', desc:'验证满足条件时可发起价格审批', scope:'市况管理→联动定价→编辑', steps:[['1','确认所有已报价零件已计算','全部已计算','',''],['2','点击"发起价格审批"','系统创建拟定状态的价格审批单','',''],['3','检查审批单','定价单状态变为"已生成价格审批单"，价格审批单零件行数据正确','','']]},
    {id:'LINK-007', name:'发起价格审批-校验失败', pre:'1.同上\n2.有零件未完成同步计算', desc:'验证未完成计算时不允许提交审批', scope:'市况管理→联动定价→编辑', steps:[['1','不执行"同步市况&计算"','有零件未计算','',''],['2','直接点击"发起价格审批"','提交失败，提示"请完成市况同步及价格计算后再提交价格审批"','','']]},
    // 2.5 价格审批与联动结果
    {id:'APPR-001', name:'审批通过后写入价格库', pre:'1.已登录GPM系统（审批人员账号）\n2.已生成联动定价审批单\n3.价格审批单处于审批流程中', desc:'验证审批通过后联动价格写入价格库', scope:'价格管理→价格审批单→审批', steps:[['1','核对定价零件及价格明细','数据正确','',''],['2','审批通过','本次联动定价结果写入价格库','',''],['3','检查价格目录','零件价格目录更新为审批通过价格，市况变动按规则统计','','']]},
    {id:'APPR-002', name:'报价单终审后自动生成投料数据', pre:'1.同上\n2.报价单终审完成且已生成价格目录', desc:'验证报价终审后自动生成零件原材料投料重量数据', scope:'市况管理→市况基础信息→零件原材料投料重量维护', steps:[['1','查询对应零件','查询成功','',''],['2','查看投料重量数据','自动生成新的生效数据，原有生效数据变为失效','',''],['3','检查生效时间','生效时间为本次报价终审完成时间','','']]},
    {id:'APPR-003', name:'报价单作废后投料数据作废', pre:'1.同上\n2.已有自动生成的投料数据', desc:'验证报价单作废时对应投料数据自动作废', scope:'询价管理→报价单→作废', steps:[['1','作废已终审的报价单','作废成功','',''],['2','查看投料重量维护页面','对应批次投料数据自动作废','','']]},
];

function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * 替换单元格 <w:t> 文本（处理 \n 换行）
 */
function replaceCellText(cellXml, newText) {
    const wtRegex = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g;
    const wts = [...cellXml.matchAll(wtRegex)];
    if (wts.length === 0) return cellXml;
    
    const first = wts[0];
    const lines = newText.split('\n');
    const escapedLines = lines.map(escapeXml);
    
    // 如果只有一行，直接替换第一个 <w:t>
    if (lines.length === 1) {
        let newCell = cellXml.substring(0, first.index) +
            first[1] + escapedLines[0] + first[3] +
            cellXml.substring(first.index + first[0].length);
        for (let j = 1; j < wts.length; j++) {
            newCell = newCell.replace(wts[j][0], wts[j][1] + '' + wts[j][3]);
        }
        return newCell;
    }
    
    // 多行：用 <w:br/> 分隔
    const joinedText = escapedLines.join('</w:t></w:r><w:r><w:rPr></w:rPr><w:t xml:space="preserve">');
    let newCell = cellXml.substring(0, first.index) +
        first[1] + joinedText + first[3] +
        cellXml.substring(first.index + first[0].length);
    for (let j = 1; j < wts.length; j++) {
        newCell = newCell.replace(wts[j][0], wts[j][1] + '' + wts[j][3]);
    }
    return newCell;
}

/**
 * 替换一行中指定单元格的文本
 * cellsToReplace: [{colIndex: text}, ...]
 */
function replaceRow(rowXml, cellsToReplace) {
    const tcRegex = /<w:tc[\s\S]*?<\/w:tc>/g;
    let result = '';
    let lastIdx = 0;
    let colIdx = 0;
    let tcMatch;
    
    while ((tcMatch = tcRegex.exec(rowXml)) !== null) {
        result += rowXml.substring(lastIdx, tcMatch.index);
        let cellXml = tcMatch[1];
        
        const replacement = cellsToReplace.find(c => c.colIdx === colIdx);
        if (replacement) {
            cellXml = replaceCellText(cellXml, replacement.text);
        }
        
        result += cellXml;
        lastIdx = tcMatch.index + tcMatch[0].length;
        colIdx++;
    }
    result += rowXml.substring(lastIdx);
    return result;
}

function buildDocx() {
    const template = new AdmZip(TEMPLATE);
    let xml = template.readAsText('word/document.xml', 'utf8');
    
    // ===== 1. 替换封面/前言文本 =====
    const textMap = {
        '广汽集团零采一体化项目': '广汽集团零采一体化项目',
        '市况联动相关': '市况联动——供应商报价 & 联动定价',
        '指导测试工作有序进行，使实施测试的有据可依；确保所实现的功能与客户预期的需求相符合。':
            '验证市况联动板块优化 v4.1 中供应商报价流程和市况联动定价功能的正确性和完整性，确保各角色操作闭环通过。',
        '在测试过程中系统每个模块可以先调用通用测试用例（新增、查询、编辑、删除、导入、导出、分页）执行测试过程，然后根据功能测试用例执行业务流程测试。':
            '黑盒测试，按用例逐步操作，验证预期结果与实际结果一致。',
        '零件认领': '零件价格联动维护',
        '新增"是否联动调价/市况补偿"字段': 'PP/钢板/铝板手工新增联动关系',
        '市况物料维护': '询价单配置（采购员发起）',
        '新增【市况物料维护】页面': '询价单价格属性选择"联动定价"',
        '原材料行情维护': '供应商填写报价单',
        '原材料行情维护移除工厂维度': '供应商查看待报价单据',
        '询价单': '市况联动定价单',
        '询价单新增控制字段': '新增市况联动定价单',
        '报价单': '价格审批与联动结果',
        '新增生管字段展示': '审批通过后写入价格库',
    };
    
    const wtRegex = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g;
    let newXml = '';
    let lastIdx = 0;
    let wtMatch;
    while ((wtMatch = wtRegex.exec(xml)) !== null) {
        newXml += xml.substring(lastIdx, wtMatch.index);
        const oldText = wtMatch[2];
        if (textMap[oldText] !== undefined) {
            newXml += wtMatch[1] + escapeXml(textMap[oldText]) + wtMatch[3];
        } else {
            newXml += wtMatch[0];
        }
        lastIdx = wtMatch.index + wtMatch[0].length;
    }
    newXml += xml.substring(lastIdx);
    xml = newXml;
    
    // ===== 2. 替换数据准备表格（表1） =====
    const tblRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
    let tblMatches = [...xml.matchAll(tblRegex)];
    
    const DATA_PREP = [
        ['1', '谷歌浏览器及用户账号', '登录 GPM 系统'],
        ['2', '市况材料主数据（铜、铝、PP等）', '已配置市况材料主数据'],
        ['3', '原材料行情数据', '已导入百川资讯/供应商来源行情'],
        ['4', '市况类型与材料关系', '已维护并生效'],
        ['5', '市况联动公式', '已配置材料单价要素+行情来源'],
        ['6', '零件认领', '已配置"是否联动调价/市况补偿"为"是"'],
        ['7', '零件价格联动维护', '已生效'],
        ['8', '测试用供应商账号', '可正常登录'],
    ];
    
    if (tblMatches.length >= 2) {
        const tblXml = tblMatches[1][0];
        const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
        const trMatches = [...tblXml.matchAll(trRegex)];
        
        if (trMatches.length >= 2) {
            const headerRow = trMatches[0][0];
            const sampleRow = trMatches[1][0];
            let newRows = '';
            for (const row of DATA_PREP) {
                newRows += replaceRow(sampleRow, [
                    {colIdx: 0, text: row[0]},
                    {colIdx: 1, text: row[1]},
                    {colIdx: 2, text: row[2]},
                ]);
            }
            const newTbl = tblXml.substring(0, trMatches[1].index) + newRows + tblXml.substring(trMatches[trMatches.length - 1].index + trMatches[trMatches.length - 1][0].length);
            xml = xml.substring(0, tblMatches[1].index) + newTbl + xml.substring(tblMatches[1].index + tblMatches[1][0].length);
        }
    }
    
    // ===== 3. 替换测试用例表格 =====
    // 重新解析（因为 xml 变了）
    tblMatches = [...xml.matchAll(tblRegex)];
    
    // 获取模板：第一个测试用例表（表2）
    if (tblMatches.length >= 3) {
        const templateTbl = tblMatches[2][0];
        const trRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
        const templateRows = [...templateTbl.matchAll(trRegex)];
        
        // Row 0-3: 2列 (标签+值)
        // Row 4: 5列表头
        // Row 5+: 5列步骤行
        const labelRow0 = templateRows[0][0]; // 用例编号行
        const labelRow1 = templateRows[1][0]; // 用例名称行
        const labelRow2 = templateRows[2][0]; // 前置条件行
        const labelRow3 = templateRows[3][0]; // 用例描述行
        const stepHeaderRow = templateRows[4][0]; // 步骤表头
        const stepRow = templateRows[5] ? templateRows[5][0] : templateRows[4][0]; // 步骤数据行
        
        // 生成所有测试用例表
        let allNewTables = '';
        for (const tc of ALL_TEST_CASES) {
            let tbl = '';
            // Row 0: 用例编号
            tbl += replaceRow(labelRow0, [{colIdx: 1, text: tc.id}]);
            // Row 1: 用例名称
            tbl += replaceRow(labelRow1, [{colIdx: 1, text: tc.name}]);
            // Row 2: 前置条件
            tbl += replaceRow(labelRow2, [{colIdx: 1, text: tc.pre}]);
            // Row 3: 用例描述
            tbl += replaceRow(labelRow3, [{colIdx: 1, text: tc.desc}]);
            // Row 4: 步骤表头（保持不变）
            tbl += stepHeaderRow;
            // Row 5+: 步骤行
            for (const step of tc.steps) {
                tbl += replaceRow(stepRow, [
                    {colIdx: 0, text: step[0]},
                    {colIdx: 1, text: step[1]},
                    {colIdx: 2, text: step[2]},
                    {colIdx: 3, text: step[3]},
                    {colIdx: 4, text: step[4]},
                ]);
            }
            // 组装完整表格：表属性 + 表网格 + 行
            const tblProps = templateTbl.substring(0, templateTbl.indexOf('<w:tr>'));
            allNewTables += '<w:tbl>' + tblProps + tbl + '</w:tbl>';
        }
        
        // 替换：表2 到 表23 是原测试用例区域
        const tcStart = tblMatches[2].index;
        const tcEnd = tblMatches[tblMatches.length - 1].index + tblMatches[tblMatches.length - 1][0].length;
        xml = xml.substring(0, tcStart) + allNewTables + xml.substring(tcEnd);
    }
    
    // ===== 写入文件 =====
    const zip = new AdmZip();
    template.getEntries().forEach(entry => {
        if (entry.entryName === 'word/document.xml') {
            zip.addFile('word/document.xml', Buffer.from(xml, 'utf8'));
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
