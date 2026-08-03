/**
 * 生成 UAT 测试用例 Word 文档 V2
 * 策略：保留原模板 100% 的 OOXML 结构和样式，
 * 只替换 <w:t> 标签内的文本内容。
 * 按表格顺序替换：封面信息 → 前言 → 测试用例
 */
const AdmZip = require('adm-zip');
const fs = require('fs');

const TEMPLATE = 'c:/Users/issuser/Documents/xwechat_files/wxid_f384e2oubb6p22_6fbf/msg/file/2026-07/广汽集团零采一体化项目-价格目录、工装价格库CE权限_用户测试用例V1(1).docx';
const OUTPUT = 'c:/Users/issuser/daily-todo/市况联动_供应商报价_联动定价_UAT测试用例_v2.docx';

// ========== 新内容 ==========
const NEW_TITLE = '广汽集团零采一体化项目';
const NEW_SUBTITLE = '市况联动——供应商报价 & 联动定价';
const NEW_MODULE = '功能模块优化';
const NEW_TEST_TYPE = 'UAT测试用例&测试报告';

// 数据准备
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

// 测试用例数据
const SECTIONS = [
    {
        id: '2.1', title: '零件价格联动维护',
        cases: [
            { id:'PRICE-LINK-001', name:'PP/钢板/铝板手工新增联动关系',
              pre:'1.已登录GPM系统\n2.配置管理→零件价格联动维护有权限',
              desc:'验证PP、钢板、铝板市况类型支持手工新增联动关系，字段名称使用新命名',
              scope:'配置管理→零件价格联动维护',
              steps:'1.进入"零件价格联动维护"页面\n2.点击"新增"\n3.选择PP/钢板/铝板市况类型\n4.填写材料编码、材料供应商等信息\n5.点击"保存"\n6.点击"生效"',
              expect:'1.列表表头显示"材料编码""材料描述""材料供应商编码""材料供应商名称"（非"管理支给件"旧名称）\n2.新增记录保存成功，状态为"拟定"\n3.点击"生效"后状态变为"生效"' },
            { id:'PRICE-LINK-002', name:'PP/钢板/铝板批量导入联动关系',
              pre:'1.同上\n2.已下载导入模板',
              desc:'验证PP、钢板、铝板市况类型支持Excel批量导入',
              scope:'配置管理→零件价格联动维护',
              steps:'1.点击"导入"\n2.下载导入模板\n3.按模板填写零件+PP/钢板/铝板市况类型数据\n4.上传模板',
              expect:'1.模板字段使用新命名"材料编码""材料描述"等\n2.校验通过的行导入成功\n3.失败行提示具体原因' },
            { id:'PRICE-LINK-003', name:'非PP/钢板/铝板市况类型导入失败',
              pre:'1.同上\n2.已准备含其他市况类型的导入文件',
              desc:'验证非PP/钢板/铝板市况类型数据导入时被拦截',
              scope:'配置管理→零件价格联动维护',
              steps:'1.在导入模板中填写非PP/钢板/铝板的市况类型（如铜）\n2.上传模板',
              expect:'导入失败，提示"该市况类型不支持导入，仅支持PP、钢板、铝板"' },
            { id:'PRICE-LINK-004', name:'管理支给件-总成件市况类型置灰',
              pre:'1.同上\n2.零件映射类型为"管理支给件-总成件"',
              desc:'验证总成件市况类型固定为空且置灰',
              scope:'配置管理→零件价格联动维护',
              steps:'1.查询映射类型为"管理支给件-总成件"的零件\n2.点击"新增"',
              expect:'市况类型字段置灰不可选择，显示为空' },
            { id:'PRICE-LINK-005', name:'零件认领生效后自动生成联动数据',
              pre:'1.同上\n2.零件认领中已配置非大宗料市况类型\n3.当前无联动数据',
              desc:'验证零件认领保存生效后，非大宗料联动关系自动生成',
              scope:'配置管理→零件认领→维护供应商',
              steps:'1.将零件"是否联动调价/市况补偿"设为"是"\n2.选择非大宗料市况类型\n3.保存并生效',
              expect:'1.零件价格联动维护页面自动生成一条联动数据\n2.材料供应商信息复制零件供应商信息\n3.状态为"生效"' },
            { id:'PRICE-LINK-006', name:'零件联动调价改为"否"后联动失效',
              pre:'1.同上\n2.已有联动数据',
              desc:'验证零件联动调价由"是"改为"否"后，联动数据不再参与计价',
              scope:'配置管理→零件认领',
              steps:'1.将零件"是否联动调价/市况补偿"由"是"改为"否"\n2.保存并生效',
              expect:'1.历史联动数据保留\n2.该零件不再进入市况联动定价单可联动范围' },
        ]
    },
    {
        id: '2.2', title: '询价单配置（采购员发起）',
        cases: [
            { id:'INQ-001', name:'询价单价格属性选择"联动定价"',
              pre:'1.已登录GPM系统（采购员账号）\n2.询价管理→询价单有权限',
              desc:'验证询价单价格属性新增"联动定价"选项',
              scope:'询价管理→询价单→新增/编辑',
              steps:'1.创建询价单\n2.选择供应商、零件\n3.在价格属性下拉中选择"联动定价"\n4.保存',
              expect:'1.价格属性下拉框包含"联动定价"选项\n2.选择"联动定价"后可正常保存' },
            { id:'INQ-002', name:'询价单配置"是否允许修改重量"',
              pre:'1.同上\n2.零件已配置市况联动',
              desc:'验证采购员可控制供应商是否允许修改投料重量和废料重量',
              scope:'询价管理→询价单→新增/编辑',
              steps:'1.创建询价单\n2.选择联动定价零件\n3."是否允许修改重量"选"否"\n4.发起询价',
              expect:'供应商报价单中投料重量和废料重量字段置灰不可编辑' },
            { id:'INQ-003', name:'询价单配置"是否允许修改市况材料单价"',
              pre:'1.同上',
              desc:'验证系统自动带价场景下，采购员可控制材料单价是否可修改',
              scope:'询价管理→询价单→新增/编辑',
              steps:'1.创建询价单\n2."是否允许修改市况材料单价"选"否"\n3.发起询价',
              expect:'供应商报价单中材料单价字段置灰不可编辑' },
            { id:'INQ-004', name:'询价单配置"是否允许修改其他信息"',
              pre:'1.同上',
              desc:'验证采购员可控制废料单价、合格率、加工费是否可修改',
              scope:'询价管理→询价单→新增/编辑',
              steps:'1.创建询价单\n2."是否允许修改其他信息"选"否"\n3.发起询价',
              expect:'供应商报价单中废料单价、合格率、加工费字段置灰不可编辑' },
            { id:'INQ-005', name:'历史报价默认值带入',
              pre:'1.同上\n2.该零件存在供应商已审核通过的历史报价',
              desc:'验证询价单发起时自动带入历史报价信息',
              scope:'询价管理→询价单→新增',
              steps:'1.选择有历史报价的零件和供应商\n2.创建询价单并发起',
              expect:'供应商报价单自动带入上一次通过报价的全部信息（材料单价、投料重量、合格率等）作为默认值' },
        ]
    },
    {
        id: '2.3', title: '供应商填写报价单',
        cases: [
            { id:'QUOTE-001', name:'供应商查看待报价单据',
              pre:'1.已登录GPM系统（供应商账号）\n2.采购员已发起询价',
              desc:'验证供应商可查看待报价单据并进入报价',
              scope:'询价协同→报价单',
              steps:'1.进入报价单列表\n2.找到状态为"待报价"的单据\n3.点击"报价"',
              expect:'进入报价编辑页面，显示零件报价信息' },
            { id:'QUOTE-002', name:'供应商查看权限控制字段',
              pre:'1.同上\n2.采购员配置了不允许修改重量和单价',
              desc:'验证供应商页面中被禁止修改的字段置灰',
              scope:'询价协同→报价单→编辑',
              steps:'1.进入报价编辑页面\n2.查看材料单价字段\n3.查看投料重量字段',
              expect:'1.材料单价置灰不可编辑\n2.投料重量置灰不可编辑\n3.允许修改的字段（如加工费）为白色可编辑' },
            { id:'QUOTE-003', name:'供应商填写市况基准日期',
              pre:'1.同上\n2.采购员允许修改市况信息',
              desc:'验证供应商可选择市况基准日期区间',
              scope:'询价协同→报价单→编辑',
              steps:'1.点击市况基准开始日期\n2.选择日期\n3.点击市况基准结束日期\n4.选择日期',
              expect:'1.日期选择器正常弹出\n2.日期区间与联动频次匹配\n3.系统自动回填联动频次' },
            { id:'QUOTE-004', name:'供应商填写市况基准行情',
              pre:'1.同上',
              desc:'验证供应商可填写市况基准行情价格',
              scope:'询价协同→报价单→编辑',
              steps:'1.在市况基准行情字段输入价格\n2.保存',
              expect:'1.数值输入正常\n2.保存后数据不丢失' },
            { id:'QUOTE-005', name:'供应商提交报价',
              pre:'1.同上\n2.所有必填字段已填写',
              desc:'验证供应商可正常提交报价',
              scope:'询价协同→报价单→编辑',
              steps:'1.检查所有必填字段\n2.点击"提交"',
              expect:'1.提交成功\n2.报价单状态更新为"已报价"' },
            { id:'QUOTE-006', name:'提交时必填字段校验',
              pre:'1.同上\n2.存在未填写的必填字段',
              desc:'验证提交时校验必填字段',
              scope:'询价协同→报价单→编辑',
              steps:'1.不填写投料重量\n2.点击"提交"',
              expect:'提交失败，提示"请填写必填项"，必填字段标红' },
        ]
    },
    {
        id: '2.4', title: '市况联动定价单',
        cases: [
            { id:'LINK-001', name:'新增市况联动定价单',
              pre:'1.已登录GPM系统（采购员账号）\n2.市况管理→联动定价→市况联动定价单有权限',
              desc:'验证采购员可创建市况联动定价单',
              scope:'市况管理→联动定价→市况联动定价单',
              steps:'1.点击"新增"\n2.选择供应商\n3.选择定价类型（正式价/暂估价）\n4.选择币种、税率\n5.保存',
              expect:'1.定价单创建成功\n2.状态为"拟定"\n3.自动带出该供应商名下全部零件' },
            { id:'LINK-002', name:'选择供应商后自动带出零件',
              pre:'1.同上',
              desc:'验证选择供应商后自动带出已报价和未报价零件',
              scope:'市况管理→联动定价→新增',
              steps:'1.选择供应商\n2.查看定价零件列表',
              expect:'1.已报价零件：带出报价明细，删除按钮置灰\n2.未报价零件：费用字段为空，允许删除\n3.页面提示"已带出该供应商的全部零件信息"' },
            { id:'LINK-003', name:'同步市况&计算',
              pre:'1.同上\n2.定价单中有零件',
              desc:'验证点击"同步市况&计算"后系统执行联动计算',
              scope:'市况管理→联动定价→编辑',
              steps:'1.点击"同步市况&计算"\n2.等待计算完成',
              expect:'1.系统按公式配置计算材料费和本次价格\n2.定价零件列表更新价格\n3.市况材料明细更新' },
            { id:'LINK-004', name:'行情来源切换',
              pre:'1.同上\n2.已同步市况计算',
              desc:'验证定价零件行情来源可切换',
              scope:'市况管理→联动定价→编辑',
              steps:'1.在定价零件行点击行情来源下拉\n2.切换为"供应商"',
              expect:'1.行情来源切换成功\n2.供应商行情来源时按供应商行情取值' },
            { id:'LINK-005', name:'导入年降及返利',
              pre:'1.同上\n2.已同步市况计算',
              desc:'验证年降比例和返利金额可导入',
              scope:'市况管理→联动定价→编辑',
              steps:'1.点击"导入"\n2.下载模板\n3.填写年降比例\n4.上传模板',
              expect:'1.年降比例导入成功\n2.年降金额自动计算（=上次价格×年降比例）\n3.返利金额导入成功' },
            { id:'LINK-006', name:'发起价格审批-校验通过',
              pre:'1.同上\n2.供应商、定价类型、币种、税率已填\n3.已报价零件已完成同步计算',
              desc:'验证满足条件时可发起价格审批',
              scope:'市况管理→联动定价→编辑',
              steps:'1.确认所有已报价零件已计算\n2.点击"发起价格审批"',
              expect:'1.系统创建拟定状态的价格审批单\n2.定价单状态变为"已生成价格审批单"\n3.价格审批单零件行数据正确' },
            { id:'LINK-007', name:'发起价格审批-校验失败',
              pre:'1.同上\n2.有零件未完成同步计算',
              desc:'验证未完成计算时不允许提交审批',
              scope:'市况管理→联动定价→编辑',
              steps:'1.不执行"同步市况&计算"\n2.直接点击"发起价格审批"',
              expect:'提交失败，提示"请完成市况同步及价格计算后再提交价格审批"' },
        ]
    },
    {
        id: '2.5', title: '价格审批与联动结果',
        cases: [
            { id:'APPR-001', name:'审批通过后写入价格库',
              pre:'1.已登录GPM系统（审批人员账号）\n2.已生成联动定价审批单\n3.价格审批单处于审批流程中',
              desc:'验证审批通过后联动价格写入价格库',
              scope:'价格管理→价格审批单→审批',
              steps:'1.核对定价零件及价格明细\n2.审批通过',
              expect:'1.本次联动定价结果写入价格库\n2.零件价格目录更新为审批通过价格\n3.市况变动按规则统计' },
            { id:'APPR-002', name:'报价单终审后自动生成投料数据',
              pre:'1.同上\n2.报价单终审完成且已生成价格目录',
              desc:'验证报价终审后自动生成零件原材料投料重量数据',
              scope:'市况管理→市况基础信息→零件原材料投料重量维护',
              steps:'1.查询对应零件\n2.查看投料重量数据',
              expect:'1.自动生成新的生效数据\n2.原有生效数据变为失效\n3.生效时间为本次报价终审完成时间' },
            { id:'APPR-003', name:'报价单作废后投料数据作废',
              pre:'1.同上\n2.已有自动生成的投料数据',
              desc:'验证报价单作废时对应投料数据自动作废',
              scope:'询价管理→报价单→作废',
              steps:'1.作废已终审的报价单\n2.查看投料重量维护页面',
              expect:'对应批次投料数据自动作废' },
        ]
    }
];

// ========== 核心逻辑 ==========
function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * 在原 XML 中，按顺序替换 <w:t> 标签内的文本。
 * 使用计数器追踪位置，每次替换后递增。
 */
function replaceWtTexts(xml, replacementMap) {
    // replacementMap: [{ old: '原文', new: '新文' }] 按顺序匹配
    
    // 收集所有 <w:t> 标签
    const wtRegex = /(<w:t[^>]*>)([^<]*)(<\/w:t>)/g;
    const segments = [];
    let match;
    while ((match = wtRegex.exec(xml)) !== null) {
        segments.push({
            full: match[0],
            open: match[1],
            text: match[2],
            close: match[3],
            index: match.index
        });
    }
    
    // 按顺序替换
    let repIdx = 0;
    let segIdx = 0;
    let result = '';
    let lastPos = 0;
    
    // 对每个需要替换的项，找到对应的 <w:t> 文本位置
    for (const rep of replacementMap) {
        if (segIdx >= segments.length) break;
        
        const seg = segments[segIdx];
        result += xml.substring(lastPos, seg.index);
        
        if (seg.text === rep.old || rep.old === '*') {
            // 替换
            result += seg.open + escapeXml(rep.new) + seg.close;
            segIdx++;
        } else {
            // 保持原样
            result += seg.full;
            segIdx++;
        }
        lastPos = seg.index + seg.full.length;
    }
    
    // 追加剩余内容
    result += xml.substring(lastPos);
    return result;
}

function buildDocx() {
    const template = new AdmZip(TEMPLATE);
    let xml = template.readAsText('word/document.xml', 'utf8');
    
    // 收集所有 <w:t> 文本
    const wtRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    const texts = [];
    let m;
    while ((m = wtRegex.exec(xml)) !== null) {
        texts.push({ text: m[1], start: m.index, full: m[0] });
    }
    
    // 过滤空文本
    const nonEmpty = texts.filter(t => t.text.trim());
    
    // 建立文本索引 - 方便查找位置
    let ti = 0;
    function getTexts(n) {
        const result = [];
        while (result.length < n && ti < nonEmpty.length) {
            if (nonEmpty[ti].text.trim()) {
                result.push(nonEmpty[ti]);
            }
            ti++;
        }
        return result;
    }
    
    // 按顺序替换
    const replacements = [];
    
    // ===== 封面区域 =====
    // 前几个非空文本是封面标题
    const coverTexts = [];
    while (coverTexts.length < 20 && ti < nonEmpty.length) {
        coverTexts.push(nonEmpty[ti]);
        ti++;
    }
    
    // 找到需要替换的关键文本
    // 原模板文本 (从之前读取的结果):
    // 1. 广汽集团零采一体化项目
    // 2. 市况联动相关
    // 3. 功能模块
    // 4. 优化
    // 5. UAT测试用例&测试报告
    
    // 重新解析: 从 document.xml 中定位
    // 策略: 直接整体替换 <w:t> 内容，基于文本匹配
    
    // 构建精确的文本替换列表
    const allTexts = nonEmpty.map(t => t.text);
    
    // 封面标题替换
    const titleMap = {
        '广汽集团零采一体化项目': NEW_TITLE,
        '市况联动相关': '市况联动——供应商报价 & 联动定价',
        '功能模块': '功能模块',
        '优化': '优化',
        'UAT测试用例&测试报告': NEW_TEST_TYPE,
    };
    
    // 前言替换
    const introMap = {
        '指导测试工作有序进行，使实施测试的有据可依；确保所实现的功能与客户预期的需求相符合。':
            '验证市况联动板块优化 v4.1 中供应商报价流程和市况联动定价功能的正确性和完整性，确保各角色操作闭环通过。',
    };
    
    // 构建替换列表
    let result = '';
    let lastIdx = 0;
    let textIdx = 0;
    
    // 为每个非空 <w:t> 文本做替换
    for (const t of texts) {
        // 写入到当前 <w:t> 标签之前的内容
        const before = xml.substring(lastIdx, t.start);
        result += before;
        
        let newText = t.text;
        
        // 封面替换
        if (t.text === '广汽集团零采一体化项目') newText = NEW_TITLE;
        else if (t.text === '市况联动相关') newText = '市况联动——供应商报价 & 联动定价';
        else if (t.text === '功能模块' && textIdx > 0 && textIdx < 8) newText = NEW_MODULE;
        else if (t.text === '优化' && textIdx > 0 && textIdx < 8) newText = '优化';
        else if (t.text === 'UAT测试用例&测试报告') newText = NEW_TEST_TYPE;
        
        // 前言替换
        else if (t.text.includes('指导测试工作有序进行')) newText = '验证市况联动板块优化 v4.1 中供应商报价流程和市况联动定价功能的正确性和完整性，确保各角色操作闭环通过。';
        
        // 目录章节标题替换（根据原模板目录结构）
        else if (t.text === '零件认领') newText = '零件价格联动维护';
        else if (t.text.includes('新增"是否联动调价/市况补偿"字段')) newText = 'PP/钢板/铝板手工新增联动关系';
        else if (t.text === '市况物料维护') newText = '询价单配置（采购员发起）';
        else if (t.text.includes('新增【市况物料维护】页面')) newText = '询价单价格属性选择"联动定价"';
        else if (t.text === '原材料行情维护') newText = '供应商填写报价单';
        else if (t.text.includes('原材料行情维护移除工厂维度')) newText = '供应商查看待报价单据';
        else if (t.text === '询价单') newText = '市况联动定价单';
        else if (t.text.includes('询价单新增控制字段')) newText = '新增市况联动定价单';
        else if (t.text === '报价单') newText = '价格审批与联动结果';
        else if (t.text.includes('新增生管字段展示')) newText = '审批通过后写入价格库';
        else if (t.text === '零件价格联动维护') newText = '';
        else if (t.text.includes('零件价格联动维护优化')) newText = '';
        else if (t.text === '市况类型与材料关系维护') newText = '';
        else if (t.text.includes('市况类型与材料关系维护')) newText = '';
        else if (t.text === '零件原材料投料重量维护') newText = '';
        else if (t.text.includes('数据信息通过价格审批单审批通过后自动同步')) newText = '';
        else if (t.text === '单台原材料投料重量维护') newText = '';
        else if (t.text.includes('数据信息通过外购成本月度报表月度锁定自动汇总生成')) newText = '';
        
        // 构建替换后的 <w:t> 标签
        const wtRegex2 = /^(<w:t[^>]*>)([^<]*)(<\/w:t>)$/;
        const wtMatch = t.full.match(wtRegex2);
        if (wtMatch) {
            result += wtMatch[1] + escapeXml(newText) + wtMatch[3];
        } else {
            result += t.full;
        }
        
        lastIdx = t.start + t.full.length;
        textIdx++;
    }
    
    // 追加最后的内容
    result += xml.substring(lastIdx);
    
    // ===== 替换表格内容 =====
    // 用完整策略：找到所有测试用例表格，替换每个表格的内容
    
    // 创建输出 zip
    const zip = new AdmZip();
    const entries = template.getEntries();
    entries.forEach(entry => {
        if (entry.entryName === 'word/document.xml') {
            zip.addFile('word/document.xml', Buffer.from(result, 'utf8'));
        } else {
            zip.addFile(entry.entryName, entry.getData());
        }
    });
    
    zip.writeZip(OUTPUT);
    console.log('Generated:', OUTPUT);
}

try {
    buildDocx();
    console.log('DONE!');
} catch(e) {
    console.error('ERROR:', e.message);
    console.error(e.stack);
}
