/**
 * 司法裁判與爭點分析 - 自動化批次執行腳本 (Batch Judgment Analyzer)
 * 
 * 最後發佈日期 (Release Date): 2026-11-22
 * 最後修改日期 (Last Modified): 2026-11-22
 * 
 * 功能:
 *   自動掃描來源目錄 (預設 judgements_org) 中的所有裁判書 (PDF/TXT)，
 *   進行要件事實、涉訟法條、核心爭點設問、法院心證金句之智慧萃取，
 *   並於目標目錄 (預設 analysis_results) 自動產出雙軌成果:
 *     1. 完整 Markdown 深度分析報告 (.md)
 *     2. 12 欄位精美 Excel 各篇分析結果摘要總表 (.xlsx)
 *     3. Excel 橫向比對矩陣 (.xlsx)
 * 
 * 用法:
 *   node scripts/batch_analyze_judgments.js
 *   node scripts/batch_analyze_judgments.js --input "judgements_org" --output "analysis_results" --topic "洗錢防制與詐欺爭點研究"
 */

const fs = require('fs');
const path = require('path');

// 模組動態安全解析（支援全域、專案根目錄與 Skill 內嵌目錄）
function resolveModule(moduleName) {
    try { return require(moduleName); } catch {}
    const candidatePaths = [
        path.resolve(__dirname, 'node_modules', moduleName),
        path.resolve(__dirname, '../node_modules', moduleName),
        path.resolve(__dirname, '../../../../scripts/node_modules', moduleName),
        path.resolve(process.cwd(), 'scripts/node_modules', moduleName),
        path.resolve(process.cwd(), 'node_modules', moduleName)
    ];
    for (const p of candidatePaths) {
        try { return require(p); } catch {}
    }
    throw new Error(`找不到套件 ${moduleName}，請先執行 npm install ${moduleName}`);
}

const ExcelJS = resolveModule('exceljs');
const { PDFParse } = resolveModule('pdf-parse');

// 動態定位專案工作區根目錄
function getWorkspaceRoot() {
    if (fs.existsSync(path.resolve(process.cwd(), 'judgements_org')) || fs.existsSync(path.resolve(process.cwd(), '.agents')) || fs.existsSync(path.resolve(process.cwd(), 'judgments_data'))) {
        return process.cwd();
    }
    let cur = __dirname;
    for (let i = 0; i < 6; i++) {
        if (fs.existsSync(path.resolve(cur, 'judgements_org')) || fs.existsSync(path.resolve(cur, '.agents')) || fs.existsSync(path.resolve(cur, 'judgments_data'))) {
            return cur;
        }
        const parent = path.dirname(cur);
        if (parent === cur) break;
        cur = parent;
    }
    return process.cwd();
}

const WORKSPACE_ROOT = getWorkspaceRoot();

// ==================== 參數配置 ====================
const CONFIG = {
    inputDir: path.resolve(WORKSPACE_ROOT, 'judgements_org'),
    outputDir: path.resolve(WORKSPACE_ROOT, 'analysis_results'),
    topic: '',
    limit: 0 // 0 為不限制
};

// 解析命令列參數
const args = process.argv.slice(2);
for (let i = 0; i < args.length; i++) {
    if (args[i] === '--input' && args[i + 1]) CONFIG.inputDir = path.resolve(args[++i]);
    if (args[i] === '--output' && args[i + 1]) CONFIG.outputDir = path.resolve(args[++i]);
    if (args[i] === '--topic' && args[i + 1]) CONFIG.topic = args[++i];
    if (args[i] === '--limit' && args[i + 1]) CONFIG.limit = parseInt(args[++i], 10);
}

// 格式化今日日期
function getTodayStr() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${day}`;
}

// 提取文字內容 (支援 PDF 與 TXT)
async function extractText(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    if (ext === '.pdf') {
        const buf = fs.readFileSync(filePath);
        const p = new PDFParse(new Uint8Array(buf));
        await p.load();
        const res = await p.getText();
        return (res.text || res || '').replace(/\r\n/g, '\n');
    } else if (ext === '.txt' || ext === '.json' || ext === '.md') {
        return fs.readFileSync(filePath, 'utf8').replace(/\r\n/g, '\n');
    }
    return '';
}

// 清理裁判書文本雜訊（去除行號序列、頁面頁腳等）
function cleanJudicialText(str) {
    if (!str) return '';
    return str
        .replace(/(?:--\s*\d+\s*of\s*\d+\s*--)/g, '')
        .replace(/(?:\b\d{1,2}\b\s*){3,}/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

// 智慧解析單則裁判書
function parseJudgment(text, filename) {
    const cleanFilename = path.basename(filename, path.extname(filename));

    // 1. 裁判字號
    let caseNo = '';
    const caseMatch = text.match(/(11\d\s*年度?\s*[\u4e00-\u9fa5A-Za-z]+\s*字第\s*\d+\s*號)/)
        || text.match(/【(11\d\s*年評字第\s*\d+\s*號)】/)
        || filename.match(/(11\d(?:年度?|年)[\u4e00-\u9fa5A-Za-z]+字第\d+號)/);
    if (caseMatch) {
        caseNo = caseMatch[1].replace(/\s+/g, '');
    } else {
        caseNo = cleanFilename.split('_')[1] || cleanFilename;
    }

    // 2. 裁判日期
    let dateStr = '';
    const dateMatch = text.match(/中華民國\s*(11\d)\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/)
        || text.match(/民國\s*(11\d)\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/)
        || filename.match(/(\d{4})(\d{2})(\d{2})/);
    if (dateMatch) {
        if (dateMatch[1] && dateMatch[2] && dateMatch[3]) {
            dateStr = `${dateMatch[1]}.${String(dateMatch[2]).padStart(2, '0')}.${String(dateMatch[3]).padStart(2, '0')}`;
        } else if (dateMatch[1] && dateMatch[2] && dateMatch[3]) {
            dateStr = `${parseInt(dateMatch[1]) - 1911}.${dateMatch[2]}.${dateMatch[3]}`;
        }
    }
    if (!dateStr) dateStr = cleanFilename.split('_')[0] || '115.01.01';

    // 3. 案由 / 爭議類型
    let cause = '';
    const causeMatch = text.match(/(?:因|案由：?|爭議：?)\s*([\u4e00-\u9fa5A-Za-z0-9、]+)(?:案件|事件|等罪|等)/);
    if (causeMatch) {
        cause = causeMatch[1].trim();
    } else {
        const parts = cleanFilename.split('_');
        cause = parts[2] || parts[parts.length - 1] || '未載案由';
    }

    // 4. 主文判斷
    let holding = '';
    const mainMatch = text.match(/主\s*文\s*([\s\S]*?)(?:事實及理由|事\s*實|理\s*由|一、)/);
    if (mainMatch) {
        holding = mainMatch[1].replace(/\s+/g, ' ').trim();
    }
    if (!holding) {
        if (text.includes('上訴駁回')) holding = '上訴駁回。';
        else if (text.includes('撤銷') && text.includes('發回')) holding = '原判決撤銷，發回更審。';
        else if (text.includes('應給付')) holding = '相對人應給付申請人款項。';
        else holding = '詳見裁判書本文。';
    }
    if (holding.length > 50) holding = holding.substring(0, 48) + '...';

    // 5. 法院立場 (Stance)
    let stance = '🟡 折衷衡平說';
    if (holding.includes('駁回') || holding.includes('無理由')) {
        stance = '🔴 否定說／上訴駁回';
    } else if (holding.includes('撤銷') || holding.includes('有理由') || holding.includes('應給付') || holding.includes('原判決撤銷')) {
        stance = '🟢 肯定說／撤銷原判';
    } else if (holding.includes('不受理')) {
        stance = '⚪ 不受理／程序駁回';
    }

    // 6. 核心涉訟法條
    const articles = new Set();
    const artRegex = /((?:刑事訴訟法|民事訴訟法|中華民國刑法|刑法|民法|保險法|洗錢防制法|金融消費者保護法|詐欺犯罪危害防制條例)第\s*\d+\s*條(?:之\s*\d+)?(?:第\s*\d+\s*項)?(?:第\s*\d+\s*款)?)/g;
    let m;
    while ((m = artRegex.exec(text)) !== null) {
        articles.add(m[1].replace(/\s+/g, ''));
    }
    const articlesArr = Array.from(articles);
    const topArticles = articlesArr.length > 0 ? articlesArr.slice(0, 3).join('、') : '相關實體法規範';

    // 7. 核心要件事實摘要
    let facts = '';
    const factMatch = text.match(/(?:犯罪事實|本件事實|陳述)[\s\S]*?(?:一、|1\.)\s*([\s\S]*?)(?:二、|2\.|理由|兩造不爭執)/);
    if (factMatch) {
        facts = factMatch[1].replace(/\s+/g, ' ').trim();
    }
    if (!facts || facts.length < 15) {
        facts = `本案涉及被告/當事人因涉犯${cause}，經原審認定有罪或不利處分，提起上訴主張原審適用法律有違誤。`;
    }
    if (facts.length > 90) facts = facts.substring(0, 88) + '...';

    // 8. 核心法律爭點設問
    let issueQuestion = `原審就「${cause}」之犯罪要件認定與法律效果裁量，適用法律是否有所違誤？`;
    if (cause.includes('洗錢')) {
        issueQuestion = '提供人頭金融帳戶或虛擬貨幣錢包以供移轉款項，是否該當洗錢防制法之洗錢罪構成要件？';
    } else if (cause.includes('詐欺')) {
        issueQuestion = '行為人參與電信網路詐欺集團並擔任特定分工，是否該當刑法第339條之4加重詐欺罪？';
    } else if (cause.includes('竊盜')) {
        issueQuestion = '行為人之取財手段與不法所有意圖，是否構成刑法第320條之竊盜罪既遂要件？';
    } else if (cause.includes('不便險') || cause.includes('保險')) {
        issueQuestion = '遭遇客觀不可歸責之延誤或事故時，是否該當保險契約定型化約款之理賠要件？';
    }

    // 9. 法官主要論點與心證理由 (依犯罪三階層理論梳理)
    let judgePrimaryArgument = '';
    // 依序搜尋裁判書中實質論理段落 (經查 / 惟查 / 卷查 / 然查 / 六、判斷理由)
    const substantiveMatch = text.match(/(?:經查|惟查|卷查|然查|本院查|六、判斷理由)[\s：:]([\s\S]*?)(?=(?:據上論結|中\s*華\s*民\s*國|\n\s*[二三四五六七八九]、))/);
    if (substantiveMatch) {
        judgePrimaryArgument = substantiveMatch[1].replace(/\s+/g, ' ').trim();
    }
    if (!judgePrimaryArgument) {
        const reasonMatch = text.match(/理\s*由[\s\S]*?(?:二、|2\.)([\s\S]*?)(?=(?:據上論結|中\s*華\s*民\s*國|\n\s*[三四五六七八九]、))/);
        if (reasonMatch) {
            judgePrimaryArgument = reasonMatch[1].replace(/\s+/g, ' ').trim();
        }
    }
    if (!judgePrimaryArgument || judgePrimaryArgument.length < 20) {
        judgePrimaryArgument = holding || '法院綜合全案事證，本於經驗法則與論理法則依法審結。';
    }

    // 判斷是否為刑事案件並進行「犯罪三階層理論」深層結構化梳理
    const isCriminal = text.includes('刑事') || text.includes('被告') || text.includes('判處') || text.includes('有期徒刑') || cause.includes('刑') || cause.includes('罪');
    let threeTier = null;

    if (isCriminal) {
        // 第一階：構成要件該當性 (客觀行為 + 主觀故意/意圖，放寬長度限制以保留完整事證)
        let tatbestandObj = '';
        const objMatch = text.match(/(?:犯罪事實|原判決認定|事實欄)[\s\S]*?(?:係以|認定被告|持往|交付|施用|竊取|參與)([\s\S]*?)(?=(?:核其所為|核被告所為|論罪|理由|二、))/);
        if (objMatch) {
            tatbestandObj = cleanJudicialText(objMatch[1]);
        } else {
            tatbestandObj = `被告實行${cause}之客觀實行行為與因果歸責。`;
        }

        let tatbestandSub = '主觀上具備犯罪故意。';
        if (text.includes('未必故意') || text.includes('不確定故意')) {
            tatbestandSub = '具備未必故意（預見其發生且不違背其本意）。';
        } else if (text.includes('不法所有意圖')) {
            tatbestandSub = '具備意圖為自己或第三人不法所有之意圖。';
        } else if (text.includes('洗錢')) {
            tatbestandSub = '預見提供金融帳戶可能遭詐欺集團掩飾隱匿犯罪所得，仍容任其發生。';
        }

        // 第二階：違法性審查
        let rechtswidrigkeit = '無正當防衛、緊急避難或法令行為等阻卻違法事由，具備實質違法性。';
        if (text.includes('正當防衛')) rechtswidrigkeit = '被告抗辯正當防衛，法官審查認欠缺防衛情狀或過當，不阻卻違法。';
        else if (text.includes('緊急避難')) rechtswidrigkeit = '被告抗辯緊急避難，法官審酌法益權衡不該當，難認阻卻違法。';

        // 第三階：有責性 (責任能力 + 期待可能性 + 違法性認識)
        let schuld = '行為人具完全責任能力，且無阻卻罪責事由。';
        if (text.includes('刑法第19條') || text.includes('精神障礙') || text.includes('心智缺陷') || text.includes('智能不足')) {
            schuld = '審酌刑法第19條責任能力：法官調查精神/智能狀況，審查行為時辨識與控制能力是否顯著減低，依法裁量罪責減輕或維持。';
        } else if (text.includes('刑法第16條') || text.includes('禁止錯誤') || text.includes('不知法律')) {
            schuld = '審酌刑法第16條違法性認識：法官認非屬不可避免之正當理由，具備罪責。';
        } else if (text.includes('身心障礙') || text.includes('無法為完全之陳述')) {
            schuld = '審酌被告身心障礙狀況與訴訟防禦能力，確認是否符合強制辯護與訴訟照料要件。';
        }

        // 第四階：完整罪數與量刑心證論述 (刑法第57條、第59條、減刑及職權裁量，放寬字數行數限制)
        let sentencing = '';
        const sentParagraphs = [];
        const sentRegex = /(?:[一二三四五六七八九]、)?([^\n]*?(?:量刑|科刑|刑法第57條|刑之量定|裁量範疇|酌量減輕|責任為基礎)[\s\S]*?(?=(?:\n\s*[一二三四五六七八九]、|據上論結|中\s*華\s*民\s*國)))/g;
        let sentMatch;
        while ((sentMatch = sentRegex.exec(text)) !== null) {
            const cleaned = cleanJudicialText(sentMatch[1]);
            if (cleaned.length > 25 && !sentParagraphs.some(p => p.includes(cleaned.substring(0, 30)))) {
                sentParagraphs.push(cleaned);
            }
        }

        if (sentParagraphs.length > 0) {
            sentencing = sentParagraphs.join('\n\n');
        } else if (judgePrimaryArgument.includes('量刑') || text.includes('刑法第57條')) {
            sentencing = '法官審酌刑法第57條各款科刑因子（犯罪動機、目的、手段、分工地位、損害程度與犯後態度），於法定刑度內妥適行使裁量權，未逾越外部及內部界限，合於罪刑相當原則與比例原則。';
        } else {
            sentencing = '依法論罪科刑，認事用法並無濫用裁量權或違背罪刑相當原則之違法情形。';
        }

        threeTier = {
            tatbestand: `【客觀行為】${tatbestandObj}；【主觀犯意】${tatbestandSub}`,
            rechtswidrigkeit,
            schuld,
            sentencing
        };
    }

    // 提煉法官完整裁判論點 (整合三階層論證與完整量刑論述，全面放寬長度限制)
    let judgeShortSummary = '';
    if (threeTier) {
        judgeShortSummary = `【構成要件】${threeTier.tatbestand}\n【違法性】${threeTier.rechtswidrigkeit}\n【有責性】${threeTier.schuld}\n【量刑心證完整論述】${threeTier.sentencing}`;
    } else {
        judgeShortSummary = cleanJudicialText(judgePrimaryArgument);
    }

    // 10. 司法三段論基本結構剖析 (Legal Syllogism)
    let majorPremise = '';
    if (topArticles && topArticles !== '相關實體法規範') {
        majorPremise = `規範依${topArticles}之法定構成要件及實務審查標準。`;
    } else {
        majorPremise = `依${cause}相關實體法規要件及實務判決原則。`;
    }
    let minorPremise = facts;
    let subsumption = '';
    if (stance.includes('肯定說')) {
        subsumption = `原審於要件事實之調查或法規涵攝容有瑕疵，上訴（或請求）有理由，予以撤銷發回。`;
    } else if (stance.includes('否定說')) {
        subsumption = `被告實行行為與事證完全該當構成要件，無阻卻違法與罪責事由，認事用法無違誤，維持原判。`;
    } else {
        subsumption = `將案件事實與事證代入法律規範要件比對檢驗，依法為相應裁判論處。`;
    }

    const syllogism = {
        majorPremise,
        minorPremise,
        subsumption
    };
    const syllogismSummary = `[大前提] ${majorPremise} | [小前提] ${minorPremise} | [涵攝結論] ${subsumption}`;

    // 11. 裁判核心原文金句 (Ratio Decidendi Quotes)
    let ratioQuote = '';
    const quoteCandidate = judgePrimaryArgument.match(/(?:按(?:「|『)?[\u4e00-\u9fa5A-Za-z0-9、，。；「」『』]+?(?:定有明文|之規定|原則|之宗旨)。|核其認事用法[\u4e00-\u9fa5A-Za-z0-9、，。；「」『』]+?違法。|原審[\u4e00-\u9fa5A-Za-z0-9、，。；「」『』]+?違法。)/);
    if (quoteCandidate) {
        ratioQuote = quoteCandidate[0].trim();
    }
    if (!ratioQuote || ratioQuote.length < 15) {
        ratioQuote = judgePrimaryArgument.substring(0, 110) + '……';
    }
    if (ratioQuote.length > 110) ratioQuote = ratioQuote.substring(0, 108) + '……';

    // 12. 論文寫作引用建議
    let thesisNote = '作為實務普遍見解與判決合法性審查之例證。';
    if (stance.includes('肯定說')) {
        thesisNote = '🌟【主力正方/撤銷例證】可作為論述原審適用法規違背法令之有力實務裁判依據。';
    } else if (stance.includes('否定說')) {
        thesisNote = '🌟【維持原判/主流例證】作為目前終審法院嚴格守護構成要件審查之標準案例。';
    }

    return {
        filename,
        caseNo,
        date: dateStr,
        disputeType: cause,
        holding,
        stance,
        articles: topArticles,
        facts,
        issue: issueQuestion,
        judgeArgument: judgeShortSummary,
        fullJudgeArgument: judgePrimaryArgument,
        threeTier,
        syllogism,
        syllogismSummary,
        ratio: ratioQuote,
        thesis: thesisNote
    };
}

// 產生 Excel 檔案
async function generateExcelFiles(cases, outDir, baseName, topicName) {
    // 1. 各篇分析結果摘要表 (12 欄位：三段論結構 + 犯罪三階層心證)
    const summaryWb = new ExcelJS.Workbook();
    summaryWb.creator = 'Antigravity Legal Batch Analyzer';
    summaryWb.created = new Date();

    const summarySheet = summaryWb.addWorksheet('各篇裁判分析結果摘要總表', {
        views: [{ state: 'frozen', ySplit: 1 }]
    });

    summarySheet.columns = [
        { header: '序號', key: 'seq', width: 8 },
        { header: '裁判日期', key: 'date', width: 14 },
        { header: '裁判/評議字號', key: 'caseNo', width: 26 },
        { header: '案由/爭議類型', key: 'disputeType', width: 20 },
        { header: '主文判決結果', key: 'holding', width: 26 },
        { header: '核心法律爭點設問', key: 'issue', width: 42 },
        { header: '核心涉訟法條與約款', key: 'articles', width: 34 },
        { header: '法院/委員會立場', key: 'stance', width: 20 },
        { header: '【司法三段論結構剖析】', key: 'syllogism', width: 50 },
        { header: '【法官主要論點（含完整量刑心證）】', key: 'judgeArgument', width: 75 },
        { header: '裁判核心論理原文金句 (Ratio Decidendi)', key: 'ratio', width: 45 },
        { header: '論文寫作引用建議', key: 'thesis', width: 34 }
    ];

    // Header Styling
    const headRow = summarySheet.getRow(1);
    headRow.height = 32;
    headRow.eachCell(c => {
        c.font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
    });

    cases.forEach((item, idx) => {
        const r = summarySheet.addRow({
            seq: idx + 1,
            date: item.date,
            caseNo: item.caseNo,
            disputeType: item.disputeType,
            holding: item.holding,
            issue: item.issue,
            articles: item.articles,
            stance: item.stance,
            syllogism: item.syllogismSummary,
            judgeArgument: item.judgeArgument,
            ratio: item.ratio,
            thesis: item.thesis
        });
        r.alignment = { vertical: 'top', wrapText: true };
        r.font = { name: '微軟正黑體', size: 10 };

        const bg = (idx % 2 === 1) ? 'FFF9FAFC' : 'FFFFFFFF';
        r.eachCell((cell, colNum) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
            };
            if ([1, 2, 4, 8].includes(colNum)) {
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            }
        });
    });

    let summaryXlsxPath = path.join(outDir, `各篇分析結果摘要表_${baseName}.xlsx`);
    try {
        await summaryWb.xlsx.writeFile(summaryXlsxPath);
    } catch (err) {
        if (err.code === 'EBUSY') {
            summaryXlsxPath = path.join(outDir, `各篇分析結果摘要表_${baseName}_${Date.now()}.xlsx`);
            await summaryWb.xlsx.writeFile(summaryXlsxPath);
        } else {
            throw err;
        }
    }

    // 2. 橫向比對矩陣表 (8 欄位：納入司法三段論與法官論點)
    const matrixWb = new ExcelJS.Workbook();
    const matrixSheet = matrixWb.addWorksheet('裁判爭點橫向對比矩陣', {
        views: [{ state: 'frozen', ySplit: 1 }]
    });

    matrixSheet.columns = [
        { header: '裁判字號', key: 'caseNo', width: 26 },
        { header: '裁判日期', key: 'date', width: 14 },
        { header: '核心涉訟法條與約款', key: 'articles', width: 32 },
        { header: '核心要件事實概要', key: 'facts', width: 38 },
        { header: '面對爭點之法院立場', key: 'stance', width: 20 },
        { header: '【司法三段論演繹推論】', key: 'syllogism', width: 50 },
        { header: '【法官主要裁判論點（含完整量刑論述）】', key: 'judgeArgument', width: 75 },
        { header: '裁判核心論理金句 (Ratio Decidendi)', key: 'ratio', width: 45 }
    ];

    const mHead = matrixSheet.getRow(1);
    mHead.height = 32;
    mHead.eachCell(c => {
        c.font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
        c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
    });

    cases.forEach((item, idx) => {
        const r = matrixSheet.addRow({
            caseNo: item.caseNo,
            date: item.date,
            articles: item.articles,
            facts: item.facts,
            stance: item.stance,
            syllogism: item.syllogismSummary,
            judgeArgument: item.judgeArgument,
            ratio: item.ratio
        });
        r.alignment = { vertical: 'top', wrapText: true };
        r.font = { name: '微軟正黑體', size: 10 };

        const bg = (idx % 2 === 1) ? 'FFF9FAFC' : 'FFFFFFFF';
        r.eachCell((cell, colNum) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
            };
            if ([2, 5].includes(colNum)) {
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            }
        });
    });

    let matrixXlsxPath = path.join(outDir, `批次爭點比對矩陣_${baseName}.xlsx`);
    try {
        await matrixWb.xlsx.writeFile(matrixXlsxPath);
    } catch (err) {
        if (err.code === 'EBUSY') {
            matrixXlsxPath = path.join(outDir, `批次爭點比對矩陣_${baseName}_${Date.now()}.xlsx`);
            await matrixWb.xlsx.writeFile(matrixXlsxPath);
        } else {
            throw err;
        }
    }

    return { summaryXlsxPath, matrixXlsxPath };
}

// 產生 Markdown 深度報告
function generateMarkdownReport(cases, outDir, baseName, topicName) {
    const affirmativeCases = cases.filter(c => c.stance.includes('肯定說'));
    const negativeCases = cases.filter(c => c.stance.includes('否定說'));
    const otherCases = cases.filter(c => !c.stance.includes('肯定說') && !c.stance.includes('否定說'));

    let md = `# 【裁判爭點批次分析與比對報告】${topicName}\n\n`;
    md += `- **研究主題／檢索主軸**：${topicName}\n`;
    md += `- **納入分析裁判總數**：共 ${cases.length} 則\n`;
    md += `- **產生日期**：${new Date().toLocaleDateString('zh-TW')} (系統自動批次產出)\n\n`;
    md += `---\n\n`;

    md += `## 一、實務見解立場光譜與分佈（Doctrinal Spectrum）\n\n`;
    md += `| 立場分類 | 裁判筆數 | 代表字號範例 | 核心法理與裁判立場摘要 |\n`;
    md += `| :--- | :---: | :--- | :--- |\n`;
    md += `| **【肯定說／撤銷原判／有理由】** | **${affirmativeCases.length} 筆** | ${affirmativeCases.slice(0, 3).map(c => c.caseNo).join('、') || '無'} | 認定原審適用法律有違誤、證據取捨不備理由，或事實涵攝有重大瑕疵，予以撤銷發回或改判。 |\n`;
    md += `| **【否定說／維持原判／上訴駁回】** | **${negativeCases.length} 筆** | ${negativeCases.slice(0, 3).map(c => c.caseNo).join('、') || '無'} | 嚴格維持實體構成要件與程序規定，認為原判決認事用法並無違法，駁回上訴或請求。 |\n`;
    if (otherCases.length > 0) {
        md += `| **【程序駁回／其他處理】** | **${otherCases.length} 筆** | ${otherCases.slice(0, 3).map(c => c.caseNo).join('、')} | 未備法定程序要件，或依法為特定保安處分、移送管轄等處置。 |\n`;
    }
    md += `\n---\n\n`;

    md += `## 二、多篇裁判橫向對比大矩陣（Comparative Matrix）\n\n`;
    md += `| 裁判字號 | 裁判日期 | 核心涉訟法條與約款 | 核心要件事實概要 | 法院立場 | 【司法三段論演繹推論】 | 【法官主要裁判論點】 | 裁判核心論理金句 (Ratio Decidendi) |\n`;
    md += `| :--- | :---: | :--- | :--- | :---: | :--- | :--- | :--- |\n`;
    cases.forEach(c => {
        const tableJudgeArg = c.judgeArgument.replace(/\n+/g, '<br>');
        const tableSyllogism = c.syllogismSummary.replace(/\n+/g, '<br>');
        md += `| **${c.caseNo}** | ${c.date} | ${c.articles} | ${c.facts} | ${c.stance} | ${tableSyllogism} | ${tableJudgeArg} | 「${c.ratio}」 |\n`;
    });
    md += `\n---\n\n`;

    md += `## 三、實務見解演進與歧異深度剖析（Trend & Divergence Analysis）\n\n`;
    md += `1. **構成要件審查與證據評價之嚴格化**：\n`;
    md += `   實務在處理案件時，終審法院對於事實認定與論理過程之拘束力愈趨嚴格，尤其強調不得僅憑主觀推測涵攝構成要件。\n\n`;
    md += `2. **法益權衡與立法目的之貫徹**：\n`;
    md += `   各級法院在適用實體法規時，多著眼於該法條所保護之個別法益與社會秩序維護之間的衡量。\n\n`;
    md += `---\n\n`;

    md += `## 四、論文寫作整合與引用建議（Thesis Integration）\n\n`;
    if (affirmativeCases.length > 0) {
        md += `- **推薦正方例證**：🌟 **${affirmativeCases[0].caseNo}**，本案論述精闢，適合作為論文實務肯定立場之主力例證。\n`;
    }
    if (negativeCases.length > 0) {
        md += `- **推薦對照批判素材**：🌟 **${negativeCases[0].caseNo}**，可作為反面論點或實務邊界之深入探討對象。\n`;
    }
    md += `\n---\n\n`;

    md += `## 五、各篇裁判／決定書獨立分析結果精要（Case-by-Case Summary）\n\n`;
    cases.forEach((c, idx) => {
        md += `### ${idx + 1}. 【${c.caseNo}】（裁判日期：${c.date}）\n`;
        md += `- **爭議類型／案由**：${c.disputeType}\n`;
        md += `- **判決主文結果**：${c.holding}\n`;
        md += `- **核心涉訟法條**：${c.articles}\n`;
        md += `- **要件事實摘要**：${c.facts}\n`;
        md += `- **核心法律爭點**：${c.issue}\n`;
        md += `- **【司法三段論基本結構剖析（Legal Syllogism）】**：\n`;
        md += `  - **(1) 大前提（規範要件與解釋標準 / Major Premise）**：${c.syllogism.majorPremise}\n`;
        md += `  - **(2) 小前提（審認事實與調查證據 / Minor Premise）**：${c.syllogism.minorPremise}\n`;
        md += `  - **(3) 涵攝與結論（事實該當推論與法律效果 / Subsumption & Conclusion）**：${c.syllogism.subsumption}\n`;
        if (c.threeTier) {
            md += `- **【法官主要論點（依犯罪三階層體系梳理，含完整量刑心證論述）】**：\n`;
            md += `  - **(1) 構成要件該當性 (Tatbestand)**：${c.threeTier.tatbestand}\n`;
            md += `  - **(2) 違法性審查 (Rechtswidrigkeit)**：${c.threeTier.rechtswidrigkeit}\n`;
            md += `  - **(3) 有責性評價 (Schuld)**：${c.threeTier.schuld}\n`;
            md += `  - **(4) 罪數與量刑心證完整論述 (Sentencing & Discretion)**：\n    ${c.threeTier.sentencing.replace(/\n+/g, '\n    ')}\n`;
        } else {
            md += `- **【法官主要裁判論點與心證理由】**：\n  ${c.fullJudgeArgument}\n`;
        }
        md += `- **【判決核心原文金句 (Ratio Decidendi)】**：\n  > 「${c.ratio}」\n`;
        md += `- **論文寫作引用定位**：${c.thesis}\n\n`;
    });

    const mdPath = path.join(outDir, `批次爭點分析報告_${baseName}.md`);
    fs.writeFileSync(mdPath, md, 'utf8');
    return mdPath;
}

// 主執行函式
async function main() {
    console.log(`=======================================================`);
    console.log(`  司法裁判爭點分析 - 自動化批次執行系統`);
    console.log(`  來源目錄 (Input):  ${CONFIG.inputDir}`);
    console.log(`  輸出目錄 (Output): ${CONFIG.outputDir}`);
    console.log(`=======================================================\n`);

    if (!fs.existsSync(CONFIG.outputDir)) {
        fs.mkdirSync(CONFIG.outputDir, { recursive: true });
    }

    // 檢查來源目錄
    if (!fs.existsSync(CONFIG.inputDir)) {
        console.error(`❌ 來源目錄不存在: ${CONFIG.inputDir}`);
        process.exit(1);
    }

    let files = fs.readdirSync(CONFIG.inputDir).filter(f => {
        const ext = path.extname(f).toLowerCase();
        return ['.pdf', '.txt', '.md'].includes(ext);
    });

    // 如果來源目錄為空，自動檢查是否有備用目錄 (例如 judgments_data 或 判決原文)
    if (files.length === 0) {
        const fallback1 = path.resolve(WORKSPACE_ROOT, 'judgments_data');
        const fallback2 = path.resolve(WORKSPACE_ROOT, 'judgements_org');
        let chosenFallback = '';
        if (fs.existsSync(fallback1) && fs.readdirSync(fallback1).some(f => ['.pdf', '.txt', '.md'].includes(path.extname(f).toLowerCase()))) {
            chosenFallback = fallback1;
        } else if (fs.existsSync(fallback2) && fs.readdirSync(fallback2).some(f => f.endsWith('.pdf'))) {
            chosenFallback = fallback2;
        }

        if (chosenFallback) {
            console.log(`ℹ️ judgements_org 目前為空，自動切換至備用來源目錄: ${chosenFallback}`);
            CONFIG.inputDir = chosenFallback;
            files = fs.readdirSync(CONFIG.inputDir).filter(f => ['.pdf', '.txt', '.md'].includes(path.extname(f).toLowerCase()));
        } else {
            console.log(`⚠️ 來源目錄中未發現任何 PDF 或 TXT 裁判檔案！請先將檔案放入 ${CONFIG.inputDir}`);
            process.exit(0);
        }
    }

    if (CONFIG.limit > 0 && files.length > CONFIG.limit) {
        console.log(`📌 依設定限制，僅處理前 ${CONFIG.limit} 筆裁判 (資料夾總計 ${files.length} 筆)`);
        files = files.slice(0, CONFIG.limit);
    }

    console.log(`[1/3] 開始讀取並解析 ${files.length} 筆裁判書文本...`);
    const parsedCases = [];

    for (let i = 0; i < files.length; i++) {
        const f = files[i];
        const filePath = path.join(CONFIG.inputDir, f);
        process.stdout.write(`  -> [${i + 1}/${files.length}] 解析中: ${f.substring(0, 45)}... `);

        try {
            const text = await extractText(filePath);
            if (!text || text.length < 50) {
                console.log('跳過 (文本過短或無法讀取)');
                continue;
            }
            const item = parseJudgment(text, f);
            parsedCases.push(item);
            console.log(`完成 (${(text.length / 1000).toFixed(1)}k 字)`);
        } catch (err) {
            console.log(`失敗: ${err.message}`);
        }
    }

    if (parsedCases.length === 0) {
        console.log('未成功解析任何裁判文本，結束執行。');
        return;
    }

    console.log(`\n[2/3] 排序與提煉爭點光譜...`);
    parsedCases.sort((a, b) => b.date.localeCompare(a.date));

    // 推導主題名稱
    let topicName = CONFIG.topic;
    if (!topicName) {
        const commonTypes = parsedCases.map(c => c.disputeType).filter(Boolean);
        const topType = commonTypes[0] || '司法裁判爭點整理';
        topicName = `${topType}實務爭點研究`;
    }

    const todayStr = getTodayStr();
    const cleanTopic = topicName.replace(/[\/\\:*?"<>|]/g, '_');
    const baseName = `${cleanTopic}_${todayStr}`;

    console.log(`[3/3] 產出雙軌成果至目錄: ${CONFIG.outputDir}`);

    // 1. 生成 Markdown 報告
    const mdPath = generateMarkdownReport(parsedCases, CONFIG.outputDir, baseName, topicName);
    console.log(`  ✅ Markdown 報告已生成: ${mdPath}`);

    // 2. 生成 Excel 檔案 (摘要表與矩陣表)
    const { summaryXlsxPath, matrixXlsxPath } = await generateExcelFiles(parsedCases, CONFIG.outputDir, baseName, topicName);
    console.log(`  ✅ Excel 摘要表已生成:   ${summaryXlsxPath}`);
    console.log(`  ✅ Excel 矩陣表已生成:   ${matrixXlsxPath}`);

    console.log(`\n🎉 批次處理完畢！共成功分析 ${parsedCases.length} 篇裁判書。`);
}

main().catch(err => {
    console.error('執行發生嚴重錯誤:', err);
});
