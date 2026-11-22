/**
 * 金融消費評議決定書自動下載腳本 (FOI Decision Auto-Downloader)
 * 目標網站: https://ods.foi.org.tw/
 * 執行環境: Node.js (v18+)
 * 
 * 執行方式:
 *   1. 互動式執行（推薦）：
 *      node scripts/download_foi_decisions.js
 *      (執行後會依序詢問關鍵字、日期等資訊，確認後才開始下載)
 * 
 *   2. 命令列參數執行：
 *      node scripts/download_foi_decisions.js --keyword "告知義務" --start "1140101" --end "1151122"
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline/promises');
const { stdin: input, stdout: output } = require('process');

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

// 取得當前民國年月日
function getTodayRoc() {
    const now = new Date();
    const rocYear = String(now.getFullYear() - 1911);
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return { rocYear, month, day, str: `${rocYear}${month}${day}` };
}

// 解析 7 碼民國年月日 (例如 1150101 -> { year: 115, month: 01, day: 01 })
function parseRocDate(dateStr, defaultVal) {
    const s = String(dateStr || '').replace(/\D/g, '');
    if (s.length >= 7) {
        return {
            year: s.substring(0, s.length - 4),
            month: s.substring(s.length - 4, s.length - 2),
            day: s.substring(s.length - 2),
            raw: s
        };
    }
    return defaultVal;
}

async function getParams() {
    const today = getTodayRoc();
    const args = process.argv.slice(2);
    let cliKeyword = '';
    let cliStart = '';
    let cliEnd = '';
    let cliOutput = '';

    for (let i = 0; i < args.length; i++) {
        if (args[i] === '--keyword' && args[i + 1]) cliKeyword = args[++i];
        if (args[i] === '--start' && args[i + 1]) cliStart = args[++i];
        if (args[i] === '--end' && args[i + 1]) cliEnd = args[++i];
        if (args[i] === '--output' && args[i + 1]) cliOutput = args[++i];
    }

    // 若未透過命令列提供關鍵字，啟動互動式提問介面
    if (!cliKeyword) {
        const rl = readline.createInterface({ input, output });
        console.log(`\n======================================================`);
        console.log(`  金融消費評議決定書下載系統 - 檢索資訊設定`);
        console.log(`======================================================\n`);

        let keyword = '';
        while (!keyword) {
            keyword = (await rl.question('👉 請輸入檢索關鍵字 (必填，例如：告知義務 / 旅遊不便險): ')).trim();
            if (!keyword) {
                console.log('   ⚠️ 關鍵字不得為空，請重新輸入！\n');
            }
        }

        const defaultStart = `${today.rocYear}0101`;
        let startDateInput = (await rl.question(`👉 請輸入起始評議決定日期 (格式如 ${defaultStart}，直接 Enter 預設為 ${defaultStart}): `)).trim();
        if (!startDateInput) startDateInput = defaultStart;

        const defaultEnd = today.str;
        let endDateInput = (await rl.question(`👉 請輸入結束評議決定日期 (格式如 ${defaultEnd}，直接 Enter 預設為今天 ${defaultEnd}): `)).trim();
        if (!endDateInput) endDateInput = defaultEnd;

        const startParsed = parseRocDate(startDateInput, { year: today.rocYear, month: '01', day: '01', raw: defaultStart });
        const endParsed = parseRocDate(endDateInput, { year: today.rocYear, month: today.month, day: today.day, raw: defaultEnd });

        console.log(`\n------------------------------------------------------`);
        console.log(`【檢索條件確認】`);
        console.log(`- 檢索關鍵字: ${keyword}`);
        console.log(`- 決定日期區間: 民國 ${startParsed.year}/${startParsed.month}/${startParsed.day} ~ ${endParsed.year}/${endParsed.month}/${endParsed.day}`);
        console.log(`------------------------------------------------------`);

        const confirm = (await rl.question('👉 確認送出檢索並下載？ [Y/n]: ')).trim().toLowerCase();
        rl.close();

        if (confirm === 'n' || confirm === 'no') {
            console.log('\n❌ 使用者取消操作，已停止下載。\n');
            process.exit(0);
        }

        return { keyword, start: startParsed, end: endParsed, outputDir: cliOutput };
    } else {
        // CLI 參數模式
        const startParsed = parseRocDate(cliStart, { year: today.rocYear, month: '01', day: '01', raw: `${today.rocYear}0101` });
        const endParsed = parseRocDate(cliEnd, { year: today.rocYear, month: today.month, day: today.day, raw: today.str });
        return { keyword: cliKeyword, start: startParsed, end: endParsed, outputDir: cliOutput };
    }
}

async function main() {
    const { keyword, start, end, outputDir: customOutput } = await getParams();

    const outputDir = customOutput ? path.resolve(customOutput) : path.resolve(getWorkspaceRoot(), 'judgements_org');

    console.log(`\n[1/4] 連線目標網站 (ods.foi.org.tw) 並取得 Session 狀態...`);
    const getRes = await fetch('https://ods.foi.org.tw/', {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    });
    const cookie = getRes.headers.get('set-cookie') || '';
    const html = await getRes.text();

    const getVal = (name) => {
        const m = html.match(new RegExp(`id="${name}" value="([^"]*)"`));
        return m ? m[1] : '';
    };

    const viewState = getVal('__VIEWSTATE');
    const viewStateGen = getVal('__VIEWSTATEGENERATOR');
    const eventValidation = getVal('__EVENTVALIDATION');

    console.log(`[2/4] 送出查詢請求 (關鍵字: "${keyword}")...`);
    const body = new URLSearchParams();
    body.append('__VIEWSTATE', viewState);
    body.append('__VIEWSTATEGENERATOR', viewStateGen);
    body.append('__EVENTVALIDATION', eventValidation);
    body.append('Foi$cph_content$ddlTypeID', '0');             // 全部類別
    body.append('Foi$cph_content$ddl_page', '100');            // 頁數容量
    body.append('Foi$cph_content$ddl_ResultKind', '0');        // 全部決定
    body.append('Foi$cph_content$hResName', '');
    body.append('Foi$cph_content$ddlVerticals', '0');          // 全部產業
    body.append('Foi$cph_content$txt_BYear', '');
    body.append('Foi$cph_content$dll_BCase', '');
    body.append('Foi$cph_content$txt_Bno', '');
    body.append('Foi$cph_content$txt_Content', keyword);
    body.append('Foi$cph_content$txt_Syear', start.year);
    body.append('Foi$cph_content$txt_Smonth', start.month);
    body.append('Foi$cph_content$txt_Sday', start.day);
    body.append('Foi$cph_content$txt_Eyear', end.year);
    body.append('Foi$cph_content$txt_Emonth', end.month);
    body.append('Foi$cph_content$txt_Eday', end.day);
    body.append('Foi$cph_content$sdate', `${start.year}${start.month}${start.day}`);
    body.append('Foi$cph_content$edate', `${end.year}${end.month}${end.day}`);
    body.append('Foi$cph_content$btn_submit', '送出查詢');

    const postRes = await fetch('https://ods.foi.org.tw/', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
            'Cookie': cookie
        },
        body: body.toString()
    });

    const postHtml = await postRes.text();

    const countMatch = postHtml.match(/id="cph_content_HFRecordCount" value="([^"]*)"/);
    const totalCount = countMatch ? parseInt(countMatch[1], 10) : 0;
    console.log(`[3/4] 檢索完成！共找到 ${totalCount} 筆符合條件之評議決定書。`);

    if (totalCount === 0) {
        console.log(`\nℹ️ 未檢索到關於「${keyword}」的評議決定書，請確認關鍵字或日期區間。\n`);
        return;
    }

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    const itemRegex = /<tr>\s*<td>\s*(\d+)\.\s*<\/td>\s*<td>\s*([^<]+)\s*<\/td>\s*<td>\s*<a href="([^"]+)"[^>]*>[\s\S]*?([0-9]+年評字第[0-9]+號)<\/a>\s*<\/td>\s*<td>\s*([0-9.]+)\s*<\/td>\s*<td>\s*([^<]+)\s*<\/td>/g;
    
    let match;
    const items = [];
    while ((match = itemRegex.exec(postHtml)) !== null) {
        items.push({
            index: match[1].trim(),
            type: match[2].trim(),
            link: match[3].trim(),
            caseNo: match[4].trim(),
            date: match[5].trim(),
            disputeType: match[6].trim()
        });
    }

    console.log(`[4/4] 檢查本機檔案並下載 PDF 決定書至:\n      ${outputDir}`);
    const downloadedFiles = [];
    let skippedCount = 0;
    let newDownloadCount = 0;

    // 先讀取輸出目錄現有檔案清單以供快速比對
    const existingFilesInDir = fs.readdirSync(outputDir);

    for (const item of items) {
        const cleanCase = item.caseNo.replace(/[\/\\:*?"<>|]/g, '_');
        const cleanType = item.type.replace(/[\/\\:*?"<>|]/g, '_');
        const cleanDispute = item.disputeType.replace(/[\/\\:*?"<>|]/g, '_');
        const filename = `${item.date}_${cleanCase}_${cleanType}_${cleanDispute}.pdf`;
        const filepath = path.join(outputDir, filename);

        // 檢查 1: 同名檔案是否已存在且大小 > 0
        // 檢查 2: 目錄下是否已有相同評議字號的 PDF (避免標題微調導致重複下載)
        const matchedExistingFile = existingFilesInDir.find(f => f.includes(cleanCase) && f.endsWith('.pdf'));

        if (matchedExistingFile) {
            const existingPath = path.join(outputDir, matchedExistingFile);
            const stat = fs.statSync(existingPath);
            if (stat.size > 0) {
                console.log(`  -> [${item.index}/${items.length}] 檔案已存在，略過下載: ${matchedExistingFile} (${(stat.size / 1024).toFixed(1)} KB)`);
                skippedCount++;
                downloadedFiles.push({
                    ...item,
                    filename: matchedExistingFile,
                    filepath: existingPath.replace(/\\/g, '/'),
                    bytes: stat.size,
                    isExisting: true
                });
                continue;
            }
        }

        const downloadUrl = `https://ods.foi.org.tw/${item.link.trim()}`;
        process.stdout.write(`  -> [${item.index}/${items.length}] 下載中: ${filename}... `);

        try {
            const res = await fetch(downloadUrl);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const buffer = await res.arrayBuffer();
            fs.writeFileSync(filepath, Buffer.from(buffer));
            console.log(`完成 (${(buffer.byteLength / 1024).toFixed(1)} KB)`);

            newDownloadCount++;
            downloadedFiles.push({
                ...item,
                filename,
                filepath: filepath.replace(/\\/g, '/'),
                bytes: buffer.byteLength,
                isExisting: false
            });
        } catch (err) {
            console.log(`失敗: ${err.message}`);
        }
    }

    // 產出 Markdown 清單
    let md = `# 金融消費評議決定書檢索清單（${keyword}）\n\n`;
    md += `- **檢索關鍵字**：${keyword}\n`;
    md += `- **評議決定日期**：民國 ${start.year}/${start.month}/${start.day} ~ ${end.year}/${end.month}/${end.day}\n`;
    md += `- **下載時間**：${new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}\n`;
    md += `- **統計摘要**：共檢索到 ${downloadedFiles.length} 筆（本次全新下載 ${newDownloadCount} 筆，已存在略過 ${skippedCount} 筆）\n\n`;
    md += `| 序號 | 決定日期 | 評議字號 | 決定類型 | 爭議類型 | 檔案狀態 | 檔案連結 |\n`;
    md += `| :---: | :---: | :--- | :--- | :--- | :---: | :--- |\n`;

    downloadedFiles.forEach((f, idx) => {
        const statusTag = f.isExisting ? '🟢 現存檔案' : '🆕 本次下載';
        md += `| ${idx + 1} | ${f.date} | ${f.caseNo} | ${f.type} | ${f.disputeType} | ${statusTag} | [${f.filename}](file:///${encodeURI(f.filepath)}) |\n`;
    });

    const indexFile = path.join(outputDir, '00_評議決定書清單.md');
    fs.writeFileSync(indexFile, md, 'utf8');
    console.log(`\n🎉 全部處理完成！共 ${downloadedFiles.length} 筆（新下載: ${newDownloadCount} 筆，略過重複: ${skippedCount} 筆）。`);
    console.log(`目錄清單已產出至:\n${indexFile}\n`);
}

main().catch(err => {
    console.error('執行發生錯誤:', err);
});
