/**
 * 爭點比對矩陣 Excel 匯出工具 (Export Matrix to Excel)
 * 功能: 將 Markdown 報告中的橫向對比矩陣自動轉換為格式精美的 .xlsx 試算表
 */

const fs = require('fs');
const path = require('path');
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

/**
 * 將包含 Markdown 表格的內容或陣列轉換為 Excel 檔案
 * @param {Array<Object>|string} data - 表格資料陣列或 Markdown 檔案路徑
 * @param {string} outputPath - 輸出的 .xlsx 檔案路徑
 * @param {string} sheetName - 工作表名稱
 */
async function exportToExcel(data, outputPath, sheetName = '裁判爭點橫向對比矩陣') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Antigravity Legal Analysis';
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet(sheetName, {
        views: [{ state: 'frozen', ySplit: 1 }] // 凍結首行
    });

    let headers = [];
    let rows = [];

    if (typeof data === 'string') {
        // 從 Markdown 檔案中解析表格
        const content = fs.readFileSync(data, 'utf8');
        const lines = content.split('\n');
        let inTable = false;
        let tableLines = [];

        for (const line of lines) {
            const trimmed = line.trim();
            if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
                if (!inTable) inTable = true;
                tableLines.push(trimmed);
            } else if (inTable) {
                if (tableLines.length >= 3) break; // 抓取第一個完整表格
                tableLines = [];
                inTable = false;
            }
        }

        if (tableLines.length >= 3) {
            headers = tableLines[0].split('|').map(c => c.trim()).filter(Boolean);
            // tableLines[1] 是分隔線 |:---|:---|
            for (let i = 2; i < tableLines.length; i++) {
                const cols = tableLines[i].split('|').map(c => c.trim().replace(/<br>/gi, '\n')).filter((_, idx, arr) => idx > 0 && idx < arr.length - 1);
                rows.push(cols);
            }
        }
    } else if (Array.isArray(data)) {
        // 直接由物件/陣列傳入
        if (data.length > 0) {
            headers = Object.keys(data[0]);
            rows = data.map(item => Object.values(item));
        }
    }

    if (headers.length === 0) {
        throw new Error('未找到可匯出的表格資料！');
    }

    // 設定欄位與寬度
    worksheet.columns = headers.map((header) => {
        let width = 20;
        if (header.includes('裁判字號') || header.includes('字號')) width = 26;
        if (header.includes('日期')) width = 14;
        if (header.includes('法條') || header.includes('條款')) width = 28;
        if (header.includes('事實')) width = 35;
        if (header.includes('立場')) width = 16;
        if (header.includes('金句') || header.includes('論理') || header.includes('理由')) width = 45;
        return { header, width };
    });

    // 填入資料行
    rows.forEach(r => worksheet.addRow(r));

    // 格式化表頭 (Navy blue 背景, 白字, 粗體, 置中)
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
        cell.font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1F4E78' }
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'medium' },
            right: { style: 'thin' }
        };
    });

    // 格式化內容行
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.alignment = { vertical: 'top', wrapText: true };
        row.font = { name: '微軟正黑體', size: 10 };

        // 斑馬紋底色
        const bgColor = rowNumber % 2 === 0 ? 'FFF9FAFC' : 'FFFFFFFF';
        row.eachCell((cell, colNumber) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColor }
            };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
            };
            // 立場與日期置中
            const header = headers[colNumber - 1] || '';
            if (header.includes('日期') || header.includes('立場')) {
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            }
        });
    });

    // 確保目標資料夾存在
    const dir = path.dirname(outputPath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    await workbook.xlsx.writeFile(outputPath);
    console.log(`Excel 矩陣已成功產出至: ${outputPath}`);
    return outputPath;
}

// 支援 CLI 執行: node scripts/export_matrix_to_excel.js <md_file_path> [output_xlsx_path]
if (require.main === module) {
    const mdPath = process.argv[2];
    const xlsxPath = process.argv[3] || (mdPath ? mdPath.replace(/\.md$/i, '.xlsx') : null);

    if (!mdPath || !xlsxPath) {
        console.log('用法: node scripts/export_matrix_to_excel.js <markdown檔案路徑> [xlsx輸出路徑]');
        process.exit(1);
    }

    exportToExcel(mdPath, xlsxPath)
        .then(() => process.exit(0))
        .catch(err => {
            console.error('匯出失敗:', err);
            process.exit(1);
        });
}

module.exports = { exportToExcel };
