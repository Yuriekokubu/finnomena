const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const targetLabels = ['รายได้รวม', 'กำไรขั้นต้น'];

// Function to convert quarter string to a comparable date format
function parseQuarter(quarter) {
    const quarterMap = { '1Q': 1, '2Q': 4, '3Q': 7, '4Q': 10 };
    const dateRegex = /(\d)Q(\d{4})/;
    const specialDateRegex = /(\d{1,2}) (.*?) (\d{2})/; // Modified regex

    const dateMatch = quarter.match(dateRegex);
    const specialDateMatch = quarter.match(specialDateRegex);

    if (dateMatch) {
        const q = parseInt(dateMatch[1]);
        const year = parseInt(dateMatch[2]);
        const month = quarterMap[dateMatch[1] + 'Q'];
        return { date: new Date(year, month - 1, 1), type: 'quarter' }; // Month is 0-indexed
    } else if (specialDateMatch) {
        // Handle special date format
        const day = parseInt(specialDateMatch[1]);
        const monthStr = specialDateMatch[2];
        let year = parseInt(specialDateMatch[3]);
        // Adjust the year to be in the 2000s
        year = 2000 + year;

        const monthMap = {
            'ม.ค.': 0, 'ก.พ.': 1, 'มี.ค.': 2, 'เม.ย.': 3, 'พ.ค.': 4, 'มิ.ย.': 5,
            'ก.ค.': 6, 'ส.ค.': 7, 'ก.ย.': 8, 'ต.ค.': 9, 'พ.ย.': 10, 'ธ.ค.': 11
        };
        const month = monthMap[monthStr];
        return { date: new Date(year, month, day), type: 'special' };
    } else {
        return null; // Return null for unknown formats
    }
}

// Function to compare quarters for sorting
function compareQuarters(a, b) {
    const parsedA = parseQuarter(a);
    const parsedB = parseQuarter(b);

    if (parsedA === null && parsedB === null) return 0;
    if (parsedA === null) return 1;
    if (parsedB === null) return -1;

    if (parsedA.type === 'special' && parsedB.type !== 'special') return 1;
    if (parsedB.type === 'special' && parsedA.type !== 'special') return -1;


    return parsedA.date - parsedB.date;
}


async function getStockFinancialSummary(symbol) {
    const browser = await puppeteer.launch({
        headless: true,
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    await page.goto(`https://www.finnomena.com/stock/${symbol}`, { waitUntil: 'networkidle2' });

    const notFoundImage = await page.$('img[alt="ไม่พบหน้าที่คุณค้นหา"]');
    if (notFoundImage) {
        console.warn(`❌ หุ้น ${symbol} ไม่พบหน้าที่คุณค้นหา — ข้าม...`);
        await browser.close();
        return { symbol, quarters: [], data: {} };
    }

    try {
        await page.waitForSelector('section[data-fn-location="stock-finance-40q"]', { timeout: 10000 });
    } catch (err) {
        console.warn(`⚠️ ไม่พบข้อมูลของหุ้น ${symbol} — ข้าม...`);
        await browser.close();
        return { symbol, quarters: [], data: {} };
    }

    await page.evaluate(async () => {
        let scrollHeight = 0;
        const distance = 2000;
        while (scrollHeight < document.body.scrollHeight) {
            window.scrollBy(0, distance);
            scrollHeight += distance;
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
    });

    const stockCode = await page.$eval(
        'div[data-alias="stock_code_financial_summary"]',
        el => el.textContent.trim()
    );

    const quarters = await page.$$eval(
        'div[data-alias^="financial_summary_header_"]',
        els => els.map(el => el.textContent.trim())
    );

    const rows = await page.$$eval(
        'table.text-colors-text-secondary tbody tr',
        trs => trs.map(tr => {
            const label = tr.querySelector('td:first-child')?.textContent?.trim();
            const values = Array.from(tr.querySelectorAll('td:not(:first-child) span')).map(
                span => span.textContent.trim()
            );
            return { label, values };
        })
    );

    const data = {};
    quarters.forEach((q, idx) => {
        data[q] = {};
        targetLabels.forEach(label => {
            const foundRow = rows.find(r => r.label === label);
            data[q][label] = foundRow ? (foundRow.values[idx] || '') : '';
        });
    });

    await browser.close();
    return { symbol: stockCode, quarters, data };
}

(async () => {
    // const symbols = ['AH', 'GYT', 'HFT', 'IHL', 'SAT', 'STANLY', 'TRU', 'AKR', 'ASEFA', 'CPT', 'HTECH', 'SELIC', 'UTP', 'CMAN', 'PTTGC', 'SCPG', 'SFLEX', 'SITHAI', 'SMPC', 'THIP', 'TPBI', '2S', 'AMC', 'BSBM', 'GJS', 'GSTEEL', 'LHK', 'MCS', 'PAP', 'PERM', 'SAM', 'TMT', 'TSTH', 'TYCN', 'NER', 'STA', 'TEGH', 'TFM', 'TRUBB', 'UVAN', 'UPOIC', 'AAI', 'APURE', 'ASIAN', 'BR', 'BRR', 'BTG', 'CBG', 'CFRESH', 'CHAO', 'COCOCO', 'CPF', 'F&D', 'HTC', 'ICHI', 'ITC', 'KCG', 'KSL', 'MALEE', 'NSL', 'OSP', 'PLUS', 'PM', 'RBF', 'SAPPE', 'SAUCE', 'SORKON', 'SUN', 'TC', 'TFG', 'TFMAMA', 'TKN', 'TU', 'TVO', 'TWPC', 'OR', 'PTG', 'SCC', 'M-CHAI', 'AHC', 'CHG', 'CMR', 'EKH', 'PHG', 'RJH', 'RPH', 'WPH', 'CPALL', 'DOHOME', 'GLOBAL', 'MEGA', 'VRANDA'];
    const symbols = ['GSTEEL', 'GJS'];
    const allData = [];
    const allQuartersSet = new Set();
    const resultMap = new Map();

    for (const symbol of symbols) {
        console.log(`📊 ดึงข้อมูล ${symbol}...`);
        const result = await getStockFinancialSummary(symbol);
        if (result) {
            resultMap.set(symbol, result);
            result.quarters.forEach(q => allQuartersSet.add(q));
        }
    }

    // Sort the quarters
    const allQuarters = Array.from(allQuartersSet).sort(compareQuarters);

    const resultRows = [];
    for (const symbol of symbols) {
        const result = resultMap.get(symbol);

        targetLabels.forEach(label => {
            const row = { 'หุ้น': symbol, 'ประเภท': label };
            for (const q of allQuarters) {
                // Check if the value exists and is not null before applying .slice()
                const value = result?.data?.[q]?.[label] || '';
                row[q] = value ? value.slice(0, 10) : ''; // Safely handle slicing
            }
            resultRows.push(row);
        });
    }

    const worksheet = XLSX.utils.json_to_sheet(resultRows);

    // 🔀 จัด merge cell เฉพาะคอลัมน์ A (หุ้น)
    worksheet['!merges'] = [];
    let mergeStart = 1; // เริ่มต้นที่ row 2 (เพราะ header row คือ 1)
    let previousSymbol = resultRows[0]?.หุ้น;

    for (let i = 0; i < resultRows.length; i++) {
        const currentSymbol = resultRows[i].หุ้น;
        if (currentSymbol !== previousSymbol) {
            if (i - mergeStart > 0) { // Make sure there are rows to merge
                worksheet['!merges'].push({
                    s: { r: mergeStart, c: 0 },
                    e: { r: i , c: 0 }
                });
            }
            mergeStart = i + 1; // Start merging from the next row
            previousSymbol = currentSymbol;
        }
    }
     // Merge the last group if needed
    if (resultRows.length - mergeStart > 0) {
        worksheet['!merges'].push({
            s: { r: mergeStart, c: 0 },
            e: { r: resultRows.length , c: 0 }
        });
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summary');

    const now = new Date();
    const pad = n => n.toString().padStart(2, '0');
    const filename = `financial_summary_${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}_${pad(now.getHours())}${pad(now.getMinutes())}.xlsx`;

    const resultFolder = path.join(__dirname, 'result');
    if (!fs.existsSync(resultFolder)) {
        fs.mkdirSync(resultFolder);
    }

    const filePath = path.join(resultFolder, filename);
    XLSX.writeFile(workbook, filePath);
    console.log(`✅ บันทึกไฟล์ Excel เรียบร้อย: ${filePath}`);
})();
