const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const targetLabels = ['รายได้รวม', 'กำไรขั้นต้น'];

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
        return { symbol, quarters: [], data: {} }; // return empty structure
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

// 🔁 ดึงหลายหุ้น + Export เป็น Excel
(async () => {
    const symbols = ['SCC', 'SCB', 'XXXX'];
    const allData = [];
    let allQuarters = new Set();

    const resultMap = new Map();

    for (const symbol of symbols) {
        console.log(`📊 ดึงข้อมูล ${symbol}...`);
        const result = await getStockFinancialSummary(symbol);
        if (result) {
            resultMap.set(symbol, result);
            result.quarters.forEach(q => allQuarters.add(q));
        }
    }

    allQuarters = Array.from(allQuarters);

    // จัด structure ให้เหมือนกันหมด
    for (const symbol of symbols) {
        const result = resultMap.get(symbol);
        const row = { 'หุ้น': symbol };

        for (const quarter of allQuarters) {
            for (const label of targetLabels) {
                const key = `${label}_${quarter}`;
                const value = result?.data?.[quarter]?.[label] || '';
                row[key] = value;
            }
        }

        allData.push(row);
    }

    // ตั้งชื่อไฟล์ตามวันเวลา
    const now = new Date();
    const pad = n => n.toString().padStart(2, '0');
    const filename = `financial_summary_${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}_${pad(now.getHours())}${pad(now.getMinutes())}.xlsx`;

    // สร้างโฟลเดอร์ 'result' ถ้ายังไม่มี
    const resultFolder = path.join(__dirname, 'result');
    if (!fs.existsSync(resultFolder)) {
        fs.mkdirSync(resultFolder);
    }

    // สร้างไฟล์ Excel
    const filePath = path.join(resultFolder, filename);
    const worksheet = XLSX.utils.json_to_sheet(allData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summary');
    XLSX.writeFile(workbook, filePath);

    console.log(`✅ บันทึกไฟล์ Excel เรียบร้อย: ${filePath}`);
})();
