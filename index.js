const puppeteer = require('puppeteer');

async function getStockFinancialSummary(symbol) {
    const browser = await puppeteer.launch({
        headless: false,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--enable-third-party-cookies' // ⬅️ Optional, usually enabled by default
        ]
    });
    const page = await browser.newPage();
    await page.goto('https://www.finnomena.com/stock', { waitUntil: 'networkidle2' });


    await page.type('input.input-search', symbol);
    await page.waitForSelector('ul.list-group-custom > li.list-group-item', { timeout: 10000 });

    const items = await page.$$('ul.list-group-custom > li.list-group-item');
    for (const item of items) {
        const dataParams = await item.evaluate(el => el.getAttribute('data-fn-params'));
        const stockName = JSON.parse(dataParams)['stock-name'];
        if (stockName && stockName.toLowerCase() === symbol.toLowerCase()) {
            await Promise.all([
                item.click(),
                page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 20000 })
            ]);
            break;
        }
    }

    await page.waitForSelector('section[data-fn-location="stock-finance-40q"]', { timeout: 30000 });

    await page.evaluate(async () => {
        let scrollHeight = 0;
        const distance = 1000;
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

    const targetLabels = ['รายได้รวม', 'กำไรขั้นต้น'];
    const finalData = {};

    quarters.forEach((quarter, index) => {
        const quarterData = {};
        rows.forEach(row => {
            if (targetLabels.includes(row.label)) {
                quarterData[row.label] = row.values[index] || '';
            }
        });
        finalData[quarter] = quarterData;
    });

    await browser.close();
    return { [stockCode]: finalData };
}

// 🔁 ดึงหลายหุ้น
(async () => {
    const symbols = ['SCC', 'SCB'];
    const results = [];

    for (const symbol of symbols) {
        console.log(`📊 กำลังดึงข้อมูลหุ้น ${symbol}...`);
        const data = await getStockFinancialSummary(symbol);
        results.push(data);
    }

    console.log('📦 ข้อมูลทั้งหมด:');
    console.log(JSON.stringify(results, null, 2));
})();
