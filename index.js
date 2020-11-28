const http = require('http');
const puppeteer = require('puppeteer');
const xl = require('exceljs');

const hostname = 'localhost';
const port = 3000;

const server = http.createServer();

server.listen(port, hostname, () => {
    start();
});

function start() {
    (async () => {
        await analizeSites();
    })()
}

async function analizeSites() {
    const browser = await puppeteer.launch({
        headless: false,
    });
    const page = await browser.newPage();

    const workbook = new xl.Workbook();
    await workbook.xlsx.readFile('file.xlsx');
    let worksheet = workbook.getWorksheet(1);

    let urlRow = 4;
    let siteUrl = worksheet.getCell(urlRow,3).toString();

    while (siteUrl) {
        await page.goto(siteUrl);
        await page.waitFor(3000);
        urlRow++;
        siteUrl = worksheet.getCell(urlRow,3).toString();
    }

}