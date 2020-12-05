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

    const urlToParseColumn = 4;
    const urlToParseRow = 2;
    const startRow = 6;
    let i = 1;
    let worksheet = workbook.getWorksheet(i);
    while (worksheet) {
        let urlToParse = worksheet.getCell(urlToParseRow,urlToParseColumn).toString();
        if (!urlToParse) {
            console.log('Страница ' + worksheet.name + ' не имеет url страницы парсинга');
        }
        console.log(worksheet.getRows(1, 1));
        worksheet.insertRows();
        let currentRow = startRow + i * 7;
        worksheet.getCell(currentRow,1).value = 'ЕЖЕМЕСЯЧНО МЕНЯЮЩИЕСЯ ДАННЫЕ';
        i++;
        worksheet = workbook.getWorksheet(i);
    }



}