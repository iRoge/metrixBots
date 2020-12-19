const http = require('http');
const puppeteer = require('puppeteer');
const xl = require('exceljs');
const axios = require('axios');

const hostname = 'localhost';
const port = 3001;

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
        args:[
            '--start-maximized',
            "--disable-notifications"
        ],
        defaultViewport: {
            width: 1920,
            height: 1080,
        }
    });
    const page = await browser.newPage();

    const workbook = new xl.Workbook();
    await workbook.xlsx.readFile('file.xlsx');

    const urlToParseColumn = 4;
    const urlToParseRow = 2;

    let collectedData = [];

    let i = 1;
    let worksheet = workbook.getWorksheet(i);
    while (worksheet) {
        let urlToParse = worksheet.getCell(urlToParseRow,urlToParseColumn).toString();
        if (!urlToParse) {
            console.log('Страница ' + worksheet.name + ' не имеет url страницы парсинга');
        }
        await page.goto(urlToParse);
        page.on('console', msg => {
            console.log(msg.text());
        });
        collectedData[i-1] = await page.evaluate(() => {
            let rankElement = document.querySelectorAll('.websiteRanks-valueContainer');
            let categoryElement = document.querySelector('li.js-categoryRank a.websiteRanks-nameText');
            let engagementElement = document.querySelectorAll('span.engagementInfo-valueNumber');
            return {
                globalRank: rankElement[0].textContent.trim().replace(',', ' '),
                countryRank: rankElement[1].textContent.trim().replace(',', ' '),
                categoryRank: rankElement[2].textContent.trim().replace(',', '.'),
                category: categoryElement.textContent.trim(),
                totalVisits: engagementElement[0].textContent.trim(),
                avgVisitsDuration: engagementElement[1].textContent.trim(),
                pagesPerVisit: engagementElement[2].textContent.trim().replace('.', ','),
                bounceRate: engagementElement[3].textContent.trim()
            };
        });
        axios.get('http://localhost:81?' + serialize({collectedData: collectedData})).then(function (response) {
            console.log(response.data);
        }).catch(function (error) {
            console.log(error);
        });
        return;
        console.log(collectedData);
        i++;
        console.log(worksheet.name + ' completed!');
        worksheet = workbook.getWorksheet(i);
    }

    await workbook.xlsx.writeFile('file1.xlsx');
    await console.log('done');
}

serialize = function(obj, prefix) {
    var str = [],
        p;
    for (p in obj) {
        if (obj.hasOwnProperty(p)) {
            var k = prefix ? prefix + "[" + p + "]" : p,
                v = obj[p];
            str.push((v !== null && typeof v === "object") ?
                serialize(v, k) :
                encodeURIComponent(k) + "=" + encodeURIComponent(v));
        }
    }
    return str.join("&");
}