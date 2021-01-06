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
    const args = [
        '--no-sandbox',
        '--user-agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3312.0 Safari/537.36"',
        '--start-maximized',
        "--disable-notifications"
    ];
    const browser = await puppeteer.launch({
        headless: false,
        args: args,
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
        await autoScroll(page);
        page.on('console', msg => {
            console.log(msg.text());
        });
        collectedData[i-1] = await page.evaluate(() => {
            let rankElement = document.querySelectorAll('.websiteRanks-valueContainer');
            let categoryElement = document.querySelector('li.js-categoryRank a.websiteRanks-nameText');
            let engagementElement = document.querySelectorAll('span.engagementInfo-valueNumber');

            let directElement = document.querySelector('li.trafficSourcesChart-item.source-direct div.trafficSourcesChart-value');
            let referralsElement = document.querySelector('li.trafficSourcesChart-item.source-referrals div.trafficSourcesChart-value');
            let searchElement = document.querySelector('li.trafficSourcesChart-item.source-search div.trafficSourcesChart-value');
            let socialElement = document.querySelector('li.trafficSourcesChart-item.source-social div.trafficSourcesChart-value');
            let mailElement = document.querySelector('li.trafficSourcesChart-item.source-mail div.trafficSourcesChart-value');
            let displayElement = document.querySelector('li.trafficSourcesChart-item.source-display div.trafficSourcesChart-value');

            let countriesInfo = {};
            let countriesBlocks = document.querySelectorAll('div.accordion.countries-list div.accordion-group');
            for (let countryBlock of countriesBlocks) {
                let countryNameBlock = countryBlock.querySelector('span.country-name');
                let countryName;
                if (!countryNameBlock) {
                    countryName = countryBlock.querySelector('span.country-container a').textContent.trim();
                } else {
                    countryName = countryNameBlock.textContent.trim();
                }

                let percentSpan = countryBlock.querySelector('span.traffic-share-valueNumber');
                let differenceSpan = countryBlock.querySelector('span.websitePage-relativeChangeNumber');
                let percent = percentSpan ? percentSpan.textContent.trim() : null;
                let difference = null;

                if (differenceSpan) {
                    if (countryBlock.querySelector('span.websitePage-relativeChange--down')) {
                        difference = '-' + differenceSpan.textContent.trim();
                    } else {
                        difference = differenceSpan.textContent.trim();
                    }
                }

                countriesInfo[countryName.toLowerCase()] = {
                    percent: percent,
                    difference: difference,
                }
            }

            let topReferringSitesInfo = [];
            let topReferringSitesBlocks = document.querySelectorAll('div.referralsSites.referring ul.websitePage-list li.websitePage-listItem');
            for (let topReferringSitesBlock of topReferringSitesBlocks) {
                let siteNameBlock = topReferringSitesBlock.querySelector('a.websitePage-listItemLink');
                let siteName;
                if (!siteNameBlock) {
                    console.log('Не найден один из блоков sitesRefferer на странице');
                } else {
                    siteName = siteNameBlock.textContent.trim();
                }

                let percentSpan = topReferringSitesBlock.querySelector('span.websitePage-trafficShare');
                let differenceSpan = topReferringSitesBlock.querySelector('span.websitePage-relativeChangeNumber');

                let percent = percentSpan ? percentSpan.textContent.trim() : null;
                let difference = null;

                if (differenceSpan) {
                    if (topReferringSitesBlock.querySelector('div.websitePage-relativeChange--down')) {
                        difference = '-' + differenceSpan.textContent.trim();
                    } else {
                        difference = differenceSpan.textContent.trim();
                    }
                }

                topReferringSitesInfo.push({
                    siteName: siteName,
                    percent: percent,
                    difference: difference,
                });
            }

            let topDestinationSitesInfo = [];
            let topDestinationSitesBlocks = document.querySelectorAll('div.referralsSites.destination ul.websitePage-list li.websitePage-listItem');
            for (let topDestinationSitesBlock of topDestinationSitesBlocks) {
                let siteNameBlock = topDestinationSitesBlock.querySelector('a.websitePage-listItemLink');
                let siteName;
                if (!siteNameBlock) {
                    console.log('Не найден один из блоков topDestinationSites на странице');
                } else {
                    siteName = siteNameBlock.textContent.trim();
                }

                let percentSpan = topDestinationSitesBlock.querySelector('span.websitePage-trafficShare');
                let differenceSpan = topDestinationSitesBlock.querySelector('span.websitePage-relativeChangeNumber');

                let percent = percentSpan ? percentSpan.textContent.trim() : null;
                let difference = null;

                if (differenceSpan) {
                    if (topDestinationSitesBlock.querySelector('div.websitePage-relativeChange--down')) {
                        difference = '-' + differenceSpan.textContent.trim();
                    } else {
                        difference = differenceSpan.textContent.trim();
                    }
                }

                topDestinationSitesInfo.push({
                    siteName: siteName,
                    percent: percent,
                    difference: difference,
                });
            }

            let organicSearchPercentBlock = document.querySelector('div.searchPie div.searchPie-text--left span.searchPie-number');
            let organicSearchPercent = organicSearchPercentBlock ? organicSearchPercentBlock.textContent.trim() : null;
            let organicSearchInfo = [];
            let organicSearchBlocks = document.querySelectorAll('div.searchKeywords-text--left li.searchKeywords-row');
            for (let organicSearchBlock of organicSearchBlocks) {
                let searchTextBlock = organicSearchBlock.querySelector('span.searchKeywords-words');
                let searchText;
                if (!searchTextBlock) {
                    console.log('Не найден один из блоков organicSearch на странице');
                } else {
                    searchText = searchTextBlock.textContent.trim();
                }

                let percentSpan = organicSearchBlock.querySelector('span.searchKeywords-trafficShare');
                let differenceSpan = organicSearchBlock.querySelector('span.websitePage-relativeChangeNumber');

                let percent = percentSpan ? percentSpan.textContent.trim() : null;
                let difference = null;

                if (differenceSpan) {
                    if (organicSearchBlock.querySelector('span.websitePage-relativeChange--down')) {
                        difference = '-' + differenceSpan.textContent.trim();
                    } else {
                        difference = differenceSpan.textContent.trim();
                    }
                }

                organicSearchInfo.push({
                    searchText: searchText,
                    percent: percent,
                    difference: difference,
                });
            }

            let paidSearchPercentBlock = document.querySelector('div.searchPie div.searchPie-text--right span.searchPie-number');
            let paidSearchPercent = paidSearchPercentBlock ? paidSearchPercentBlock.textContent.trim() : null;
            let paidSearchInfo = [];
            let paidSearchBlocks = document.querySelectorAll('div.searchKeywords-text--right li.searchKeywords-row');
            for (let paidSearchBlock of paidSearchBlocks) {
                let searchTextBlock = paidSearchBlock.querySelector('span.searchKeywords-words');
                let searchText;
                if (!searchTextBlock) {
                    console.log('Не найден один из блоков organicSearch на странице');
                } else {
                    searchText = searchTextBlock.textContent.trim();
                }

                let percentSpan = paidSearchBlock.querySelector('span.searchKeywords-trafficShare');
                let differenceSpan = paidSearchBlock.querySelector('span.websitePage-relativeChangeNumber');

                let percent = percentSpan ? percentSpan.textContent.trim() : null;
                let difference = null;

                if (differenceSpan) {
                    if (paidSearchBlock.querySelector('span.websitePage-relativeChange--down')) {
                        difference = '-' + differenceSpan.textContent.trim();
                    } else {
                        difference = differenceSpan.textContent.trim();
                    }
                }

                paidSearchInfo.push({
                    searchText: searchText,
                    percent: percent,
                    difference: difference,
                });
            }

            let socialInfo = {};
            let socialBlocks = document.querySelectorAll('div.socialSection ul.socialList li.socialItem');
            for (let socialBlock of socialBlocks) {
                let socialNameBlock = socialBlock.querySelector('a.socialItem-title');
                let socialName;
                if (!socialNameBlock) {
                    console.log('Не найден один из блоков social на странице');
                } else {
                    socialName = socialNameBlock.textContent.trim();
                }

                let percentSpan = socialBlock.querySelector('div.socialItem-value');
                let percent = percentSpan ? percentSpan.textContent.trim() : null;

                socialInfo[socialName.toLowerCase()] = {
                    percent: percent,
                }
            }

            return {
                globalRank: rankElement[0] ? rankElement[0].textContent.trim().replace(',', ' ') : null,
                countryRank: rankElement[1] ? rankElement[1].textContent.trim().replace(',', ' ') : null,
                categoryRank: rankElement[2] ? rankElement[2].textContent.trim().replace(',', '.') : null,
                category: categoryElement ? categoryElement.textContent.trim() : null,
                totalVisits: engagementElement[0] ? engagementElement[0].textContent.trim() : null,
                avgVisitsDuration: engagementElement[1] ? engagementElement[1].textContent.trim() : null,
                pagesPerVisit: engagementElement[2] ? engagementElement[2].textContent.trim().replace('.', ',') : null,
                bounceRate: engagementElement[3] ? engagementElement[3].textContent.trim() : null,
                directPercent: directElement ? directElement.textContent.trim() : null,
                referralsPercent: referralsElement ? referralsElement.textContent.trim() : null,
                searchPercent: searchElement ? searchElement.textContent.trim() : null,
                socialPercent: socialElement ? socialElement.textContent.trim() : null,
                mailPercent: mailElement ? mailElement.textContent.trim() : null,
                displayPercent: displayElement ? displayElement.textContent.trim() : null,
                countriesInfo: countriesInfo,
                topReferringSitesInfo: topReferringSitesInfo,
                topDestinationSitesInfo: topDestinationSitesInfo,

                organicSearchPercent: organicSearchPercent,
                organicSearchInfo: organicSearchInfo,
                paidSearchPercent: paidSearchPercent,
                paidSearchInfo: paidSearchInfo,
                socialInfo: socialInfo,
            };
        });
        console.log(collectedData);
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

function serialize(obj, prefix){
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

async function autoScroll(page){
    await page.evaluate(async () => {
        await new Promise((resolve, reject) => {
            var totalHeight = 0;
            var distance = 100;
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;

                if(totalHeight >= scrollHeight){
                    clearInterval(timer);
                    resolve();
                }
            }, 100);
        });
    });
}