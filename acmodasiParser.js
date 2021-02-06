const http = require('http');
const puppeteer = require('puppeteer');
const xl = require('exceljs');
const fs = require('fs');
const mailer = require('nodemailer');

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
        "--disable-notifications",
        '--disable-web-security',
        // '--proxy-server=188.113.190.7:80'
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
    const countOfPages = 68;
    await page.setDefaultNavigationTimeout(0);

    // page.on('console', msg => {
    //     console.log(msg.text());
    // });

    const workbook = new xl.Workbook();
    await workbook.xlsx.readFile('acmodasi.xlsx');
    let worksheet = workbook.getWorksheet(1);

    await page.setCookie(
        {
            "name": "acmodasi",
            "value": "0280b176042078360ebaf0a7aa43830f",
            "domain": "www.acmodasi.ru",
            "path": "/",
            "expires": 1633622615.629729,
            "httpOnly": false,
            "secure": false,
            "session": false,
        }
    );

    for (let i = 10, row = 1; i <= countOfPages; i++) {
        await page.goto('https://www.acmodasi.ru/index.php?action=search&todo=advanced&base=&un=&vo=4&vd=17&pvo=&pvd=&us=0&uc=&c=219&eda=0&ac=0&edm=0&mo=0&edd=0&da=0&eds=0&si=0&bt=&et=&hf=&ht=&wf=&wt=&hc=0&lf=0&lt=0&eye=0&cf=&ct=&sf=&st=&pp=3&on=1&sp=0&id=0&ph=0&p=' + i);

        let ankets = await page.$$('.user_card_image_box');
        for (let anket of ankets) {
            await anket.click();
            await page.waitForTimeout(2000);
            let child = {};
            let tabContent = await page.$('#pills-tabContent .col-md-6.col-sm-12.mt-4');
            child['name'] = await tabContent.$$eval('div.row', (elements) => {
                for (let e of elements) {
                    let propName = e.querySelector('div.table_label').innerHTML;
                    if (propName === 'ФИО') {
                        return e.querySelector('div.table_value').innerHTML;
                    }
                }
            });
            child['gender'] = await tabContent.$$eval('div.row', (elements) => {
                for (let e of elements) {
                    let propName = e.querySelector('div.table_label').innerHTML;
                    if (propName === 'Пол') {
                        return e.querySelector('div.table_value').innerHTML;
                    }
                }
            });
            await page.click('#p_contacts');
            await page.waitForTimeout(1000);

            if (await page.evaluate(() => { return !!document.getElementById("show_contact_buton") })) {
                await page.click('#show_contact_buton');
            }
            await page.waitForTimeout(500);
            let contactRows = await page.$$('div#user_contacts div.row');
            child['phone'] = null;
            for (let contactRow of contactRows) {
                let nameOfService = await contactRow.$eval('div.table_label', (e) => e.innerHTML);
                if (nameOfService === 'Whatsapp' || nameOfService === 'Телефон:' || nameOfService === 'Viber') {
                    child['phone'] = await (await contactRow.$eval('div.table_value a', (e) => e.innerHTML)).replace(/\s|\(|\)|-/g, '');
                    break;
                }
            }
            if (child.hasOwnProperty('name') && child.hasOwnProperty('gender') && child.hasOwnProperty('phone')) {
                worksheet.getCell(row, 1).value = child['name'];
                worksheet.getCell(row, 2).value = child['phone'];
                worksheet.getCell(row, 3).value = child['gender'];
            }
            row++;
            // let form = page.$eval('form input.btn.orange_button.mt-2[type=submit]', (e) => e.parentElement);
            // await form.type('textarea.form-control', 'Приглашаем молодого музыканта на конкурс. Подробности: kidfest.ru');
            // await form.click('textarea.form-control');
            await page.click('div.modal-footer button.btn');
            await page.waitForTimeout(1000);
        }
        await console.log('Page number ' + i + ' completed!');
        break;
    }
    await workbook.xlsx.writeFile('acmodasi.xlsx');

    await fs.readFile("./acmodasi.xlsx", function (err, data) {
        let transporter = mailer.createTransport({
            host: 'smtp.yandex.ru',
            port: 465,
            secure: true,
            auth: {
                user: 'iroge-27@yandex.ru',
                pass: 'V04@r70n1V'
            }
        });
        transporter.sendMail({
            sender: 'iroge-27@yandex.ru',
            to: 'iroge-27@yandex.ru',
            subject: 'Данные acmodasi',
            body: '',
            attachments: [{'filename': 'acmodasi.xlsx', 'content': data}]
        });
    });
}

function serialize(obj, prefix)
{
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

async function autoScroll(page)
{
    return await page.evaluate(async () => {
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
            }, 55);
        });
    });
}

async function dealWithRecaptcha(page)
{
    let sitekey = await page.evaluate(() => {
        let captchaBlock = document.querySelector('div.g-recaptcha');
        return captchaBlock.dataset.sitekey;
    });
    let data = {
        key: '695355b02869d2f575b6e89201672a71',
        googlekey: sitekey,
        method: 'userrecaptcha',
        pageurl: page.url(),
        json: 1,
    };
    console.log('Sending request to RuCaptcha...');
    let response = await axios.get('https://rucaptcha.com/in.php?' + serialize(data), {headers: {'Content-Type': 'application/x-www-form-urlencoded'}});

    if (response.data.status) {
        let data = {
            key: '695355b02869d2f575b6e89201672a71',
            action: 'get',
            id: response.data.request,
            json: 1,
        };
        return await new Promise((resolve, reject) => {
            let timer = setInterval(() => {
                (async () => {
                    let response = await axios.get('https://rucaptcha.com/res.php?' + serialize(data), {headers: {'Content-Type': 'application/x-www-form-urlencoded'}});
                    if (response.data) {
                        if (response.data.status === 1) {
                            clearInterval(timer);
                            await page.evaluate(() => {
                                let textarea = document.querySelector('div.g-recaptcha textarea[id="g-recaptcha-response"]');
                                textarea.style = '';
                            });
                            await page.type('div.g-recaptcha textarea[id="g-recaptcha-response"]', response.data.request, {delay: 15});
                            await page.evaluate((code) => {
                                let textarea = document.querySelector('div.g-recaptcha textarea[id="g-recaptcha-response"]');
                                let input = document.createElement('input');
                                input.setAttribute('type', 'button');
                                input.setAttribute('id', 'handleCaptcha');
                                input.setAttribute('onclick', 'handleCaptcha("' + code + '")');
                                input.style.width = 400;
                                input.style.height = 400;
                                input.style.position = 'absolute';
                                input.style.zIndex = 100000;
                                textarea.after(input, textarea);
                                // handleCaptcha(code);
                            }, response.data.request);
                            await page.click('input[id=handleCaptcha]');
                            await resolve();
                        } else if (response.data.status !== 0) {
                            await reject();
                        }
                    }
                })();
            }, 1000)
        }).then(() => {
            return true;
        }).catch(() => {
            return false;
        });
    } else {
        console.log('При отправке запроса в RuCaptcha произошла ошибка');
        return false;
    }
}