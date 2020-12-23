const xlsx = require('xlsx');
const puppeteer = require('puppeteer');

/*
Promise.prototype.finally = function(onFinally) {
  return this.then(
    // onFulfilled 
    res => Promise.resolve(onFinally()).then(() => res),
    // onRejected 
    err => Promise.resolve(onFinally()).then(() => { throw err; })
  );
};
*/


const workbook = xlsx.readFile('Credentials.xlsx');
var firstSheet = workbook.SheetNames[0];

var credentialsObjArr = xlsx.utils.sheet_to_json(workbook.Sheets[firstSheet]);
console.log(credentialsObjArr);


credentialsObjArr.forEach((obj) => {

    var browser;
    (async () => {




        browser = await puppeteer.launch({
            headless: false
        //    args: ['--proxy-server=http://10.24.19.246:8080']
        });




        let mail = obj.Email;
        let pwd = obj.Password;

        const page = await browser.newPage();

        await page.goto('https://developer.servicenow.com/app.do#!/home', {
            timeout: 0
        });

        await page.waitFor('#dp-hdr-login-link')

        await Promise.all([page.click('#dp-hdr-login-link'), page.waitForNavigation({
            timeout: 0
        })]);
        await page.waitFor('#username');
		
        await page.$eval('#username', (elem, mail) => {
            elem.value = mail
        }, mail);

		await page.click('#usernameSubmitButton',{timeout:0});
		
        await page.$eval('#password', (elem, pwd) => {
            elem.value = pwd;
        }, pwd);

		await page.waitFor(1000);
        await Promise.all([page.click('#submitButton'), page.waitForNavigation({
            timeout: 0
        })]);
		
		console.log('70');
        
		await page.waitFor('#dp-hdr-br-manage-link');
        let result = await page.$('#dp-hdr-br-manage-link');


        await page.click('#dp-side-menu-toggler-sm', {
            timeout: 0
        });

        await page.click('#dp-hdr-sm-manage-link', {
            timeout: 0
        });
        await page.click('#dp-hdr-sm-link-instance', {
            timeout: 0
        });
        await page.waitFor(10000);
        let isWakenHandle = await page.$('[class="instance-info-state ng-binding"]');

        const val = await (await isWakenHandle.getProperty('textContent')).jsonValue();

        if (val.includes('Online')) {
            console.log('Here');
            let extendElement = await page.$('dp-instance-extend-button-disabled');
            if (extendElement == undefined)
                page.click('#dp-instance-extend-button')
        } else {
            console.log('Here')
            await page.click('#instanceWakeUpBtn');

        }




        //      browser.close();


    })()
    .catch((err) => {
        console.log("Error catched ==============================> " + err);
        console.log("--------------------------------------------------------------------------------------------");

        //  browser.close();
    })




});

//browser.close();