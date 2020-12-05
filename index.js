require("dotenv").config();
const puppeteer = require("puppeteer");
var checkForParticipantsInteval;
var isLastMeeting = false;

async function mainFunction(){
    var browser = await puppeteer.launch({headless:false, args:["--start-maximized","--use-fake-ui-for-media-stream"],defaultViewport:null});
    var page = await browser.newPage();
    signin();

    async function signin(){
        await page.goto("https://teams.microsoft.com/_#/calendarv2");
        await page.waitForSelector("input[type='email']");
        await page.focus("input[type='email']");
        await page.keyboard.type(process.env.ID);
        await page.click("input[type='submit']");
        await page.waitForSelector("input[type='password']");
        await page.focus("input[type='password']");
        await page.keyboard.type(process.env.PASSWORD);
        await page.waitForTimeout(1000);
        await page.click("input[type='submit']");
        await page.waitForSelector("#idSIButton9");
        await page.click("#idSIButton9");
        await page.waitForSelector(".use-app-lnk");
        await page.click(".use-app-lnk");
        joinFirstMeeting();
    }

    async function joinFirstMeeting(){
        await page.waitForSelector("button[title='Switch your calendar view']");
        await page.click("button[title='Switch your calendar view']");
        await page.waitForSelector("button[name='Week']");
        await page.click("button[name='Week']");
        var date = myDate();
        await page.waitForXPath(`//div[text()='${date.date} ${date.month}']`);
        var parent1 = await page.$x(`//div[text()='${date.date} ${date.month}']`);
        var parent2 = (await parent1[0].$x('following-sibling::*'))[1];
        var child1 = await parent2.$("div");
        var child2 =  await child1.$("div");
        var firstClassTiming = await page.evaluate(el=>el.getAttribute("title"),child2);
        firstClassTiming = firstClassTiming.split("from ");
        firstClassTiming=firstClassTiming[1];
        firstClassTiming = firstClassTiming.split(" to ");
        firstClassTiming = firstClassTiming[0];
        var firstClassHours = (firstClassTiming.split(":"))[0];
        firstClassHours = parseInt(firstClassHours);
        var firstClassMins = (firstClassTiming.split(":"))[1];
        firstClassMins = parseInt(firstClassMins);
        var formatedTime = formatTime(firstClassHours,firstClassMins);
        joinMeeting({startHours:formatedTime.startHours,startMins:formatedTime.startMins,endHours:formatedTime.endHours,endMins:formatedTime.endMins,date:date.date,month:date.month});
    }

    async function joinMeeting(time){
        await idleTill(time.startHours,time.startMins);
        console.log(`join: ${time.date} ${time.month} ${time.startHours}:${time.startMins} to ${time.date} ${time.month} ${time.endHours}:${time.endMins}`);
        try{
            await page.waitForXPath(`//div[contains(@aria-label,'${time.date} ${time.month} ${time.startHours}:${time.startMins} to ${time.date} ${time.month} ${time.endHours}:${time.endMins}') and @role='button' and not(contains(@aria-label,'Canceled'))]`);
        }catch(e){
            nextMeeting(time.endHours,time.endMins);
            return;
        }
        var meetingButton = await page.$x(`//div[contains(@aria-label,'${time.date} ${time.month} ${time.startHours}:${time.startMins} to ${time.date} ${time.month} ${time.endHours}:${time.endMins}') and @role='button' and not(contains(@aria-label,'Canceled'))]`);
        await meetingButton[0].click();
        const example_parent = (await meetingButton[0].$x('..'))[0]; // Element Parent
        const example_siblings = await example_parent.$x('following-sibling::*'); 
        if(example_siblings.length==0){
            isLastMeeting=true;
        }
        
        await page.waitForXPath("//span[text() = 'Join']");
        var joinButton = await page.$x("//span[text()='Join']");
        await joinButton[0].click();

        //the bot was crashing without this ... L O L
        await page.waitForTimeout(1500);

        await page.waitForXPath("//button[text()='Join now']");
        var muteMicrophone = await page.$("span[title='Mute microphone']");
        if(muteMicrophone!=null){
            await muteMicrophone.click();
        }

        var joinNow = await page.$x("//button[text()='Join now']");
        await joinNow[0].click();

        //click show participants button so that we can get the numver of attendees
        await page.waitForXPath("//button[@aria-label='Show participants']");
        var showParticipants = await page.$x("//button[@aria-label='Show participants']");
        await showParticipants[0].click();
        console.log("joined class at:"+myDate().hours+":"+myDate().minutes);
        // check if prticipants are less than 25 after w8ing 20 mins since at the begining of the class paticipants might be less than 25 and we dont want to leave the meeting immediately
        setTimeout(() => {
          checkForParticipantsInteval = setInterval(() => {checkForParticipants(time.endHours,time.endMins)},1000*10);
        },10000);
    }

    async function checkForParticipants(endHours,endMins){
        await page.waitForXPath("//button[contains(@aria-label,'Attendees')]");
        var elHandle = await page.$x("//button[contains(@aria-label,'Attendees')]");
        var attendees = await page.evaluate((el) => el.getAttribute("aria-label"),elHandle[0]);
        attendees = attendees.replace("Attendees ","");
        var attendees = parseInt(attendees);
        if(attendees<25){
            hangup(endHours,endMins);
        }
        clearInterval(checkForParticipantsInteval);
    }

    async function hangup(endHours,endMins){
        var hangupButton = await page.$("#hangup-button");
        await hangupButton.click();
        console.log("left class at:"+myDate().hours+":"+myDate().minutes);
        if(!isLastMeeting){
        await page.goto("https://teams.microsoft.com/_#/calendarv2");
        nextMeeting(endHours,endMins);
        }else{
            browser.close();
        }
    }

    function nextMeeting(endHours,endMins){
        var date = myDate().date;
        var month = myDate().month;
        var startNext = endMins+10;
        if(startNext>=60){
            startNext = startNext-60;
            endHours++;
        }
        var formatedTime = formatTime(endHours,startNext);
        joinMeeting({startHours:formatedTime.startHours,startMins:formatedTime.startMins,endHours:formatedTime.endHours,endMins:formatedTime.endMins,date:date,month:month});
    }

    async function idleTill(tillHours,tillMins){
        if (typeof tillHours == "string"){
            tillHours=parseInt(tillHours);
        }
        if (typeof tillMins == "string"){
            tillMins=parseInt(tillMins);
        }
        var hours = myDate().hours;
        var minutes = myDate().minutes;
        var timeoutInMinutes = 1000*60*(tillMins-minutes);
        if(timeoutInMinutes==0){
            var timeoutInhours = 1000*60*60*(tillHours-hours);
        }else{
            var timeoutInhours = 1000*60*60*(tillHours-(hours+1));
        }
        var timeOut = timeoutInhours+timeoutInMinutes;
        console.log("timeout:"+timeOut);
        if(timeOut>0){
            setTimeout(() => {
                return Promise.resolve();
            },timeOut);
        }else{
            return Promise.resolve();
        }
    }

}

function myDate(){
    var full_date = new Date();
    var date = full_date.getDate();
    var hours = full_date.getHours();
    var minutes = full_date.getMinutes();
    var am_pm = "am";
    var month = full_date.toLocaleString('default', { month: 'long' });
    return {date:date,hours:hours,minutes:minutes,am_pm:am_pm,month:month};
}

function formatTime(hours,minutes){
    var startHours = hours < 10 ? hours.toLocaleString("en-US",{minimumIntegerDigits: 2, useGrouping:false}) : hours;
    var startMins = minutes < 10 ? minutes.toLocaleString("en-US",{minimumIntegerDigits: 2, useGrouping:false}) : minutes;
    var endMins = minutes+30;
    var endHours = hours;

    if( endMins >= 60){
        var difference = 60 - startMins;
        endMins = 30 - difference;
        endMins = endMins < 10 ? endMins.toLocaleString('en-US',{minimumIntegerDigits: 2, useGrouping:false}) : endMins;
        endHours++;
    }

    endHours = endHours < 10 ? endHours.toLocaleString("en-US",{minimumIntegerDigits: 2, useGrouping:false}) : endHours;

    return {
        startHours:startHours,
        startMins:startMins,
        endHours:endHours,
        endMins:endMins,
    }
}

mainFunction();