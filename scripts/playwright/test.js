async page => {

    console.log("STEP 1: Opening Worklog");

    await page.getByRole(
        "button",
        { name: "Worklog" }
    ).click();

    await page.waitForTimeout(1500);


    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );


    console.log("STEP 2: Reading freshly opened values");

    const afterOpen = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    console.log(
        "AFTER OPEN:",
        JSON.stringify(afterOpen)
    );


    console.log("STEP 3: Setting Start ONLY");

    await start.click();

    await start.fill("08:00");


    const afterStart = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    console.log(
        "AFTER START:",
        JSON.stringify(afterStart)
    );


    await page.waitForTimeout(1500);


    const afterWait = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    console.log(
        "AFTER WAIT:",
        JSON.stringify(afterWait)
    );


    return {
        afterOpen,
        afterStart,
        afterWait
    };
}