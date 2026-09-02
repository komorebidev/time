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


    console.log("STEP 2: Values immediately after opening");

    const afterOpen = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    console.log(
        "AFTER OPEN:",
        JSON.stringify(afterOpen)
    );


    // ------------------------------------------------------------
    // WORKLOG EDITOR
    // ------------------------------------------------------------

    console.log("STEP 3: Filling worklog editor");

    const editor = page.locator(
        '[contenteditable="true"]'
    ).first();

    await editor.waitFor({
        state: "visible",
        timeout: 10000
    });

    await editor.fill(
        "Diagnostic worklog test"
    );

    await page.waitForTimeout(500);


    const afterEditor = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    console.log(
        "AFTER EDITOR:",
        JSON.stringify(afterEditor)
    );


    // ------------------------------------------------------------
    // START
    // ------------------------------------------------------------

    console.log("STEP 4: Setting Start");

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
        afterEditor,
        afterStart,
        afterWait
    };
}