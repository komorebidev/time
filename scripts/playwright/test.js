async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    console.log(
        "BEFORE:",
        await start.inputValue(),
        await end.inputValue()
    );

    await start.fill("08:00");

    console.log(
        "AFTER START:",
        await start.inputValue(),
        await end.inputValue()
    );

    await page.waitForTimeout(1500);

    console.log(
        "AFTER WAIT:",
        await start.inputValue(),
        await end.inputValue()
    );
}