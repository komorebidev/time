async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    console.log("START BEFORE:", await start.inputValue());

    await start.fill("08:00");

    console.log("START IMMEDIATELY:", await start.inputValue());

    await page.waitForTimeout(1500);

    console.log("START AFTER 1.5 SEC:", await start.inputValue());
}