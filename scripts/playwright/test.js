async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const before = await start.inputValue();

    console.log("BEFORE:", before);

    await start.click();

    console.log(
        "AFTER CLICK:",
        await start.inputValue(),
        "focused:",
        await start.evaluate(
            el => document.activeElement === el
        )
    );

    await start.fill("09:00");

    console.log(
        "AFTER FILL:",
        await start.inputValue()
    );

    await page.waitForTimeout(100);

    console.log(
        "AFTER 100MS:",
        await start.inputValue()
    );

    await page.waitForTimeout(1500);

    console.log(
        "AFTER 1.5 SEC:",
        await start.inputValue()
    );

    return {
        before,
        afterFill: await start.inputValue()
    };
}