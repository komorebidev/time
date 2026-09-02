async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const before = await start.inputValue();

    await start.fill("08:00");

    const immediately = await start.inputValue();

    await page.waitForTimeout(1500);

    const after = await start.inputValue();

    return {
        before,
        immediately,
        after
    };
}