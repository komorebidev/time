async page => {

    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    const report = {};

    // ------------------------------------------------------------
    // Initial
    // ------------------------------------------------------------

    report.initial = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // Set START
    // ------------------------------------------------------------

    await start.fill("08:00");

    report.afterStart = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // Set END
    // ------------------------------------------------------------

    await end.fill("09:00");

    report.afterEnd = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // Open STATUS
    // ------------------------------------------------------------

    const statusCombobox = page.getByRole(
        "combobox",
        { name: "Status *" }
    );

    await statusCombobox.click();

    await page.waitForTimeout(500);

    report.afterStatusOpen = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // Select STATUS
    // ------------------------------------------------------------

    const statusOption = page.locator(
        ".Select__option",
        { hasText: "Completed (On Hold)" }
    ).last();

    await statusOption.click();

    await page.waitForTimeout(500);

    report.afterStatus = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // Open CHARGE TYPE
    // ------------------------------------------------------------

    const chargeCombobox = page.getByRole(
        "combobox",
        { name: "Charge Type *" }
    );

    await chargeCombobox.click();

    await page.waitForTimeout(500);

    report.afterChargeOpen = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // Select CHARGE TYPE
    // ------------------------------------------------------------

    const chargeOption = page.locator(
        ".Select__option",
        { hasText: "Internal Work" }
    ).last();

    await chargeOption.click();

    await page.waitForTimeout(500);

    report.afterCharge = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    // ------------------------------------------------------------
    // FINAL
    // ------------------------------------------------------------

    report.final = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };


    return report;
}

