async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    // Set Start
    await start.fill("08:00");

    const afterStart = await start.inputValue();

    // Set End
    await end.fill("09:00");

    const afterEnd = await start.inputValue();

    // Open Charge Type
    const chargeCombobox = page.getByRole(
        "combobox",
        { name: "Charge Type *" }
    );

    await chargeCombobox.click();

    await page.waitForTimeout(500);

    // Select Internal Work
    const chargeOption = page.locator(
        '.Select__option',
        { hasText: "Internal Work" }
    ).last();

    await chargeOption.click();

    await page.waitForTimeout(1000);

    const afterCharge = await page
        .locator(
            'input[name="actionarrivaldate_time"]'
        )
        .inputValue();

    return {
        afterStart,
        afterEnd,
        afterCharge
    };
}