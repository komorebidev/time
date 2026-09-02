async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    await start.fill("08:00");

    const afterStart = await start.inputValue();

    const statusCombobox = page.getByRole(
        "combobox",
        { name: "Status *" }
    );

    await statusCombobox.click();

    await page.waitForTimeout(500);

    const statusOption = page.locator(
        '.Select__option',
        { hasText: "Completed (On Hold)" }
    ).last();

    await statusOption.click();

    await page.waitForTimeout(1000);

    const afterStatus = await page
        .locator(
            'input[name="actionarrivaldate_time"]'
        )
        .inputValue();

    return {
        afterStart,
        afterStatus
    };
}