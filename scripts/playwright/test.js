async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    // Set both times
    await start.fill("08:00");
    await end.fill("09:00");

    const beforeSave = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };

    // Find the actual visible Save button
    const saveButton = page.getByRole(
        "button",
        { name: "Save", exact: true }
    );

    await saveButton.waitFor({
        state: "visible",
        timeout: 10000
    });

    // Use Playwright's normal click
    await saveButton.click();

    await page.waitForTimeout(2000);

    return {
        beforeSave,
        urlAfterSave: page.url()
    };
}