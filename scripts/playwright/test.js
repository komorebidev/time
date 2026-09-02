async page => {

    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    // Set known values
    await start.fill("08:00");
    await end.fill("09:00");

    // Verify immediately before save
    const beforeSave = {
        start: await start.inputValue(),
        end: await end.inputValue()
    };

    console.log("BEFORE SAVE:", beforeSave);

    // Save
    const saveButton = page.getByRole(
        "button",
        {
            name: "Save",
            exact: true
        }
    );

    await saveButton.click();

    // Give Halo time to process the save
    await page.waitForTimeout(2000);

    console.log("URL:", page.url());

    return {
        beforeSave,
        urlAfterSave: page.url()
    };
}
