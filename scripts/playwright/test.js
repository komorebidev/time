async page => {

    // ---------------------------------------------------------
    // OPEN WORKLOG
    // ---------------------------------------------------------

    await page.getByRole(
        "button",
        { name: "Worklog" }
    ).click();

    await page.waitForTimeout(1500);


    // ---------------------------------------------------------
    // WORKLOG TEXT
    // ---------------------------------------------------------

    const editor = page.locator(
        '[contenteditable="true"]'
    ).first();

    await editor.waitFor({
        state: "visible",
        timeout: 10000
    });

    await editor.fill(
        "TEST WORKLOG"
    );


    // ---------------------------------------------------------
    // START / END
    // ---------------------------------------------------------

    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    const end = page.locator(
        'input[name="actioncompletiondate_time"]'
    );

    await start.fill("08:00");

    await end.fill("09:00");


    // ---------------------------------------------------------
    // STATUS
    // ---------------------------------------------------------

    const statusCombobox = page.getByRole(
        "combobox",
        { name: "Status *" }
    );

    await statusCombobox.click();

    await page.waitForTimeout(500);

    await page.locator(
        '.Select__option',
        { hasText: "Completed (On Hold)" }
    ).last().click();

    await page.waitForTimeout(500);


    // ---------------------------------------------------------
    // CHARGE TYPE
    // ---------------------------------------------------------

    const chargeCombobox = page.getByRole(
        "combobox",
        { name: "Charge Type *" }
    );

    await chargeCombobox.click();

    await page.waitForTimeout(500);

    await page.locator(
        '.Select__option',
        { hasText: "Internal Work" }
    ).last().click();

    await page.waitForTimeout(500);


    // ---------------------------------------------------------
    // CHECK VALUES BEFORE SAVE
    // ---------------------------------------------------------

    const beforeSave = {
        start: await page.locator(
            'input[name="actionarrivaldate_time"]'
        ).inputValue(),

        end: await page.locator(
            'input[name="actioncompletiondate_time"]'
        ).inputValue()
    };


    // ---------------------------------------------------------
    // SAVE
    // ---------------------------------------------------------

    await page.getByRole(
        "button",
        { name: "Save", exact: true }
    ).click();

    await page.waitForTimeout(2000);


    return {
        beforeSave,
        urlAfterSave: page.url()
    };
}