async page => {
    const rows = page.locator('.ag-row, tr, [role="row"]');
    const count = await rows.count();
    for (let i = 0; i < count; i++) {
        await rows.nth(i).hover();
        await page.waitForTimeout(200); // Let tooltip or ID render
    }
    console.log("Hovered over all rows.");
}