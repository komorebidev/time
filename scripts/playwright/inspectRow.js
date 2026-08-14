async page => {
    const info = await page.evaluate(() => {
        const rows = document.querySelectorAll('.ag-row, tr, [role="row"]');
        return Array.from(rows).map(row => {
            return {
                outerHTML: row.outerHTML.substring(0, 300), // Check attributes
                text: row.innerText.replace(/\n/g, ' | ')
            };
        });
    });
    console.log(JSON.stringify(info, null, 2));
}