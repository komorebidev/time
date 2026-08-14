async page => {
    // Look for elements containing the ticket prefix or table cells
    const count = await page.evaluate(() => {
        const elements = Array.from(document.querySelectorAll('*'));
        const matches = elements.filter(el => {
            const text = el.textContent ? el.textContent.trim() : '';
            return /^00\d{5}$/.test(text);
        });
        return matches.map(el => ({ tag: el.tagName, text: el.textContent.trim() }));
    });
    
    console.log("Found matches:", JSON.stringify(count, null, 2));
}