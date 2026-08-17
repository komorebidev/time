(async (page) => {
    const inputsInfo = await page.evaluate(() => {
        const allInputs = Array.from(document.querySelectorAll('input, select, textarea'));
        return allInputs.map((el, index) => {
            // Get the closest form group or parent row text
            const container = el.closest('.form-group, .row, .col, div') || el.parentElement;
            return {
                index: index,
                tag: el.tagName,
                id: el.id,
                name: el.name,
                type: el.type,
                placeholder: el.placeholder,
                ariaLabel: el.getAttribute('aria-label'),
                containerText: container ? container.innerText.trim().replace(/\s+/g, ' ') : ''
            };
        });
    });

    console.log("Found " + inputsInfo.length + " total fields.");
    // Print all fields that have some text or placeholder
    inputsInfo.forEach(i => {
        if (i.containerText || i.placeholder || i.id || i.name) {
            console.log(`[Index ${i.index}] Tag: ${i.tag} | ID: ${i.id} | Name: ${i.name} | Placeholder: ${i.placeholder} | Text: ${i.containerText.substring(0, 100)}`);
        }
    });
})(page);