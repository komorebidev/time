await (async page => {
    await page.evaluate(() => {
        // Replace with the exact selector or condition you found from your inspection
        // e.g., finding the input by placeholder or by looping through inputs and checking surrounding labels
        const inputs = Array.from(document.querySelectorAll('input'));
        const startInput = inputs.find(el => {
            const label = el.closest('.form-group, .row, div')?.textContent || '';
            return label.includes('Job Start');
        });

        if (startInput) {
            // Use native setter to bypass React state suppression
            const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
            nativeSetter.call(startInput, "09:00");
            
            startInput.dispatchEvent(new Event('input', { bubbles: true }));
            startInput.dispatchEvent(new Event('change', { bubbles: true }));
            console.log("Successfully updated Job Start!");
        } else {
            console.log("Job Start input not found by container text.");
        }
    });
})(page);