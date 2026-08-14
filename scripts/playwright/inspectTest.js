async page => {
    const matches = await page.evaluate(() => {
        const results = [];
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        let node;
        while (node = walker.nextNode()) {
            const val = (node.nodeValue || "").trim();
            if (/^00\d{5}$/.test(val)) {
                results.push({
                    text: val,
                    tag: node.parentElement ? node.parentElement.tagName : 'UNKNOWN',
                    parentClass: node.parentElement ? node.parentElement.className : ''
                });
            }
        }
        return results;
    });
    
    console.log("Found text nodes:", JSON.stringify(matches, null, 2));
}