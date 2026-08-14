async page => {
    const tickets = await page.evaluate(() => {
        const results = [];
        const seenIds = new Set();
        
        // Find all elements that look like a ticket row or contain ticket numbers
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        let node;
        while (node = walker.nextNode()) {
            const val = (node.nodeValue || "").trim();
            const match = val.match(/^00\d{5}$/);
            if (match) {
                const ticketId = match[0];
                if (!seenIds.has(ticketId)) {
                    seenIds.add(ticketId);
                    
                    // Traverse up to find the container row with the full row details
                    let container = node.parentElement;
                    let bestText = val;
                    while (container && container !== document.body) {
                        const text = container.innerText || "";
                        if (text.includes(ticketId) && text.length > 20) {
                            bestText = text;
                            if (text.includes("Service Request") || text.includes("Incident") || text.includes("Project Support") || text.split('\n').length > 3) {
                                break;
                            }
                        }
                        container = container.parentElement;
                    }
                    
                    results.push({
                        id: ticketId,
                        title: bestText.replace(/\n/g, ' | ').trim().substring(0, 120)
                    });
                }
            }
        }
        return results;
    });

    console.log(JSON.stringify(tickets, null, 2));
}