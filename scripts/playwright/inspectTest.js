async page => {
    const tickets = await page.evaluate(() => {
        // Put your test code here and return an array
        return [{id: "0026229", title: "Test Ticket"}];
    });
    console.log(JSON.stringify(tickets, null, 2));
}