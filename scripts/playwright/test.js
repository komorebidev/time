async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    return {
        count: await start.count(),
        visible: await start.isVisible(),
        enabled: await start.isEnabled(),
        value: await start.inputValue(),
        id: await start.getAttribute("id"),
        name: await start.getAttribute("name"),
        type: await start.getAttribute("type")
    };
}
