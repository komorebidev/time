async page => {
    const start = page.locator(
        'input[name="actionarrivaldate_time"]'
    );

    await start.click();

    return await start.evaluate(el => ({
        value: el.value,
        active: document.activeElement === el,
        activeId: document.activeElement?.id,
        activeName: document.activeElement?.getAttribute("name"),
        activeType: document.activeElement?.type
    }));
}