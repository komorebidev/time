import subprocess
import tempfile
import os

SESSION = "halo"
CLI = "playwright-cli"

# This snippet runs inside the browser via page.evaluate()
js_code = """
async page => {
    // Inspect all inputs on the page to find where Job Start / Job End are located
    const inputsInfo = await page.evaluate(() => {
        const allInputs = Array.from(document.querySelectorAll('input'));
        return allInputs.map((el, index) => {
            // Find parent label or container text
            let labelText = "";
            let parent = el.parentElement;
            for (let i = 0; i < 4 && parent; i++) {
                labelText += " " + parent.innerText;
                parent = parent.parentElement;
            }
            return {
                index: index,
                id: el.id,
                name: el.name,
                type: el.type,
                className: el.className,
                placeholder: el.placeholder,
                ariaLabel: el.getAttribute('aria-label'),
                surroundingText: labelText.trim().substring(0, 150).replace(/\\s+/g, ' ')
            };
        });
    });

    console.log("Found " + inputsInfo.length + " total inputs on page.");
    
    // Filter and print inputs that look relevant
    const relevant = inputsInfo.filter(i => 
        i.surroundingText.toLowerCase().includes('job start') || 
        i.surroundingText.toLowerCase().includes('job end') ||
        i.surroundingText.toLowerCase().includes('start time') ||
        i.surroundingText.toLowerCase().includes('end time')
    );

    console.log("--- RELEVANT INPUTS FOUND ---");
    console.log(JSON.stringify(relevant, null, 2));
}
"""

with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False, encoding="utf-8") as f:
    f.write(js_code)
    temp_path = f.name

try:
    result = subprocess.run([CLI, f"--s={SESSION}", "run-code", f"--filename={temp_path}"], capture_output=True, text=True, encoding="utf-8")
    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)
finally:
    os.remove(temp_path)