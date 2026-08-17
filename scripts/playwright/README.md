# Playwright

* Automation for browser tabs
* Create ticket notes by scripting
* Run the ticketPlaywright.py

## Features

* Scrape assigned ticket view
* Manual ticket number entry

## Prereqs

### Mac

```bash
#!/bin/bash

set -e

echo "=== Playwright macOS Setup ==="
echo

# Check for Homebrew
if command -v brew >/dev/null 2>&1; then
    echo "✓ Homebrew is already installed."
else
    echo "✗ Homebrew is not installed."
    read -r -p "Do you want to install Homebrew? [y/N]: " answer

    case "$answer" in
        [yY]|[yY][eE][sS])
            echo "Installing Homebrew..."
            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
            ;;
        *)
            echo "Homebrew is required. Exiting."
            exit 1
            ;;
    esac
fi

echo
echo "Installing Python and Node.js..."
brew install python node

echo
echo "Installing Python Playwright..."
python3 -m pip install playwright --break-system-packages

echo
echo "Installing official Playwright CLI..."
npm install -g @playwright/cli@latest

echo
echo "✓ Installation complete!"
echo
echo "Installed:"
echo "  • Python"
echo "  • Node.js"
echo "  • Python Playwright"
echo "  • Playwright CLI"
echo
echo "Playwright browsers were NOT installed."
```


## Protocol

* Uses Playwright extension

## Edge debugging

* Might need to set as on
* edge://inspect

## Run command

```powershell
python ticketPlaywright.py
```

### Future improvements

* Need to set environment variable somehow
* It asks at every launch to take control of the browser tab
* Would be nicer to skip that manual work
* The scraper is also positional scraping so it may become inaccurate in some edge cases