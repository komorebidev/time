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
python3 -m pip install playwright

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