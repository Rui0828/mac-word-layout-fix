#!/bin/bash
# Install mac-word-layout-fix for the current user.
#   ./install.sh                 install the Scripts-menu fix (and the add-in if built)
#   ./install.sh --fonts <dir>   also install licensed .ttf/.ttc fonts from <dir>
set -euo pipefail

cd "$(dirname "$0")"

OFFICE="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized"
SCRIPTS_DIR="$HOME/Library/Application Scripts/com.microsoft.Word"
STARTUP_DIR="$OFFICE/Startup.localized/Word"
ADDIN="dist/MacWordLayoutFix.dotm"

if [ ! -d "/Applications/Microsoft Word.app" ]; then
    echo "Microsoft Word is not installed." >&2
    exit 1
fi

# 1. Scripts menu (manual fix)
mkdir -p "$SCRIPTS_DIR"
osacompile -o "$SCRIPTS_DIR/Fix HTML Table Layout.scpt" src/FixLayout.applescript
echo "Installed: Word > Scripts menu > Fix HTML Table Layout"

# 2. Add-in (automatic fix on open)
if [ -f "$ADDIN" ]; then
    mkdir -p "$STARTUP_DIR"
    cp "$ADDIN" "$STARTUP_DIR/"
    echo "Installed: add-in $STARTUP_DIR/$(basename "$ADDIN")"
else
    echo "Skipped add-in: $ADDIN not found (see README: Build the add-in)"
fi

# 3. Fonts (optional)
if [ "${1:-}" = "--fonts" ]; then
    FONT_SRC="${2:?usage: ./install.sh --fonts <dir>}"
    mkdir -p "$HOME/Library/Fonts"
    found=0
    for f in "$FONT_SRC"/*.ttf "$FONT_SRC"/*.ttc "$FONT_SRC"/*.TTF "$FONT_SRC"/*.TTC; do
        [ -f "$f" ] || continue
        cp "$f" "$HOME/Library/Fonts/"
        echo "Installed font: $(basename "$f")"
        found=1
    done
    [ "$found" = 1 ] || echo "No .ttf/.ttc files found in $FONT_SRC"
fi

echo "Done. Restart Word to apply."
