#!/bin/bash
# Remove mac-word-layout-fix. Fonts are left in place (remove them in Font Book if needed).
set -euo pipefail

OFFICE="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized"

rm -f "$HOME/Library/Application Scripts/com.microsoft.Word/Fix HTML Table Layout.scpt"
rm -f "$OFFICE/Startup.localized/Word/MacWordLayoutFix.dotm"

echo "Removed. Restart Word to apply."
