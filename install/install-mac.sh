#!/bin/bash
# One-command installer for the Trace & Color Excel add-in (Microsoft 365 for Mac).
#   curl -fsSL https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/install/install-mac.sh | bash
set -e
DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$DIR"
curl -fsSL "https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/manifest.xml" \
  -o "$DIR/excel-trace-color.manifest.xml"
echo "Installed! Quit Excel (Cmd+Q), reopen it, and look for the \"Trace & Color\" tab."
