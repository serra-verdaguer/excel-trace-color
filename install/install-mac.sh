#!/bin/bash
# One-command installer for the Trace & Color Excel add-in (Microsoft 365 for Mac).
#   curl -fsSL https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/install/install-mac.sh | bash
set -e
DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$DIR"
curl -fsSL "https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/manifest.xml" \
  -o "$DIR/excel-trace-color.manifest.xml"
echo "Installed!"
echo "1. Quit Excel completely (Cmd+Q) and reopen it with any workbook."
echo "2. Insert > Add-ins > My Add-ins (click the small arrow) > under 'Developer Add-ins' click 'Trace & Color'."
echo "3. The 'Trace & Color' tab appears on the ribbon."
