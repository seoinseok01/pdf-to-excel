#!/bin/zsh
# Print fixtures/xlsx/*.xlsx to PDF with Microsoft Excel (macOS) and LibreOffice.
#   fixtures/pdf/excel/<name>.pdf        Excel "Save as PDF"
#   fixtures/pdf/libreoffice/<name>.pdf  LibreOffice headless export
# Excel is sandboxed, so files go through its container folder. The first run asks for
# permission to control Excel.
set -e
cd "$(dirname "$0")/.."
BOX="$HOME/Library/Containers/com.microsoft.Excel/Data/pdftoexcel"
mkdir -p "$BOX" fixtures/pdf/excel fixtures/pdf/libreoffice
for f in fixtures/xlsx/*.xlsx; do
  name=$(basename "$f" .xlsx)
  cp "$f" "$BOX/$name.xlsx"
  rm -f "$BOX/$name.pdf"
  osascript <<OSA
tell application "Microsoft Excel"
  set wb to open workbook workbook file name (POSIX file "$BOX/$name.xlsx" as text)
  save workbook as wb filename (POSIX file "$BOX/$name.pdf" as text) file format PDF file format
  close wb saving no
end tell
OSA
  cp "$BOX/$name.pdf" "fixtures/pdf/excel/$name.pdf"
  soffice --headless --convert-to pdf --outdir fixtures/pdf/libreoffice "$f" >/dev/null 2>&1
  echo "printed $name"
done
