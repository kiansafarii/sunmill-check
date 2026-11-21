# Receipt check

A small web page for checking receipts on a phone, with no sign-in and no AI.

1. Open the page and tap **Open check file**. Pick the check file (a `.zip` or `.json` made by the receipts app), for example from a shared Google Drive folder.
2. Go through each receipt: swipe right if it's correct, swipe left to fix it. Pinch or double-tap the photo to zoom.
3. Tap **Send checked file back** so your changes can be imported into the receipts app, or **Download Excel** to get the finished spreadsheet.

This repository holds only the page's code. The check file, with its figures and photos, is opened on your phone and stays there; nothing is uploaded.

Files:
- `index.html`: the checking screen
- `export.js`: builds the Excel file in the browser from the template inside the check file
