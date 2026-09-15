# PDIoT iOS build guide

`Building-the-PDIoT-app-on-iPhone.docx` — a start-to-finish guide for a complete
beginner: install Xcode, install Flutter via VS Code, download the source,
prepare the iPhone, sign in Xcode, and run the app on the device.

## Adding the screenshots

The document currently contains 19 numbered placeholder panels. Each panel names
the exact file it expects, for example:

    Save as:  docs/screenshots/01-app-store-xcode.png

Take the screenshot, name the file exactly as the panel says, drop it in
`docs/screenshots/`, then rebuild:

    python3 -m venv /tmp/docvenv
    /tmp/docvenv/bin/pip install python-docx Pillow
    /tmp/docvenv/bin/python docs/build_doc.py

Any placeholder whose image now exists is replaced by the real screenshot with a
`Figure N` caption. Placeholders with no image stay as placeholders, so you can
add them a few at a time.

### Capturing on the Mac
`Cmd + Shift + 4`, then press `Space`, then click the window you want. The file
lands on your Desktop; rename and move it.

### Capturing on the iPhone
Press the Side button and Volume Up together, then AirDrop the image to the Mac.
Five of the placeholders (11, 12, 18, 19 and the Trust dialog) are iPhone
screens and can only be captured this way.

## Regenerating from scratch

`build_doc.py` is the single source of truth for the document — edit the script,
not the `.docx`, so that changes survive the next rebuild.
