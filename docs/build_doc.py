#!/usr/bin/env python3
"""
Build the PDIoT iOS build guide as a Word document.

Screenshots: if docs/screenshots/<id>.png exists it is embedded; otherwise a
labelled placeholder panel is written telling the reader exactly what to
capture. Re-running the script after dropping images in swaps them in.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

REPO = "/Users/s2255740/Documents/respeck-pdiot"
SHOTS = os.path.join(REPO, "docs", "screenshots")
OUT = os.path.join(REPO, "docs", "Building-the-PDIoT-app-on-iPhone.docx")

# ---------------------------------------------------------------- palette
BRAND      = "02569B"   # Flutter blue
BRAND_DARK = "013A6B"
ACCENT     = "13B9FD"
TEXT       = "1A1A1A"
MUTED      = "5A6672"
RULE       = "D8DEE6"

CODE_BG, CODE_BORDER = "F5F7FA", "D8DEE6"
INFO_BG,  INFO_BAR   = "EAF3FB", "02569B"
TIP_BG,   TIP_BAR    = "E9F5EC", "1E7B34"
WARN_BG,  WARN_BAR   = "FFF6E5", "C77700"
DANGER_BG, DANGER_BAR= "FDECEA", "B3261E"
PH_BG,    PH_BORDER  = "F0F4F8", "A8BACD"

BODY_FONT = "Calibri"
CODE_FONT = "Consolas"

fig_counter = [0]

# ---------------------------------------------------------------- helpers
def shade(element, color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    element.append(shd)


def cell_shade(cell, color):
    shade(cell._tc.get_or_add_tcPr(), color)


def set_cell_margins(cell, top=100, start=140, bottom=100, end=140):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for tag, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{tag}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tcPr.append(mar)


def table_borders(table, color=RULE, sz=4, kinds=("top", "left", "bottom", "right", "insideH", "insideV")):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for kind in kinds:
        el = OxmlElement(f"w:{kind}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for kind in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{kind}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        borders.append(el)
    tblPr.append(borders)


def left_bar(table, color, sz=24):
    """Thick coloured left edge, hairline elsewhere - the callout look."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for kind in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{kind}")
        if kind == "left":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tblPr.append(borders)


def run(p, text, *, font=BODY_FONT, size=11, bold=False, italic=False,
        color=TEXT, mono=False):
    r = p.add_run(text)
    r.font.name = CODE_FONT if mono else font
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = RGBColor.from_string(color)
    if mono:
        rPr = r._element.get_or_add_rPr()
        rf = rPr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rPr.append(rf)
        for a in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(a), CODE_FONT)
    return r


def para(doc, text="", *, size=11, bold=False, italic=False, color=TEXT,
         space_after=8, space_before=0, align=None, indent=None, line=1.25):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Inches(indent)
    if text:
        run(p, text, size=size, bold=bold, italic=italic, color=color)
    return p


def rich(doc, parts, *, size=11, space_after=8, indent=None, line=1.25):
    """parts: list of (text, style) where style in {'', 'b', 'i', 'code'}"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if indent is not None:
        pf.left_indent = Inches(indent)
    for text, style in parts:
        if style == "code":
            r = run(p, text, size=size - 0.5, mono=True, color="B3261E")
        else:
            r = run(p, text, size=size, bold=("b" in style), italic=("i" in style))
    return p


def rule(doc, color=RULE, space_before=6, space_after=10, sz=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    bdr.append(bottom)
    pPr.append(bdr)
    return p


def h1(doc, text, *, number=None):
    doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    if number:
        run(p, number + "  ·  ", size=26, bold=True, color=ACCENT)
    run(p, text, size=26, bold=True, color=BRAND_DARK)
    p.style = doc.styles["Heading 1"]
    for r in p.runs:
        r.font.name = BODY_FONT
    rule(doc, color=ACCENT, sz=12, space_before=2, space_after=14)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 2"]
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run(p, text, size=15, bold=True, color=BRAND)
    return p


def h3(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 3"]
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run(p, text, size=12.5, bold=True, color=TEXT)
    return p


def bullet(doc, text, *, level=0, bold_lead=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Inches(0.3 + 0.25 * level)
    p.paragraph_format.line_spacing = 1.2
    if bold_lead:
        run(p, bold_lead, size=11, bold=True)
    run(p, text, size=11)
    return p


def numbered(doc, n, text_parts):
    """A numbered step: coloured number then rich text."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(0.42)
    pf.first_line_indent = Inches(-0.42)
    pf.space_after = Pt(7)
    pf.line_spacing = 1.25
    run(p, f"{n}.", size=11, bold=True, color=BRAND)
    run(p, "\t", size=11)
    for text, style in text_parts:
        if style == "code":
            run(p, text, size=10.5, mono=True, color="B3261E")
        else:
            run(p, text, size=11, bold=("b" in style), italic=("i" in style))
    return p


def code(doc, lines, *, caption=None):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    c = t.cell(0, 0)
    cell_shade(c, CODE_BG)
    set_cell_margins(c, top=120, bottom=120, start=160, end=160)
    table_borders(t, color=CODE_BORDER, sz=4)
    c.text = ""
    first = True
    for ln in lines:
        p = c.paragraphs[0] if first else c.add_paragraph()
        first = False
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        if ln.startswith("$ "):
            run(p, "$ ", size=10, mono=True, color="8A94A0")
            run(p, ln[2:], size=10, mono=True, color="0B3D62", bold=True)
        elif ln.startswith("#"):
            run(p, ln, size=10, mono=True, color="6A737D", italic=True)
        else:
            run(p, ln, size=10, mono=True, color="333333")
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(8)
    if caption:
        cp = doc.add_paragraph()
        cp.paragraph_format.space_after = Pt(10)
        run(cp, caption, size=9, italic=True, color=MUTED)
    return t


def callout(doc, kind, title, body_parts):
    styles = {
        "info":   (INFO_BG, INFO_BAR, "i"),
        "tip":    (TIP_BG, TIP_BAR, "Tip"),
        "warn":   (WARN_BG, WARN_BAR, "!"),
        "danger": (DANGER_BG, DANGER_BAR, "!"),
    }
    bg, bar, _ = styles[kind]
    t = doc.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    cell_shade(c, bg)
    set_cell_margins(c, top=140, bottom=140, start=180, end=180)
    left_bar(t, bar)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    run(p, title, size=11, bold=True, color=bar)
    bp = c.add_paragraph()
    bp.paragraph_format.space_after = Pt(0)
    bp.paragraph_format.line_spacing = 1.2
    for text, style in body_parts:
        if style == "code":
            run(bp, text, size=10.5, mono=True, color="B3261E")
        else:
            run(bp, text, size=10.5, bold=("b" in style), italic=("i" in style))
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(10)
    return t


def screenshot(doc, sid, what, how, device="Mac"):
    """Embed docs/screenshots/<sid>.png if present, else a placeholder panel."""
    fig_counter[0] += 1
    n = fig_counter[0]
    path = os.path.join(SHOTS, sid + ".png")

    if os.path.exists(path):
        with Image.open(path) as im:
            w, h = im.size
        max_w = 6.0
        width = min(max_w, 6.0)
        doc.add_picture(path, width=Inches(width))
        pic_p = doc.paragraphs[-1]
        pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic_p.paragraph_format.space_before = Pt(4)
        pic_p.paragraph_format.space_after = Pt(3)
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(12)
        run(cap, f"Figure {n} — ", size=9, bold=True, color=MUTED)
        run(cap, what, size=9, italic=True, color=MUTED)
        return

    t = doc.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    cell_shade(c, PH_BG)
    set_cell_margins(c, top=200, bottom=200, start=180, end=180)
    table_borders(t, color=PH_BORDER, sz=8)
    c.text = ""
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    run(p, f"[ SCREENSHOT {n} ]", size=11, bold=True, color=BRAND)
    p2 = c.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(3)
    run(p2, what, size=10.5, bold=True, color=TEXT)
    p3 = c.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(3)
    run(p3, how, size=9.5, italic=True, color=MUTED)
    p4 = c.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_after = Pt(0)
    run(p4, f"Save as:  docs/screenshots/{sid}.png", size=9, mono=True, color=MUTED)
    run(p4, f"     ({device})", size=9, italic=True, color=MUTED)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(12)


def kv_table(doc, rows, widths=(2.0, 4.0), header=None):
    ncols = len(rows[0])
    t = doc.add_table(rows=0, cols=ncols)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_borders(t, color=RULE, sz=4)
    if header:
        hr = t.add_row()
        for i, htxt in enumerate(header):
            c = hr.cells[i]
            cell_shade(c, BRAND)
            set_cell_margins(c)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run(p, htxt, size=10, bold=True, color="FFFFFF")
    for idx, row in enumerate(rows):
        r = t.add_row()
        for i, val in enumerate(row):
            c = r.cells[i]
            if idx % 2 == 1:
                cell_shade(c, "F7F9FB")
            set_cell_margins(c)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            mono = val.startswith("`") and val.endswith("`")
            txt = val.strip("`")
            run(p, txt, size=10, mono=mono, bold=(i == 0 and ncols == 2))
    for i, w in enumerate(widths[:ncols]):
        for row in t.rows:
            row.cells[i].width = Inches(w)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(10)
    return t


def add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, "PDIoT — Building the Respeck app on iPhone     ", size=8.5, color=MUTED)
    r = p.add_run()
    r.font.size = Pt(8.5)
    r.font.name = BODY_FONT
    r.font.color.rgb = RGBColor.from_string(MUTED)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r._element.addnext(fld)


def add_toc(doc):
    p = doc.add_paragraph()
    r = p.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-2" \\h \\z \\u'
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "Right-click here and choose “Update Field” to build the table of contents."
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r._element.append(fld)
    r._element.append(instr)
    r._element.append(sep)
    r._element.append(placeholder)
    r._element.append(end)


# ---------------------------------------------------------------- document
doc = Document()

st = doc.styles["Normal"]
st.font.name = BODY_FONT
st.font.size = Pt(11)
st.font.color.rgb = RGBColor.from_string(TEXT)
st._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)

sec = doc.sections[0]
sec.page_width = Inches(8.5)
sec.page_height = Inches(11)
sec.left_margin = Inches(1.0)
sec.right_margin = Inches(1.0)
sec.top_margin = Inches(0.9)
sec.bottom_margin = Inches(0.9)
add_page_number_footer(sec)

# ------------------------------------------------- title page
for _ in range(4):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_after = Pt(0)
run(p, "PDIoT COURSE GUIDE", size=11, bold=True, color=ACCENT)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(0)
run(p, "Building the Respeck", size=38, bold=True, color=BRAND_DARK)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
run(p, "Recording App", size=38, bold=True, color=BRAND_DARK)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(14)
run(p, "on your iPhone", size=38, bold=True, color=ACCENT)

rule(doc, color=ACCENT, sz=18, space_before=4, space_after=16)

para(doc, "A complete, start-to-finish walkthrough for absolute beginners.",
     size=14, color=MUTED, space_after=4)
para(doc, "No previous experience with Flutter, Xcode or the command line is assumed.",
     size=12, color=MUTED, space_after=26)

kv_table(doc, [
    ["Application", "PDIoT Respeck Recording App (Flutter)"],
    ["Target device", "iPhone running iOS 16 or later"],
    ["You will need", "A Mac, an iPhone, a USB cable and a free Apple ID"],
    ["Time required", "About 2–3 hours, mostly waiting for downloads"],
    ["Cost", "Free — no paid Apple Developer account required"],
], widths=(1.8, 4.2))

doc.add_page_break()

# ------------------------------------------------- TOC
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(4)
run(p, "Contents", size=22, bold=True, color=BRAND_DARK)
rule(doc, color=ACCENT, sz=12, space_after=12)
add_toc(doc)

callout(doc, "info", "How to use this document",
        [("Work through the parts ", ""), ("in order", "b"),
         (". Each part ends with a checkpoint telling you exactly what should have happened "
          "before you move on. If something goes wrong, jump to ", ""),
         ("Part 8 — Troubleshooting", "b"),
         (", which lists every error message a beginner normally hits, and what to do about each one.", "")])

# ------------------------------------------------- Before you begin
h1(doc, "Before you begin")

para(doc, "This guide takes you from a completely fresh Mac to the PDIoT recording app "
          "running on your own iPhone, collecting live data from a Respeck sensor. "
          "Every command you need to type is shown in full, and every step explains "
          "what it does and why it matters.")

h2(doc, "What you need")

kv_table(doc, [
    ["A Mac", "Any Mac made in roughly the last five years. Apple silicon (M1–M4) or Intel both work. "
              "You must be able to install software, so it should be your own machine rather than a locked-down lab Mac."],
    ["macOS", "Recent enough to install the current Xcode from the App Store. If your Mac will not offer you "
              "the latest Xcode, update macOS first via  System Settings  ▸  General  ▸  Software Update."],
    ["Free disk space", "At least 40 GB. Xcode alone is around 20 GB once installed, and the iOS platform "
                        "support files add several more."],
    ["An iPhone", "Running iOS 16 or later. Developer Mode, which you will switch on in Part 4, only exists from iOS 16 onwards."],
    ["A USB cable", "To connect the iPhone to the Mac. A cable that charges the phone from the Mac will work. "
                    "Wireless debugging is possible later but is far more fiddly to set up first time."],
    ["An Apple ID", "Any ordinary Apple ID — the one you already use for the App Store is fine. "
                    "You do not need to pay for the Apple Developer Program."],
    ["Repository access", "The app's source code lives in a private GitHub repository. You need a GitHub "
                          "account that has been granted access to it before you can download the code in Part 3."],
], widths=(1.5, 4.5), header=["Item", "Details"])

h2(doc, "A note on the two apps you are installing")

para(doc, "Beginners often get confused about why two different programs are needed. They do different jobs:")

bullet(doc, "is Apple's own development tool. It contains the iOS SDK, the compiler for iPhone apps, "
            "the device management tools, and the code-signing machinery that lets an app run on real hardware. "
            "Nothing can be installed onto an iPhone without it.", bold_lead="Xcode ")
bullet(doc, "is a lightweight, friendly code editor. Together with the Flutter extension it gives you an easy "
            "way to install Flutter itself, edit the app's Dart code, and run everyday commands. "
            "It is much pleasanter to work in than Xcode for day-to-day editing.", bold_lead="Visual Studio Code ")

para(doc, "You will use VS Code to install Flutter and prepare the project, then hand over to Xcode for the "
          "one job only Xcode can do: signing the app with your Apple ID and installing it onto the phone.",
     space_before=4)

callout(doc, "tip", "Getting comfortable with Terminal",
        [("Several steps ask you to type a command into ", ""), ("Terminal", "b"),
         (". To open it, press ", ""), ("Cmd + Space", "b"), (", type ", ""), ("Terminal", "b"),
         (", and press Return. Type commands exactly as shown, then press Return to run them. "
          "The ", ""), ("$", "code"),
         (" at the start of a command in this guide is just the prompt — do not type it.", "")])

# ------------------------------------------------- Part 1 Xcode
h1(doc, "Install Xcode", number="Part 1")

para(doc, "Xcode is a large download — typically 8–10 GB compressed, expanding to around 20 GB. "
          "Start this early and let it run while you read ahead; on a home connection it can easily take an hour.")

h2(doc, "Step 1.1 — Install Xcode from the Mac App Store")

numbered(doc, 1, [("Open the ", ""), ("App Store", "b"), (" app on your Mac.", "")])
numbered(doc, 2, [("Search for ", ""), ("Xcode", "b"), (" and open its page.", "")])
numbered(doc, 3, [("Click ", ""), ("Get", "b"), (", then ", ""), ("Install", "b"),
                  (". You may be asked for your Apple ID password.", "")])
numbered(doc, 4, [("Wait for the download to finish. The App Store will show a progress bar; "
                   "it is safe to keep using your Mac meanwhile.", "")])

screenshot(doc, "01-app-store-xcode",
           "The Xcode page in the Mac App Store, showing the Get / Install button",
           "Open App Store, search for Xcode, then press Cmd + Shift + 4 followed by Space and click the App Store window.")

callout(doc, "warn", "If the App Store will not let you install Xcode",
        [("The App Store always offers the newest Xcode, which sometimes needs a newer macOS than you are running. "
          "If you see ", ""), ("“Requires macOS 26.0 or later”", "i"),
         (" or similar, update macOS first through  System Settings ▸ General ▸ Software Update, then try again. "
          "If you cannot update, older Xcode versions can be downloaded manually from ", ""),
         ("developer.apple.com/download/all", "b"), (" with a free Apple ID.", "")])

h2(doc, "Step 1.2 — Open Xcode once and let it finish setting up")

para(doc, "The App Store download is only half the installation. The first launch unpacks additional "
          "components and must be allowed to complete.")

numbered(doc, 1, [("Open ", ""), ("Xcode", "b"), (" from your Applications folder.", "")])
numbered(doc, 2, [("Accept the licence agreement when prompted.", "")])
numbered(doc, 3, [("If a dialog appears offering to install additional required components, click ", ""),
                  ("Install", "b"), (" and enter your Mac password.", "")])
numbered(doc, 4, [("Wait until Xcode reaches its normal welcome window, then leave it open — you will need it in Part 5.", "")])

screenshot(doc, "02-xcode-first-launch",
           "Xcode's first-launch dialog installing additional required components",
           "Capture the dialog as it appears on first launch with Cmd + Shift + 4 then Space.")

h2(doc, "Step 1.3 — Install the Xcode command-line tools")

para(doc, "These are the compilers and helper programs that Flutter calls behind the scenes. "
          "The Flutter documentation lists this as the first prerequisite on macOS, because it also "
          "provides Git, which you may want later.")

code(doc, ["$ xcode-select --install"],
     caption="A dialog will open asking you to confirm. Click Install, then Done when it finishes. "
             "If it replies that the tools are already installed, that is fine — move on.")

screenshot(doc, "03-command-line-tools",
           "The “The command line developer tools … install now?” confirmation dialog",
           "Run the command above, then capture the dialog that appears.")

h2(doc, "Step 1.4 — Point the tools at Xcode and run its first-launch tasks")

para(doc, "This tells macOS to use the full Xcode installation rather than the minimal command-line tools, "
          "and completes any setup Xcode still needs. It is taken verbatim from the official Flutter iOS "
          "setup instructions.")

code(doc, ["$ sudo sh -c 'xcode-select -s /Applications/Xcode.app/Contents/Developer && xcodebuild -runFirstLaunch'"],
     caption="Because this starts with sudo, Terminal will ask for your Mac login password. "
             "Nothing appears on screen as you type it — that is normal. Type it and press Return.")

h2(doc, "Step 1.5 — Agree to the Xcode licence")

code(doc, ["$ sudo xcodebuild -license"],
     caption="Press Space to page through the agreement, then type  agree  and press Return.")

h2(doc, "Step 1.6 — Download the iOS platform support files")

para(doc, "Xcode ships without the iOS-specific build files by default. This command fetches them. "
          "It is another sizeable download, so again, be patient.")

code(doc, ["$ xcodebuild -downloadPlatform iOS"])

callout(doc, "tip", "Checkpoint for Part 1",
        [("Xcode opens without prompting you to install anything further, and all four commands above "
          "have completed without errors. You do not need to verify anything else yet — ", ""),
         ("flutter doctor", "code"), (" in Part 2 will confirm the whole toolchain at once.", "")])

# ------------------------------------------------- Part 2 VS Code + Flutter
h1(doc, "Install VS Code and Flutter", number="Part 2")

para(doc, "Flutter is the framework the PDIoT app is written in. The friendliest way to install it is to let "
          "the VS Code Flutter extension do the work, which is the approach the official Flutter "
          "“quick start” documentation now recommends.")

h2(doc, "Step 2.1 — Install Visual Studio Code")

numbered(doc, 1, [("In your web browser, go to ", ""), ("code.visualstudio.com", "b"), (".", "")])
numbered(doc, 2, [("Click the large download button. The site detects macOS automatically; "
                   "if it offers a choice, pick the ", ""), ("Apple Silicon", "b"), (" build for M-series Macs or ", ""),
                  ("Intel", "b"), (" for older ones. The ", ""), ("Universal", "b"), (" build works on both.", "")])
numbered(doc, 3, [("Open the downloaded ", ""), (".zip", "code"), (" file, then drag ", ""),
                  ("Visual Studio Code", "b"), (" into your ", ""), ("Applications", "b"), (" folder.", "")])
numbered(doc, 4, [("Open it. If macOS warns that it was downloaded from the internet, click ", ""),
                  ("Open", "b"), (" to confirm.", "")])

screenshot(doc, "04-vscode-download",
           "The Visual Studio Code download page with the macOS download button",
           "Open code.visualstudio.com in your browser and capture the window.")

h2(doc, "Step 2.2 — Install the Flutter extension")

numbered(doc, 1, [("In VS Code, click the ", ""), ("Extensions", "b"),
                  (" icon in the left-hand sidebar — it looks like four small squares — "
                   "or press ", ""), ("Cmd + Shift + X", "b"), (".", "")])
numbered(doc, 2, [("Type ", ""), ("Flutter", "b"), (" into the search box.", "")])
numbered(doc, 3, [("Select the extension named ", ""), ("Flutter", "b"), (" published by ", ""),
                  ("Dart Code", "b"), (" — it is the official one and will be the first result.", "")])
numbered(doc, 4, [("Click ", ""), ("Install", "b"),
                  (". The Dart extension is installed automatically alongside it.", "")])

screenshot(doc, "05-vscode-flutter-extension",
           "The Flutter extension by Dart Code in the VS Code Extensions panel",
           "Open the Extensions panel, search Flutter, and capture the VS Code window.")

h2(doc, "Step 2.3 — Let VS Code download the Flutter SDK")

para(doc, "This is the clever part: rather than downloading and unpacking Flutter by hand and editing "
          "configuration files, the extension does all of it for you, including the PATH setup that "
          "trips up most beginners.")

numbered(doc, 1, [("Press ", ""), ("Cmd + Shift + P", "b"),
                  (" to open the Command Palette (also available at  View ▸ Command Palette).", "")])
numbered(doc, 2, [("Type ", ""), ("flutter", "b"), (" and choose ", ""), ("Flutter: New Project", "b"), (".", "")])
numbered(doc, 3, [("VS Code will notice Flutter is missing and offer ", ""),
                  ("Locate the Flutter SDK on your computer", "b"), (". Click ", ""), ("Download SDK", "b"), (".", "")])
numbered(doc, 4, [("A ", ""), ("Select Folder for Flutter SDK", "b"),
                  (" dialog appears. Choose a simple location you will remember — your home folder is ideal. "
                   "Avoid folder names containing spaces, and avoid Desktop, Documents or Downloads if they "
                   "are synced to iCloud, as syncing can corrupt the SDK.", "")])
numbered(doc, 5, [("Click ", ""), ("Clone Flutter", "b"), (". A notification reads ", ""),
                  ("“Downloading the Flutter SDK. This may take a few minutes.”", "i"),
                  (" Let it finish — it is a large download.", "")])
numbered(doc, 6, [("When prompted, click ", ""), ("Add SDK to PATH", "b"),
                  (". You should see ", ""), ("“The Flutter SDK was added to your PATH”", "i"), (".", "")])
numbered(doc, 7, [("If a Google Analytics notice appears, click ", ""), ("OK", "b"), (".", "")])
numbered(doc, 8, [("Quit and reopen VS Code, and close every open Terminal window. "
                   "This matters: PATH changes only take effect in newly opened terminals.", "")])

screenshot(doc, "06-vscode-download-sdk",
           "The “Locate the Flutter SDK” prompt with the Download SDK button",
           "Capture the VS Code window when the prompt appears.")

screenshot(doc, "07-vscode-sdk-path",
           "The “The Flutter SDK was added to your PATH” notification",
           "Capture the VS Code notification in the lower-right corner.")

callout(doc, "warn", "If VS Code did not offer to add Flutter to your PATH",
        [("You can add it yourself. Open Terminal and run the two commands below, replacing the path if you "
          "installed Flutter somewhere other than your home folder. Then close and reopen Terminal.", "")])

code(doc, ['$ echo \'export PATH="$HOME/flutter/bin:$PATH"\' >> ~/.zshrc',
           "$ source ~/.zshrc"],
     caption="On macOS the shell is zsh, so ~/.zshrc is the correct file. See Part 8 if flutter is still “command not found”.")

h2(doc, "Step 2.4 — Check the installation with flutter doctor")

para(doc, "This command inspects your whole setup and reports what is working. It is the single most "
          "useful diagnostic in Flutter, and you will be asked to run it whenever you seek help.")

code(doc, ["$ flutter doctor"])

para(doc, "You are looking for a green tick beside ", space_after=4)
bullet(doc, "Flutter", bold_lead="[✓] ")
bullet(doc, "Xcode — develop for iOS and macOS", bold_lead="[✓] ")

para(doc, "Some warnings are expected and completely harmless for this project:", space_before=6)

kv_table(doc, [
    ["Android toolchain — Android SDK not found",
     "Ignore. You are building for iPhone. Android support is a separate installation you do not need."],
    ["CocoaPods not installed",
     "Ignore for this project. See Appendix A for the full explanation — this app uses Swift Package Manager "
     "instead, and has been confirmed to build with no CocoaPods present at all."],
    ["Android Studio not installed",
     "Ignore, for the same reason as the Android toolchain warning. The official docs explicitly say this is "
     "safe to ignore when targeting iOS."],
], widths=(2.3, 3.7), header=["Warning you may see", "What to do"])

screenshot(doc, "08-flutter-doctor",
           "Terminal showing flutter doctor output with Flutter and Xcode ticked",
           "Run flutter doctor in Terminal, then press Cmd + Shift + 4, Space, and click the Terminal window.")

callout(doc, "tip", "Checkpoint for Part 2",
        [("Typing ", ""), ("flutter doctor", "code"),
         (" in a brand-new Terminal window works, and both Flutter and Xcode show a green tick. "
          "Warnings about Android and CocoaPods are expected.", "")])

# ------------------------------------------------- Part 3 source
h1(doc, "Get the app's source code", number="Part 3")

callout(doc, "warn", "This is a private repository",
        [("The PDIoT app's code is in a ", ""), ("private", "b"),
         (" GitHub repository. You must be signed in to GitHub with an account that has been granted "
          "access, otherwise the page will show a 404 error. If you get one, ask the repository owner "
          "to add you as a collaborator.", "")])

h2(doc, "Step 3.1 — Download the code as a ZIP file")

numbered(doc, 1, [("Sign in to ", ""), ("github.com", "b"), (" in your browser.", "")])
numbered(doc, 2, [("Go to the repository page: ", ""), ("github.com/passarac/respeck-pdiot", "b"), (".", "")])
numbered(doc, 3, [("Click the green ", ""), ("Code", "b"), (" button near the top right.", "")])
numbered(doc, 4, [("Choose ", ""), ("Download ZIP", "b"), (".", "")])
numbered(doc, 5, [("Open your Downloads folder and double-click the ZIP to unpack it. "
                   "You will get a folder called ", ""), ("respeck-pdiot-main", "code"), (".", "")])
numbered(doc, 6, [("Move that folder somewhere sensible and permanent — your home folder is a good choice. "
                   "Avoid paths containing spaces where you can.", "")])

screenshot(doc, "09-github-download-zip",
           "The GitHub Code button expanded, showing Download ZIP",
           "Open the repository page, click Code, and capture the browser window.")

h2(doc, "Step 3.2 — Open the project in VS Code")

numbered(doc, 1, [("In VS Code choose  File ▸ Open Folder…", "")])
numbered(doc, 2, [("Select the unpacked project folder and click ", ""), ("Open", "b"), (".", "")])
numbered(doc, 3, [("If VS Code asks whether you trust the authors of the files, click ", ""),
                  ("Yes, I trust the authors", "b"), (" — otherwise the Flutter tooling stays disabled.", "")])

para(doc, "You should now see the project's folders in the sidebar, including ", space_after=4)
bullet(doc, "— the Dart source code of the app itself", bold_lead="lib/ ")
bullet(doc, "— the iPhone-specific project that Xcode opens", bold_lead="ios/ ")
bullet(doc, "— the list of the app's dependencies", bold_lead="pubspec.yaml ")

screenshot(doc, "10-vscode-project-open",
           "The PDIoT project open in VS Code with lib/ and ios/ visible in the sidebar",
           "Capture the whole VS Code window after opening the folder.")

h2(doc, "Step 3.3 — Download the project's dependencies")

para(doc, "The app relies on external packages — for Bluetooth, for reading QR codes, for file storage. "
          "They are listed in pubspec.yaml but are not included in the download, so fetch them now.")

para(doc, "Open a terminal inside VS Code with  Terminal ▸ New Terminal, then run:", space_after=6)

code(doc, ["$ cd ~/respeck-pdiot-main", "$ flutter pub get"],
     caption="Adjust the first line if you put the folder somewhere else. A quick way to get the path right: "
             "type  cd  followed by a space, then drag the folder from Finder into the Terminal window.")

callout(doc, "tip", "Checkpoint for Part 3",
        [("flutter pub get", "code"),
         (" finishes with “Got dependencies!” and no red error text.", "")])

# ------------------------------------------------- Part 4 iPhone
h1(doc, "Prepare your iPhone", number="Part 4")

para(doc, "An iPhone will not run software from an unknown developer until you explicitly allow it. "
          "There are three separate permissions involved, and beginners often miss one. Do all three in order.")

h2(doc, "Step 4.1 — Connect the phone and trust the Mac")

numbered(doc, 1, [("Plug the iPhone into the Mac with the USB cable.", "")])
numbered(doc, 2, [("Unlock the phone.", "")])
numbered(doc, 3, [("A ", ""), ("Trust This Computer?", "b"), (" alert appears on the phone. Tap ", ""),
                  ("Trust", "b"), (" and enter your passcode.", "")])

callout(doc, "info", "If the Trust alert never appears",
        [("Unlock the phone first, then unplug and replug the cable. If it still does not appear, "
          "the cable may be charge-only — try a different one. A cable that charges but never shows "
          "this dialog is the single most common cause of a phone that Xcode cannot see.", "")])

screenshot(doc, "11-iphone-trust", "The “Trust This Computer?” alert on the iPhone",
           "On the iPhone press the Side button and Volume Up together to screenshot, then AirDrop it to your Mac.",
           device="iPhone")

h2(doc, "Step 4.2 — Turn on Developer Mode")

para(doc, "Developer Mode is a security feature introduced in iOS 16. Without it, a self-built app cannot "
          "launch at all. The menu item only appears after the phone has been connected to a Mac running "
          "Xcode at least once, so do Step 4.1 first.")

numbered(doc, 1, [("On the iPhone open ", ""), ("Settings", "b"), (".", "")])
numbered(doc, 2, [("Go to ", ""), ("Privacy & Security", "b"), (".", "")])
numbered(doc, 3, [("Scroll to the bottom and tap ", ""), ("Developer Mode", "b"), (".", "")])
numbered(doc, 4, [("Switch it ", ""), ("On", "b"), (".", "")])
numbered(doc, 5, [("The phone asks to restart. Tap ", ""), ("Restart", "b"), (".", "")])
numbered(doc, 6, [("After it restarts, unlock it and confirm ", ""), ("Turn On", "b"), (" when asked.", "")])

screenshot(doc, "12-iphone-developer-mode",
           "Settings ▸ Privacy & Security ▸ Developer Mode switched on",
           "Screenshot on the iPhone and AirDrop it to your Mac.",
           device="iPhone")

callout(doc, "warn", "No Developer Mode entry in Settings?",
        [("It only appears once the phone has been attached to a Mac with Xcode installed. Make sure Xcode "
          "is fully installed and has been opened at least once, then reconnect the phone and look again. "
          "On iOS 15 and earlier the setting does not exist and is not needed.", "")])

h2(doc, "Step 4.3 — Confirm the Mac can see the phone")

code(doc, ["$ flutter devices"],
     caption="Your iPhone should be listed by name. If only macOS and Chrome appear, revisit Steps 4.1 and 4.2, "
             "and see Part 8.")

screenshot(doc, "13-flutter-devices",
           "Terminal output of flutter devices listing the connected iPhone",
           "Run the command in Terminal and capture the window.")

callout(doc, "tip", "Checkpoint for Part 4",
        [("The iPhone appears in the output of ", ""), ("flutter devices", "code"),
         (", and Developer Mode is switched on in the phone's settings.", "")])

# ------------------------------------------------- Part 5 signing
h1(doc, "Set up signing in Xcode", number="Part 5")

para(doc, "Apple requires every app to be cryptographically signed before it will run on a physical device. "
          "This proves the app came from an identifiable developer. Using a free Apple ID, Xcode can create "
          "that signature for you automatically.")

h2(doc, "Step 5.1 — Prepare the iOS project first (do not skip this)")

para(doc, "Before you open Xcode, run the command below. It generates configuration files inside the "
          "ios/ folder that tell Xcode where the Flutter code is and how to build it. These files are "
          "deliberately not stored in the repository because they contain paths specific to your own Mac.")

code(doc, ["$ cd ~/respeck-pdiot-main", "$ flutter build ios --config-only"])

callout(doc, "danger", "Why this step really matters",
        [("If you skip it, Xcode may build using a stale configuration left over from somebody else's machine "
          "and silently compile the ", ""), ("wrong copy of the app's code", "b"),
         (". Everything appears to work, but your changes never show up and you end up debugging a file "
          "that is not being built. Running this command takes seconds and removes the risk entirely. "
          "Run it again any time you move or rename the project folder.", "")])

h2(doc, "Step 5.2 — Open the Xcode workspace")

para(doc, "Note carefully which file to open. The ios/ folder contains both a ", space_after=4)
rich(doc, [("Runner.xcodeproj", "code"), (" and a ", ""), ("Runner.xcworkspace", "code"),
           (". Always open the ", ""), ("workspace", "b"),
           (" — this is what the official Flutter deployment documentation instructs, and opening the wrong "
            "one causes confusing build failures on many projects.", "")])

para(doc, "The easiest way is from the terminal, which avoids picking the wrong file by mistake:", space_before=6, space_after=6)

code(doc, ["$ open ios/Runner.xcworkspace"],
     caption="Alternatively, in Finder open the project folder, then ios, then double-click Runner.xcworkspace "
             "— the icon is white and blue, not the plain blue of the .xcodeproj.")

h2(doc, "Step 5.3 — Select the Runner target")

numbered(doc, 1, [("In the left sidebar, click the blue ", ""), ("Runner", "b"),
                  (" icon at the very top. This opens the project editor.", "")])
numbered(doc, 2, [("In the panel that appears, find the ", ""), ("TARGETS", "b"), (" list and select ", ""),
                  ("Runner", "b"), (". Make sure you pick the target under TARGETS, not the entry under PROJECT.", "")])
numbered(doc, 3, [("Click the ", ""), ("Signing & Capabilities", "b"), (" tab along the top.", "")])

screenshot(doc, "14-xcode-signing-tab",
           "Xcode showing the Runner target with the Signing & Capabilities tab selected",
           "With Runner.xcworkspace open, navigate to the tab and capture the Xcode window.")

h2(doc, "Step 5.4 — Add your Apple ID and choose the team")

numbered(doc, 1, [("Tick ", ""), ("Automatically manage signing", "b"),
                  (" if it is not already ticked. Xcode then creates and renews certificates for you.", "")])
numbered(doc, 2, [("Open the ", ""), ("Team", "b"), (" dropdown and choose ", ""), ("Add an Account…", "b"), (".", "")])
numbered(doc, 3, [("Sign in with your Apple ID. Two-factor authentication codes go to your other Apple devices.", "")])
numbered(doc, 4, [("Close the Accounts window when you are done.", "")])
numbered(doc, 5, [("Back on Signing & Capabilities, open the ", ""), ("Team", "b"),
                  (" dropdown again and pick the entry labelled ", ""), ("(Personal Team)", "b"), (".", "")])

screenshot(doc, "15-xcode-add-account",
           "The Xcode Accounts window after signing in with an Apple ID",
           "Capture the sheet that appears after Add an Account…")

callout(doc, "info", "What “Personal Team” means",
        [("A Personal Team is the free tier that comes with any Apple ID. It lets you install your own apps "
          "onto your own devices. It cannot be used to publish to the App Store, and it carries the "
          "limits described below — but for the PDIoT coursework it is all you need.", "")])

h2(doc, "Step 5.5 — Fix the bundle identifier if Xcode complains")

para(doc, "The bundle identifier uniquely names the app across the whole of Apple's ecosystem. "
          "The project ships with com.specknet.pdiot, and because free Personal Teams must register a "
          "genuinely unique identifier, Xcode will refuse it if somebody else on your course has already "
          "claimed it.")

para(doc, "If you see a red error reading ", space_after=4)
rich(doc, [("“Failed to register bundle identifier”", "i"), (" or ", ""),
           ("“The app identifier cannot be registered to your development team”", "i"),
           (", change it to something unique to you:", "")])

numbered(doc, 1, [("Click the ", ""), ("General", "b"), (" tab.", "")])
numbered(doc, 2, [("Find ", ""), ("Bundle Identifier", "b"), (" under Identity.", "")])
numbered(doc, 3, [("Change it to something personal, for example ", ""),
                  ("com.s1234567.pdiot", "code"), (" using your own student number.", "")])
numbered(doc, 4, [("Return to ", ""), ("Signing & Capabilities", "b"),
                  (" and confirm the error has cleared. Xcode regenerates the profile automatically.", "")])

screenshot(doc, "16-xcode-bundle-id",
           "The Bundle Identifier field on the General tab",
           "Capture the General tab with the Identity section visible.")

callout(doc, "info", "The 7-day limit on free accounts",
        [("Apps signed with a free Personal Team stop working after ", ""), ("7 days", "b"),
         (". When that happens the app will simply refuse to open. The fix is to connect the phone and press "
          "Run in Xcode again, which re-signs it for another 7 days. Your recorded data is not affected. "
          "You are also limited to 3 such apps on one device at a time. The paid Apple Developer Program "
          "($99/year) raises the limit to a year, but is not required for this coursework.", "")])

callout(doc, "tip", "Checkpoint for Part 5",
        [("The Signing & Capabilities tab shows your name next to Team, a Personal Team is selected, and "
          "there are no red error messages in the signing section.", "")])

# ------------------------------------------------- Part 6 run
h1(doc, "Build and run on your iPhone", number="Part 6")

h2(doc, "Step 6.1 — Choose your iPhone as the destination")

numbered(doc, 1, [("Look at the toolbar at the top of the Xcode window. Next to the ", ""),
                  ("Runner", "b"), (" scheme there is a destination menu, which usually reads ", ""),
                  ("Any iOS Device", "b"), (" or the name of a simulator.", "")])
numbered(doc, 2, [("Click it and select your iPhone by name under ", ""), ("iOS Device", "b"), (".", "")])

callout(doc, "warn", "Do not choose a Simulator",
        [("The PDIoT app talks to a Respeck sensor over Bluetooth Low Energy, and the iOS Simulator has no "
          "Bluetooth hardware. The app may launch in a simulator but it will never find a sensor. "
          "You must use a real iPhone.", "")])

screenshot(doc, "17-xcode-device-picker",
           "The Xcode destination menu with the connected iPhone selected",
           "Click the destination menu in the Xcode toolbar and capture the window.")

h2(doc, "Step 6.2 — Press Run")

numbered(doc, 1, [("Click the ▶ ", ""), ("Run", "b"), (" button, or press ", ""), ("Cmd + R", "b"), (".", "")])
numbered(doc, 2, [("The first build takes several minutes. Xcode's status bar shows progress. Keep the phone "
                   "unlocked and connected throughout.", "")])
numbered(doc, 3, [("If macOS asks for permission to use your signing key, enter your Mac password and choose ", ""),
                  ("Always Allow", "b"), (" to avoid being asked repeatedly.", "")])

h2(doc, "Step 6.3 — Trust the developer certificate on the phone")

para(doc, "On the very first install the app will fail to launch, and the phone shows ", space_after=4)
rich(doc, [("“Untrusted Developer”", "i"),
           (". This is expected and is not a mistake on your part. Your Apple ID is not yet trusted "
            "on the device. To fix it:", "")])

numbered(doc, 1, [("On the iPhone open ", ""), ("Settings", "b"), (" ▸ ", ""), ("General", "b"),
                  (" ▸ ", ""), ("VPN & Device Management", "b"), (".", "")])
numbered(doc, 2, [("Under ", ""), ("Developer App", "b"), (", tap the entry showing your Apple ID.", "")])
numbered(doc, 3, [("Tap ", ""), ("Trust “…”", "b"), (", then ", ""), ("Trust", "b"), (" again to confirm.", "")])
numbered(doc, 4, [("Go back to Xcode and press ", ""), ("Run", "b"), (" once more.", "")])

screenshot(doc, "18-iphone-device-management",
           "Settings ▸ General ▸ VPN & Device Management showing the Trust option",
           "Screenshot on the iPhone and AirDrop it to your Mac.",
           device="iPhone")

h2(doc, "Step 6.4 — Success")

para(doc, "The app launches on the iPhone and shows the PDIoT home screen with a Connect button, "
          "acceleration readings, two dropdown menus for activity and social signal, and buttons to "
          "start and stop recording.")

screenshot(doc, "19-app-running",
           "The PDIoT app running on the iPhone home screen",
           "Screenshot on the iPhone once the app has launched, and AirDrop it to your Mac.",
           device="iPhone")

callout(doc, "tip", "From now on you can use VS Code instead",
        [("Xcode is only needed for signing. Once the app has been installed successfully once, you can build "
          "and run from VS Code or the terminal with ", ""), ("flutter run", "code"),
         (", which is faster and gives you hot reload — your code changes appear on the phone in about a "
          "second without a full rebuild. Return to Xcode only when the 7-day signature expires.", "")])

# ------------------------------------------------- Part 7 using
h1(doc, "Using the app", number="Part 7")

h2(doc, "Step 7.1 — Pair with your Respeck")

numbered(doc, 1, [("Tap ", ""), ("Settings", "b"), (" at the bottom of the app's home screen.", "")])
numbered(doc, 2, [("Enter your ", ""), ("Subject ID", "b"), (" in the first field.", "")])
numbered(doc, 3, [("Enter the ", ""), ("Respeck UUID", "b"), (" in the second, or tap ", ""),
                  ("Scan QR code", "b"), (" and point the camera at the label on the back of the sensor.", "")])
numbered(doc, 4, [("Tap ", ""), ("Save settings", "b"), (".", "")])

callout(doc, "info", "A note specific to iPhone",
        [("On iOS, Apple hides a Bluetooth device's true hardware address from apps, so the Respeck is "
          "identified by an ID it broadcasts in its advertising data instead. This requires Respeck "
          "firmware 6AM or newer. An older sensor that works on Android may not be detectable on an iPhone.", "")])

h2(doc, "Step 7.2 — Record data")

numbered(doc, 1, [("Make sure the Respeck is awake, then tap ", ""), ("Connect", "b"),
                  (". Live acceleration values and the battery level appear within a few seconds.", "")])
numbered(doc, 2, [("Choose the ", ""), ("Activity", "b"), (" and ", ""), ("Social signal", "b"),
                  (" you are about to perform from the dropdowns.", "")])
numbered(doc, 3, [("Tap ", ""), ("Start recording", "b"),
                  (". The status line shows a running count of samples written and seconds elapsed.", "")])
numbered(doc, 4, [("When finished, tap ", ""), ("Stop recording", "b"), (".", "")])

h2(doc, "Step 7.3 — Retrieve the recorded CSV files")

para(doc, "On iPhone the app writes its CSV files into its own documents folder. To get them off the phone:")

numbered(doc, 1, [("Open the ", ""), ("Files", "b"), (" app on the iPhone.", "")])
numbered(doc, 2, [("Go to ", ""), ("On My iPhone", "b"), (" and look for the ", ""), ("PDIoT", "b"), (" folder.", "")])
numbered(doc, 3, [("Select the CSV files and share them to your Mac with AirDrop, or to cloud storage.", "")])

para(doc, "Each file is named with your subject ID, the activity, the social signal, a UTC timestamp and "
          "the Respeck ID, and contains one row per sample with the phone timestamp, sensor timestamp, "
          "packet and sample sequence numbers, and the x, y and z acceleration in g.", space_before=4)

# ------------------------------------------------- Part 8 troubleshooting
h1(doc, "Troubleshooting", number="Part 8")

para(doc, "These are the problems beginners hit most often, with the actual fix for each.")

h2(doc, "Command line and Flutter")

kv_table(doc, [
    ["zsh: command not found: flutter",
     "The SDK is not on your PATH, or you are using a Terminal window that was open before it was added. "
     "Close every Terminal window and open a new one. If it persists, add the PATH line from Step 2.3 "
     "manually and run  source ~/.zshrc."],
    ["flutter doctor hangs or is very slow the first time",
     "Normal on first run — it is downloading the Dart SDK and build tools. Leave it for a few minutes."],
    ["“Flutter failed to create a directory” or permission errors",
     "You installed the SDK somewhere macOS protects. Move it to your home folder and repeat Step 2.3."],
    ["Xcode warnings about an obsolete Java or Android SDK",
     "Irrelevant to iOS builds. Ignore them."],
], widths=(2.4, 3.6), header=["Symptom", "Fix"])

h2(doc, "Device problems")

kv_table(doc, [
    ["The iPhone does not appear in Xcode or flutter devices",
     "Work through these in order: unlock the phone; confirm you tapped Trust in Step 4.1; confirm Developer "
     "Mode is on (Step 4.2); try a different USB cable, as charge-only cables are a very common cause; "
     "try a different USB port; restart both the phone and the Mac."],
    ["“Developer Mode disabled”",
     "Complete Step 4.2. The phone must be restarted for the setting to take effect."],
    ["The phone shows as “unavailable” or “Preparing device for development”",
     "Xcode is copying debug support files onto the phone. Wait for it to finish — on a first connection "
     "this can take several minutes. Keep the phone unlocked."],
], widths=(2.4, 3.6), header=["Symptom", "Fix"])

h2(doc, "Signing and installation")

kv_table(doc, [
    ["“Signing for Runner requires a development team”",
     "No team is selected. Return to Step 5.4 and choose your Personal Team from the Team dropdown."],
    ["“Failed to register bundle identifier”",
     "Somebody else has already claimed com.specknet.pdiot. Change the Bundle Identifier to something "
     "unique, as described in Step 5.5."],
    ["“Untrusted Developer” on the phone",
     "Expected on the first install. Trust your certificate via Settings ▸ General ▸ VPN & Device "
     "Management, as described in Step 6.3."],
    ["“Unable to install… The maximum number of apps for free development profiles has been reached”",
     "A free Apple ID allows only 3 self-signed apps per device. Delete one you no longer need and try again."],
    ["The app stops opening after about a week",
     "The free 7-day signature has expired. Connect the phone and press Run in Xcode again to renew it. "
     "This is normal, not a fault."],
], widths=(2.4, 3.6), header=["Symptom", "Fix"])

h2(doc, "Build failures")

kv_table(doc, [
    ["Xcode builds successfully but your code changes do nothing",
     "The generated configuration is stale or points at a different folder. Quit Xcode, run  "
     "flutter build ios --config-only  in the project folder (Step 5.1), then reopen the workspace."],
    ["“Module not found” or missing plugin errors",
     "Run  flutter pub get  then  flutter build ios --config-only  and reopen Xcode. If it persists, "
     "only then consider installing CocoaPods — see Appendix A."],
    ["“CocoaPods not installed” in flutter doctor",
     "Harmless for this project. See Appendix A."],
    ["Build errors immediately after unzipping the project",
     "You probably skipped  flutter pub get  in Step 3.3."],
], widths=(2.4, 3.6), header=["Symptom", "Fix"])

h2(doc, "App behaviour")

kv_table(doc, [
    ["“Please pair with a Respeck first”",
     "No Respeck UUID has been saved. Complete Step 7.1."],
    ["The Respeck is never found during scanning",
     "Check the sensor is awake and charged, that its firmware is 6AM or newer (older firmware is not "
     "detectable on iOS), and that the UUID was entered exactly, including colons."],
    ["Bluetooth permission was denied by mistake",
     "Open Settings ▸ PDIoT on the iPhone and switch Bluetooth back on for the app."],
], widths=(2.4, 3.6), header=["Symptom", "Fix"])

# ------------------------------------------------- Appendix A
h1(doc, "Appendix A — Why you do not need CocoaPods", number="A")

para(doc, "Almost every Flutter iOS tutorial you will find online, including Apple- and Flutter-authored "
          "documentation, tells you to install CocoaPods. For this project you genuinely do not need it, "
          "and this appendix explains why so you can ignore the warning with confidence.")

para(doc, "CocoaPods is a dependency manager for Apple platforms. Historically, Flutter used it to attach "
          "plugins with native iOS code — Bluetooth, camera, file storage and so on — to the Xcode project.")

para(doc, "Flutter has since migrated to Apple's own ", space_after=4)
rich(doc, [("Swift Package Manager", "b"),
           (" (SPM), which is built directly into Xcode and requires no extra installation. This project "
            "uses that newer mechanism: its Xcode project references a generated Swift package that pulls "
            "in every native plugin the app needs, including the Bluetooth plugin that talks to the Respeck.", "")])

para(doc, "This has been verified on this project rather than assumed. A complete iOS build was run on a Mac "
          "with no CocoaPods installed at all, and it succeeded, producing a working Runner.app. No Podfile "
          "exists in the repository, and the iOS build configuration contains no CocoaPods references.",
     space_before=4)

callout(doc, "info", "So what should you do about the warning?",
        [("Ignore it. ", "b"),
         ("flutter doctor prints the CocoaPods warning generically for anyone targeting iOS; it does not "
          "inspect whether your particular project needs it. Only if you hit a “module not found” error "
          "that survives the fixes in Part 8 should you install CocoaPods, and in that case follow the "
          "official guide at guides.cocoapods.org.", "")])

# ------------------------------------------------- Appendix B
h1(doc, "Appendix B — Command reference", number="B")

para(doc, "Every command used in this guide, collected for easy reference.")

h2(doc, "One-off setup")
kv_table(doc, [
    ["`xcode-select --install`", "Install the Xcode command-line tools (Step 1.3)"],
    ["`sudo sh -c 'xcode-select -s /Applications/Xcode.app/Contents/Developer && xcodebuild -runFirstLaunch'`",
     "Point the tools at Xcode and finish its setup (Step 1.4)"],
    ["`sudo xcodebuild -license`", "Read and accept the Xcode licence (Step 1.5)"],
    ["`xcodebuild -downloadPlatform iOS`", "Download iOS platform support (Step 1.6)"],
], widths=(3.1, 2.9), header=["Command", "Purpose"])

h2(doc, "Everyday use")
kv_table(doc, [
    ["`flutter doctor`", "Check the whole toolchain and report problems"],
    ["`flutter devices`", "List every device Flutter can currently see"],
    ["`flutter pub get`", "Download the project's dependencies"],
    ["`flutter build ios --config-only`", "Regenerate the Xcode configuration without building"],
    ["`open ios/Runner.xcworkspace`", "Open the project in Xcode"],
    ["`flutter run`", "Build and run on the connected device, with hot reload"],
    ["`flutter clean`", "Delete all build output; a good first step when builds misbehave"],
], widths=(3.1, 2.9), header=["Command", "Purpose"])

# ------------------------------------------------- Appendix C
h1(doc, "Appendix C — Official sources", number="C")

para(doc, "The instructions in this guide were cross-checked against the following official documentation. "
          "If anything here ever conflicts with these pages, trust the official pages — Flutter and Xcode "
          "both change frequently.")

kv_table(doc, [
    ["Flutter — Quick start (VS Code install method)",
     "docs.flutter.dev/install/quick"],
    ["Flutter — iOS setup on macOS",
     "docs.flutter.dev/platform-integration/ios/install-ios/vscode"],
    ["Flutter — Adding Flutter to your PATH",
     "docs.flutter.dev/install/add-to-path"],
    ["Flutter — Build and release an iOS app",
     "docs.flutter.dev/deployment/ios"],
    ["Apple — Enabling Developer Mode on a device",
     "developer.apple.com/documentation/xcode/enabling-developer-mode-on-a-device"],
    ["Apple — Automatically managing signing",
     "developer.apple.com/library/ios/qa/qa1814"],
    ["CocoaPods — Installation guide (only if ever needed)",
     "guides.cocoapods.org/using/getting-started.html"],
], widths=(2.9, 3.1), header=["Document", "Address"])

rule(doc, color=RULE, space_before=18, space_after=8)
para(doc, "End of guide.", size=10, italic=True, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER)

doc.save(OUT)

present = sum(1 for n in range(1, fig_counter[0] + 1))
have = len([f for f in os.listdir(SHOTS) if f.endswith(".png")]) if os.path.isdir(SHOTS) else 0
print(f"Saved: {OUT}")
print(f"Figures referenced: {fig_counter[0]}   images currently present: {have}")
