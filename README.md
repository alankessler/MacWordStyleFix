docx-fix-styles — fix Word style picker rendering issues in .docx files.

Fixes two problems caused by Pages exports (and some older Word versions):

1. w14:textOutline blocks — verbose "no outline" markers on every style
   definition that bloat the XML and can confuse Word's renderer.

2. Normal style excessive line spacing — Pages sets atLeast 455 twips
   (32pt minimum) for 9pt text. Word's style picker preview rows are
   too short for this, causing all style names to be clipped/unreadable.

Usage:
    docx-fix-styles <file.docx> [-o output.docx]
