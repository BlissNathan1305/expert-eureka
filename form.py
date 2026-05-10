│ """
   2   │ Word Document Formatter Script
   3   │ Formats a Word document with specific chapter headers, subhe
       │ ading bolding, citation italicization, and text alignment
   4   │ Usage: python format_word.py input.docx output.docx
   5   │ """
   6   │
   7   │ from docx import Document
   8   │ from docx.shared import Pt, Cm
   9   │ from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACI
       │ NG
  10   │ from docx.oxml.ns import qn
  11   │ import re
  12   │ import sys
  13   │ import os
  14   │
  15   │ def set_run_font(run, font_name='Times New Roman', font_size
       │ =12, bold=False, italic=False):
  16   │     """Set font properties for a run"""
  17   │     run.font.name = font_name
  18   │     run.font.size = Pt(font_size)
  19   │     run.font.bold = bold
  20   │     run.font.italic = italic
  21   │     # For Asian characters support
  22   │     run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
  23   │
  24   │ def italicize_et_al(paragraph):
  25   │     """Find 'et al' in paragraph and italicize it"""
  26   │     text = paragraph.text
  27   │     if 'et al' in text.lower():
  28   │         # Store original formatting
  29   │         alignment = paragraph.alignment
  30   │         line_spacing = paragraph.paragraph_format.line_spaci
       │ ng
  31   │         space_before = paragraph.paragraph_format.space_befo
       │ re
  32   │         space_after = paragraph.paragraph_format.space_after
  33   │         first_line_indent = paragraph.paragraph_format.first
       │ _line_indent
  34   │
  35   │         # Clear existing runs
  36   │         paragraph.clear()
  37   │
  38   │         # Process text with regex to find 'et al' (case inse
       │ nsitive)
  39   │         parts = re.split(r'(et al\.?|et al\b)', text, flags=
       │ re.IGNORECASE)
: