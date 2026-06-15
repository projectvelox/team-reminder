#!/usr/bin/env python3
"""Generate docs/DayReminders-Overview.pptx — a 9-slide deck for announcing
Day Reminders to the team.

Modern hero-style layout: each content slide leads with the screenshot
sized to its native aspect ratio, with a short caption beneath. Dark
title + closing slides for visual rhythm. Real Picture objects backed by
the captured screenshots — placeholders only kick in when a PNG is
missing.

Re-run after capturing screenshots or after content changes:
    py tools/build-overview-pptx.py
"""

from pathlib import Path
from io import BytesIO

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "DayReminders-Overview.pptx"
SCREENSHOTS_DIR = ROOT / "docs" / "screenshots"

# Palette
ACCENT = RGBColor(0x38, 0xAE, 0xEB)
INK_DEEP = RGBColor(0x0E, 0x1F, 0x2E)
INK_DARK = RGBColor(0x18, 0x2C, 0x3C)
TEXT_PRIMARY = RGBColor(0x1F, 0x2D, 0x3D)
TEXT_MUTED = RGBColor(0x70, 0x80, 0x95)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT_BLUE = RGBColor(0xB0, 0xC8, 0xDC)

SCREENSHOT_FILES = {
    1: "01-left-rail.png",
    2: "02-add-form.png",
    3: "03-lines-view.png",
    4: "04-week-view.png",
    5: "05-group-by-client.png",
    6: "06-proactive-card.png",
}


# ---------- helpers ----------

def _load_font(name, size):
    try:
        return ImageFont.truetype(name, size)
    except Exception:
        try:
            return ImageFont.truetype("segoeui.ttf", size)
        except Exception:
            return ImageFont.load_default()


def make_placeholder_png(idx, caption, w=1600, h=900):
    """Gray dashed-border placeholder used when a real screenshot is missing."""
    img = Image.new("RGB", (w, h), (236, 240, 244))
    draw = ImageDraw.Draw(img)
    border_color = (180, 192, 205)
    seg = 24
    for x in range(8, w - 8, seg * 2):
        draw.line([(x, 8), (min(x + seg, w - 8), 8)], fill=border_color, width=3)
        draw.line([(x, h - 8), (min(x + seg, w - 8), h - 8)], fill=border_color, width=3)
    for y in range(8, h - 8, seg * 2):
        draw.line([(8, y), (8, min(y + seg, h - 8))], fill=border_color, width=3)
        draw.line([(w - 8, y), (w - 8, min(y + seg, h - 8))], fill=border_color, width=3)
    title_font = _load_font("segoeuib.ttf", 96)
    sub_font = _load_font("segoeui.ttf", 38)
    title_text = f"Screenshot {idx:02d}"
    tb = draw.textbbox((0, 0), title_text, font=title_font)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    draw.text(((w - tw) // 2, h // 2 - th - 30), title_text, fill=(60, 75, 95), font=title_font)
    max_w = int(w * 0.85)
    words = caption.split()
    lines, cur = [], ""
    for word in words:
        test = (cur + " " + word).strip()
        if draw.textbbox((0, 0), test, font=sub_font)[2] <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    y = h // 2 + 30
    for line in lines:
        lb = draw.textbbox((0, 0), line, font=sub_font)
        draw.text(((w - (lb[2] - lb[0])) // 2, y), line, fill=(110, 124, 138), font=sub_font)
        y += int((lb[3] - lb[1]) * 1.4)
    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def fit_picture(slide, idx, caption, slot_left, slot_top, slot_w, slot_h, shadow=False):
    """Add a picture fitted inside the slot at its native aspect ratio, centered."""
    real = SCREENSHOTS_DIR / SCREENSHOT_FILES.get(idx, "")
    use_real = real.exists() and real.stat().st_size > 0
    if use_real:
        with Image.open(real) as im:
            iw, ih = im.size
        img_ar = iw / ih
        slot_ar = slot_w / slot_h
        if img_ar > slot_ar:
            w = slot_w
            h = int(slot_w / img_ar)
        else:
            h = slot_h
            w = int(slot_h * img_ar)
        l = slot_left + (slot_w - w) // 2
        t = slot_top + (slot_h - h) // 2
        pic = slide.shapes.add_picture(str(real), l, t, w, h)
    else:
        png_bytes = make_placeholder_png(idx, caption)
        pic = slide.shapes.add_picture(BytesIO(png_bytes), slot_left, slot_top, slot_w, slot_h)
    return pic


def add_text(slide, text, left, top, width, height, *, size=18, color=TEXT_PRIMARY, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name="Segoe UI"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    parts = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(parts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = name
        p.alignment = align
    return tb


def add_dark_bg(slide, prs, color=INK_DEEP):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()


def add_section_title(slide, prs, text):
    """Top title + small accent bar underneath — replaces the v1 colored title bar."""
    add_text(slide, text, Inches(0.7), Inches(0.55), prs.slide_width - Inches(1.4), Inches(0.8),
             size=30, color=TEXT_PRIMARY, bold=True)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.25), Inches(0.55), Inches(0.06))
    bar.fill.solid(); bar.fill.fore_color.rgb = ACCENT; bar.line.fill.background()


def add_centered_caption(slide, prs, text, top=Inches(6.65), size=14, color=TEXT_MUTED):
    add_text(slide, text, Inches(0.7), top, prs.slide_width - Inches(1.4), Inches(0.5),
             size=size, color=color, align=PP_ALIGN.CENTER)


def add_footer(slide, prs):
    add_text(slide, "Day Reminders v1.4.6",
             Inches(0.5), prs.slide_height - Inches(0.4),
             prs.slide_width - Inches(1.0), Inches(0.3),
             size=9, color=TEXT_MUTED, align=PP_ALIGN.RIGHT)


# ---------- slides ----------

def slide_title(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_dark_bg(s, prs)
    # accent bar above the title
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2.05), Inches(0.7), Inches(0.09))
    bar.fill.solid(); bar.fill.fore_color.rgb = ACCENT; bar.line.fill.background()
    add_text(s, "Day Reminders",
             Inches(1), Inches(2.3), Inches(11.3), Inches(1.6),
             size=84, color=WHITE, bold=True)
    add_text(s, "Reminders, where you already work.",
             Inches(1), Inches(4.0), Inches(11.3), Inches(0.8),
             size=30, color=ACCENT)
    add_text(s, "v1.4.6  ·  Kation Technologies  ·  June 2026",
             Inches(1), Inches(6.6), Inches(11.3), Inches(0.5),
             size=13, color=SOFT_BLUE)
    return s


def slide_pitch(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(s, "Stop forgetting things in five different apps.",
             Inches(1), Inches(2.7), Inches(11.3), Inches(1.8),
             size=46, color=TEXT_PRIMARY, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, "Day Reminders is a tab + bot inside Teams.\nYou add things; the bot pings you in chat before each one's due.",
             Inches(1.5), Inches(4.6), Inches(10.3), Inches(1.5),
             size=20, color=TEXT_MUTED, align=PP_ALIGN.CENTER)
    add_footer(s, prs)
    return s


def slide_where_it_lives(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_section_title(s, prs, "Find it in your Teams left rail.")
    # Two pictures: left-rail icon (tall/narrow) + lines view (wide)
    pic_top = Inches(1.7)
    pic_h = Inches(4.7)
    gap = Inches(0.3)
    left_w = Inches(3.6)
    right_w = prs.slide_width - Inches(1.0) - left_w - gap
    fit_picture(s, 1, "Teams left rail",
                Inches(0.5), pic_top, left_w, pic_h)
    fit_picture(s, 3, "Lines view",
                Inches(0.5) + left_w + gap, pic_top, right_w, pic_h)
    add_centered_caption(s, prs,
        "Alarm-clock icon in the rail.   →   Two top tabs: Reminders (the UI) and Chat (where the bot pings you).")
    add_footer(s, prs)
    return s


def slide_adding(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_section_title(s, prs, "Adding a reminder.")
    fit_picture(s, 2, "Add form",
                Inches(0.5), Inches(1.65), prs.slide_width - Inches(1.0), Inches(4.85))
    add_centered_caption(s, prs,
        "Title  ·  Client (autocompletes from your past)  ·  Date (defaults to today)  ·  Time (optional)  ·  + Details for notes / links / sub-tasks")
    add_footer(s, prs)
    return s


def slide_three_ways(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_section_title(s, prs, "Three ways to add one.")

    col_w = (prs.slide_width - Inches(1.4) - Inches(0.6)) // 3  # 3 cols + 2 gaps of 0.3"
    gap = Inches(0.3)
    col_top = Inches(2.0)
    left0 = Inches(0.7)

    columns = [
        ("In the tab", "for thoughtful adding",
         "Fill the form at the top of the Reminders tab — title, client, date, time, details. Hit Add."),
        ("Mid-conversation", "for when you remember while replying",
         "Click the · · · under any Teams message box → Day Reminders → Quick add reminder. Type, submit, back to your reply."),
        ("In the bot chat", "for typing-fast moments",
         "/add 5pm tomorrow #qc Review batch 14\n\nDates can be tomorrow, mon, fri, 6/20, or a full YYYY-MM-DD."),
    ]
    for i, (title, subtitle, body) in enumerate(columns):
        l = left0 + (col_w + gap) * i
        # numeric badge
        badge = s.shapes.add_shape(MSO_SHAPE.OVAL, l, col_top, Inches(0.55), Inches(0.55))
        badge.fill.solid(); badge.fill.fore_color.rgb = ACCENT; badge.line.fill.background()
        tf = badge.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = str(i + 1); p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(20); p.font.bold = True
        p.font.color.rgb = WHITE; p.font.name = "Segoe UI"
        add_text(s, title, l + Inches(0.75), col_top - Inches(0.05),
                 col_w - Inches(0.75), Inches(0.6),
                 size=22, color=TEXT_PRIMARY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, subtitle, l, col_top + Inches(0.85), col_w, Inches(0.5),
                 size=12, color=ACCENT, bold=True)
        add_text(s, body, l, col_top + Inches(1.45), col_w, Inches(3.2),
                 size=14, color=TEXT_PRIMARY)
    add_footer(s, prs)
    return s


def slide_week_view(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_section_title(s, prs, "Am I bombarded this week, or free?")
    fit_picture(s, 4, "Week view",
                Inches(0.5), Inches(1.65), prs.slide_width - Inches(1.0), Inches(4.85))
    add_centered_caption(s, prs,
        "Mon–Sun grid by due date. Today's column highlighted. Click any empty day to add for that date.")
    add_footer(s, prs)
    return s


def slide_organize(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_section_title(s, prs, "Group by client. Or tag. Or just leave it flat.")
    fit_picture(s, 5, "Group by client",
                Inches(0.5), Inches(1.65), prs.slide_width - Inches(1.0), Inches(4.85))
    add_centered_caption(s, prs,
        "The Group toggle cycles off → tag → client. Sections per group, each header colored to match. Click any chip to filter to just that one.")
    add_footer(s, prs)
    return s


def slide_notifications(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_section_title(s, prs, "When it's due, a card pops in your chat.")
    fit_picture(s, 6, "Proactive card",
                Inches(0.5), Inches(1.65), prs.slide_width - Inches(1.0), Inches(4.85))
    add_centered_caption(s, prs,
        "Mark done. Or snooze 15m, 1h, Tomorrow. End-of-day check-in lists anything still open.")
    add_footer(s, prs)
    return s


def slide_whats_next_and_cta(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_dark_bg(s, prs)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1.15), Inches(0.7), Inches(0.08))
    bar.fill.solid(); bar.fill.fore_color.rgb = ACCENT; bar.line.fill.background()
    add_text(s, "Try it now.",
             Inches(1), Inches(1.4), Inches(11.3), Inches(1.3),
             size=64, color=WHITE, bold=True)
    add_text(s, "Alarm-clock icon in your Teams left rail.",
             Inches(1), Inches(2.85), Inches(11.3), Inches(0.7),
             size=26, color=ACCENT)
    # Coming-next block
    add_text(s, "COMING NEXT",
             Inches(1), Inches(4.2), Inches(11.3), Inches(0.4),
             size=13, color=SOFT_BLUE, bold=True)
    add_text(s, "v1.5  ·  Sharing.",
             Inches(1), Inches(4.6), Inches(11.3), Inches(0.7),
             size=30, color=WHITE, bold=True)
    add_text(s, "Assign reminders to teammates. Set per-tag default share lists so #QC auto-shares with the QC team without picking recipients each time.",
             Inches(1), Inches(5.4), Inches(11.3), Inches(1.0),
             size=16, color=SOFT_BLUE)
    add_text(s, "Bugs, requests, 'this is annoying' → ping Josh, or drop them in the team channel.",
             Inches(1), Inches(6.7), Inches(11.3), Inches(0.5),
             size=14, color=SOFT_BLUE)
    return s


def build_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide_title(prs)
    slide_pitch(prs)
    slide_where_it_lives(prs)
    slide_adding(prs)
    slide_three_ways(prs)
    slide_week_view(prs)
    slide_organize(prs)
    slide_notifications(prs)
    slide_whats_next_and_cta(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes, {len(prs.slides)} slides)")


if __name__ == "__main__":
    build_deck()
