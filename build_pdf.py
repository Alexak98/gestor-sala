"""Convert MANUAL.md to a nicely styled PDF using reportlab Platypus."""
import re
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    PageBreak,
    Table,
    TableStyle,
    KeepTogether,
    HRFlowable,
    ListFlowable,
    ListItem,
)

ROOT = Path(__file__).parent
SRC = ROOT / "MANUAL.md"
OUT = ROOT / "MANUAL.pdf"

# Colours tuned to the app's palette
INK = HexColor("#1a1d23")
INK2 = HexColor("#5b6270")
INK3 = HexColor("#8a909b")
GREEN = HexColor("#22a06b")
PINK = HexColor("#e03956")
LINE = HexColor("#ececf0")
BG_SOFT = HexColor("#f7f7f9")
YELLOW_BG = HexColor("#fff7d6")

# Base font — use built-in Helvetica; replace emojis to avoid rendering boxes.
FONT = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
FONT_IT = "Helvetica-Oblique"


def strip_emojis(text: str) -> str:
    """Remove emoji sequences that would render as tofu in built-in fonts."""
    # Also drop common ones we use as section prefixes
    emoji_pattern = re.compile(
        "["
        "\U0001F300-\U0001F5FF"
        "\U0001F600-\U0001F64F"
        "\U0001F680-\U0001F6FF"
        "\U0001F700-\U0001F77F"
        "\U0001F780-\U0001F7FF"
        "\U0001F800-\U0001F8FF"
        "\U0001F900-\U0001F9FF"
        "\U0001FA00-\U0001FA6F"
        "\U0001FA70-\U0001FAFF"
        "\U00002702-\U000027B0"
        "\U000024C2-\U0001F251"
        "\U0001F1E0-\U0001F1FF"
        "\u2600-\u26FF"
        "\u2700-\u27BF"
        "\u2B00-\u2BFF"
        "\u2100-\u214F"
        "\uFE00-\uFE0F"
        "]+",
        flags=re.UNICODE,
    )
    cleaned = emoji_pattern.sub("", text)
    # Also strip specific typographic chars that might show as boxes
    # Keep em-dash em –, arrows →, bullets • and × (all covered in Helvetica)
    return cleaned


def md_inline_to_reportlab(text: str) -> str:
    """Convert a single line of markdown inline formatting to reportlab Paragraph markup."""
    text = strip_emojis(text).strip()
    # Escape reportlab special chars
    text = text.replace("&", "&amp;")
    # Bold **x** → <b>x</b>
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    # Italic *x* (but not ** already handled) — reportlab uses <i>
    text = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"<i>\1</i>", text)
    # Inline code `x` → <font face="Courier">x</font>
    text = re.sub(r"`([^`]+)`", r'<font face="Courier" color="#c92a2a">\1</font>', text)
    # Links [text](url) → just the text, bold
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"<b>\1</b>", text)
    return text


# ---------- Paragraph styles ----------
styles = getSampleStyleSheet()

BASE = ParagraphStyle(
    "Base",
    parent=styles["Normal"],
    fontName=FONT,
    fontSize=10.5,
    leading=16,
    textColor=INK,
    alignment=TA_LEFT,
    spaceAfter=6,
)

H1 = ParagraphStyle(
    "H1", parent=BASE, fontName=FONT_BOLD, fontSize=26, leading=30,
    spaceBefore=0, spaceAfter=6, textColor=INK,
)
SUB = ParagraphStyle(
    "Sub", parent=BASE, fontName=FONT, fontSize=12, leading=18,
    textColor=INK2, spaceAfter=18,
)
H2 = ParagraphStyle(
    "H2", parent=BASE, fontName=FONT_BOLD, fontSize=17, leading=22,
    textColor=INK, spaceBefore=18, spaceAfter=8,
)
H3 = ParagraphStyle(
    "H3", parent=BASE, fontName=FONT_BOLD, fontSize=13, leading=18,
    textColor=GREEN, spaceBefore=10, spaceAfter=4,
)
BODY = ParagraphStyle("Body", parent=BASE)
BODY_BULLET = ParagraphStyle(
    "Bullet", parent=BASE, leftIndent=14, bulletIndent=2, spaceAfter=3,
)
CALLOUT = ParagraphStyle(
    "Callout", parent=BASE, textColor=HexColor("#926800"),
    backColor=YELLOW_BG, borderPadding=(8, 10, 8, 10),
    borderRadius=8, leftIndent=4, rightIndent=4, spaceAfter=10,
)
FOOTER = ParagraphStyle(
    "Footer", parent=BASE, fontSize=8.5, textColor=INK3, alignment=TA_CENTER,
)

# ---------- Parse the markdown ----------
lines = SRC.read_text(encoding="utf-8").splitlines()

story = []
i = 0
first_h1_consumed = False


def flush_table(rows):
    """Build a rounded table flowable from parsed rows."""
    if not rows:
        return None
    header = rows[0]
    body = rows[1:]
    data = [[Paragraph(md_inline_to_reportlab(c), BASE) for c in header]]
    for row in body:
        data.append([Paragraph(md_inline_to_reportlab(c), BASE) for c in row])
    tbl = Table(data, hAlign="LEFT", colWidths=[4 * cm, 12 * cm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BG_SOFT),
        ("TEXTCOLOR", (0, 0), (-1, 0), INK),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return tbl


def consume_list(lines, i):
    items = []
    while i < len(lines):
        m = re.match(r"^[\-\*]\s+(.*)$", lines[i])
        n = re.match(r"^\d+\.\s+(.*)$", lines[i])
        if m or n:
            raw = (m or n).group(1)
            items.append(
                ListItem(
                    Paragraph(md_inline_to_reportlab(raw), BODY),
                    leftIndent=18, value="\u2022",
                )
            )
            i += 1
        elif lines[i].strip() == "":
            # allow one blank line inside list
            if i + 1 < len(lines) and (re.match(r"^[\-\*]\s+", lines[i + 1]) or re.match(r"^\d+\.\s+", lines[i + 1])):
                i += 1
            else:
                break
        else:
            break
    bullet_type = "bullet"
    return ListFlowable(
        items, bulletType=bullet_type, start=None,
        leftIndent=14, bulletFontName=FONT, bulletFontSize=10,
    ), i


def consume_table(lines, i):
    rows = []
    while i < len(lines) and lines[i].strip().startswith("|"):
        raw = lines[i].strip().strip("|")
        # Skip alignment row like |---|---|
        if re.match(r"^[\s\-:|]+$", raw):
            i += 1
            continue
        cells = [c.strip() for c in raw.split("|")]
        rows.append(cells)
        i += 1
    return flush_table(rows), i


while i < len(lines):
    line = lines[i]
    stripped = line.strip()

    # Horizontal rule
    if re.match(r"^---+\s*$", stripped):
        story.append(Spacer(1, 6))
        story.append(HRFlowable(width="100%", thickness=0.6, color=LINE))
        story.append(Spacer(1, 6))
        i += 1
        continue

    # H1
    if stripped.startswith("# "):
        text = stripped[2:].strip()
        if not first_h1_consumed:
            # Treat the first H1 as the cover title
            story.append(Spacer(1, 10))
            story.append(Paragraph(md_inline_to_reportlab(text), H1))
            first_h1_consumed = True
        else:
            story.append(Paragraph(md_inline_to_reportlab(text), H2))
        i += 1
        continue

    # H2
    if stripped.startswith("## "):
        text = stripped[3:].strip()
        story.append(Paragraph(md_inline_to_reportlab(text), H2))
        i += 1
        continue

    # H3
    if stripped.startswith("### "):
        text = stripped[4:].strip()
        story.append(Paragraph(md_inline_to_reportlab(text), H3))
        i += 1
        continue

    # Blockquote (callout)
    if stripped.startswith(">"):
        chunk = []
        while i < len(lines) and lines[i].strip().startswith(">"):
            chunk.append(lines[i].strip().lstrip(">").strip())
            i += 1
        text = " ".join(chunk)
        story.append(Paragraph(md_inline_to_reportlab(text), CALLOUT))
        continue

    # Table
    if stripped.startswith("|"):
        tbl, i = consume_table(lines, i)
        if tbl:
            story.append(Spacer(1, 4))
            story.append(tbl)
            story.append(Spacer(1, 4))
        continue

    # List (ordered or unordered)
    if re.match(r"^[\-\*]\s+", stripped) or re.match(r"^\d+\.\s+", stripped):
        lst, i = consume_list(lines, i)
        story.append(lst)
        story.append(Spacer(1, 4))
        continue

    # Blank line
    if stripped == "":
        i += 1
        continue

    # Paragraph (possibly multi-line)
    buf = [stripped]
    i += 1
    while i < len(lines) and lines[i].strip() and not (
        lines[i].startswith("#")
        or lines[i].lstrip().startswith(">")
        or lines[i].lstrip().startswith("|")
        or re.match(r"^[\-\*]\s+", lines[i])
        or re.match(r"^\d+\.\s+", lines[i])
        or re.match(r"^---+\s*$", lines[i].strip())
    ):
        buf.append(lines[i].strip())
        i += 1
    para_text = " ".join(buf)
    # Insert subtitle (first paragraph after title) with SUB style
    if len(story) == 2 and first_h1_consumed:
        story.append(Paragraph(md_inline_to_reportlab(para_text), SUB))
    else:
        story.append(Paragraph(md_inline_to_reportlab(para_text), BODY))

# Footer on every page
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT, 8.5)
    canvas.setFillColor(INK3)
    footer = f"Gestor de Sala · Manual de uso · Página {doc.page}"
    canvas.drawCentredString(A4[0] / 2, 12 * mm, footer)
    canvas.restoreState()


doc = SimpleDocTemplate(
    str(OUT),
    pagesize=A4,
    leftMargin=22 * mm,
    rightMargin=22 * mm,
    topMargin=20 * mm,
    bottomMargin=22 * mm,
    title="Gestor de Sala — Manual de uso",
    author="Yurest",
    subject="Guía de uso para empleados",
)
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f"OK → {OUT}")
