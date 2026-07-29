"""Generate a portfolio thumbnail showing Excel spreadsheet preview."""
import json
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent / "output"

with open(OUTPUT_DIR / "tonies_enriched_2026-07-19_23-54-33.json") as f:
    stores = json.load(f)

WIDTH, HEIGHT = 1200, 800
img = Image.new("RGB", (WIDTH, HEIGHT), "#FFFFFF")
draw = ImageDraw.Draw(img)

try:
    font = ImageFont.truetype("calibri.ttf", 14)
    font_bold = ImageFont.truetype("calibrib.ttf", 14)
    font_title = ImageFont.truetype("calibrib.ttf", 22)
    font_small = ImageFont.truetype("calibri.ttf", 12)
except:
    font = ImageFont.load_default()
    font_bold = font
    font_title = font
    font_small = font

# Title bar
draw.rectangle([0, 0, WIDTH, 50], fill="#1B5E20")
draw.text((20, 12), "Tonies USA Retailer Data Extraction — 4,503 Stores", fill="#FFFFFF", font=font_title)

# Stats bar
draw.rectangle([0, 50, WIDTH, 85], fill="#E8F5E9")
has_website = sum(1 for s in stores if s.get("Website URL"))
has_email = sum(1 for s in stores if s.get("Email Address"))
has_contact = sum(1 for s in stores if s.get("Vendor/Wholesale Contact Form URL"))
stats_text = f"Coverage:  {len(stores)} retailers  |  {has_website} websites ({has_website*100//len(stores)}%)  |  {has_email} emails  |  {has_contact} contact forms ({has_contact*100//len(stores)}%)"
draw.text((20, 60), stats_text, fill="#1B5E20", font=font_small)

# Column headers
headers = ["Retailer Name", "Website URL", "Email Address", "Contact Form URL", "City", "State"]
col_widths = [220, 280, 200, 260, 140, 60]
col_x = [20]
for w in col_widths[:-1]:
    col_x.append(col_x[-1] + w)

header_y = 95
draw.rectangle([0, header_y, WIDTH, header_y + 30], fill="#C41E3A")
for i, header in enumerate(headers):
    draw.text((col_x[i] + 5, header_y + 6), header, fill="#FFFFFF", font=font_bold)

# Data rows
row_height = 28
visible_rows = 22
for row_idx in range(visible_rows):
    if row_idx >= len(stores):
        break
    store = stores[row_idx]
    y = header_y + 30 + row_idx * row_height
    bg = "#F5F5F5" if row_idx % 2 == 0 else "#FFFFFF"
    draw.rectangle([0, y, WIDTH, y + row_height], fill=bg)

    vals = [
        store.get("Retailer Name", "")[:28],
        store.get("Website URL", "")[:35],
        store.get("Email Address", "")[:25],
        store.get("Vendor/Wholesale Contact Form URL", "")[:32],
        store.get("City", "")[:16],
        store.get("State", ""),
    ]
    for i, val in enumerate(vals):
        color = "#0563C1" if i in (1, 3) and val.startswith("http") else "#333333"
        draw.text((col_x[i] + 5, y + 5), val, fill=color, font=font)

    # Grid lines
    draw.line([(0, y + row_height), (WIDTH, y + row_height)], fill="#E0E0E0")

# Vertical grid lines
for x in col_x:
    draw.line([(x, header_y), (x, header_y + 30 + visible_rows * row_height)], fill="#E0E0E0")

# Fade-out at bottom
for i in range(40):
    alpha = int(255 * i / 40)
    y = HEIGHT - 40 + i
    draw.rectangle([0, y, WIDTH, y + 1], fill=(255, 255, 255, alpha))

# Footer
draw.rectangle([0, HEIGHT - 35, WIDTH, HEIGHT], fill="#F5F5F5")
draw.text((20, HEIGHT - 28), f"Total: {len(stores)} unique retailers across 49 states + DC  |  Python + ChannelSight API + DuckDuckGo enrichment", fill="#666666", font=font_small)

out_path = OUTPUT_DIR / "portfolio_thumbnail.png"
img.save(str(out_path), "PNG")
print(f"Saved thumbnail to {out_path}")
print(f"Size: {img.size}")
