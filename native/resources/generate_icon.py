#!/usr/bin/env python3
"""Generate Nexel app icon — Excel-style: bold green 'N' panel over spreadsheet doc."""

from PIL import Image, ImageDraw, ImageFont
import os
import subprocess


SIZE = 1024
CORNER_RADIUS = 180


def rounded_rectangle_mask(size, radius):
    mask = Image.new('L', (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return mask


def draw_shadow(img, bbox, radius, offset=6, blur_steps=18, alpha=30):
    """Draw a soft drop shadow."""
    ov = Image.new('RGBA', img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    x1, y1, x2, y2 = bbox
    for i in range(blur_steps):
        a = int(alpha * (1 - i / blur_steps))
        d.rounded_rectangle(
            [x1 + i - 1, y1 + i + offset, x2 + i - 1, y2 + i + offset],
            radius=radius + i, fill=(0, 0, 0, a)
        )
    return Image.alpha_composite(img, ov)


def draw_icon(size=SIZE):
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # === BACKGROUND: rich green gradient ===
    for y in range(size):
        t = y / size
        r = int(19 + 8 * t)
        g = int(115 + 18 * t)
        b = int(62 + 12 * t)
        for x in range(size):
            img.putpixel((x, y), (r, g, b, 255))

    draw = ImageDraw.Draw(img)

    # ======================================================
    #  DOCUMENT PAGE — right side, slightly behind the N tile
    # ======================================================
    doc_left = int(size * 0.32)
    doc_top = int(size * 0.10)
    doc_right = int(size * 0.92)
    doc_bottom = int(size * 0.90)
    doc_rad = int(size * 0.04)

    # Shadow for document
    img = draw_shadow(img, [doc_left, doc_top, doc_right, doc_bottom],
                      doc_rad, offset=8, blur_steps=22, alpha=35)
    draw = ImageDraw.Draw(img)

    # Document background
    draw.rounded_rectangle([doc_left, doc_top, doc_right, doc_bottom],
                           radius=doc_rad, fill=(255, 255, 255, 250))

    # --- Spreadsheet grid on the document ---
    margin_l = doc_left + int(size * 0.025)
    margin_r = doc_right - int(size * 0.025)
    margin_t = doc_top + int(size * 0.025)
    margin_b = doc_bottom - int(size * 0.025)

    grid_w = margin_r - margin_l
    grid_h = margin_b - margin_t

    ncols = 3
    nrows = 8
    col_w = grid_w / ncols
    row_h = grid_h / nrows

    # Header row — light green
    hdr_h = row_h * 0.9
    hdr_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    hd = ImageDraw.Draw(hdr_ov)
    hd.rounded_rectangle(
        [margin_l, margin_t, margin_r, int(margin_t + hdr_h)],
        radius=8, fill=(33, 135, 75, 50)
    )
    img = Image.alpha_composite(img, hdr_ov)
    draw = ImageDraw.Draw(img)

    # Alternating row stripes
    stripe_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    sd = ImageDraw.Draw(stripe_ov)
    for r in range(1, nrows):
        if r % 2 == 0:
            y1 = int(margin_t + r * row_h)
            y2 = int(margin_t + (r + 1) * row_h)
            if y2 <= margin_b:
                sd.rectangle([margin_l, y1, margin_r, y2], fill=(33, 135, 75, 12))
    img = Image.alpha_composite(img, stripe_ov)
    draw = ImageDraw.Draw(img)

    # Grid lines
    line_color = (180, 195, 185, 50)
    # Horizontal
    for r in range(1, nrows):
        y = int(margin_t + r * row_h)
        if y < margin_b:
            draw.line([(margin_l, y), (margin_r, y)], fill=line_color, width=1)
    # Vertical
    for c in range(1, ncols):
        x = int(margin_l + c * col_w)
        draw.line([(x, margin_t), (x, margin_b)], fill=line_color, width=1)

    # Colored data bars in some cells (like mini bar chart in cells)
    bar_colors = [
        (52, 168, 83, 120),   # green
        (66, 133, 244, 120),  # blue
        (234, 67, 53, 100),   # red
        (251, 188, 4, 110),   # yellow
        (52, 168, 83, 90),
        (66, 133, 244, 100),
    ]
    bar_widths = [0.75, 0.55, 0.85, 0.40, 0.65, 0.50]
    bar_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bar_ov)

    for r_idx, (bw_frac, bcolor) in enumerate(zip(bar_widths, bar_colors)):
        r = r_idx + 2  # start from row 2
        if r >= nrows:
            break
        # Draw in the last column
        c = ncols - 1
        cx1 = int(margin_l + c * col_w) + 6
        cy1 = int(margin_t + r * row_h) + 4
        cx2 = cx1 + int((col_w - 12) * bw_frac)
        cy2 = int(margin_t + (r + 1) * row_h) - 4
        bd.rounded_rectangle([cx1, cy1, cx2, cy2], radius=3, fill=bcolor)

    img = Image.alpha_composite(img, bar_ov)
    draw = ImageDraw.Draw(img)

    # ======================================================
    #  GREEN "N" TILE — left side, overlapping the document
    # ======================================================
    tile_left = int(size * 0.06)
    tile_top = int(size * 0.22)
    tile_right = int(size * 0.48)
    tile_bottom = int(size * 0.78)
    tile_rad = int(size * 0.055)

    # Shadow for tile
    img = draw_shadow(img, [tile_left, tile_top, tile_right, tile_bottom],
                      tile_rad, offset=10, blur_steps=24, alpha=50)
    draw = ImageDraw.Draw(img)

    # Tile gradient: dark green top to slightly brighter green bottom
    tile_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    # Create tile with gradient
    tile_w = tile_right - tile_left
    tile_h = tile_bottom - tile_top
    tile_img = Image.new('RGBA', (tile_w, tile_h), (0, 0, 0, 0))
    for y in range(tile_h):
        t = y / tile_h
        r = int(24 + 14 * t)
        g = int(120 + 30 * t)
        b = int(55 + 20 * t)
        for x in range(tile_w):
            tile_img.putpixel((x, y), (r, g, b, 255))

    # Apply rounded mask to tile
    tile_mask = Image.new('L', (tile_w, tile_h), 0)
    tm_draw = ImageDraw.Draw(tile_mask)
    tm_draw.rounded_rectangle([0, 0, tile_w - 1, tile_h - 1], radius=tile_rad, fill=255)
    tile_img.putalpha(tile_mask)

    # Paste tile onto overlay
    tile_ov.paste(tile_img, (tile_left, tile_top))
    img = Image.alpha_composite(img, tile_ov)
    draw = ImageDraw.Draw(img)

    # --- Draw the "N" letter ---
    try:
        # Try bold font first
        font_n = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.28))
    except Exception:
        font_n = ImageFont.load_default()

    letter = "N"
    bb = draw.textbbox((0, 0), letter, font=font_n)
    lw, lh = bb[2] - bb[0], bb[3] - bb[1]
    lx = tile_left + (tile_w - lw) // 2
    ly = tile_top + (tile_h - lh) // 2 - int(size * 0.01)

    # White letter with subtle shadow
    draw.text((lx + 2, ly + 3), letter, fill=(0, 0, 0, 40), font=font_n)
    draw.text((lx, ly), letter, fill=(255, 255, 255, 250), font=font_n)

    # Apply rounded app icon mask
    mask = rounded_rectangle_mask(size, CORNER_RADIUS)
    img.putalpha(mask)
    return img


def create_icns(png_path, icns_path):
    iconset_dir = png_path.replace('.png', '.iconset')
    os.makedirs(iconset_dir, exist_ok=True)
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    src = Image.open(png_path)
    for s in sizes:
        resized = src.resize((s, s), Image.LANCZOS)
        resized.save(os.path.join(iconset_dir, f'icon_{s}x{s}.png'))
        if s <= 512:
            double = src.resize((s * 2, s * 2), Image.LANCZOS)
            double.save(os.path.join(iconset_dir, f'icon_{s}x{s}@2x.png'))
    subprocess.run(['iconutil', '-c', 'icns', iconset_dir, '-o', icns_path], check=True)
    import shutil
    shutil.rmtree(iconset_dir)


if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon = draw_icon()
    png_path = os.path.join(script_dir, 'icon.png')
    icns_path = os.path.join(script_dir, 'AppIcon.icns')
    icon.save(png_path, 'PNG')
    print(f"Saved {png_path}")
    create_icns(png_path, icns_path)
    print(f"Saved {icns_path}")
