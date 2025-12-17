#!/usr/bin/env python3
"""
Generate placeholder icons for the EML to PDF Converter.

This script creates simple placeholder icons that can be replaced
with custom designs later.

Requires: pip install pillow
"""

import os
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("Pillow is required to generate icons.")
    print("Install with: pip install pillow")
    sys.exit(1)


def create_base_icon(size=1024):
    """Create a base icon image."""
    # Create image with gradient background
    img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    # Background rounded rectangle
    padding = size // 8
    corner_radius = size // 6

    # Draw rounded rectangle background (blue gradient effect)
    for i in range(size - 2 * padding):
        color_val = int(66 + (i / (size - 2 * padding)) * 30)
        draw.rectangle(
            [padding, padding + i, size - padding, padding + i + 1],
            fill=(52, color_val, 168, 255)
        )

    # Draw envelope shape
    envelope_padding = size // 4
    envelope_height = size // 3

    # Envelope body
    draw.rectangle(
        [envelope_padding, size // 3, size - envelope_padding, size // 3 + envelope_height],
        fill=(255, 255, 255, 255),
        outline=(200, 200, 200, 255),
        width=size // 50
    )

    # Envelope flap (triangle)
    flap_points = [
        (envelope_padding, size // 3),
        (size // 2, size // 3 + envelope_height // 2),
        (size - envelope_padding, size // 3)
    ]
    draw.polygon(flap_points, fill=(240, 240, 240, 255))
    draw.line(flap_points[:2], fill=(200, 200, 200, 255), width=size // 50)
    draw.line(flap_points[1:], fill=(200, 200, 200, 255), width=size // 50)

    # PDF badge
    badge_size = size // 3
    badge_x = size - envelope_padding - badge_size // 2
    badge_y = size // 3 + envelope_height - badge_size // 2

    draw.ellipse(
        [badge_x - badge_size // 2, badge_y - badge_size // 2,
         badge_x + badge_size // 2, badge_y + badge_size // 2],
        fill=(220, 53, 69, 255)
    )

    # PDF text
    try:
        # Try to use a built-in font
        font_size = badge_size // 3
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)
    except (IOError, OSError):
        try:
            font = ImageFont.truetype("arial.ttf", badge_size // 3)
        except (IOError, OSError):
            font = ImageFont.load_default()

    # Draw "PDF" text
    text = "PDF"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    text_x = badge_x - text_width // 2
    text_y = badge_y - text_height // 2 - bbox[1]
    draw.text((text_x, text_y), text, fill=(255, 255, 255, 255), font=font)

    return img


def save_png(img, path, size=512):
    """Save as PNG at specified size."""
    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    resized.save(path, 'PNG')
    print(f"Created: {path}")


def save_ico(img, path):
    """Save as Windows ICO with multiple sizes."""
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    icons = [img.resize(size, Image.Resampling.LANCZOS) for size in sizes]
    icons[0].save(path, format='ICO', sizes=sizes)
    print(f"Created: {path}")


def save_icns(img, path):
    """Save as macOS ICNS."""
    # ICNS requires specific sizes
    # We'll save as PNG and note that conversion is needed
    png_path = path.replace('.icns', '.png')
    img.save(png_path, 'PNG')
    print(f"Created: {png_path}")
    print(f"  Note: Convert to .icns using: iconutil -c icns {png_path}")
    print(f"  Or use an online converter for {path}")


def main():
    """Generate all icon formats."""
    script_dir = Path(__file__).parent
    os.chdir(script_dir)

    print("Generating placeholder icons...")
    print()

    # Create base icon
    base = create_base_icon(1024)

    # Save in different formats
    save_png(base, 'icon.png', 512)
    save_ico(base, 'icon.ico')

    # For ICNS, save as PNG first (needs manual conversion)
    base.save('icon_1024.png', 'PNG')
    print(f"Created: icon_1024.png")
    print(f"  To create icon.icns on macOS:")
    print(f"    mkdir icon.iconset")
    print(f"    sips -z 16 16 icon_1024.png --out icon.iconset/icon_16x16.png")
    print(f"    sips -z 32 32 icon_1024.png --out icon.iconset/icon_16x16@2x.png")
    print(f"    sips -z 32 32 icon_1024.png --out icon.iconset/icon_32x32.png")
    print(f"    sips -z 64 64 icon_1024.png --out icon.iconset/icon_32x32@2x.png")
    print(f"    sips -z 128 128 icon_1024.png --out icon.iconset/icon_128x128.png")
    print(f"    sips -z 256 256 icon_1024.png --out icon.iconset/icon_128x128@2x.png")
    print(f"    sips -z 256 256 icon_1024.png --out icon.iconset/icon_256x256.png")
    print(f"    sips -z 512 512 icon_1024.png --out icon.iconset/icon_256x256@2x.png")
    print(f"    sips -z 512 512 icon_1024.png --out icon.iconset/icon_512x512.png")
    print(f"    sips -z 1024 1024 icon_1024.png --out icon.iconset/icon_512x512@2x.png")
    print(f"    iconutil -c icns icon.iconset")

    print()
    print("Icon generation complete!")
    print("Replace these placeholders with your custom icons for production builds.")


if __name__ == "__main__":
    main()
