"""
generate_assets.py
===================
One-off script that draws the NVISH monogram logo and window/exe icon and
saves them as static files under gui/assets/. This is NOT run by the
application itself -- it was run once to produce the committed asset
files (nvish_logo.png, nvish_icon.ico) so that end users (and
PyInstaller, at build time) never need network access or a design tool
to get a professional-looking, on-brand icon.

Re-run this only if you want to change the logo's colours/text:

    python gui/assets/generate_assets.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ASSETS_DIR = Path(__file__).resolve().parent

# NVISH brand palette (deep navy + teal accent -- a clean, professional
# "enterprise software" look that reads well at both large and small
# sizes, including as a 16x16 taskbar icon).
NAVY = (11, 42, 71)          # #0B2A47
NAVY_DARK = (7, 28, 48)      # #071C30
TEAL = (0, 168, 150)         # #00A896
WHITE = (255, 255, 255)


def _load_font(size: int) -> ImageFont.FreeTypeFont:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def draw_monogram(size: int) -> Image.Image:
    """Draw a rounded-square navy badge with a teal 'N' accent bar and
    white 'NV' monogram -- used both as the in-app header logo and (at
    smaller sizes) as the window/taskbar icon."""
    scale = 4  # supersample for crisp edges, then downscale
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    corner = int(s * 0.22)
    draw.rounded_rectangle([0, 0, s - 1, s - 1], radius=corner, fill=NAVY)

    # Subtle teal accent stripe along the bottom edge.
    stripe_h = int(s * 0.14)
    draw.rounded_rectangle(
        [0, s - stripe_h, s - 1, s - 1],
        radius=corner,
        fill=TEAL,
    )
    # Re-cover the top corners of the stripe so only a clean bottom band
    # shows (rounded rect trick: draw a plain rectangle over the top half
    # of the stripe area, above the rounding radius).
    draw.rectangle([0, s - stripe_h, s - 1, s - stripe_h + corner], fill=NAVY)
    draw.rounded_rectangle([0, s - stripe_h, s - 1, s - 1], radius=corner, fill=TEAL)
    draw.rectangle([0, s - int(stripe_h * 0.55), s - 1, s - 1], fill=TEAL)

    font = _load_font(int(s * 0.46))
    text = "NV"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(
        ((s - tw) / 2 - bbox[0], (s - th) / 2 - bbox[1] - s * 0.04),
        text,
        font=font,
        fill=WHITE,
    )

    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    logo = draw_monogram(256)
    logo.save(ASSETS_DIR / "nvish_logo.png")

    icon_sizes = [16, 24, 32, 48, 64, 128, 256]
    icon_images = [draw_monogram(s) for s in icon_sizes]
    icon_images[-1].save(
        ASSETS_DIR / "nvish_icon.ico",
        format="ICO",
        sizes=[(s, s) for s in icon_sizes],
    )
    print(f"Wrote {ASSETS_DIR / 'nvish_logo.png'}")
    print(f"Wrote {ASSETS_DIR / 'nvish_icon.ico'}")


if __name__ == "__main__":
    main()
