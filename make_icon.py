from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


def make_icon() -> None:
    size = 1024
    img = Image.new("RGBA", (size, size), (7, 17, 34, 255))
    draw = ImageDraw.Draw(img)

    # Soft radial glow.
    glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(glow)
    for r, alpha in ((420, 18), (320, 28), (220, 40), (140, 56)):
        box = (size // 2 - r, size // 2 - r, size // 2 + r, size // 2 + r)
        gdraw.ellipse(box, fill=(36, 116, 255, alpha))
    img = Image.alpha_composite(img, glow.filter(ImageFilter.GaussianBlur(20)))
    draw = ImageDraw.Draw(img)

    # Rounded tile border.
    draw.rounded_rectangle((36, 36, size - 36, size - 36), radius=180, outline=(73, 132, 255, 180), width=10)

    # Logbook pages.
    left_page = [(260, 300), (500, 260), (500, 760), (260, 800)]
    right_page = [(520, 260), (760, 300), (760, 800), (520, 760)]
    draw.polygon(left_page, fill=(18, 34, 64, 255), outline=(100, 170, 255, 180))
    draw.polygon(right_page, fill=(20, 40, 74, 255), outline=(100, 170, 255, 180))

    # Center fold.
    draw.line((510, 260, 510, 770), fill=(66, 124, 226, 180), width=6)

    # Jet silhouette.
    jet = [(220, 500), (450, 455), (560, 405), (780, 360), (610, 500), (780, 640), (560, 595), (450, 545)]
    draw.polygon(jet, fill=(215, 232, 255, 245))

    # Check mark.
    draw.line((315, 635, 390, 710), fill=(255, 198, 66, 255), width=28)
    draw.line((390, 710, 470, 560), fill=(255, 198, 66, 255), width=28)

    # Pencil on right page.
    draw.polygon([(605, 640), (700, 545), (730, 575), (635, 670)], fill=(60, 219, 255, 255))
    draw.polygon([(730, 575), (756, 549), (770, 589), (744, 615)], fill=(148, 245, 255, 255))

    base = Path(__file__).resolve().parent
    png_path = base / "app_icon.png"
    ico_path = base / "app_icon.ico"

    img.save(png_path)
    img.save(ico_path, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"Created: {png_path.name}, {ico_path.name}")


if __name__ == "__main__":
    make_icon()
