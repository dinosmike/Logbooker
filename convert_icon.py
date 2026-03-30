from pathlib import Path

from PIL import Image, ImageFilter, ImageOps


def convert_png_to_multi_ico(png_path: Path, ico_path: Path) -> None:
    img = Image.open(png_path).convert("RGBA")
    # Trim transparent margins so symbol occupies more icon area in thumbnails.
    alpha = img.getchannel("A")
    bbox = alpha.getbbox()
    if bbox:
        img = img.crop(bbox)

    # Normalize to square with a small transparent margin.
    side = max(img.size)
    padded_side = int(side * 1.08)
    img = ImageOps.pad(img, (padded_side, padded_side), method=Image.Resampling.LANCZOS, color=(0, 0, 0, 0))

    # Create a crisp master at 1024 for better downsampling quality.
    master = img.resize((1024, 1024), Image.Resampling.LANCZOS)
    master = master.filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=2))
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (96, 96), (128, 128), (256, 256)]

    # Pillow builds all ICO layers from this master using `sizes`.
    master.save(ico_path, format="ICO", sizes=sizes)


if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    src = base / "iconMY.png"
    dst = base / "iconMY.ico"
    if not src.exists():
        raise SystemExit("iconMY.png not found")
    convert_png_to_multi_ico(src, dst)
    print(f"Created {dst.name}")
