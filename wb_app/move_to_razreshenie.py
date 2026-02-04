import os
from pathlib import Path
from PIL import Image

SRC_DIR = Path(r"C:\korob\wb_makets_po_odnomu")
DST_DIR = Path(r"C:\korob\wb_makets_po_odnomu_2307x1128")

TARGET_W, TARGET_H = 2307, 1128

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}


def contain_to_canvas(img: Image.Image, target_w: int, target_h: int) -> Image.Image:
    """
    Масштабируем с сохранением пропорций (ничего не режем),
    кладём по центру на белый холст target_w x target_h.
    """
    img = img.convert("RGB")
    iw, ih = img.size

    scale = min(target_w / iw, target_h / ih)
    new_w = max(1, int(round(iw * scale)))
    new_h = max(1, int(round(ih * scale)))

    resized = img.resize((new_w, new_h), Image.LANCZOS)

    canvas = Image.new("RGB", (target_w, target_h), "white")
    x = (target_w - new_w) // 2
    y = (target_h - new_h) // 2
    canvas.paste(resized, (x, y))
    return canvas


def main():
    if not SRC_DIR.exists():
        raise FileNotFoundError(f"Не найдена папка: {SRC_DIR}")

    DST_DIR.mkdir(parents=True, exist_ok=True)

    total = 0
    converted = 0
    skipped = 0
    errors = 0

    for p in SRC_DIR.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in IMAGE_EXTS:
            continue

        total += 1
        out_path = DST_DIR / (p.stem + ".jpg")

        try:
            with Image.open(p) as im:
                out = contain_to_canvas(im, TARGET_W, TARGET_H)
                out.save(out_path, "JPEG", quality=95, subsampling=0)
            converted += 1
        except Exception as e:
            errors += 1
            print(f"[ERR] {p.name}: {e}")

    print("==== ГОТОВО ====")
    print(f"Источник: {SRC_DIR}")
    print(f"Результат: {DST_DIR}")
    print(f"Файлов найдено: {total}")
    print(f"Сконвертировано: {converted}")
    print(f"Пропущено: {skipped}")
    print(f"Ошибок: {errors}")


if __name__ == "__main__":
    main()
