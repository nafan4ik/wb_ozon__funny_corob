import os
import re
from collections import Counter
from datetime import datetime

import pdfplumber
from PIL import Image
import numpy as np

def make_unique_sheet_path(dir_path: str, base_name: str, ext: str = ".jpg") -> str:
    base_name = base_name.strip()
    if not base_name:
        base_name = "SHEET"

    path = os.path.join(dir_path, base_name + ext)
    if not os.path.exists(path):
        return path

    i = 2
    while True:
        path = os.path.join(dir_path, f"{base_name} ({i}){ext}")
        if not os.path.exists(path):
            return path
        i += 1

def norm_key(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").replace("\xa0", " ")).upper()


def norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()

def compose_sheets(
    pool: list[tuple[str, str]],
    sheets_dir: str,
    gap: int = 4,
):
    """
    pool: список (article, image_path)
    Собираем листы по 3.
    Если остался 1 или 2 — добиваем белой заглушкой(и).
    Имя файла: ART1+ART2+ART3.jpg (и если занято -> (2), (3)...)
    """
    os.makedirs(sheets_dir, exist_ok=True)

    for i in range(0, len(pool), 3):
        batch = pool[i:i + 3]

        articles = [a for a, _ in batch]
        images = [Image.open(p).convert("RGB") for _, p in batch]

        # добивка белым до 3
        if len(images) < 3:
            widths = [im.width for im in images]
            heights = [im.height for im in images]
            base_w = max(widths) if widths else 1000
            base_h = int(np.median(heights)) if heights else 400

            while len(images) < 3:
                images.append(Image.new("RGB", (base_w, base_h), "white"))
                articles.append("EMPTY")

        W = max(im.width for im in images)
        H = sum(im.height for im in images) + gap * (len(images) - 1)

        sheet = Image.new("RGB", (W, H), "white")

        y = 0
        for im in images:
            x = (W - im.width) // 2
            sheet.paste(im, (x, y))
            y += im.height + gap

        base_name = "+".join(articles)
        out_path = make_unique_sheet_path(sheets_dir, base_name, ext=".jpg")
        sheet.save(out_path)



def cleanup_postavka_leave_images_only(dest_dir: str) -> int:
    """
    Чистим только корень папки поставки, подпапки не трогаем.
    """
    IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
    removed = 0
    for f in os.listdir(dest_dir):
        p = os.path.join(dest_dir, f)
        if not os.path.isfile(p):
            continue
        _, ext = os.path.splitext(f)
        if ext.lower() not in IMAGE_EXTS:
            os.remove(p)
            removed += 1
    return removed


# ======================= WB =======================
def extract_wb_articles(pdf_path: str) -> list[str]:
    """
    Возвращает список артикулов WB, где каждый элемент = 1 штука.
    """
    articles = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table or not table[0]:
                    continue
                headers = table[0]
                if "Артикул продавца" not in headers:
                    continue
                idx = headers.index("Артикул продавца")
                for row in table[1:]:
                    if row and len(row) > idx and row[idx]:
                        articles.append(norm_key(row[idx]))
    return articles


def wb_needed_counter(wb_articles: list[str]) -> Counter:
    return Counter(norm_key(a) for a in wb_articles)


def build_wb_single_index(makets_dir: str) -> dict[str, str]:
    """
    WB одиночные: имя файла БЕЗ расширения = артикул продавца (ключ).
    """
    IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
    idx: dict[str, str] = {}

    for f in os.listdir(makets_dir):
        p = os.path.join(makets_dir, f)
        if not os.path.isfile(p):
            continue
        name, ext = os.path.splitext(f)
        if ext.lower() not in IMAGE_EXTS:
            continue
        idx[norm_key(name)] = p

    return idx


def build_pool_from_index(
    need: Counter,
    index: dict[str, str],
    allow_contains: bool = True
) -> tuple[list[tuple[str, str]], list[str]]:
    """
    Возвращает pool: [(article, path), ...] ровно нужное кол-во.
    """
    pool: list[tuple[str, str]] = []
    not_found: list[str] = []

    for article, qty in need.items():
        key = norm_key(article)
        path = index.get(key)

        if path is None and allow_contains:
            for k, p in index.items():
                if key in k:
                    path = p
                    break

        if path is None:
            not_found.append(article)
            continue

        for _ in range(int(qty)):
            pool.append((key, path))

    return pool, not_found



# ======================= OZON =======================
def ozon_pdf_to_df(pdf_path: str):
    import pandas as pd

    row_ship_re = re.compile(r"^\s*(\d+)\s+(\d{6,}-\d{4}-\d)\b")
    ship_re = re.compile(r"\b\d{6,}-\d{4}-\d\b")
    art_qty_re = re.compile(r"\b(\d{9})\b\s+(\d{1,3})\b")

    rows = []
    current_shipment = None
    current_row_no = None

    with pdfplumber.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            for line_no, line in enumerate(text.splitlines(), start=1):
                raw = norm_spaces(line)
                if not raw:
                    continue

                m = row_ship_re.search(raw)
                if m:
                    current_row_no = int(m.group(1))
                    current_shipment = m.group(2)
                else:
                    m2 = ship_re.search(raw)
                    if m2:
                        current_shipment = m2.group(0)

                pairs = art_qty_re.findall(raw)
                for art, qty_s in pairs:
                    if not current_shipment:
                        continue
                    rows.append({
                        "page": pno,
                        "line_no": line_no,          # ✅ добавили номер строки на странице
                        "row_no": current_row_no,
                        "shipment": current_shipment,
                        "article": art,
                        "qty": int(qty_s),
                        "raw_line": raw
                    })

    df = pd.DataFrame(rows)

    # ✅ ВАЖНО: НЕ ДЕЛАЕМ drop_duplicates, иначе можно потерять реальные повторяющиеся позиции
    return df



def ozon_need_from_df(df) -> Counter:
    """
    OZON: из df делаем need(article -> total_qty)
    (shipment не нужен для печати листов, только для аналитики/отчёта).
    """
    need = Counter()
    if df is None or df.empty:
        return need

    for _, r in df.iterrows():
        art = str(r["article"])
        qty = int(r["qty"])
        need[norm_key(art)] += qty

    return need


def build_ozon_single_index(makets_dir: str) -> dict[str, str]:
    """
    OZON одиночные: вытаскиваем 9-значный артикул из имени файла.
    Если в названии несколько 9-значных — берём первый.
    """
    IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
    idx: dict[str, str] = {}

    for f in os.listdir(makets_dir):
        p = os.path.join(makets_dir, f)
        if not os.path.isfile(p):
            continue
        name, ext = os.path.splitext(f)
        if ext.lower() not in IMAGE_EXTS:
            continue

        found = re.findall(r"\b\d{9}\b", name)
        if not found:
            continue

        art = norm_key(found[0])
        # если несколько файлов на один артикул — оставим первый (можно поменять логику)
        idx.setdefault(art, p)

    return idx


# ======================= RUN =======================
def run(
    wb_pdf: str,
    wb_singles_dir: str,
    ozon_pdf: str,
    ozon_singles_dir: str,
    gap: int = 4,
    save_ozon_xlsx: bool = True,
    ozon_xlsx_name: str = "ozon_table.xlsx",
) -> dict:
    """
    Главная функция для UI.
    WB и OZON оба работают через папки одиночных макетов и оба собирают листы по 3.
    """
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    dest_dir = rf"C:\korob\postavka-{now}"
    os.makedirs(dest_dir, exist_ok=True)

    wb_sheets_dir = os.path.join(dest_dir, "wb")
    ozon_sheets_dir = os.path.join(dest_dir, "ozon")

    os.makedirs(wb_sheets_dir, exist_ok=True)
    os.makedirs(ozon_sheets_dir, exist_ok=True)

    # ===== WB =====
    wb_articles = extract_wb_articles(wb_pdf)
    wb_need = wb_needed_counter(wb_articles)

    wb_index = build_wb_single_index(wb_singles_dir)
    wb_pool, wb_nf = build_pool_from_index(wb_need, wb_index, allow_contains=True)

    compose_sheets(wb_pool, wb_sheets_dir, gap=gap)
    wb_sheets_count = (len(wb_pool) + 2) // 3

    # ===== OZON =====
    oz_df = ozon_pdf_to_df(ozon_pdf)

    if save_ozon_xlsx:
        import openpyxl  # noqa: F401
        oz_df.to_excel(os.path.join(dest_dir, ozon_xlsx_name), index=False)

    oz_need = ozon_need_from_df(oz_df)
    oz_index = build_ozon_single_index(ozon_singles_dir)
    oz_pool, oz_nf = build_pool_from_index(oz_need, oz_index, allow_contains=True)

    compose_sheets(oz_pool, ozon_sheets_dir, gap=gap)
    oz_sheets_count = (len(oz_pool) + 2) // 3

    removed = cleanup_postavka_leave_images_only(dest_dir)

    return {
        "dest_dir": dest_dir,

        "wb_sheets_dir": wb_sheets_dir,
        "wb_rows": len(wb_articles),
        "wb_pool": len(wb_pool),
        "wb_sheets": wb_sheets_count,
        "wb_not_found": wb_nf,

        "ozon_sheets_dir": ozon_sheets_dir,
        "oz_rows": len(oz_df),
        "oz_pool": len(oz_pool),
        "oz_sheets": oz_sheets_count,
        "oz_not_found": oz_nf,

        "cleanup_removed": removed,
    }
