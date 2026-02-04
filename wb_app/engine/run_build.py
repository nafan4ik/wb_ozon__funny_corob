import os
import re
from collections import Counter
from datetime import datetime
from datetime import datetime, timedelta
import shutil

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

def is_paired_marker(s: str) -> bool:
    """
    Парность по маркерам в тексте: ПАРКР или ПАРН (без учёта регистра).
    """
    x = (s or "").upper()
    return ("ПАРКР" in x) or ("ПАРН" in x)



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

def cleanup_old_postavki(base_dir: str = r"C:\korob", keep_days: int = 2):
    """
    Удаляет папки вида postavka-YYYY-MM-DD_HH-MM-SS, если они старше keep_days.
    Если имя не парсится — fallback по mtime.
    """
    prefix = "postavka-"
    now = datetime.now()
    deadline = now - timedelta(days=keep_days)

    if not os.path.isdir(base_dir):
        return 0

    removed = 0
    for name in os.listdir(base_dir):
        if not name.startswith(prefix):
            continue

        full = os.path.join(base_dir, name)
        if not os.path.isdir(full):
            continue

        # пробуем распарсить дату из имени
        dt = None
        ts = name[len(prefix):]
        try:
            dt = datetime.strptime(ts, "%Y-%m-%d_%H-%M-%S")
        except Exception:
            dt = None

        # fallback: по времени изменения папки
        if dt is None:
            try:
                dt = datetime.fromtimestamp(os.path.getmtime(full))
            except Exception:
                continue

        if dt < deadline:
            try:
                shutil.rmtree(full, ignore_errors=True)
                removed += 1
            except Exception:
                pass

    return removed


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

def build_pool_from_index_with_pairs(
    need: Counter,
    index: dict[str, str],
    allow_contains: bool = True,
    pair_by_marker: bool = True,
) -> tuple[list[tuple[str, str]], list[str]]:
    """
    Универсальный пул:
    - если есть пара ART(1) и ART(2) в индексе -> добавляем 2*qty
    - иначе обычный ART -> qty

    pair_by_marker=True: сначала пробуем маркер ПАРКР/ПАРН (в article),
    но для OZON можно оставить True — не мешает.
    """
    pool: list[tuple[str, str]] = []
    not_found: list[str] = []

    for article, qty in need.items():
        raw_article = str(article)
        key = norm_key(raw_article)

        # 1) Сначала проверим "парность": либо по маркеру, либо по факту существования (1)/(2)
        k1 = norm_key(raw_article + "(1)")
        k2 = norm_key(raw_article + "(2)")

        p1 = index.get(k1)
        p2 = index.get(k2)

        if allow_contains:
            if p1 is None:
                for k, p in index.items():
                    if k1 in k:
                        p1 = p
                        break
            if p2 is None:
                for k, p in index.items():
                    if k2 in k:
                        p2 = p
                        break

        has_pair_files = (p1 is not None and p2 is not None)
        has_pair_marker = is_paired_marker(raw_article) if pair_by_marker else False

        if has_pair_files or has_pair_marker:
            # если маркер есть, но файлов пары нет — считаем ошибкой
            if not has_pair_files:
                miss = []
                if p1 is None:
                    miss.append("(1)")
                if p2 is None:
                    miss.append("(2)")
                not_found.append(f"{key} missing {','.join(miss)}")
                continue

            for _ in range(int(qty)):
                pool.append((key, p1))
                pool.append((key, p2))
            continue

        # 2) обычный режим
        path = index.get(key)
        if path is None and allow_contains:
            for k, p in index.items():
                if key in k:
                    path = p
                    break

        if path is None:
            not_found.append(key)
            continue

        for _ in range(int(qty)):
            pool.append((key, path))

    return pool, not_found


def build_wb_pool_from_index(
    need: Counter,
    index: dict[str, str],
    allow_contains: bool = True,
) -> tuple[list[tuple[str, str]], list[str]]:
    """
    WB пул: [(article_for_name, path), ...]
    Для парных артикулов ищем ART(1) и ART(2) и добавляем 2*qty элементов.
    """
    pool: list[tuple[str, str]] = []
    not_found: list[str] = []

    for article, qty in need.items():
        raw_article = str(article)
        key = norm_key(raw_article)

        if is_paired_marker(raw_article):
            k1 = norm_key(raw_article + "(1)")
            k2 = norm_key(raw_article + "(2)")

            p1 = index.get(k1)
            p2 = index.get(k2)

            # fallback по вхождению (на случай если в имени есть ещё что-то)
            if allow_contains:
                if p1 is None:
                    for k, p in index.items():
                        if k1 in k:
                            p1 = p
                            break
                if p2 is None:
                    for k, p in index.items():
                        if k2 in k:
                            p2 = p
                            break

            if p1 is None or p2 is None:
                miss = []
                if p1 is None:
                    miss.append("(1)")
                if p2 is None:
                    miss.append("(2)")
                not_found.append(f"{key} missing {','.join(miss)}")
                continue

            # на каждую "штуку" добавляем пару (1)+(2)
            for _ in range(int(qty)):
                pool.append((key, p1))
                pool.append((key, p2))

        else:
            # обычный (не парный)
            path = index.get(key)

            if path is None and allow_contains:
                for k, p in index.items():
                    if key in k:
                        path = p
                        break

            if path is None:
                not_found.append(key)
                continue

            for _ in range(int(qty)):
                pool.append((key, path))

    return pool, not_found



# ======================= OZON =======================
def ozon_pdf_to_df(pdf_path: str):
    import pandas as pd

    row_ship_re = re.compile(r"^\s*(\d+)\s+(\d{6,}-\d{4}-\d)\b")
    ship_re = re.compile(r"\b\d{6,}-\d{4}-\d\b")
    art_qty_re = re.compile(r"\b(\d{8,9})\b\s+(\d{1,3})\b")

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
    OZON одиночные:
    - вытаскиваем артикул 8-9 цифр
    - если в имени есть (1) или (2) -> ключ ART(1)/ART(2)
    - иначе ключ ART
    """
    IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
    idx: dict[str, str] = {}

    art_re = re.compile(r"\b(\d{8,9})\b")
    pair_re = re.compile(r"\((\s*[12]\s*)\)")

    for f in os.listdir(makets_dir):
        p = os.path.join(makets_dir, f)
        if not os.path.isfile(p):
            continue

        name, ext = os.path.splitext(f)
        if ext.lower() not in IMAGE_EXTS:
            continue

        m_art = art_re.search(name)
        if not m_art:
            continue

        art = m_art.group(1)
        m_pair = pair_re.search(name)

        if m_pair:
            n = m_pair.group(1).strip()  # 1 или 2
            key = norm_key(f"{art}({n})")
        else:
            key = norm_key(art)

        idx[key] = p

    return idx


def ozon_pair_exists_in_index(article: str, index: dict[str, str]) -> tuple[bool, bool]:
    """
    Возвращает (has_1, has_2) для ART(1)/ART(2) в индексе.
    """
    k1 = norm_key(str(article) + "(1)")
    k2 = norm_key(str(article) + "(2)")
    return (k1 in index), (k2 in index)
def build_ozon_pool_from_index(
    need: Counter,
    index: dict[str, str],
    allow_contains: bool = True,
) -> tuple[list[tuple[str, str]], list[str]]:
    pool: list[tuple[str, str]] = []
    not_found: list[str] = []

    for article, qty in need.items():
        raw = str(article)
        key = norm_key(raw)

        k1 = norm_key(raw + "(1)")
        k2 = norm_key(raw + "(2)")

        p1 = index.get(k1)
        p2 = index.get(k2)

        if allow_contains:
            if p1 is None:
                for k, p in index.items():
                    if k1 in k:
                        p1 = p
                        break
            if p2 is None:
                for k, p in index.items():
                    if k2 in k:
                        p2 = p
                        break

        # Считаем парным, если в папке есть хотя бы один из (1)/(2)
        if p1 is not None or p2 is not None:
            miss = []
            if p1 is None:
                miss.append("(1)")
            if p2 is None:
                miss.append("(2)")

            if miss:
                not_found.append(f"{key} missing {','.join(miss)}")
                continue

            for _ in range(int(qty)):
                pool.append((key, p1))
                pool.append((key, p2))
            continue

        # обычный
        path = index.get(key)
        if path is None and allow_contains:
            for k, p in index.items():
                if key in k:
                    path = p
                    break

        if path is None:
            not_found.append(key)
            continue

        for _ in range(int(qty)):
            pool.append((key, path))

    return pool, not_found



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
    # перед созданием новой поставки чистим старые
    cleanup_old_postavki(r"C:\korob", keep_days=2)
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
    wb_pool, wb_nf = build_pool_from_index_with_pairs(wb_need, wb_index, allow_contains=True, pair_by_marker=True)

    compose_sheets(wb_pool, wb_sheets_dir, gap=gap)
    wb_sheets_count = (len(wb_pool) + 2) // 3

    # ===== OZON =====
    oz_df = ozon_pdf_to_df(ozon_pdf)

    if save_ozon_xlsx:
        import openpyxl  # noqa: F401
        oz_df.to_excel(os.path.join(dest_dir, ozon_xlsx_name), index=False)

    oz_need = ozon_need_from_df(oz_df)
    oz_index = build_ozon_single_index(ozon_singles_dir)
    sample = [k for k in oz_index.keys() if "(1)" in k or "(2)" in k][:20]
    print("OZON pair keys sample:", sample)
    print("OZON index size:", len(oz_index))

    oz_pool, oz_nf = build_ozon_pool_from_index(oz_need, oz_index, allow_contains=True)


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
