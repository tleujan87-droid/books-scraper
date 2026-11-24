"""
Учебный парсер для портфолио.
Собирает все книги с books.toscrape.com (сайт создан специально для практики парсинга)
и сохраняет их в Excel: название, цена, рейтинг, наличие, ссылка.

Запуск:  python parser.py
Результат: books.xlsx рядом со скриптом
"""

import csv
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/catalogue/"
START_PAGE = "page-1.html"
OUTPUT = Path(__file__).with_name("books.xlsx")
DELAY_SECONDS = 0.5  # пауза между страницами, чтобы не нагружать сайт

RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

session = requests.Session()
session.headers["User-Agent"] = "Mozilla/5.0 (portfolio scraper demo)"


def parse_page(url: str):
    """Возвращает список книг со страницы и ссылку на следующую страницу (или None)."""
    response = session.get(url, timeout=20)
    response.raise_for_status()
    response.encoding = "utf-8"
    soup = BeautifulSoup(response.text, "html.parser")

    books = []
    for card in soup.select("article.product_pod"):
        link = card.h3.a
        rating_word = card.select_one("p.star-rating")["class"][1]
        price_text = card.select_one("p.price_color").get_text(strip=True)
        books.append({
            "Название": link["title"],
            "Цена (£)": float(price_text.replace("£", "").replace("Â", "")),
            "Рейтинг": RATINGS.get(rating_word, 0),
            "Наличие": card.select_one("p.availability").get_text(strip=True),
            "Ссылка": urljoin(url, link["href"]),
        })

    next_btn = soup.select_one("li.next a")
    next_url = urljoin(url, next_btn["href"]) if next_btn else None
    return books, next_url


def save_to_excel(books, path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Книги"

    headers = list(books[0].keys())
    ws.append(headers)
    for book in books:
        ws.append([book[h] for h in headers])

    # оформление: жирная шапка, фильтр, закреплённая строка, ширина колонок
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    widths = {"Название": 60, "Цена (£)": 12, "Рейтинг": 10, "Наличие": 14, "Ссылка": 70}
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 15)

    wb.save(path)


def save_to_csv(books, path: Path):
    # utf-8-sig — чтобы Excel правильно открывал кириллицу в CSV
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(books[0].keys()))
        writer.writeheader()
        writer.writerows(books)


def main():
    all_books = []
    url = urljoin(BASE_URL, START_PAGE)
    page = 1
    while url:
        books, url = parse_page(url)
        all_books.extend(books)
        print(f"Страница {page}: +{len(books)} книг (всего {len(all_books)})")
        page += 1
        time.sleep(DELAY_SECONDS)

    save_to_excel(all_books, OUTPUT)
    save_to_csv(all_books, OUTPUT.with_suffix(".csv"))
    print(f"\nГотово! Сохранено {len(all_books)} книг в {OUTPUT.name} и {OUTPUT.with_suffix('.csv').name}")


if __name__ == "__main__":
    main()
