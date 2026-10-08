"""Scrape every book on books.toscrape.com (a sandbox site built for scraping practice) into a clean Excel file.

Usage:
    python scrape_books.py              # writes output/books.xlsx and output/books.csv

The Excel file has two sheets: "Books" (one row per book: title, category, price, rating, stock, link) and
"Summary" (books and average price per category). Requests are spaced out so the site isn't hammered.
"""
import time
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://books.toscrape.com/"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
session = requests.Session()
session.headers["User-Agent"] = "portfolio-demo-scraper (learning project)"


def get(url):
    time.sleep(0.2)  # be polite: at most ~5 requests a second
    response = session.get(url, timeout=30)
    response.raise_for_status()
    response.encoding = "utf-8"
    return BeautifulSoup(response.text, "html.parser")


def categories():
    """Return (name, url) for every category in the sidebar."""
    soup = get(BASE)
    links = soup.select("div.side_categories ul li ul li a")
    return [(a.get_text(strip=True), urljoin(BASE, a["href"])) for a in links]


def books_in(category, url):
    """Follow the category's 'next' links and yield one dict per book."""
    while url:
        soup = get(url)
        for card in soup.select("article.product_pod"):
            link = card.select_one("h3 a")
            yield {
                "title": link["title"],
                "category": category,
                "price_gbp": float(card.select_one("p.price_color").get_text().replace("£", "").replace("Â", "")),
                "rating": RATINGS[card.select_one("p.star-rating")["class"][1]],
                "in_stock": "In stock" in card.select_one("p.availability").get_text(),
                "url": urljoin(url, link["href"]),
            }
        next_link = soup.select_one("li.next a")
        url = urljoin(url, next_link["href"]) if next_link else None


def main():
    rows = []
    cats = categories()
    for i, (name, url) in enumerate(cats, 1):
        found = list(books_in(name, url))
        rows += found
        print(f"[{i:>2}/{len(cats)}] {name:<22} {len(found):>3} books")

    books = pd.DataFrame(rows).drop_duplicates(subset="url").sort_values(["category", "title"])
    summary = (books.groupby("category")
                    .agg(books=("title", "count"), avg_price_gbp=("price_gbp", "mean"), avg_rating=("rating", "mean"))
                    .round(2).sort_values("books", ascending=False).reset_index())

    out = Path("output"); out.mkdir(exist_ok=True)
    books.to_csv(out / "books.csv", index=False)
    with pd.ExcelWriter(out / "books.xlsx", engine="openpyxl") as xl:
        books.to_excel(xl, sheet_name="Books", index=False)
        summary.to_excel(xl, sheet_name="Summary", index=False)
        for sheet, df in (("Books", books), ("Summary", summary)):
            ws = xl.sheets[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col, name in zip(ws.columns, df.columns):
                width = max(len(str(name)), *(len(str(v)) for v in df[name].head(200))) + 2
                ws.column_dimensions[col[0].column_letter].width = min(width, 60)

    print(f"\n{len(books)} books in {books['category'].nunique()} categories -> output/books.xlsx")


if __name__ == "__main__":
    main()
