"""
PROJECT: Web Scraping and Dataset Creation
TASK: Extract book data from a public website.

Website: https://books.toscrape.com/

Libraries:
- requests
- BeautifulSoup
- pandas
- openpyxl

Output:
- books_dataset.csv
- books_dataset.xlsx
- scraping_summary.txt
"""

import time
import logging
from pathlib import Path
from urllib.parse import urljoin

import requests
import pandas as pd

from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://books.toscrape.com/"
OUTPUT_DIR = Path("output")

CSV_FILE = OUTPUT_DIR / "books_dataset.csv"
EXCEL_FILE = OUTPUT_DIR / "books_dataset.xlsx"
SUMMARY_FILE = OUTPUT_DIR / "scraping_summary.txt"

# Set to None to scrape all available catalogue pages.
# For testing, use 2 or 3.
MAX_PAGES = None

REQUEST_DELAY = 0.5

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


#  LOGGING CONFIGURATION

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logger = logging.getLogger(__name__)


#  CREATE REQUEST SESSION


def create_session():
    """
    Create a requests session with retry support.
    """

    session = requests.Session()

    retry_strategy = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )

    adapter = HTTPAdapter(max_retries=retry_strategy)

    session.mount("https://", adapter)
    session.mount("http://", adapter)

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    })

    return session


#  FETCH WEBPAGE


def fetch_page(session, url):
    """
    Download an HTML page and return its BeautifulSoup object.
    """

    try:
        response = session.get(url, timeout=20)

        response.raise_for_status()

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        return soup

    except requests.RequestException as error:
        logger.error("Failed to fetch %s: %s", url, error)
        return None



#  EXTRACT BOOK RATING


def extract_rating(book_element):
    """
    Convert a book's CSS rating class into a numeric score.
    """

    rating_element = book_element.select_one(
        "p.star-rating"
    )

    if rating_element is None:
        return None

    rating_classes = rating_element.get("class", [])

    rating_mapping = {
        "One": 1,
        "Two": 2,
        "Three": 3,
        "Four": 4,
        "Five": 5
    }

    for class_name in rating_classes:
        if class_name in rating_mapping:
            return rating_mapping[class_name]

    return None


#  EXTRACT BOOK INFORMATION


def extract_books(soup, page_url):
    """
    Extract all book records from one catalogue page.
    """

    books = []

    book_elements = soup.select("article.product_pod")

    for book in book_elements:

        # Book title
        title_element = book.select_one("h3 a")

        title = None
        book_url = None

        if title_element:

            title = title_element.get("title", "").strip()

            relative_url = title_element.get("href", "")

            if relative_url:
                book_url = urljoin(page_url, relative_url)

        # Price
        price_element = book.select_one(".price_color")

        price_text = None
        price = None

        if price_element:

            price_text = price_element.get_text(strip=True)

            # Remove currency symbol and convert to float.
            numeric_price = (
                price_text
                .replace("£", "")
                .replace("Â", "")
                .strip()
            )

            try:
                price = float(numeric_price)
            except ValueError:
                price = None

        # Availability
        availability_element = book.select_one(
            ".availability"
        )

        availability = None

        if availability_element:

            availability = " ".join(
                availability_element.stripped_strings
            )

        # Rating
        rating = extract_rating(book)

        # Image URL
        image_element = book.select_one("img")

        image_url = None

        if image_element:

            image_source = (
                image_element.get("src")
                or image_element.get("data-src")
            )

            if image_source:
                image_url = urljoin(
                    page_url,
                    image_source
                )

        # Create record
        record = {
            "title": title,
            "price_gbp": price,
            "price_original": price_text,
            "rating": rating,
            "availability": availability,
            "product_url": book_url,
            "image_url": image_url,
            "source_page": page_url
        }

        books.append(record)

    return books



#  FIND NEXT PAGE


def get_next_page(soup, current_url):
    """
    Find the next catalogue page URL.
    """

    next_link = soup.select_one("li.next a")

    if next_link is None:
        return None

    next_href = next_link.get("href")

    if not next_href:
        return None

    return urljoin(current_url, next_href)



#  SCRAPE ALL PAGES


def scrape_books():
    """
    Scrape book data from multiple catalogue pages.
    """

    session = create_session()

    all_books = []

    current_url = BASE_URL
    pages_scraped = 0

    try:

        while current_url:

            if (
                MAX_PAGES is not None
                and pages_scraped >= MAX_PAGES
            ):
                break

            logger.info(
                "Scraping page %s: %s",
                pages_scraped + 1,
                current_url
            )

            soup = fetch_page(
                session,
                current_url
            )

            if soup is None:
                logger.warning(
                    "Stopping because a page could not be fetched."
                )
                break

            page_books = extract_books(
                soup,
                current_url
            )

            if not page_books:
                logger.warning(
                    "No books found on the current page."
                )
                break

            all_books.extend(page_books)

            pages_scraped += 1

            logger.info(
                "Books collected from this page: %s",
                len(page_books)
            )

            current_url = get_next_page(
                soup,
                current_url
            )

            if current_url:
                time.sleep(REQUEST_DELAY)

    finally:
        session.close()

    logger.info(
        "Total pages scraped: %s",
        pages_scraped
    )

    logger.info(
        "Total records collected: %s",
        len(all_books)
    )

    return all_books



#  CLEAN DATASET


def clean_dataset(data):
    """
    Clean scraped records and remove duplicates.
    """

    df = pd.DataFrame(data)

    if df.empty:
        logger.warning("Dataset is empty.")
        return df

    # Remove extra spaces from text columns.
    text_columns = [
        "title",
        "price_original",
        "availability",
        "product_url",
        "image_url",
        "source_page"
    ]

    for column in text_columns:

        if column in df.columns:

            df[column] = df[column].apply(
                lambda value: (
                    value.strip()
                    if isinstance(value, str)
                    else value
                )
            )

    # Convert numeric columns.
    df["price_gbp"] = pd.to_numeric(
        df["price_gbp"],
        errors="coerce"
    )

    df["rating"] = pd.to_numeric(
        df["rating"],
        errors="coerce"
    )

    # Remove records without a title.
    df = df.dropna(
        subset=["title"]
    )

    # Remove duplicate books using product URL.
    df = df.drop_duplicates(
        subset=["product_url"],
        keep="first"
    )

    # Create a unique identifier.
    df.insert(
        0,
        "book_id",
        range(1, len(df) + 1)
    )

    # Normalize availability text.
    df["availability"] = (
        df["availability"]
        .fillna("Unknown")
        .str.replace(r"\s+", " ", regex=True)
    )

    # Sort by price.
    df = df.sort_values(
        by="price_gbp",
        ascending=False,
        na_position="last"
    )

    df = df.reset_index(drop=True)

    # Reassign IDs after sorting.
    df["book_id"] = range(1, len(df) + 1)

    return df



#  ANALYZE DATASET


def analyze_dataset(df):
    """
    Generate useful summary statistics.
    """

    if df.empty:
        return "No data available for analysis."

    total_books = len(df)

    average_price = df["price_gbp"].mean()

    minimum_price = df["price_gbp"].min()

    maximum_price = df["price_gbp"].max()

    average_rating = df["rating"].mean()

    rating_counts = (
        df["rating"]
        .value_counts()
        .sort_index()
    )

    most_expensive_book = df.loc[
        df["price_gbp"].idxmax()
    ] if df["price_gbp"].notna().any() else None

    cheapest_book = df.loc[
        df["price_gbp"].idxmin()
    ] if df["price_gbp"].notna().any() else None

    available_books = df[
        df["availability"].str.contains(
            "In stock",
            case=False,
            na=False
        )
    ]

    report = []

    report.append("=" * 60)
    report.append("WEB SCRAPING PROJECT - SUMMARY REPORT")
    report.append("=" * 60)

    report.append(
        f"Total books collected: {total_books}"
    )

    report.append(
        f"Average price: £{average_price:.2f}"
        if pd.notna(average_price)
        else "Average price: Not available"
    )

    report.append(
        f"Minimum price: £{minimum_price:.2f}"
        if pd.notna(minimum_price)
        else "Minimum price: Not available"
    )

    report.append(
        f"Maximum price: £{maximum_price:.2f}"
        if pd.notna(maximum_price)
        else "Maximum price: Not available"
    )

    report.append(
        f"Average rating: {average_rating:.2f}/5"
        if pd.notna(average_rating)
        else "Average rating: Not available"
    )

    report.append(
        f"Books marked in stock: {len(available_books)}"
    )

    report.append("\nRating distribution:")

    for rating, count in rating_counts.items():
        report.append(
            f"Rating {int(rating)}: {count} books"
        )

    if most_expensive_book is not None:

        report.append(
            "\nMost expensive book: "
            f"{most_expensive_book['title']} "
            f"(£{most_expensive_book['price_gbp']:.2f})"
        )

    if cheapest_book is not None:

        report.append(
            "Cheapest book: "
            f"{cheapest_book['title']} "
            f"(£{cheapest_book['price_gbp']:.2f})"
        )

    report.append("=" * 60)

    return "\n".join(report)



#  SAVE DATASET


def save_dataset(df):
    """
    Export the cleaned dataset to CSV and Excel.
    """

    if df.empty:
        logger.error(
            "No records available. Files were not created."
        )
        return False

    # Save CSV.
    df.to_csv(
        CSV_FILE,
        index=False,
        encoding="utf-8-sig"
    )

    # Save Excel.
    df.to_excel(
        EXCEL_FILE,
        index=False,
        engine="openpyxl"
    )

    logger.info(
        "CSV saved to: %s",
        CSV_FILE.resolve()
    )

    logger.info(
        "Excel saved to: %s",
        EXCEL_FILE.resolve()
    )

    return True



#  SAVE SUMMARY REPORT


def save_summary(report):
    """
    Save analysis results to a text file.
    """

    with open(
        SUMMARY_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(report)

    logger.info(
        "Summary saved to: %s",
        SUMMARY_FILE.resolve()
    )



#  MAIN FUNCTION

def main():

    print("\n" + "=" * 60)
    print("       PYTHON WEB SCRAPING PROJECT")
    print("=" * 60)

    # Step 1: Scrape website.
    raw_data = scrape_books()

    if not raw_data:

        print(
            "\nNo data was collected. "
            "Check your internet connection and try again."
        )

        return

    # Step 2: Clean data.
    df = clean_dataset(raw_data)

    if df.empty:

        print("\nThe cleaned dataset is empty.")
        return

    # Step 3: Export dataset.
    saved = save_dataset(df)

    # Step 4: Analyze data.
    report = analyze_dataset(df)

    print("\n")
    print(report)

    # Step 5: Save summary.
    save_summary(report)

    # Step 6: Show sample records.
    print("\nFirst 5 records:")
    print(
        df.head(5).to_string(index=False)
    )

    if saved:

        print("\n" + "=" * 60)
        print("PROJECT COMPLETED SUCCESSFULLY!")
        print("=" * 60)

        print(f"\nCSV file: {CSV_FILE.resolve()}")
        print(f"Excel file: {EXCEL_FILE.resolve()}")
        print(f"Report file: {SUMMARY_FILE.resolve()}")



#  PROGRAM ENTRY POINT


if __name__ == "__main__":
    main()