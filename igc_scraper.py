"""
IGC Market at a Glance Web Scraper
Website: https://www.igc.int/en/default.aspx

Features:
- Robust Error Handling & HTTP Retries (Exponential Backoff)
- Scrapes daily commodity price graph data for Wheat, Maize, Barley, Soyabeans, Rice
- Exports to JSON, CSV, and multi-sheet Excel

Author: Antigravity AI
"""

import os
import sys
import json
import logging
import argparse
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import pandas as pd

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)

class IGCScraper:
    BASE_URL = "https://www.igc.int/en/default.aspx"
    
    TABS = [
        ("Wheat", "WheatPriceButton"),
        ("Maize", "MaizePriceButton"),
        ("Barley", "BarleyPriceButton"),
        ("Soyabeans", "SoyabeansPriceButton"),
        ("Rice", "RicePriceButton")
    ]

    def __init__(self, max_retries: int = 3, backoff_factor: float = 1.0):
        self.session = requests.Session()
        
        # Configure Retries for Network Resilience
        retries = Retry(
            total=max_retries,
            backoff_factor=backoff_factor,
            status_forcelist=[500, 502, 503, 504],
            raise_on_status=False
        )
        adapter = HTTPAdapter(max_retries=retries)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
        })

    def _extract_hidden_inputs(self, soup: BeautifulSoup) -> dict:
        """Extract hidden ASP.NET form inputs required for PostBack."""
        data = {}
        for tag in soup.find_all('input', type='hidden'):
            name = tag.get('name')
            val = tag.get('value', '')
            if name:
                data[name] = val
        return data

    def scrape_all(self) -> dict:
        """Scrape market data across all 5 commodity groups with Error Handling."""
        logging.info(f"Fetching main page: {self.BASE_URL}")
        try:
            res = self.session.get(self.BASE_URL, timeout=30)
            res.raise_for_status()
        except Exception as e:
            logging.error(f"Failed to fetch IGC main page: {e}")
            raise RuntimeError(f"Network error while connecting to IGC website: {e}")

        soup = BeautifulSoup(res.text, 'html.parser')
        result_data = {}

        for group_name, btn_id in self.TABS:
            logging.info(f"Scraping group: {group_name} (Triggering {btn_id})...")
            
            try:
                # Prepare ASP.NET postback parameters
                form_data = self._extract_hidden_inputs(soup)
                form_data['__EVENTTARGET'] = btn_id
                form_data['__EVENTARGUMENT'] = ''

                # Execute PostBack with Timeout and Error Handling
                res_post = self.session.post(self.BASE_URL, data=form_data, timeout=30)
                res_post.raise_for_status()

                soup = BeautifulSoup(res_post.text, 'html.parser')
                table = soup.find('table', id='GridViewHiddenPrices')

                if not table:
                    logging.warning(f"Could not find GridViewHiddenPrices table for group '{group_name}'")
                    continue

                headers = [th.get_text(strip=True) for th in table.find_all('th')]
                rows = []
                for tr in table.find_all('tr')[1:]:
                    cols = [td.get_text(strip=True) for td in tr.find_all('td')]
                    if cols:
                        rows.append(cols)

                sub_commodities = headers[1:]
                daily_records = []

                for row in rows:
                    date_str = row[0]
                    prices = {}
                    for idx, comm in enumerate(sub_commodities):
                        val_str = row[idx + 1] if idx + 1 < len(row) else None
                        try:
                            val_num = float(val_str) if val_str and val_str != '-' else None
                        except (ValueError, TypeError):
                            val_num = val_str
                        prices[comm] = val_num

                    daily_records.append({
                        "date": date_str,
                        "prices": prices
                    })

                result_data[group_name] = {
                    "sub_commodities": sub_commodities,
                    "daily_prices": daily_records
                }
                logging.info(f"Successfully scraped {len(daily_records)} daily records for {group_name} ({len(sub_commodities)} sub-commodities)")

            except Exception as e:
                logging.error(f"Error scraping group {group_name}: {e}")
                # Continue scraping other groups even if one group fails

        if not result_data:
            raise RuntimeError("Scraper completed but 0 commodity groups were scraped successfully.")

        return result_data

    @staticmethod
    def to_flat_dataframe(scraped_data: dict) -> pd.DataFrame:
        flat_rows = []
        for group, info in scraped_data.items():
            for record in info["daily_prices"]:
                date_val = record["date"]
                for sub_comm, price_val in record["prices"].items():
                    flat_rows.append({
                        "Group": group,
                        "SubCommodity": sub_comm,
                        "Date": date_val,
                        "Price_USD": price_val
                    })
        return pd.DataFrame(flat_rows)

    @staticmethod
    def save_to_json(scraped_data: dict, filepath: str = "igc_market_data.json"):
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(scraped_data, f, ensure_ascii=False, indent=2)
        logging.info(f"Saved JSON data to {filepath}")

    @staticmethod
    def save_to_csv(df: pd.DataFrame, filepath: str = "igc_market_data.csv"):
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
        logging.info(f"Saved CSV data to {filepath}")

    @staticmethod
    def save_to_excel(scraped_data: dict, df_flat: pd.DataFrame, filepath: str = "igc_market_data.xlsx"):
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df_flat.to_excel(writer, sheet_name="All_Data", index=False)
            for group, info in scraped_data.items():
                group_rows = []
                for record in info["daily_prices"]:
                    row = {"Date": record["date"]}
                    row.update(record["prices"])
                    group_rows.append(row)
                df_group = pd.DataFrame(group_rows)
                df_group.to_excel(writer, sheet_name=group, index=False)
        logging.info(f"Saved Excel data workbook to {filepath}")


def main():
    parser = argparse.ArgumentParser(description="Scrape IGC Market at a Glance commodity price graph data.")
    parser.add_argument("--json", default="igc_market_data.json", help="Path for JSON output")
    parser.add_argument("--csv", default="igc_market_data.csv", help="Path for CSV output")
    parser.add_argument("--excel", default="igc_market_data.xlsx", help="Path for Excel output")
    args = parser.parse_args()

    try:
        scraper = IGCScraper()
        data = scraper.scrape_all()

        scraper.save_to_json(data, args.json)
        df_flat = scraper.to_flat_dataframe(data)
        scraper.save_to_csv(df_flat, args.csv)
        scraper.save_to_excel(data, df_flat, args.excel)

        print("\nScraping Completed Successfully!")
    except Exception as e:
        logging.critical(f"Scraper program terminated with error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
