import requests
import pandas as pd
import logging
import time
import random
from datetime import datetime
from typing import List, Dict, Optional, Set
import sys
import os
from bs4 import BeautifulSoup

# Configure logging for GitHub Actions
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

class TransfermarktScraper:
    def __init__(self):
        self.players_data = []
        self.session = requests.Session()
        self.seen_players: Set[str] = set()  # Track seen players to avoid duplicates
        
        # Set up output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.output_filename = f"transfermarkt_u28_players_{timestamp}.xlsx"
        
        # Predefined search URLs
        self.search_urls = [
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781586",  # U28 players 200-30M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781634",  # U28 players 30-20M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781645",  # U28 players 20-15M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781662",  # U28 players 15-12M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781673",  # U28 players 12-10M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781689",  # U28 players 10-8M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781698",  # U28 players 8-7M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781717",  # U28 players 7-6M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781836",  # U28 players 6-5.1M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782210", # U24 players 5-4M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782216", # U24 players 4-3.5M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782224", # U24 players 3.5-3M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782248", # U24 players 3-2.5M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782272", # U24 players 2.5-2.2M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782279", # U24 players 2.2-2.1M
        ]
        
        # More realistic User-Agent
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Cache-Control': 'max-age=0',
            'Referer': 'https://www.transfermarkt.com/'
        }
        
        self.session.headers.update(self.headers)
        self.last_request_time = 0

    def wait_between_requests(self):
        """Wait a random time between requests - reduced for GitHub Actions"""
        time_since_last = time.time() - self.last_request_time
        # Reduced wait time for CI environment
        min_wait = 15 if os.getenv('GITHUB_ACTIONS') else 30
        if time_since_last < min_wait:
            time.sleep(min_wait - time_since_last)
        self.last_request_time = time.time()

    def make_request(self, url: str, max_retries: int = 3) -> Optional[requests.Response]:
        """Make an HTTP request with wait time between requests"""
        retry_count = 0
        while retry_count < max_retries:
            try:
                self.wait_between_requests()
                response = self.session.get(url, timeout=30)
                
                if response.status_code == 200:
                    return response
                elif response.status_code == 503:
                    wait_time = 120 * (retry_count + 1)  # Reduced wait time for CI
                    logging.warning(f"Server busy (503). Waiting {wait_time/60} minutes...")
                    time.sleep(wait_time)
                elif response.status_code == 429:
                    wait_time = 300  # 5 minutes for rate limiting
                    logging.warning(f"Rate limited (429). Waiting {wait_time/60} minutes...")
                    time.sleep(wait_time)
                else:
                    logging.error(f"HTTP Error {response.status_code}")
                    time.sleep(60 * (retry_count + 1))
                
                retry_count += 1
                    
            except Exception as e:
                logging.error(f"Request error: {str(e)}")
                time.sleep(60 * (retry_count + 1))
                retry_count += 1
                
        return None

    def extract_instagram(self, profile_url: str) -> str:
        """Extract Instagram link from player profile"""
        try:
            response = self.make_request(profile_url)
            if not response:
                return ''
                
            soup = BeautifulSoup(response.text, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if 'instagram.com' in href:
                    return href
            return ''
            
        except Exception as e:
            logging.error(f"Error extracting Instagram: {str(e)}")
            return ''

    def save_data_to_file(self):
        """Save current data to file (overwrites existing file)"""
        if not self.players_data:
            return
            
        try:
            df = pd.DataFrame(self.players_data)
            # Sort by value descending to see highest valued players first
            df = df.sort_values('Value (M€)', ascending=False)
            
            with pd.ExcelWriter(self.output_filename, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Players')
                
                worksheet = writer.sheets['Players']
                worksheet.column_dimensions['A'].width = 30  # Name
                worksheet.column_dimensions['B'].width = 15  # Value
                worksheet.column_dimensions['C'].width = 50  # Transfermarkt
                worksheet.column_dimensions['D'].width = 50  # Instagram
                
                # Add hyperlinks for Excel format
                for idx, row in enumerate(df.itertuples(), start=2):
                    if row.Transfermarkt:
                        worksheet.cell(row=idx, column=3).hyperlink = row.Transfermarkt
                        worksheet.cell(row=idx, column=3).style = "Hyperlink"
                    if row.Instagram:
                        worksheet.cell(row=idx, column=4).hyperlink = row.Instagram
                        worksheet.cell(row=idx, column=4).style = "Hyperlink"
            
            logging.info(f"Data saved to {self.output_filename} - Total players: {len(self.players_data)}")
            
        except Exception as e:
            logging.error(f"Error saving data to file: {str(e)}")

    def get_players_from_page(self, base_url: str, page: int) -> List[Dict]:
        """Get players from a specific page of a search URL"""
        url = f"{base_url}?page={page}"
        
        logging.info(f"Getting players from page {page} of search")
        
        response = self.make_request(url)
        if not response:
            return []
        
        soup = BeautifulSoup(response.text, 'html.parser')
        players = []
        
        # Look for the responsive table structure
        for row in soup.select('table.items tbody tr'):
            try:
                # Find name and profile link
                name_cell = row.select_one('td.hauptlink a')
                if not name_cell:
                    continue
                    
                name = name_cell.text.strip()
                profile_url = f"https://www.transfermarkt.com{name_cell['href']}"
                
                # Find market value - look for the cell with currency
                value_cell = None
                for cell in row.select('td'):
                    if '€' in cell.text and 'm' in cell.text.lower():
                        value_cell = cell
                        break
                
                if not value_cell:
                    continue
                    
                value_text = value_cell.text.strip().replace('€', '').replace('m', '').strip()
                try:
                    value = float(value_text)
                except:
                    continue
                
                # Create unique identifier for player
                player_id = f"{name}|{profile_url}"
                
                # Skip if we've already seen this player
                if player_id in self.seen_players:
                    logging.info(f"Skipping duplicate: {name}")
                    continue
                
                # Add to seen players
                self.seen_players.add(player_id)
                
                # Get Instagram
                instagram = self.extract_instagram(profile_url)
                
                # Add player data
                player_data = {
                    'Name': name,
                    'Value (M€)': value,
                    'Transfermarkt': profile_url,
                    'Instagram': instagram
                }
                players.append(player_data)
                
                # Add to main data and save periodically (not after every player in CI)
                self.players_data.append(player_data)
                
                # Save every 10 players in CI environment, every player otherwise
                if os.getenv('GITHUB_ACTIONS'):
                    if len(self.players_data) % 10 == 0:
                        self.save_data_to_file()
                else:
                    self.save_data_to_file()
                
                logging.info(f"Added: {name} - {value}M€ - Instagram: {instagram} - Total players: {len(self.players_data)}")
                
                # Add GitHub Actions timeout protection
                if os.getenv('GITHUB_ACTIONS') and len(self.players_data) > 500:
                    logging.warning("GitHub Actions timeout protection: Stopping at 500 players")
                    return players
                
            except Exception as e:
                logging.error(f"Error processing player: {str(e)}")
                continue
        
        return players

    def scrape_search_url(self, base_url: str, search_index: int) -> int:
        """Scrape all pages from a single search URL"""
        logging.info(f"\n=== SCRAPING SEARCH {search_index + 1}/{len(self.search_urls)} ===")
        logging.info(f"URL: {base_url}")
        
        page = 1
        players_found = 0
        
        max_pages = 5 if os.getenv('GITHUB_ACTIONS') else 10  # Reduced for CI
        
        while page <= max_pages:
            players = self.get_players_from_page(base_url, page)
            
            if not players:
                logging.info(f"No players found on page {page}. Moving to next search.")
                break
            
            players_found += len(players)
            logging.info(f"Page {page} completed. Found {len(players)} players. Total from this search: {players_found}")
            
            page += 1
            
            # Reduced wait time for CI
            wait_time = random.uniform(15, 25) if os.getenv('GITHUB_ACTIONS') else random.uniform(30, 45)
            time.sleep(wait_time)
            
            # GitHub Actions timeout protection
            if os.getenv('GITHUB_ACTIONS') and len(self.players_data) > 500:
                logging.warning("GitHub Actions timeout protection: Stopping search")
                break
        
        logging.info(f"Search {search_index + 1} completed. Total players found: {players_found}")
        return players_found

    def scrape(self):
        """Main scraping method that processes all search URLs"""
        logging.info(f"Starting scraping of all U28 player searches")
        logging.info(f"Output file: {self.output_filename}")
        logging.info(f"Total searches to process: {len(self.search_urls)}")
        
        if os.getenv('GITHUB_ACTIONS'):
            logging.info("Running in GitHub Actions environment")
        
        total_players_found = 0
        
        try:
            for i, search_url in enumerate(self.search_urls):
                players_from_search = self.scrape_search_url(search_url, i)
                total_players_found += players_from_search
                
                # GitHub Actions timeout protection
                if os.getenv('GITHUB_ACTIONS') and len(self.players_data) > 500:
                    logging.warning("GitHub Actions timeout protection: Stopping all searches")
                    break
                
                # Wait between different searches (except after the last one)
                if i < len(self.search_urls) - 1:
                    wait_time = 20 if os.getenv('GITHUB_ACTIONS') else 30
                    logging.info(f"Waiting {wait_time} seconds before next search...")
                    time.sleep(wait_time)
                
        except Exception as e:
            logging.error(f"General scraping error: {str(e)}")
        finally:
            if self.players_data:
                # Final save
                self.save_data_to_file()
                
                # Create final summary
                df = pd.DataFrame(self.players_data)
                df = df.sort_values('Value (M€)', ascending=False)
                
                logging.info(f"\n=== FINAL SUMMARY ===")
                logging.info(f"Total players scraped: {len(self.players_data)}")
                logging.info(f"Value range: {df['Value (M€)'].min():.1f}M€ - {df['Value (M€)'].max():.1f}M€")
                logging.info(f"Players with Instagram: {len(df[df['Instagram'] != ''])}")
                logging.info(f"Final output saved to: {self.output_filename}")
                logging.info(f"Top 10 most valuable players:")
                for idx, player in df.head(10).iterrows():
                    logging.info(f"  {player['Name']}: {player['Value (M€)']}M€")

def main():
    scraper = TransfermarktScraper()
    
    try:
        scraper.scrape()
    except Exception as e:
        logging.error(f"Error in main script: {str(e)}")
        sys.exit(1)  # Exit with error code for GitHub Actions

if __name__ == "__main__":
    main()
