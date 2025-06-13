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
import asyncio
import aiohttp
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# Configure logging for GitHub Actions
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

class TurboTransfermarktScraper:
    def __init__(self):
        self.players_data = []
        self.session = requests.Session()
        self.seen_players: Set[str] = set()
        self.data_lock = threading.Lock()  # Thread safety
        
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
        
        # AGGRESSIVE headers with multiple user agents
        self.user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/120.0.0.0'
        ]
        
        self.base_headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Cache-Control': 'max-age=0',
            'Referer': 'https://www.transfermarkt.com/',
            'DNT': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin'
        }

    def get_random_headers(self):
        """Get random headers to avoid detection"""
        headers = self.base_headers.copy()
        headers['User-Agent'] = random.choice(self.user_agents)
        return headers

    def make_aggressive_request(self, url: str, max_retries: int = 2) -> Optional[requests.Response]:
        """SUPER AGGRESSIVE request with minimal delays"""
        for attempt in range(max_retries):
            try:
                # Minimal delay - only 1-3 seconds!
                time.sleep(random.uniform(1, 3))
                
                headers = self.get_random_headers()
                response = requests.get(url, headers=headers, timeout=15)
                
                if response.status_code == 200:
                    return response
                elif response.status_code == 429:
                    # Rate limited - wait just 30 seconds instead of minutes
                    logging.warning(f"Rate limited. Quick wait...")
                    time.sleep(30)
                elif response.status_code == 503:
                    # Server busy - wait just 60 seconds
                    logging.warning(f"Server busy. Quick wait...")
                    time.sleep(60)
                else:
                    logging.error(f"HTTP Error {response.status_code}")
                    time.sleep(5)  # Very short wait
                    
            except Exception as e:
                logging.error(f"Request error: {str(e)}")
                time.sleep(5)  # Very short wait
                
        return None

    def extract_instagram_fast(self, profile_url: str) -> str:
        """FAST Instagram extraction - bail out quickly if not found"""
        try:
            response = self.make_aggressive_request(profile_url)
            if not response:
                return ''
                
            # Quick search in raw HTML instead of full parsing
            html_text = response.text.lower()
            if 'instagram.com' not in html_text:
                return ''
            
            # Only parse if Instagram is mentioned
            soup = BeautifulSoup(response.text, 'html.parser')
            for a in soup.find_all('a', href=True, limit=50):  # Limit search
                href = a['href']
                if 'instagram.com' in href:
                    return href
            return ''
            
        except Exception as e:
            logging.error(f"Error extracting Instagram: {str(e)}")
            return ''

    def process_player_row(self, row) -> Optional[Dict]:
        """Process a single player row - optimized"""
        try:
            # Find name and profile link
            name_cell = row.select_one('td.hauptlink a')
            if not name_cell:
                return None
                
            name = name_cell.text.strip()
            profile_url = f"https://www.transfermarkt.com{name_cell['href']}"
            
            # Find market value - look for the cell with currency
            value_cell = None
            for cell in row.select('td'):
                cell_text = cell.text
                if '€' in cell_text and 'm' in cell_text.lower():
                    value_cell = cell
                    break
            
            if not value_cell:
                return None
                
            value_text = value_cell.text.strip().replace('€', '').replace('m', '').strip()
            try:
                value = float(value_text)
            except:
                return None
            
            # Create unique identifier for player
            player_id = f"{name}|{profile_url}"
            
            # Thread-safe duplicate check
            with self.data_lock:
                if player_id in self.seen_players:
                    return None
                self.seen_players.add(player_id)
            
            return {
                'name': name,
                'value': value,
                'profile_url': profile_url,
                'player_id': player_id
            }
            
        except Exception as e:
            logging.error(f"Error processing player row: {str(e)}")
            return None

    def get_players_from_page_turbo(self, base_url: str, page: int) -> List[Dict]:
        """TURBO speed page processing"""
        url = f"{base_url}?page={page}"
        
        response = self.make_aggressive_request(url)
        if not response:
            return []
        
        soup = BeautifulSoup(response.text, 'html.parser')
        rows = soup.select('table.items tbody tr')
        
        # Process rows in parallel using ThreadPoolExecutor
        players = []
        with ThreadPoolExecutor(max_workers=10) as executor:  # AGGRESSIVE: 10 threads
            future_to_row = {executor.submit(self.process_player_row, row): row for row in rows}
            
            for future in as_completed(future_to_row):
                try:
                    player_data = future.result()
                    if player_data:
                        players.append(player_data)
                except Exception as e:
                    logging.error(f"Error in thread: {str(e)}")
        
        return players

    def process_players_instagram(self, players: List[Dict]) -> List[Dict]:
        """Process Instagram links in parallel"""
        def get_instagram_for_player(player_data):
            instagram = self.extract_instagram_fast(player_data['profile_url'])
            return {
                'Name': player_data['name'],
                'Value (M€)': player_data['value'],
                'Transfermarkt': player_data['profile_url'],
                'Instagram': instagram
            }
        
        # SUPER AGGRESSIVE: Process Instagram links in parallel
        final_players = []
        with ThreadPoolExecutor(max_workers=20) as executor:  # 20 threads for Instagram!
            future_to_player = {executor.submit(get_instagram_for_player, player): player for player in players}
            
            for future in as_completed(future_to_player):
                try:
                    final_player = future.result()
                    final_players.append(final_player)
                    
                    # Thread-safe data addition
                    with self.data_lock:
                        self.players_data.append(final_player)
                    
                    if len(final_players) % 10 == 0:
                        logging.info(f"Processed {len(final_players)} players with Instagram...")
                        
                except Exception as e:
                    logging.error(f"Error processing Instagram: {str(e)}")
        
        return final_players

    def save_data_to_file(self):
        """Thread-safe save"""
        with self.data_lock:
            if not self.players_data:
                return
                
            try:
                df = pd.DataFrame(self.players_data)
                df = df.sort_values('Value (M€)', ascending=False)
                
                with pd.ExcelWriter(self.output_filename, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Players')
                    
                    worksheet = writer.sheets['Players']
                    worksheet.column_dimensions['A'].width = 30
                    worksheet.column_dimensions['B'].width = 15
                    worksheet.column_dimensions['C'].width = 50
                    worksheet.column_dimensions['D'].width = 50
                
                logging.info(f"Data saved to {self.output_filename} - Total players: {len(self.players_data)}")
                
            except Exception as e:
                logging.error(f"Error saving data to file: {str(e)}")

    def scrape_search_turbo(self, base_url: str, search_index: int) -> int:
        """TURBO scrape a single search URL"""
        logging.info(f"\n=== TURBO SCRAPING SEARCH {search_index + 1}/{len(self.search_urls)} ===")
        
        all_players_from_search = []
        max_pages = 8  # Reasonable limit
        
        # Get all pages from this search in parallel!
        with ThreadPoolExecutor(max_workers=5) as executor:  # 5 pages at once!
            future_to_page = {
                executor.submit(self.get_players_from_page_turbo, base_url, page): page 
                for page in range(1, max_pages + 1)
            }
            
            for future in as_completed(future_to_page):
                try:
                    page_players = future.result()
                    all_players_from_search.extend(page_players)
                    page_num = future_to_page[future]
                    logging.info(f"Page {page_num} completed: {len(page_players)} players")
                except Exception as e:
                    logging.error(f"Error processing page: {str(e)}")
        
        # Now process Instagram for all players from this search
        if all_players_from_search:
            self.process_players_instagram(all_players_from_search)
            self.save_data_to_file()
        
        logging.info(f"Search {search_index + 1} completed: {len(all_players_from_search)} players")
        return len(all_players_from_search)

    def scrape_turbo(self):
        """MAIN TURBO SCRAPING METHOD"""
        logging.info("🚀 STARTING TURBO SCRAPING MODE! 🚀")
        logging.info(f"Output file: {self.output_filename}")
        
        start_time = time.time()
        
        try:
            # Process multiple searches in parallel!
            with ThreadPoolExecutor(max_workers=3) as executor:  # 3 searches at once!
                future_to_search = {
                    executor.submit(self.scrape_search_turbo, url, i): (url, i) 
                    for i, url in enumerate(self.search_urls[:6])  # First 6 searches only for speed
                }
                
                total_players = 0
                for future in as_completed(future_to_search):
                    try:
                        players_found = future.result()
                        total_players += players_found
                        url, index = future_to_search[future]
                        logging.info(f"✅ Search {index + 1} DONE: {players_found} players")
                    except Exception as e:
                        logging.error(f"❌ Search failed: {str(e)}")
            
        except Exception as e:
            logging.error(f"General scraping error: {str(e)}")
        finally:
            elapsed_time = time.time() - start_time
            
            if self.players_data:
                self.save_data_to_file()
                
                df = pd.DataFrame(self.players_data)
                df = df.sort_values('Value (M€)', ascending=False)
                
                logging.info(f"\n🎉 TURBO SCRAPING COMPLETE! 🎉")
                logging.info(f"⏱️  Total time: {elapsed_time/60:.1f} minutes")
                logging.info(f"👥 Total players: {len(self.players_data)}")
                logging.info(f"💰 Value range: {df['Value (M€)'].min():.1f}M€ - {df['Value (M€)'].max():.1f}M€")
                logging.info(f"📸 Instagram links: {len(df[df['Instagram'] != ''])}")
                logging.info(f"⚡ Players per minute: {len(self.players_data)/(elapsed_time/60):.1f}")

def main():
    scraper = TurboTransfermarktScraper()
    
    try:
        scraper.scrape_turbo()
    except Exception as e:
        logging.error(f"Error in main script: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
