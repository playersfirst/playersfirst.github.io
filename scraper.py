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
import json
import argparse

# Configure logging for GitHub Actions
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

class MatrixTurboScraper:
    def __init__(self, worker_id: int = 0, total_workers: int = 1):
        self.worker_id = worker_id
        self.total_workers = total_workers
        self.players_data = []
        self.session = requests.Session()
        self.seen_players: Set[str] = set()
        
        # Set up output filename with worker ID
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.output_filename = f"transfermarkt_worker_{worker_id}_{timestamp}.xlsx"
        self.json_output = f"transfermarkt_worker_{worker_id}_{timestamp}.json"
        
        # ALL search URLs - will be distributed among workers
        self.all_search_urls = [
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781586",  # U28 players 200-30M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781634",  # U28 players 30-20M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781645",  # U28 players 20-15M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781662",  # U28 players 15-12M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781673",  # U28 players 12-10M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781689",  # U28 players 10-8M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781698",  # U28 players 8-7M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781717",  # U28 players 7-6M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55781836",  # U28 players 6-5.1M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782210",  # U24 players 5-4M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782216",  # U24 players 4-3.5M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782224",  # U24 players 3.5-3M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782248",  # U24 players 3-2.5M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782272",  # U24 players 2.5-2.2M
            "https://www.transfermarkt.com/detailsuche/spielerdetail/suche/55782279",  # U24 players 2.2-2.1M
        ]
        
        # Distribute URLs among workers
        self.my_search_urls = self.distribute_urls()
        
        # Different user agents for each worker
        self.user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/120.0.0.0',
            'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/121.0',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15'
        ]
        
        # Worker-specific headers
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
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': self.user_agents[worker_id % len(self.user_agents)]  # Worker-specific UA
        }

    def distribute_urls(self) -> List[str]:
        """Distribute URLs among workers - OPTIMIZED for 15 workers = 15 URLs"""
        if self.total_workers == 15 and len(self.all_search_urls) == 15:
            # Perfect 1:1 ratio - each worker gets exactly 1 URL
            logging.info(f"Worker {self.worker_id}: PERFECT DISTRIBUTION - Assigned URL index {self.worker_id}")
            return [self.all_search_urls[self.worker_id]]
        else:
            # Fallback to original distribution logic
            urls_per_worker = len(self.all_search_urls) // self.total_workers
            remainder = len(self.all_search_urls) % self.total_workers
            
            start_idx = self.worker_id * urls_per_worker
            end_idx = start_idx + urls_per_worker
            
            if self.worker_id < remainder:
                start_idx += self.worker_id
                end_idx += self.worker_id + 1
            else:
                start_idx += remainder
                end_idx += remainder
            
            my_urls = self.all_search_urls[start_idx:end_idx]
            logging.info(f"Worker {self.worker_id}: Assigned {len(my_urls)} URLs (indices {start_idx}-{end_idx-1})")
            return my_urls

    def get_worker_headers(self):
        """Get worker-specific headers with random variations"""
        headers = self.base_headers.copy()
        
        # Add some randomness but keep worker identity
        if random.random() > 0.5:
            headers['Accept-Language'] = random.choice(['en-US,en;q=0.9', 'en-GB,en;q=0.9', 'en;q=0.9'])
        
        # Worker-specific delays
        worker_delay = 0.5 + (self.worker_id * 0.3)  # Stagger workers
        return headers, worker_delay

    def make_smart_request(self, url: str, max_retries: int = 3) -> Optional[requests.Response]:
        """Smart request with worker-specific timing"""
        headers, base_delay = self.get_worker_headers()
        
        for attempt in range(max_retries):
            try:
                # Worker-specific delay to avoid collision
                delay = base_delay + random.uniform(0.5, 2.0)
                time.sleep(delay)
                
                response = requests.get(url, headers=headers, timeout=20)
                
                if response.status_code == 200:
                    return response
                elif response.status_code == 429:
                    # Rate limited - exponential backoff per worker
                    wait_time = (2 ** attempt) * (1 + self.worker_id * 0.5)
                    logging.warning(f"Worker {self.worker_id}: Rate limited, waiting {wait_time:.1f}s")
                    time.sleep(wait_time)
                elif response.status_code == 503:
                    # Server busy - longer wait with worker offset
                    wait_time = 30 + (self.worker_id * 10)
                    logging.warning(f"Worker {self.worker_id}: Server busy, waiting {wait_time}s")
                    time.sleep(wait_time)
                else:
                    logging.error(f"Worker {self.worker_id}: HTTP Error {response.status_code}")
                    time.sleep(5 + self.worker_id)
                    
            except Exception as e:
                logging.error(f"Worker {self.worker_id}: Request error: {str(e)}")
                time.sleep(5 + self.worker_id)
                
        return None

    def extract_instagram_fast(self, profile_url: str) -> str:
        """Fast Instagram extraction"""
        try:
            response = self.make_smart_request(profile_url)
            if not response:
                return ''
                
            # Quick search in raw HTML
            html_text = response.text.lower()
            if 'instagram.com' not in html_text:
                return ''
            
            soup = BeautifulSoup(response.text, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if 'instagram.com' in href:
                    return href
            return ''
            
        except Exception as e:
            logging.error(f"Worker {self.worker_id}: Error extracting Instagram: {str(e)}")
            return ''

    def process_player_row(self, row) -> Optional[Dict]:
        """Process a single player row"""
        try:
            name_cell = row.select_one('td.hauptlink a')
            if not name_cell:
                return None
                
            name = name_cell.text.strip()
            profile_url = f"https://www.transfermarkt.com{name_cell['href']}"
            
            # Find market value
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
            
            player_id = f"{name}|{profile_url}"
            
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
            logging.error(f"Worker {self.worker_id}: Error processing player row: {str(e)}")
            return None

    def get_players_from_page(self, base_url: str, page: int) -> List[Dict]:
        """Get players from a single page"""
        url = f"{base_url}?page={page}"
        
        response = self.make_smart_request(url)
        if not response:
            return []
        
        soup = BeautifulSoup(response.text, 'html.parser')
        rows = soup.select('table.items tbody tr')
        
        players = []
        for row in rows:
            player_data = self.process_player_row(row)
            if player_data:
                players.append(player_data)
        
        return players

    def scrape_search_sequential(self, base_url: str, search_index: int) -> int:
        """Scrape a single search URL sequentially (no threading within worker)"""
        logging.info(f"Worker {self.worker_id}: Scraping search {search_index + 1}")
        
        all_players_from_search = []
        max_pages = 10  # Limit pages for speed
        
        for page in range(1, max_pages + 1):
            page_players = self.get_players_from_page(base_url, page)
            all_players_from_search.extend(page_players)
            logging.info(f"Worker {self.worker_id}: Page {page} -> {len(page_players)} players")
            
            if len(page_players) == 0:
                break  # No more players
        
        # Process Instagram for all players
        final_players = []
        for i, player in enumerate(all_players_from_search):
            instagram = self.extract_instagram_fast(player['profile_url'])
            final_player = {
                'Name': player['name'],
                'Value (M€)': player['value'],
                'Transfermarkt': player['profile_url'],
                'Instagram': instagram
            }
            final_players.append(final_player)
            self.players_data.append(final_player)
            
            if (i + 1) % 5 == 0:
                logging.info(f"Worker {self.worker_id}: Processed {i + 1}/{len(all_players_from_search)} players")
        
        self.save_data_to_files()
        logging.info(f"Worker {self.worker_id}: Search {search_index + 1} completed: {len(final_players)} players")
        return len(final_players)

    def save_data_to_files(self):
        """Save data to both Excel and JSON"""
        if not self.players_data:
            return
            
        try:
            df = pd.DataFrame(self.players_data)
            df = df.sort_values('Value (M€)', ascending=False)
            
            # Save Excel
            with pd.ExcelWriter(self.output_filename, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Players')
                
                worksheet = writer.sheets['Players']
                worksheet.column_dimensions['A'].width = 30
                worksheet.column_dimensions['B'].width = 15
                worksheet.column_dimensions['C'].width = 50
                worksheet.column_dimensions['D'].width = 50
            
            # Save JSON for merging
            with open(self.json_output, 'w', encoding='utf-8') as f:
                json.dump(self.players_data, f, ensure_ascii=False, indent=2)
            
            logging.info(f"Worker {self.worker_id}: Data saved - {len(self.players_data)} players")
            
        except Exception as e:
            logging.error(f"Worker {self.worker_id}: Error saving data: {str(e)}")

    def scrape_matrix_worker(self):
        """Main scraping method for matrix worker"""
        logging.info(f"🚀 Worker {self.worker_id}/{self.total_workers} STARTING!")
        logging.info(f"📋 Assigned URLs: {len(self.my_search_urls)}")
        
        start_time = time.time()
        total_players = 0
        
        try:
            for i, url in enumerate(self.my_search_urls):
                players_found = self.scrape_search_sequential(url, i)
                total_players += players_found
                logging.info(f"✅ Worker {self.worker_id}: Search {i + 1}/{len(self.my_search_urls)} DONE: {players_found} players")
                
        except Exception as e:
            logging.error(f"Worker {self.worker_id}: General error: {str(e)}")
        finally:
            elapsed_time = time.time() - start_time
            
            if self.players_data:
                self.save_data_to_files()
                
                logging.info(f"\n🎉 Worker {self.worker_id} COMPLETE! 🎉")
                logging.info(f"⏱️  Time: {elapsed_time/60:.1f} minutes")
                logging.info(f"👥 Players: {len(self.players_data)}")
                logging.info(f"📸 Instagram: {len([p for p in self.players_data if p['Instagram']])}")

def main():
    parser = argparse.ArgumentParser(description='Matrix Turbo Scraper')
    parser.add_argument('--worker-id', type=int, default=0, help='Worker ID (0-based)')
    parser.add_argument('--total-workers', type=int, default=1, help='Total number of workers')
    
    args = parser.parse_args()
    
    scraper = MatrixTurboScraper(args.worker_id, args.total_workers)
    
    try:
        scraper.scrape_matrix_worker()
    except Exception as e:
        logging.error(f"Error in worker {args.worker_id}: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
