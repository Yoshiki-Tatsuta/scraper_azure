from deep_translator import GoogleTranslator
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import requests

class AzureUpdatesScraper:
    def __init__(self, get_limit):
        self.get_url = 'https://azure.microsoft.com/ja-jp/updates/'
        self.get_limit = get_limit
        self.begin_url = 'https://azure.microsoft.com'
        self.translator = GoogleTranslator(source='auto', target='ja')
        self.full_urls = []

    def scrape_url(self):
        try:
            response = requests.get(self.get_url)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            urls = []
            for link in soup.find_all('a', {'data-test-element': 'update-entry-link'}):
                if len(urls) < self.get_limit:
                    urls.append(link.get('href'))
                else:
                    break
            urls.reverse()
            self.full_urls = [self.begin_url+url for url in urls]
            print(self.full_urls)
        except requests.exceptions.RequestException as e:
            print(f"要求失敗: {e}")
        except Exception as e:
            print(f"エラーが発生しました。: {e}")

    def excel_write(self):
        try:
            excel_file_path = 'test.xlsx'
            wb = load_workbook(excel_file_path)
            ws = wb.active
            for start_cell in range(1, ws.max_row+2):
                if ws[f'A{start_cell}'].value is None:
                    print(f'{start_cell}行目から書き込みます')
                    for i, url in enumerate(self.full_urls, start_cell):
                        response = requests.get(url)
                        soup = BeautifulSoup(response.text, 'html.parser')
                        # 題名取得
                        h1_tag = soup.find('h1')
                        # 日付取得
                        date_div_tag = soup.find('div', class_='row row-size3 column')
                        date_h6_tag = date_div_tag.find('h6')
                        date = date_h6_tag.text.split(':')[-1].strip()
                        dt = datetime.strptime(date, '%m月 %d, %Y')
                        date_tag = dt.strftime('%Y/%m/%d')
                        # タグの取得
                        tags_ul = soup.find('ul', class_='tags')
                        tags_li = tags_ul.find_all('li')
                        tags = [li.text for li in tags_li]
                        tags_str = ','.join(tags)
                        # 本文の取得
                        main_column_tags = soup.find('div', class_='column small-12')
                        main_tags = main_column_tags.find('div', class_='row column')
                        main_tag = main_tags.text.strip()
                        # 本文の日本語翻訳
                        main_tag_ja = self.translator.translate(main_tag)
                        # セルに書き込み
                        if h1_tag and date_tag and tags_str and main_tag and main_tag_ja:
                            ws[f'A{i}'] = h1_tag.text
                            ws[f'B{i}'] = date_tag
                            ws[f'C{i}'] = tags_str
                            ws[f'D{i}'] = main_tag
                            ws[f'E{i}'] = main_tag_ja
                        else:
                            print('なにかの値がとれていません。')
                    break
            wb.save(excel_file_path)
            wb.close()
            return print('取得した値の書き込み完了')
        except Exception as e:
            print(f"エラーが発生しました。: {e}")

# 実行
get_limit_input = int(input('取得したい数を入力してください。：  '))
scraper_obj = AzureUpdatesScraper(get_limit_input)
scraper_obj.scrape_url()
scraper_obj.excel_write()