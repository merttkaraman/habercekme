import os
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(
    filename='habercekme_selenium_log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

unwanted_terms = [
    'masthead', 'ntv.com.tr', 'google-play', 'app-store', 'Huawei App Gallery',
    'logo', 'banner', 'sponsor', 'advertisement'
]

def is_unwanted(title, unwanted_terms):
    title_lower = title.lower()
    for term in unwanted_terms:
        if term.lower() in title_lower:
            return True
    return False

def get_news_links(img_elements):
    news = []
    seen_titles = set()

    for img in img_elements:
        alt_text = img.get_attribute('alt')
        title_text = img.get_attribute('title')

        if alt_text:
            alt_text = alt_text.strip()
            if not is_unwanted(alt_text, unwanted_terms) and alt_text not in seen_titles:
                try:
                    parent = img.find_element(By.XPATH, './ancestor::a')
                    href = parent.get_attribute('href')
                    if href:
                        news.append({'Başlık': alt_text, 'Link': href})
                        seen_titles.add(alt_text)
                except Exception as e:
                    logging.error(f"Link bulunamadı: {e}")

        if title_text:
            title_text = title_text.strip()
            if not is_unwanted(title_text, unwanted_terms) and title_text not in seen_titles:
                try:
                    parent = img.find_element(By.XPATH, './ancestor::a')
                    href = parent.get_attribute('href')
                    if href:
                        news.append({'Başlık': title_text, 'Link': href})
                        seen_titles.add(title_text)
                except Exception as e:
                    logging.error(f"Link bulunamadı: {e}")

    return news

def get_description(driver, link):
    try:
        driver.get(link)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, 'h2'))
        )
        h2_element = driver.find_element(By.TAG_NAME, 'h2')
        description = h2_element.text.strip()
        return description
    except Exception as e:
        logging.error(f"Linke giderken veya açıklamayı çekerken bir hata oluştu: {e}")
        return ""

service = Service(ChromeDriverManager().install())

options = webdriver.ChromeOptions()
options.add_argument('--headless')
driver = webdriver.Chrome(service=service, options=options)

try:
    driver.get('https://www.ntv.com.tr/')

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, 'img'))
    )

    img_elements = driver.find_elements(By.TAG_NAME, 'img')

    news = get_news_links(img_elements)

    if not news:
        logging.info(
            "Haber başlıkları bulunamadı veya tüm başlıklar istenmeyen terimler içeriyor.")
    else:
        valid_news = []

        for entry in news:
            title = entry['Başlık']
            link = entry['Link']
            description = get_description(driver, link)
            if description:
                valid_news.append({
                    'Başlık': title,
                    'Açıklama': description,
                    'Link': link
                })

        if not valid_news:
            logging.info("Tüm haber başlıklarının açıklamaları bulunamadı.")
        else:
            df_news = pd.DataFrame(valid_news)

            df_news = df_news[['Başlık', 'Açıklama', 'Link']]

            excel_dosyasi = 'habercekme_selenium.xlsx'

            with pd.ExcelWriter(excel_dosyasi, engine='xlsxwriter') as writer:
                df_news.to_excel(writer, index=False, sheet_name='Haberler')

                workbook = writer.book
                worksheet = writer.sheets['Haberler']

                table_format = {
                    'columns': [
                        {'header': 'Başlık', 'width': 50},
                        {'header': 'Açıklama', 'width': 100},
                        {'header': 'Link', 'width': 50},
                    ],
                    'style': 'Table Style Medium 9'
                }

                (max_row, max_col) = df_news.shape
                worksheet.add_table(0, 0, max_row, max_col - 1, table_format)

                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#7fb809',
                    'border': 1
                })
                for col_num, value in enumerate(df_news.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                for i, column in enumerate(df_news.columns):
                    column_length = max(df_news[column].astype(str).map(len).max(), len(column)) + 2
                    worksheet.set_column(i, i, column_length)

            logging.info(
                f"Haberler ve açıklamaları başarıyla '{excel_dosyasi}' dosyasına kaydedildi ve biçimlendirildi.")

except Exception as e:
    logging.error(f"Bir hata oluştu: {e}")
finally:
    driver.quit()
