from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import time
from googletrans import Translator
import asyncio
import datetime
from pycbrf.toolbox import ExchangeRates

options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")

s = Service(executable_path='C:/Users/azaza/PycharmProjects/PythonProject13/chromedriver-win64/chromedriver.exe')
driver = webdriver.Chrome(service=s, options=options)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    'source': '''
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_JSON;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
    '''
})

def get_euro_to_rub_rate():
    rates = ExchangeRates()
    return float(rates['EUR'].value)

euro_to_rub = get_euro_to_rub_rate()
print("Нынешний курс", euro_to_rub)
def convert_to_rubles(price_in_euro):

    adjusted_price = adjust_price(price_in_euro)
    return round(adjusted_price * euro_to_rub, 2)

async def translate_description(description, src_lang='de', dest_lang='ru'):
    translator = Translator()
    try:
        print(description)
        translated = await translator.translate(description, src=src_lang, dest=dest_lang)
        return translated.text
    except Exception as e:
        print(f"Error during translation: {e}")
        return "Перевод недоступен"

def clean_and_filter_sizes_and_prices(sizes, prices, tones):
    filtered_data = {}

    for size, price in zip(sizes, prices):
        if "Duftset" in size:
            continue

        if "ml" in size:
            size_value = size.split("ml")[0].strip()
            size_cleaned = f"Тон:{', '.join(tones)};Объём:{size_value} мл"
        elif "g" in size:
            size_value = size.split("g")[0].strip()
            size_cleaned = f"Тон:{', '.join(tones)};Объём:{size_value} г"
        else:
            continue

        try:
            price = float(price)
        except ValueError:
            continue

        if size_cleaned not in filtered_data or price > filtered_data[size_cleaned]:
            filtered_data[size_cleaned] = price

    filtered_sizes = list(filtered_data.keys())
    filtered_prices = list(filtered_data.values())

    return filtered_sizes, filtered_prices

def adjust_price(price_in_euro):
    if price_in_euro < 20:
        price_in_euro *= 2.2
    elif 20 <= price_in_euro < 30:
        price_in_euro *= 2.0
    elif 30 <= price_in_euro < 50:
        price_in_euro *= 1.7
    elif 50 <= price_in_euro < 60:
        price_in_euro *= 1.6
    elif 60 <= price_in_euro < 70:
        price_in_euro *= 1.5
    else:
        price_in_euro *= 1.4

    return price_in_euro

def parse_photos(driver):
    photos = []
    try:
        large_picture_element = driver.find_element(By.CSS_SELECTOR,
                                                    'picture.ProductPreviewSliderstyle__Picture-sc-195u70x-3 img')
        large_picture_element.click()
        time.sleep(1)
        while True:
            try:
                image_elements = driver.find_elements(By.CSS_SELECTOR,
                                                      'picture.ProductPreviewSliderstyle__Picture-sc-195u70x-3 img')

                for image_element in image_elements:
                    photo_url = image_element.get_attribute('src')
                    if photo_url and '/product/' not in photo_url and photo_url not in photos:
                        photos.append(photo_url)

                next_button_selector = 'div.swiper-button-next'
                next_button = driver.find_element(By.CSS_SELECTOR, next_button_selector)
                if next_button.is_enabled():
                    next_button.click()
                    time.sleep(1)
                else:
                    break

            except Exception as e:
                print(f"Error while navigating carousel: {e}")
                break

        close_button = driver.find_element(By.CSS_SELECTOR, 'button[aria-label="Close"]')
        close_button.click()
        time.sleep(1)
    except Exception as e:
        print(f"Error parsing photos: {e}")
        photos = ["N/A"]
    return photos

async def main():
    try:
        driver.maximize_window()
        driver.get('https://www.flaconi.at/foundation/')
        time.sleep(5)
        scroll_pause_time = 2
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(scroll_pause_time)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        product_elements = driver.find_elements(By.CSS_SELECTOR, 'div[data-product-list-id] a[data-testid="card"]')
        product_links = [elem.get_attribute('href') for elem in product_elements if elem.get_attribute('href')]
        parent_uid_counter = 1
        data = []
        for product_link in product_links:
            driver.get(product_link)
            await asyncio.sleep(3)

            try:
                brand_element = driver.find_element(By.CSS_SELECTOR, 'a[data-qa-block="product_brand_name"]')
                brand_name = brand_element.get_attribute('title')
            except Exception:
                brand_name = "N/A"

            try:
                name_element = driver.find_element(By.CSS_SELECTOR, 'span[data-qa-block="product_name"]')
                product_name = name_element.text.strip()
            except Exception:
                product_name = "N/A"

            full_product_name = f"{brand_name} {product_name}".strip()

            # try:
            #     description_element = driver.find_element(By.CSS_SELECTOR, 'span.O3VZd.pdp-product-info-details')
            #     product_description = description_element.text.strip()
            # except Exception:
            #     product_description = "N/A"

            # translated_description = await translate_description(product_description)

            sizes = []
            try:
                size_elements = driver.find_elements(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')

                for size_element in size_elements:
                    try:
                        size_span = size_element.find_element(By.CSS_SELECTOR,
                                                              'span[data-nc="typography"]:nth-of-type(1)')
                        unit_span = size_element.find_element(By.CSS_SELECTOR,
                                                              'span[data-nc="typography"]:nth-of-type(2)')

                        size_value = size_span.text.strip()
                        size_unit = unit_span.text.strip()
                        if size_value and size_unit:
                            sizes.append(f"{size_value}{size_unit}")
                    except Exception as inner_e:
                        print(f"Ошибка при обработке элемента: {inner_e}")
            except Exception as e:
                print(f"Ошибка при парсинге размеров: {e}")
                sizes = ["N/A"]

            print(sizes)

            prices = []
            try:
                price_elements = driver.find_elements(By.CSS_SELECTOR, 'span[data-qa-block="product_variant_price"]')
                for price_element in price_elements:
                    price_text = price_element.text.strip()
                    if price_text:
                        price_in_euro = float(price_text.replace("€", "").replace(",", "."))
                        price_in_rub = convert_to_rubles(price_in_euro)
                        prices.append(price_in_rub)
            except Exception:
                prices = ["N/A"]

            tones = []
            try:
                tone_elements = driver.find_elements(By.CSS_SELECTOR, 'div.ProductColorSelectorstyle__Header-sc-1ozarh3-0 span[data-nc="typography"]')

                for element in tone_elements:
                    tone_text = element.text.strip()
                    if tone_text:
                        tones.append(tone_text)
            except Exception as e:
                print(f"Ошибка при парсинге тонов: {e}")
                tones = ["N/A"]

            sizes, prices = clean_and_filter_sizes_and_prices(sizes, prices, tones)
            photos = parse_photos(driver)

            if photos:
                data.append({
                    'Brand': '',
                    'Name': full_product_name,
                    'Text': '',
                    'Link': product_link,
                    'Title': full_product_name,
                    'Photo': photos[0],
                    'Editions': sizes[0] if sizes else "N/A",
                    'Price': prices[0] if prices else "N/A",
                    'Parent UID': parent_uid_counter
                })

                for photo in photos[1:]:
                    data.append(
                        {'Brand': '', 'Name': full_product_name, 'Text': '', 'Link': '', 'Title': '', 'Photo': photo,
                         'Editions': '', 'Price': '', 'Parent UID': ''})

                for size, price in zip(sizes[1:], prices[1:]):
                    data.append(
                        {'Brand': '', 'Name': full_product_name, 'Text': '', 'Link': '', 'Title': '', 'Photo': '',
                         'Editions': size, 'Price': price, 'Parent UID': parent_uid_counter})
                if sizes and prices:
                    data.append({
                        'Brand': '',
                        'Name': full_product_name,
                        'Text': '',
                        'Link': '',
                        'Title': '',
                        'Photo': '',
                        'Editions': '',
                        'Price': '',
                        'Parent UID': '',
                        'Category': 'Тональный крем',
                        'Characteristics:Среднее время доставки': '3 недели'
                    })
                if sizes and prices:
                    data.append({
                        'Brand': brand_name,
                        'Name': full_product_name,
                        'Text': f"Ссылка на оригинальный товар: {product_link}",
                        'Link': '',
                        'Title': '',
                        'Photo': '',
                        'Editions': '',
                        'Price': '',
                        'Parent UID': ''
                    })

            parent_uid_counter += 1

        df = pd.DataFrame(data,
                          columns=['Name', 'Text', 'Link', 'Title', 'Editions', 'Price', 'Photo',
                                   'Parent UID', 'Category', 'Characteristics:Среднее время доставки', 'Brand'])

        current_time = datetime.datetime.now().strftime('%d-%m-%Y_%H-%M-%S')
        filename = f'Foundation_{current_time}.xlsx'

        df.to_excel(filename, index=False)
        print(f"Данные успешно сохранены в файл '{filename}'")

    except Exception as ex:
        print(f"Ошибка: {ex}")
    finally:
        driver.close()
        driver.quit()

if __name__ == "__main__":
    asyncio.run(main())
