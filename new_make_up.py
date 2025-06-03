from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import time
import asyncio
import datetime
from pycbrf.toolbox import ExchangeRates
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from itertools import zip_longest
from selenium.common.exceptions import NoSuchElementException
import random
import secrets

import methods

options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")

s = Service(executable_path=methods.chrome_driver)
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
    adjusted_price = methods.adjust_price(price_in_euro)
    return round(adjusted_price * euro_to_rub+3, 2)

def parse_main_photo(driver):
    main_photos = []
    try:
        image_element = driver.find_element(By.CSS_SELECTOR,
            'picture.ProductPreviewSliderstyle__Picture-sc-195u70x-3 img')
        photo_url = image_element.get_attribute('src')

        if photo_url and '/product/' not in photo_url:
            main_photos.append(photo_url)
        else:
            main_photos.append("N/A")
    except Exception as e:
        print(f"Error extracting main photo: {e}")
        main_photos.append("N/A")

    return main_photos

def parse_tone_volume(driver):
    try:
        volume_block = driver.find_element(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')
        spans = volume_block.find_elements(By.TAG_NAME, 'span')
        volume_text = ' '.join([span.text for span in spans if span.text.strip()])
        volume_text = methods.clean_size(volume_text.lower())
        return  volume_text # например, "18 ml"
    except Exception as e:
        print(f"Error extracting volume: {e}")
        return "N/A"

def parse_tone_label(driver):
    try:
        # 1. Пробуем взять из aria-label, где явно указан выбранный тон
        tone_element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((
                By.XPATH,
                '//div[@role="button" and @aria-pressed="true" and contains(@aria-label, "selected")]'
            ))
        )
        aria_label = tone_element.get_attribute("aria-label")
        if aria_label:
            # Убираем " selected" с конца
            return aria_label.replace(" selected", "").strip()
    except:
        pass

    try:
        # 2. Резервный способ — берем span из верхнего блока, если он не содержит мусор
        tone_element = driver.find_element(
            By.XPATH,
            '//div[contains(@class, "VBguL")]//span[@data-nc="typography"]'
        )
        text = tone_element.text.strip()
        if text and len(text.split()) >= 2:
            return text
    except:
        pass

    return "N/A"

def parse_price(driver):
    try:
        price_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR,
                'span[data-qa-block="product_variant_price"]'
            ))
        )
        price_text = price_element.text.strip()
        if price_text:
            price_in_euro = float(price_text.replace("€", "").replace(",", "."))
            price_in_euro = methods.adjust_price(price_in_euro)
            price_in_rub = convert_to_rubles(price_in_euro) * 1.024
        #return price_text.replace('\xa0', ' ')  # заменяем неразрывный пробел на обычный
        return price_in_rub  # заменяем неразрывный пробел на обычный
    except Exception as e:
        print(f"Ошибка при получении цены: {e}")
        return "N/A"

def parse_photos(driver):
    photos = []
    try:
        # Открытие карусели (если требуется)
        large_picture_element = driver.find_element(By.CSS_SELECTOR,
                                                    'picture.ProductPreviewSliderstyle__Picture-sc-195u70x-3 img')
        large_picture_element.click()
        time.sleep(1)

        # Сбор всех изображений
        while True:
            try:
                # Получение всех изображений в карусели
                image_elements = driver.find_elements(By.CSS_SELECTOR,
                                                      'picture.ProductPreviewSliderstyle__Picture-sc-195u70x-3 img')

                # Извлекаем URL изображения с каждого элемента
                for image_element in image_elements:
                    photo_url = image_element.get_attribute('src')
                    # Исключаем URL из каталога /product/
                    if photo_url and '/product/' not in photo_url and photo_url not in photos:
                        photos.append(photo_url)

                # Переход к следующему изображению
                next_button_selector = 'div.swiper-button-next'
                next_button = driver.find_element(By.CSS_SELECTOR, next_button_selector)
                if next_button.is_enabled():
                    next_button.click()
                    time.sleep(1)
                else:
                    break  # Если кнопка недоступна, завершаем цикл

            except Exception as e:
                print(f"Error while navigating carousel: {e}")
                break

        close_button = driver.find_element(By.CSS_SELECTOR, 'button[aria-label="Close"]') # Закрытие карусели
        close_button.click()
        time.sleep(1)
    except Exception as e:
        print(f"Error parsing photos: {e}")
        photos = ["N/A"]

    return photos

def parse_default_item_without_tones(driver):
    sizes = []
    try:
        size_elements = driver.find_elements(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')
        for size_element in size_elements:
            try:
                size_span = size_element.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(1)')
                unit_span = size_element.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(2)')
                size_value = size_span.text.strip()
                size_unit = unit_span.text.strip()
                if size_value and size_unit:
                    sizes.append(f"{size_value}{size_unit}")
            except Exception as inner_e:
                print(f"Ошибка при обработке элемента размера: {inner_e}")
    except Exception as e:
        print(f"Ошибка при парсинге размеров: {e}")
        sizes = ["N/A"]

    print("Найденные размеры:", sizes)

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

    return list(dict.fromkeys(sizes)), list(dict.fromkeys(prices))

def parse_tones_photo(driver):
    all_photos = []
    all_volumes = []
    all_tone_labes = []
    all_prices = []

    # Изначально выбираем первый тон
    try:
        first_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.swiper-slide [role="button"]'))
        )
        first_button.click()
        time.sleep(0.5)
    except Exception as e:
        print(f"Не удалось кликнуть по первому тону: {e}")

    # Получаем количество тонов
    tone_count = len(driver.find_elements(By.CSS_SELECTOR, '.swiper-slide [role="button"]'))
    if tone_count == 0:
        print("ТОВАР БЕЗ ТОНОВ",parse_default_item_without_tones(driver))
        size, price = parse_default_item_without_tones(driver)
        size = [methods.clean_size(s) for s in size]
        main_photo = parse_main_photo(driver)
        all_photos.extend(main_photo)
        print(size)
        print(price)
        return all_photos, size, '', price

    for i in range(tone_count):
        try:
            # Заново получаем кнопки на каждом шаге
            buttons = driver.find_elements(By.CSS_SELECTOR, '.swiper-slide [role="button"]')
            btn = buttons[i]

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            driver.execute_script("arguments[0].click();", btn)

            time.sleep(0.4)  # дожидаемся обновления фото

            # Проверка наличия других объемов
            volume_buttons = driver.find_elements(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')
            if len(volume_buttons) <= 1:
                if is_add_to_cart_button_disabled(driver):
                    continue
                main_photo = parse_main_photo(driver)
                volumes = [parse_tone_volume(driver)]
                tone_label = [parse_tone_label(driver)]
                price = [parse_price(driver)]
                print(tone_label)
                print(volumes)
                print(price)
                all_photos.extend(main_photo)
                all_volumes.extend(volumes)
                all_tone_labes.extend(tone_label)
                all_prices.extend(price)

            else:
                for btn in volume_buttons[1:]:

                    print(len(volume_buttons))
                    btn.click()
                    time.sleep(2)
                    if is_add_to_cart_button_disabled(driver):
                        continue
                    main_photo = parse_main_photo(driver)
                    print(f"[Тон {i + 1}] Фото: {main_photo}")
                    volumes = [parse_tone_volume(driver)]
                    tone_label = [parse_tone_label(driver)]
                    price = [parse_price(driver)]
                    print(tone_label)
                    print(volumes)
                    print(price)
                    all_photos.extend(main_photo)
                    all_volumes.extend(volumes)
                    all_tone_labes.extend(tone_label)
                    all_prices.extend(price)

        except Exception as e:
            print(f"Ошибка при обработке тона {i + 1}: {e}")

    return all_photos,all_volumes,all_tone_labes,all_prices

def is_add_to_cart_button_disabled(driver):
    xpath = '//span[normalize-space(text())="Bald wieder verfügbar"]'
    try:
        driver.find_element(By.XPATH, xpath)
        return True
    except NoSuchElementException:
        return False


def go_to_next_page(driver):
    try:
        time.sleep(2)
        next_button = driver.find_element(By.XPATH,
            '//a[contains(@class, "Paginationstyle__NextPage") and contains(@href, "offset=")]')

        next_href = next_button.get_attribute("href")
        print(f"[INFO] Найдена ссылка на следующую страницу: {next_href}")

        if next_href:
            if next_href.startswith("/"):
                next_href = "https://www.flaconi.de" + next_href

            print(f"[INFO] Переход по адресу: {next_href}")
            driver.get(next_href)
            time.sleep(5)
            return True
        else:
            print("[WARN] Кнопка 'Следующая страница' есть, но ссылка пустая.")
            return False

    except Exception as e:
        print(f"[ERROR] Кнопка 'Следующая страница' не найдена: {e}")
        return False



# Основная асинхронная функция
async def main():
    try:
        driver.maximize_window()
        # Открываем первую страницу каталога
        current_url = 'https://www.flaconi.de/lippenstift/'
        driver.get(current_url)

        # Извлекаем category_slug из current_url, не изменяя саму переменную
        category_slug = current_url.rstrip('/').split('/')[-1]
        time.sleep(5)
        data = []

        # Цикл по страницам каталога
        while True:
            # Ждем, чтобы страница полностью загрузилась
            time.sleep(3)

            # Скроллим страницу до конца для загрузки всех товаров
            scroll_pause_time = 2
            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(scroll_pause_time)
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

            #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            #    ТУТ ВЫБИРАЕМ КАКИЕ ТОВАРЫ ЧЕРЕЗ ИНДЕКСЫ В МАССИВЕ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            #    Собираем ссылки на товары на текущей странице!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            product_elements = driver.find_elements(
                By.CSS_SELECTOR,
                'div[data-qa-block="product-section"] a[data-nc="card"]'
            )
            product_links = [elem.get_attribute('href') for elem in product_elements][:3]

            print(f"Найдено {len(product_links)} товаров на текущей странице.")
            catalog_url = driver.current_url  # сохраняем URL каталога

            # Обрабатываем каждый товар
            for product_link in product_links:
                driver.get(product_link)
                await asyncio.sleep(3)
                print("Обработка товара:", product_link)
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

                try:
                    brand_type_element = driver.find_element(By.CSS_SELECTOR,
                                                             'span[data-qa-block="product_brand_tyoe"]')
                    brand_type = brand_type_element.text.strip()
                    brand_type = methods.process_brand_type(brand_type)  # Обрабатываем значение
                except Exception:
                    brand_type = ""

                full_product_name = f"{brand_name} {product_name}".strip()
                if brand_type:
                    full_product_name += f" {brand_type}"

                sizes = []
                try:
                    size_elements = driver.find_elements(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')
                    for size_element in size_elements:
                        try:
                            size_span = size_element.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(1)')
                            unit_span = size_element.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(2)')
                            size_value = size_span.text.strip()
                            size_unit = unit_span.text.strip()
                            if size_value and size_unit:
                                sizes.append(f"{size_value}{size_unit}")
                        except Exception as inner_e:
                            print(f"Ошибка при обработке элемента размера: {inner_e}")
                except Exception as e:
                    print(f"Ошибка при парсинге размеров: {e}")
                    sizes = ["N/A"]

                print("Найденные размеры:", sizes)

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

                #sizes, prices = methods.clean_and_filter_sizes_and_prices(sizes, prices)
                photos = parse_photos(driver)  # Первые пять основных фото
                #tones_and_prices = parse_tones(driver) # Тона и цены
                tone_photos, volumes, tones, prices = parse_tones_photo(driver)  # Главная фотка тона и при переключении тонов
                external_id = random.randint(10**9, 10**10 - 1)
                parent_uid_counter = random.randint(10**9, 10**10 - 1)
                first_row = True
                # Цикл для первых главных фоток
                for photo in photos:
                    data.append({
                        'Brand': brand_name,
                        'Name': full_product_name,
                        'Text': f"Ссылка на австрийский сайт: {product_link}" if first_row else  '',
                        'Link': product_link if first_row else '',
                        'Title': full_product_name if first_row else '',
                        'Photo': photo or '',  # подставим пустую строку, если фото нет
                        'Editions': '',
                        'Price': '',
                        'Parent UID': parent_uid_counter,
                        'External ID': parent_uid_counter if first_row else  '',
                        'Category': f"Макияж;{brand_type};" if first_row else  '' ,
                        'Characteristics:Среднее время доставки': '2-3 недели'  if first_row else  '',

                    })
                    first_row = False
                for i, j, k, l in zip_longest(tone_photos, volumes, tones, prices, fillvalue=''):
                    external_id = secrets.randbelow(10 ** 10 - 10 ** 9) + 10 ** 9
                    data.append({
                        'Brand': brand_name,
                        'Name': full_product_name,
                        'Text': '',
                        'Link': product_link if first_row else '',
                        'Title': full_product_name,
                        'Photo': i or '',
                        'Editions': f"Тон:{k};{j}" if k else j,
                        'Price': l,
                        'Parent UID': parent_uid_counter,
                        'External ID': external_id
                    })


                # data.append({
                #     'Brand': '',
                #     'Name': full_product_name,
                #     'Text': '',
                #     'Link': '',
                #     'Title': '',
                #     'Photo': '',
                #     'Editions': '',
                #     'Price': '',
                #     'Parent UID': parent_uid_counter,
                #     'Category': f"Макияж;{brand_type};",
                #     'Characteristics:Среднее время доставки': '2-3 недели',
                #     'External ID': ''
                # })

                # data.append({
                #     'Brand': brand_name,
                #     'Name': full_product_name,
                #     'Text': f"Ссылка на австрийский сайт: {product_link}",
                #     'Link': '',
                #     'Title': '',
                #     'Photo': '',
                #     'Editions': '',
                #     'Price': '',
                #     'Parent UID': parent_uid_counter,
                #     'External ID': ''
                # })

                driver.get(catalog_url)  # возвращаемся в каталог
                time.sleep(3)

            # Перед поиском кнопки "Следующая страница" делаем скролл вниз
            time.sleep(3)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            #time.sleep(2)
            success = go_to_next_page(driver)
            if not success:
                break



        # Сохранение данных в Excel
        df = pd.DataFrame(data, columns=[
            'Name', 'Text', 'Link', 'Title', 'Editions', 'Price', 'Photo',
            'Parent UID', 'External ID', 'Category', 'Characteristics:Среднее время доставки', 'Brand'
        ])
        current_time = datetime.datetime.now().strftime('%d-%m-%Y_%H-%M-%S')
        filename = f'{category_slug}_{current_time}.xlsx'
        df.to_excel(filename, index=False)
        print(f"Данные успешно сохранены в файл '{filename}'")

    except Exception as ex:
        print(f"Ошибка: {ex}")
    finally:
        driver.close()
        driver.quit()

# Запуск основной асинхронной функции
if __name__ == "__main__":
    asyncio.run(main())

