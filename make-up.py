from pandas.core.internals.blocks import external_values
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

import methods
import openpyxl

options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")

s = Service(executable_path='C:/Users/evzhe/Parser-Flaconi/chromedriver-win64/chromedriver.exe')
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
    return round(adjusted_price * euro_to_rub, 2)

def click_different_volumes_2(driver):
    try:
        volume_buttons = driver.find_elements(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')

        if len(volume_buttons) <= 1:
            print("Только один вариант объема, пропускаем")
            return None, None

        volumes_data = []
        seen_sizes = set()  # Для отслеживания уже добавленных размеров

        for btn in volume_buttons:
            try:
                btn.click()
                time.sleep(2)

                size_element = btn.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(1)')
                unit_element = btn.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(2)')
                size_value = size_element.text.strip()
                size_unit = unit_element.text.strip()
                size = f"{size_value}{size_unit}"

                # Приводим размер к человекочитаемому виду
                if "ml" in size:
                    size_cleaned = f"Объём:{size_value} мл"
                elif "stk" in size.lower():
                    size_cleaned = f"Объём:{size_value} шт"
                elif "g" in size:
                    size_cleaned = f"Объём:{size_value} г"
                else:
                    size_cleaned = size

                # Добавляем только уникальные размеры
                if size_cleaned not in seen_sizes:
                    volumes_data.append({'size': size_cleaned})
                    seen_sizes.add(size_cleaned)

            except Exception as e:
                print(f"Ошибка при обработке варианта объема: {e}")
                continue

        return volumes_data

    except Exception as e:
        print(f"Ошибка при поиске вариантов объемов: {e}")
        return None, None

def click_different_volumes_1(driver):
    try:
        volume_buttons = driver.find_elements(By.CSS_SELECTOR, 'div[data-qa-block="product_variant_quantity"]')

        if len(volume_buttons) <= 1:
            print("Только один вариант объема, пропускаем")
            return None, None

        volumes_data = []
        seen_sizes = set()  # Для отслеживания уже добавленных размеров

        for btn in volume_buttons:
            try:
                #btn.click()
                #time.sleep(2)

                size_element = btn.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(1)')
                unit_element = btn.find_element(By.CSS_SELECTOR, 'span[data-nc="typography"]:nth-of-type(2)')
                size_value = size_element.text.strip()
                size_unit = unit_element.text.strip()
                size = f"{size_value}{size_unit}"

                # Приводим размер к человекочитаемому виду
                if "ml" in size:
                    size_cleaned = f"Объём:{size_value} мл"
                elif "stk" in size.lower():
                    size_cleaned = f"Объём:{size_value} шт"
                elif "g" in size:
                    size_cleaned = f"Объём:{size_value} г"
                else:
                    size_cleaned = size

                # Добавляем только уникальные размеры
                if size_cleaned not in seen_sizes:
                    volumes_data.append({'size': size_cleaned})
                    seen_sizes.add(size_cleaned)

            except Exception as e:
                print(f"Ошибка при обработке варианта объема: {e}")
                continue

        return volumes_data

    except Exception as e:
        print(f"Ошибка при поиске вариантов объемов: {e}")
        return None, None

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

def parse_tones_photo(driver):
    buttons = driver.find_elements(By.CSS_SELECTOR, '.swiper-slide [role="button"]')
    all_photos = []
    for btn in buttons:
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", btn)  # если требуется скролл
            btn.click()
            main_photo = parse_main_photo(driver)
            print("ПЕРВЫЙ ПРИНТ:", main_photo)
            time.sleep(0.3)  # задержка, чтобы успевала отработать анимация
            all_photos.extend(main_photo)  # Добавляем найденные фото в общий список
        except Exception as e:
            print(f"Ошибка при клике: {e}")

    return all_photos

def parse_tones(driver):
    try:
        # Открываем все варианты цветов
        view_more = driver.find_element(
            By.CSS_SELECTOR,
            'button[data-qa-block="view-more-colors"]'
        )
        view_more.click()
        time.sleep(1)  # ждём появления элементов

        # Получаем тона
        tone_elements = driver.find_elements(
            By.CSS_SELECTOR,
            'span.ColorSelectorVariantstyle__Content-sc-jlekvj-2[data-nc="typography"]'
        )
        raw_tones = [el.text.strip() for el in tone_elements]
        codes = [t.split(" - ")[0].replace("Nr. ", "") for t in raw_tones]

        # Получаем цены
        price_elements = driver.find_elements(
            By.CSS_SELECTOR,
            'span.ColorSelectorPricestyle__PriceBox-sc-6eo1ts-0[data-nc="typography"]'
        )
        raw_prices = [el.text.strip().replace("€", "").replace(",", ".") for el in price_elements]

        # Конвертируем в рубли
        prices_in_rub = []
        for price in raw_prices:
            try:
                euro = float(price)
                rub = convert_to_rubles(euro)
                prices_in_rub.append(rub)
            except:
                prices_in_rub.append("N/A")

        # Собираем результат
        tones_and_prices = []
        for code, rub_price in zip(codes, prices_in_rub):
            tones_and_prices.append({
                "Edition": f"Тон:{code}",
                "Price": rub_price
            })

        # === Вместо закрытия ищем и нажимаем кнопку "Auswählen" ===
        try:
            auswahl_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    '//button[.//span[text()="Auswählen"]]'
                ))
            )
            auswahl_button.click()
            time.sleep(1)
        except Exception as e:
            print(f"Не удалось нажать кнопку 'Auswählen': {e}")

        print("Тона и цены:", tones_and_prices)
        return tones_and_prices

    except Exception as e:
        print(f"Error parsing tones: {e}")
        return []

# Основная асинхронная функция
async def main():
    try:
        driver.maximize_window()
        # Открываем первую страницу каталога
        driver.get('https://www.flaconi.at/make-up/')
        time.sleep(5)
        parent_uid_counter = 1
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

            # Собираем ссылки на товары на текущей странице
            product_elements = driver.find_elements(By.CSS_SELECTOR, 'a.Linkstyle__A-sc-16w5a4n-0')
            product_links = [elem.get_attribute('href') for elem in product_elements][8:10]

            print(f"Найдено {len(product_links)} товаров на текущей странице.")

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

                sizes, prices = methods.clean_and_filter_sizes_and_prices(sizes, prices)
                photos = parse_photos(driver)
                tones_and_prices = parse_tones(driver)
                tone_photos = parse_tones_photo(driver)
                print("ВТОРОЙ ПРИНТ", tone_photos)
                new_sizes = click_different_volumes_1(driver)
                print(len(new_sizes))

                if tones_and_prices:
                    # кликаем для наполнения new_sizes
                    click_different_volumes_2(driver)

                    external_id = parent_uid_counter + 1
                    first_row = True

                    # Вместо zip используем zip_longest
                    for tp, photo_url in zip_longest(tones_and_prices, tone_photos, fillvalue=None):
                        if tp is None:
                            continue  # пропускаем, если данных о тоне нет

                        external_id += 1
                        data.append({
                            'Brand': brand_name,
                            'Name': full_product_name,
                            'Text': '',
                            'Link': product_link if first_row else '',
                            'Title': full_product_name if first_row else '',
                            'Photo': photo_url or '',  # подставим пустую строку, если фото нет
                            'Editions': tp['Edition'] + ";" + (sizes[0] if sizes else "Объём:N/A"),
                            'Price': tp['Price'],
                            'Parent UID': parent_uid_counter,
                            'External ID': external_id
                        })
                        first_row = False
                    data.append({
                        'Brand': '',
                        'Name': full_product_name,
                        'Text': '',
                        'Link': '',
                        'Title': '',
                        'Photo': '',
                        'Editions': '',
                        'Price': '',
                        'Parent UID': ''
                    })

                    for i in range(1, len(new_sizes)):
                        tones_and_prices = []
                        tones_and_prices = parse_tones(driver)
                        for tp in tones_and_prices:
                            data.append({
                                'Brand': '',
                                'Name': full_product_name,
                                'Text': '',
                                'Link': '',
                                'Title': '',
                                'Photo': '',
                                'Editions': tp['Edition'] + ";" + new_sizes[i]['size'],
                                'Price': tp['Price'],
                                'Parent UID': parent_uid_counter
                            })

                    # # теперь дополнительные фото — по одной строке с пустыми всеми остальными полями
                    # for extra_photo in photos[1:]:
                    #     data.append({
                    #         'Brand': '',
                    #         'Name': full_product_name,
                    #         'Text': '',
                    #         'Link': '',
                    #         'Title': '',
                    #         'Photo': extra_photo,
                    #         'Editions': '',
                    #         'Price': '',
                    #         'Parent UID': parent_uid_counter
                    #     })

                    for tone_photo in tone_photos[:]:
                        data.append({
                            'Brand': '',
                            'Name': full_product_name,
                            'Text': '',
                            'Link': '',
                            'Title': '',
                            'Photo': tone_photo,
                            'Editions': '',
                            'Price': '',
                            'Parent UID': parent_uid_counter
                        })

                    # и, как раньше, добавляем категория/ссылка в Text
                    data.append({
                        'Brand': brand_name,
                        'Name': full_product_name,
                        'Text': f"Ссылка на австрийский сайт: {product_link}",
                        'Link': '',
                        'Title': '',
                        'Photo': '',
                        'Editions': '',
                        'Price': '',
                        'Parent UID': '',
                        'Category': f"Макияж;{brand_type};",
                        'Characteristics:Среднее время доставки': '2-3 недели'
                    })

                parent_uid_counter += 1

            # Перед поиском кнопки "Следующая страница" делаем скролл вниз
            time.sleep(3)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

            # Ищем кнопку по XPath
            next_buttons = driver.find_elements(By.XPATH,
                                                "//a[contains(@class, 'Paginationstyle__NextPage-sc-d38xli-3')]")
            if next_buttons:
                next_href = next_buttons[0].get_attribute("href")
                print(f"Найден next_href: {next_href}")
                if next_href:
                    # Если ссылка относительная – добавляем домен
                    if next_href.startswith('/'):
                        next_page_url = f"https://www.flaconi.at{next_href}"
                    else:
                        next_page_url = next_href
                    print("Переход на следующую страницу:", next_page_url)
                    driver.get(next_page_url)
                    time.sleep(5)
                else:
                    print("Ссылка на следующую страницу отсутствует. Выходим из цикла.")
                    break
            else:
                print("Кнопка 'Следующая страница' не найдена. Выходим из цикла.")
                break

        # Сохранение данных в Excel
        df = pd.DataFrame(data, columns=[
            'Name', 'Text', 'Link', 'Title', 'Editions', 'Price', 'Photo',
            'Parent UID', 'External ID', 'Category', 'Characteristics:Среднее время доставки', 'Brand'
        ])
        current_time = datetime.datetime.now().strftime('%d-%m-%Y_%H-%M-%S')
        filename = f'Make_up_{current_time}.xlsx'
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

