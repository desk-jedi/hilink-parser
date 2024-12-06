#
# ---------------------------------------------------------------------------
#                      Скрипт парсинга данных модемов
#                           Автор: Антон Степанов
#                          Версия: 4.01 (ProxyLab)
#                             Дата: 2024-12-05
# ---------------------------------------------------------------------------
# Описание:
#    Данный скрипт предназначен для автоматизации процесса получения
#    информации о модемах, используя Selenium для веб-автоматизации
#    и Beautiful Soup для парсинга данных. Скрипт получает локальные
#    IP-адреса модемов и извлекает важные данные, такие как модель,
#    серийный номер и мобильный номер, сохраняет их в лог-файл.
# ---------------------------------------------------------------------------
#


from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from socket import gethostname, getaddrinfo
from bs4 import BeautifulSoup
from openpyxl import Workbook
import datetime


def get_IP() -> dict:
    """
    Функция get_IP() получает список локальных IP-адресов компьютера и
    формирует ссылки для доступа к информации о модемах.

    Возвращает:
    list: Список URL-адресов для извлечения данных о модемах.

    Исключения:
    SystemExit: Завершает выполнение программы, если не удалось
    найти ни одного локального IP-адреса в диапазоне 192.168.x.x.
    """
    host = gethostname()  # Получаем имя текущего хоста.

    # Получаем все IP-адреса, связанные с именем хоста.

    ip_addresses = [addr[4][0] for addr in getaddrinfo(host, 0)]

    temp = []  # Список для хранения подходящих IP-адресов.

    # Фильтруем IP-адреса и оставляем только локальные.

    for url in ip_addresses:
        if url.find("192.168") != -1:
            temp.append(url)
    # Проверяем, были ли найдены подходящие IP-адреса.

    if temp:
        print(f"Подходящих IP адресов в системе найдено: {len(temp) - 1}")
    else:
        print("Подходящих IP адресов в системе не обнаружено")
        exit()  # Завершаем выполнение программы.
    # Формируем список URL для доступа к странице устройства.

    temp = ["http://" + url[:-2] + "/html/deviceinformation.html" for url in temp]

    return temp


def create_table(data_sheet) -> None:
    """
    Процедура create_table создает Excel-таблицу и сохраняет в ней данные о модемах.
    Аргументы:
    data_sheet (list): Данные о модемах для записи в таблицу.
    """
    workbook = Workbook()  # Создаем новый Excel-файл.
    sheet = workbook.active  # Получаем активный лист.
    sheet.title = "Данные модемов"  # Устанавливаем название листа.

    headers = [
        "IP-адрес",
        "Модель",
        "Серийный номер",
        "IMEI",
        "ICCID/SIM",
        "Моб. номер",
    ]

    sheet.append(headers)  # Добавляем заголовки в таблицу.

    # Заполняем таблицу данными о модемах.

    for row in data_sheet:
        sheet.append(row)
    workbook.save("modem_sheet.xlsx")  # Сохраняем таблицу в файл.


def main() -> None:
    """
    Основная функция, осуществляющая выполнение парсинга данных
    модемах и сохранение в Excel-файл.
    """
    print("Скрипт парсинга данных модемов")
    print("Версия 4.02 от 05.12.24\n")

    urls = get_IP()  # Получаем список IP-адресов модемов.
    data_sheet = []  # Контейнер для сохранения данных о модемах.
    count_address = 0  # Счетчик адресов.

    # Проходим по каждому URL для получения данных о модемах.

    for url in urls:
        count_address += 1  # Увеличиваем счетчик адресов.

        url_format = url.replace("http://", "").replace(
            "/html/deviceinformation.html", ""
        )  # Форматируем URL.

        try:
            driver = webdriver.Firefox()  # Инициализируем веб-драйвер.
        except WebDriverException:
            print("В системе отсутствует браузер FireFox или возникла другая ошибка")
        try:
            if count_address != 1:  # Пропускаем первый адрес, так как это, Mikrotik
                driver.get(url)  # Переходим по URL.

                html = driver.page_source  # Получаем HTML-код страницы.
            else:
                continue  # Пропускаем первый адрес.
        except WebDriverException:
            print(
                f"[ {count_address - 1} ] Не удалось получить данные по адресу {url_format}. Проверьте корректность URL и доступность сайта"
            )
            continue  # Переходим к следующему адресу.
        block = BeautifulSoup(html, "lxml")  # Парсим HTML-код.

        info_values = block.find_all(
            "td", class_="info_value"
        )  # Извлекаем нужные данные.

        output_name = (
            "output-" + str(datetime.datetime.now())[:10] + ".log"
        )  # Название лог-файла.

        try:
            with open(output_name, "a+", encoding="utf-8") as output:
                # Записываем данные в лог-файл.

                output.write(f"{url_format}:\n")
                output.write(f"Модель: {info_values[0].text}\n")
                output.write(f"Серийный номер: {info_values[1].text}\n")
                output.write(f"IMEI: {info_values[2].text}\n")
                output.write(f"ICCID/SIM: {info_values[4].text}\n")
                output.write(f"Моб. номер: {info_values[5].text}\n\n")

                # Добавляем данные в таблицу.

                data_sheet.append(
                    [
                        url_format,
                        info_values[0].text,
                        info_values[1].text,
                        info_values[2].text,
                        info_values[4].text,
                        info_values[5].text,
                    ]
                )

                output.close()  # Закрываем лог-файл.
            print(
                f"[ {count_address - 1} ] Данные по адресу {url_format} успешно получены"
            )
        except Exception:
            print(
                f"[ {count_address - 1} ] Не удалось получить данные по адресу {url_format}. Проверьте корректность URL и доступность сайта"
            )
            continue  # Переходим к следующему адресу.
        driver.close()  # Закрываем веб-драйвер.
    create_table(data_sheet)  # Создаем таблицу с данными о модемах.


if __name__ == "__main__":
    main()  # Запускаем основную функцию.
