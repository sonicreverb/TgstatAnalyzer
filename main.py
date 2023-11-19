import configparser
import time

import tables_managment as tables

from datetime import datetime
from api_utils import get_API_response, token_key

key_words = []
config_options = {}

try:
    # ЧТЕНИЕ ПАРАМЕТРОВ ИЗ ФАЙЛА КОНФИГУРАЦИИ
    config = configparser.ConfigParser()
    config.read('search_params.ini', encoding='utf-8')
    section_name = 'search_params'

    # отдельно получаем остальные опции поиска из конфига
    all_params = dict(config.items(section_name))

    for key, value in all_params.items():
        if key == 'key_words':
            key_words = value.split(', ')
        else:
            # приведение ключей конфига из типа нижнего регистра к типу camelCase, принимаемому API
            key_modifed = key.split('_')
            key_modifed = key_modifed[0] + ''.join(key.capitalize() for key in key_modifed[1:])
            if value == 'None' or value == 'False':
                value = '0'
            elif value == 'True':
                value = '1'

            config_options[key_modifed] = value

    print(f'[ЧТЕНИЕ ПАРАМЕТРОВ ИЗ ФАЙЛА КОНФИГУРАЦИИ] Данные успешно получены!'
          f'\n[ЧТЕНИЕ ПАРАМЕТРОВ ИЗ ФАЙЛА КОНФИГУРАЦИИ] Ключевые слова: {key_words}'
          f'\n[ЧТЕНИЕ ПАРАМЕТРОВ ИЗ ФАЙЛА КОНФИГУРАЦИИ] Параметры: {config_options}\n')
except Exception as ex:
    print(f"[ЧТЕНИЕ ПАРАМЕТРОВ ИЗ ФАЙЛА КОНФИГУРАЦИИ] Возникла ошибка! ({ex}).")

OUTPUT_FILENAME = 'TgstatAnalyzer_output.xlsx'


def main():
    # проверка на валидность чтения из конфига
    if not key_words or not config_options:
        print("[MAIN] Ошибка! Невозможно выполнить парсинг ввиду невалидного чтения файла конфигураций.")
        return None

    # список полученных ответов от API.
    responses_list = {}

    # определяем дату последней найденной публикации для каждого ключевого слова в таблице
    keywordsXLSX_data = tables.xlsx_connector.get_column_data(9)
    publish_data_XLSX = tables.xlsx_connector.get_column_data(8)
    last_publish_data_dict = {}

    for row_num in range(1, len(keywordsXLSX_data)):
        date_xlsx = publish_data_XLSX[row_num]
        current_keyword = keywordsXLSX_data[row_num]

        if not isinstance(date_xlsx, str):
            timestamp = date_xlsx.timestamp()
        else:
            datetime_object = datetime.strptime(date_xlsx, "%d/%m/%Y %H:%M:%S")
            timestamp = datetime.timestamp(datetime_object)

        if current_keyword not in last_publish_data_dict:
            last_publish_data_dict[current_keyword] = timestamp
        else:
            if last_publish_data_dict[current_keyword] < timestamp:
                last_publish_data_dict[current_keyword] = timestamp

    # проход по каждому ключевому слову
    for search_text in key_words:
        if config_options.get('endDate', '0') == '0':
            if search_text in last_publish_data_dict:
                config_options['endDate'] = int(time.time())
                config_options['startDate'] = int(last_publish_data_dict.get(search_text, 0))
        responses_list[search_text] = []
        try:
            current_offset = 0

            # формирование параметров для обращения к API
            search_params = {'token': token_key, 'q': search_text, 'offset': '0'}
            search_params.update(config_options)
            search_params['limit'] = 50

            if search_params['startDate'] != '0' and search_params['endDate'] != '0':
                print(f"[MAIN] Поиск публикаций во временном промежутке с"
                      f" [{datetime.fromtimestamp(float(search_params['startDate']))}] "
                      f"по [{datetime.fromtimestamp(float(search_params['endDate']))}].")
            else:
                print(f'[MAIN] Поиск публикаций во временном промежутке по умолчанию.')

            # ракировка limit условного (result_posts_num = тотальное кол-во постов)
            # и limit API (для API максимум 50 постов за запрос)
            result_posts_num = int(config_options['limit'])

            if result_posts_num == 0:
                result_posts_num = 50  # по умолчанию тотальное кол-во постов равно 50

            while (result_posts_num / (current_offset + 1)) >= 1:
                res = get_API_response('https://api.tgstat.ru/posts/search', search_params).json()
                responses_list[search_text].append(res)
                print(f'[MAIN] Запрос по ключевым словам "{search_text}" вернул состояние {res.get("status")}.', end='')
                if res.get("response", None):
                    if res['response'].get('count') != 50:
                        print(f"[{current_offset + res['response'].get('count') }/{res['response']['total_count']}] "
                              f"(limit = {result_posts_num}).",
                              end='')
                        print('\n', end='')
                        break
                    print(f"[{current_offset}/{res['response'].get('total_count')}] "
                          f"(limit = {result_posts_num}).", end='')
                print('\n', end='')
                current_offset += 50
                search_params['offset'] = current_offset
        except Exception as _ex:
            print(f"[MAIN] Во время попытки получить данные по ключевому слову {search_text} возникла ошибка! ({_ex}).")

    output_data = []
    # запись в таблицу
    for search_key in responses_list:
        for resp in responses_list[search_key]:
            try:
                data = resp.get('response', {}).get('items', None)
                if not data:
                    print(f"[АНАЛИЗ ПУБЛИКАЦИИ] Информация не найдена. API response status ({resp}")

                else:
                    for elem in data:
                        try:
                            info_dict = tables.get_item_data(elem)
                            info_dict['keyWord'] = search_key
                            output_data.append(info_dict)
                        except Exception as _ex:
                            print([f'[АНАЛИЗ ПУБЛИКАЦИИ] Во время попытки получить информацию из публикации было '
                                   f'вызвано исключение ({_ex})'])

            except Exception as _ex:
                print(f"Во время обработки полученного ответа от API было вызвано исключение ({_ex}).")
    if output_data:
        tables.write_to_excel(OUTPUT_FILENAME, output_data)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        print(f'Критическая ошибка! {ex}')
    print('Для завершения программы нажмите ENTER...')
    input()
