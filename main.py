import configparser
import pickle
import tables_managment as tables
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


def main():
    # проверка на валидность чтения из конфига
    if not key_words or not config_options:
        print("[MAIN] Ошибка! Невозможно выполнить парсинг ввиду невалидного чтения файла конфигураций.")
        return None

    # список полученных ответов от API.
    responses_list = []

    # проход по каждому ключевому слову
    for search_text in key_words:
        try:
            current_offset = 0

            # формирование параметров для обращения к API
            search_params = {'token': token_key, 'q': search_text, 'offset': '0'}
            search_params.update(config_options)
            search_params['limit'] = 50

            # ракировка limit условного (result_posts_num = тотальное кол-во постов)
            # и limit API (для API максимум 50 постов за запрос)
            result_posts_num = int(config_options['limit'])

            if result_posts_num == 0:
                result_posts_num = 50  # по умолчанию тотальное кол-во постов равно 50

            while (result_posts_num / (current_offset + 1)) >= 1:
                res = get_API_response('https://api.tgstat.ru/posts/search', search_params).json()
                responses_list.append(res)
                print(f'[MAIN] Запрос по ключевому слову ({search_text}) вернул состояние {res.get("status")}.', end='')
                print(res)
                if res.get("response", None):
                    if res['response'].get('count') != 50:
                        print(f"[{current_offset + res['response'].get('count') }/{res['response']['total_count']}]")
                        break
                    print(f"[{current_offset}/{res['response'].get('total_count')}]")

                current_offset += 50
                search_params['offset'] = current_offset
        except Exception as _ex:
            print(f"[MAIN] Во время попытки получить данные по ключевому слову {search_text} возникла ошибка! ({_ex}).")

    with open('tables_managment/responses.pkl', 'wb') as file:
        pickle.dump(responses_list, file)

    output_data = []
    # запись в таблицу
    for resp in responses_list:
        try:
            print(resp)
            data = resp.get('response', {}).get('items', None)
            if not data:
                print(f"API response bad status ({resp.get('status')}).")

            else:
                for elem in data:
                    try:
                        output_data.append(tables.get_item_data(elem))
                    except Exception as _ex:
                        print([f'[АНАЛИЗ ПУБЛИКАЦИИ] Во время попытки получить информацию из публикации было '
                               f'вызвано исключение ({_ex})'])

        except Exception as _ex:
            print(f"Во время обработки полученного ответа от API было вызвано исключение ({_ex}).")
    if output_data:
        tables.write_to_excel('TgstatAnalyzer_output.xlsx', output_data)


if __name__ == "__main__":
    # main()
    with open('tables_managment/responses.pkl', 'rb') as file:
        li = []
        for loaded_variable in pickle.load(file):
            print(loaded_variable)
            for item2 in loaded_variable.get('response', {}).get('items', [None]):
                if item2:
                    li.append(tables.get_item_data(item2))
                else:
                    print("ERROR!")

        tables.write_to_excel('test_output.xlsx', li)
