from datetime import datetime
import openpyxl
import os


# принимает на вход API class Post, возвращает словарь с необходимыми хар-ками публикации
def get_item_data(item):
    # ссылка на сообщение
    post_url = "https://" + item['link']

    # содержимое сообщения
    post_text = item['text']

    # ссылка на фото
    if item['media'].get('media_type') == 'mediaPhoto':
        post_image_url = item['media']['file_url']
        if not post_image_url:
            post_image_url = "TgstatAPI return null img source."
    else:
        post_image_url = 'None'

    # ссылка на видео
    if item['media'].get('media_type', None) == 'mediaDocument' and item['media'].get('mime_type', '0') == 'video/mp4':
        post_video_url = item['media']['file_url']
        if not post_video_url:
            post_video_url = "TgstatAPI return null video source."
    else:
        post_video_url = 'None'

    # ссылка на источник
    post_source_link = post_url[:post_url.rfind('/')]

    # дата публикации
    timestamp = item['date']
    dt_object = datetime.utcfromtimestamp(timestamp)
    post_date = dt_object.strftime('%d/%m/%Y %H:%M:%S')

    return {'postURL': post_url, 'postText': post_text, 'postImgURL': post_image_url, 'postVideoURL': post_video_url,
            'postSourceURL': post_source_link, 'postPublishDate': post_date}


def write_to_excel(filename, data):
    # проверяем, существует ли файл
    if os.path.exists(filename):
        # если файл существует, открываем его и выбираем активный лист
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
    else:
        # если файл не существует, создаем новую книгу Excel и выбираем активный лист
        wb = openpyxl.Workbook()
        ws = wb.active

    # добавляем заголовки таблицы
    headers = list(data[0].keys())
    for col_num, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_num, value=header)

    # записываем данные в конец таблицы
    for row_num, row_data in enumerate(data, ws.max_row + 1):
        for col_num, key in enumerate(headers, 1):
            ws.cell(row=row_num, column=col_num, value=row_data[key])

    # сохраняем книгу Excel
    wb.save(filename)

    print(f'[ЗАПИСЬ В ЭКСЕЛЬ] Данные успешно записаны в файл {filename}')


# import pickle
#
# # Загрузка переменной из файла
# with open('responses.pkl', 'rb') as file:
#     loaded_variable_list = pickle.load(file)
#
# li = []
# for loaded_variable in loaded_variable_list:
#     for item2 in loaded_variable['response']['items']:
#         li.append(get_item_data(item2))
#
# write_to_excel('test_output.xlsx', li)
