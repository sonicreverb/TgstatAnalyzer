import requests


# возвращает результат GET-запроса с передачей параметров
def get_API_response(url: str, parameters: dict):
    response = requests.get(url, params=parameters)
    return response
