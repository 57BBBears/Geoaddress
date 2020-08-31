#Gets cordinates by address and create xls file for Yandex map constucture service
#Last modifed 2020-04
import pandas as pd
import requests
import time


def loadDataFile(fileName):
    fileInput = fileName
    fileFormat = fileName[fileName.rfind('.')+1:]

    while fileInput.lower() != 'exit':
        try:
            if fileFormat == 'xls':
                data = pd.read_excel(fileName)
                break
            elif fileFormat == 'csv':
                data = pd.read_csv(fileName)
                break
        except OSError:
            print(f'Ошибка чтения файла. Проверьте, что файл "{fileName}" находится в папке с программой. ')
            fileInput = input('Введите другое название файла с раширением или "Y" для повтора. "Exit" для отмены: ')

            if fileInput.lower() == 'y':
                continue
            elif fileInput.lower() == 'exit':
                raise OSError("Не удалось загрузить файл. ")
            else:
                fileName = fileInput

    return data

def getCords(address='', apikey='', url='https://geocode-maps.yandex.ru/1.x/'):

    if apikey and address:

        try:
            response = requests.get(url, params={'apikey': apikey,
                                                 'format': 'json',
                                                 'results': '1',
                                                 'll': '30.312733,59.940073',
                                                 'geocode': address},
                                    timeout=(2, 2))
            response.raise_for_status()

        except requests.exceptions.Timeout as request_error:
            print(f'Ошибка соеднинения. {request_error}')
            time.sleep(10)
            return getCords(address=address, apikey=apikey, url=url)
        except requests.RequestException as request_error:
            print(f'Ошибка соеднинения. {request_error}')
            time.sleep(10)
            return getCords(address=address, apikey=apikey, url=url)

        json = response.json()
            
        if 'error' in json:
            return json['error']+': ' + json['message']
        else:
            try:
                cords = json['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos']
                return cords

            except Exception as read_error:
                print('Ошибка чтения данных json. '+str(type(read_error).__name__)+': '+str(read_error))
                return 'Ошибка чтения данных json. '+str(type(read_error).__name__)+': '+str(read_error)


    elif not apikey:
        exit('Ошибка. Не указан API ключ ')
    else:
        exit('Ошибка. Не указан адрес для поиска')


def geoAddress():
    print('Geo Address 1.1')

    try:
        addrFile = loadDataFile('address.xls')
    except OSError as error:
        print(error)
        exit('Конец программы. ')

    print('Загружены данные: ')
    print(addrFile)

    print('\nПолучение данных. Может занять некоторое время..')

    if 'geometry_name' in addrFile.columns:
        data = addrFile[['name', 'geometry_name']]

        # Get cordinations instead of address and devide into 2 columns
        if 'city' in addrFile.columns:
            addrCol = addrFile[['city', 'geometry_name']].apply(lambda x: ', '.join(x.astype(str)), axis=1)
        else:
            addrCol = addrFile['geometry_name'].apply(lambda x: str(x))

        cords = addrCol.apply(getCords, apikey='829ef111-8b8c-4dd2-802b-f6dfd6b03327')
        #cords = addrFile['geometry_name'].apply(lambda x: [1, 2])
        #cords = cords.apply(lambda x: x['cords'].split(' ', expand=True))
        #cords.columns = ['Широта', 'Долгота']
        #cords.columns('Широта', 'Долгота')
        #cords = pd.DataFrame(cords)
        #data['cords'] = cords
        altLat = cords.str.split(' ', expand=True)
        data = pd.concat([altLat, data], axis=1)
        data = data[[1, 0, 'geometry_name', 'name']]
        data['Номер метки'] = range(1, len(data.index)+1)
        data.columns = ['Широта', 'Долгота', 'Описание', 'Подпись', 'Номер метки']
        print(data)

        if not data.empty:
            try:
                data2File = 'address_map.xls'
                data.to_excel(data2File, index=False)

                print(f'Готово. Проверьте файл {data2File}')

            except Exception as write_error:
                exit('Ошибка записи в файл. '+str(write_error))
        else:
            print('Данных нет. Что-то пошло не так...')

    else:
        exit('Проверьте наличие столбца "geometry_name" в файле')


geoAddress()

