"""
GeoAddress v. 1.2

Gets cordination by address and create xlsx file for Yandex map constucture service

Attributes: xls file with 'name' and 'geometry_name' i.g. address columns

Author: @57BBBears

Last modifed 2020-09
"""
import pandas as pd
import requests
import time
import folium as fo
import html


def loadDataFile(fileName):
    fileInput = fileName
    fileFormat = fileName[fileName.rfind('.')+1:]

    while fileInput.lower() != 'exit':
        try:
            if fileFormat == 'xls' or fileFormat == 'xlsx':
                try:
                    data = pd.read_excel(fileName)
                except FileNotFoundError:
                    # if no xlsx file, try xls
                    if fileFormat == 'xlsx':
                        data = pd.read_excel(fileName[:-1])
                        break
                else:
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


def getCords(address='', apikey='', url='https://geocode-maps.yandex.ru/1.x/', timeout=2):
    """
    Get coordinates by address using Yandex map API

    :param str address: City, street, house
    :param str apikey: Yandex map api key
    :param str url: Yandex api service url
    :param int timeout: connection time and default pause between connection problems
    :return: cords for address
    """

    if not address:
        print('Ошибка. Пустой адрес')
        return 'Ошибка. Пустой адрес'

    if apikey and address:
        # try to get data while pause less 36 sec or 7 times
        pause = timeout
        t = 0
        while pause <= 64:

            try:
                response = requests.get(url, params={'apikey': apikey,
                                                     'format': 'json',
                                                     'results': '1',
                                                     'll': '30.312733,59.940073',
                                                     'geocode': address},
                                        timeout=(timeout, timeout))
                response.raise_for_status()

            except requests.exceptions.Timeout as request_error:
                print(f'Ошибка соеднинения. {request_error}')
                if t == 2:
                    timeout += 1
                    t = 0
                t += 1
                time.sleep(pause)
                pause *= 2
            except requests.RequestException as request_error:
                print(f'Ошибка соеднинения. {request_error}')
                time.sleep(pause)
                pause *= 2
            else:
                break
        else:
            print(f'Ошибка. Превышено количество попыток подключения. Адрес: {address} ')
            return 'Ошибка. Превышено количество попыток подключения'

        json = response.json()
        # check if there is a response
        if 'error' in json:
            return json['error'] + ': ' + json['message']
        else:
            try:
                cords = json['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos']
                return cords

            except Exception as read_error:
                print('Ошибка. В чтении данных json. ' + str(type(read_error).__name__) + ': ' + str(read_error))
                return 'Ошибка. В чтении данных json. ' + str(type(read_error).__name__) + ': ' + str(read_error)

    elif not apikey:
        exit('Ошибка. Не указан API ключ ')
    else:
        exit('Ошибка. Не указан адрес для поиска')

def draw_map(data):
    print('Рисуем карту...')

    map = fo.Map()

    markers = fo.FeatureGroup(name='Geo Address')

    data_lat = list(data['Широта'])
    data_long = list(data['Долгота'])
    data_descr = list(data['Описание'])
    data_name = list(data['Подпись'])

    for lat, long, name, descr in zip(data_lat, data_long, data_name, data_descr):
        try:
            lat = float(lat)
            long = float(long)
        except ValueError:
            continue

        markers.add_child(fo.Marker(location=[lat, long], popup=html.escape(str(name).replace('`', "'").replace('\\', '/'))+'<br/>'+html.escape(str(descr).replace('`', "'").replace('\\', '/')), icon=fo.Icon(color='blue')))

    map.add_child(markers)
    map.save('map.html')

    print('Готово. Карта в map.html')


def geoAddress(file: str = 'address.xlsx', apikey: str = '', **kwargs):
    """
    Geo Address 1.2
    :param str file: name of the input file. Dedault: address.xlsx
    :param kwargs: parameters for output file
    :param str city: additional region/city for address (optional)
    :param str sign: format for icon sign on the map (optional)
        Formats: name - name from file, number - numbers 1 - n (default)
    :param str name: format for icon description (optional).
        Formats: name_address - 'name \n address', address - duplicate address, default value from input file
    :return: generates address_map.xml file
    """
    print('Geo Address 1.2')

    if not apikey:
        exit('Укажите API ключ!')

    try:
        addrFile = loadDataFile(file)
        addrFile.rename(str.lower, axis='columns', inplace=True)
    except OSError as error:
        print(error)
        exit('Конец программы. ')

    print('Загружены данные: \n', addrFile)

    print('\nПолучение данных. Может занять некоторое время..')

    if 'geometry_name' in addrFile.columns:
        if 'name' in addrFile.columns:
            data = addrFile.loc[:, ('name', 'geometry_name')]
        else:
            data = pd.DataFrame(addrFile.loc[:, 'geometry_name'])
            data['name'] = ''


        errors = data[data['geometry_name'].isnull()].copy().fillna('')
        errors['error'] = 'Ошибка. Пустой адрес'
        #errors = errors[['error', 'name', 'geometry_name']] if name
        data = data[data['geometry_name'].notnull()].copy().fillna('')

        data_names = data['name']

        # Get name format for output by params
        if 'name' in kwargs:
            if kwargs['name'] == 'name_address':
                data['name'] = pd.Series(map(lambda x, y: str(x)+'\n'+str(y), data['name'], data['geometry_name']))
            elif kwargs['name'] == 'address':
                data['name'] = data['geometry_name']

        # Get coordinates instead of address and devide into 2 columns
        if 'city' in addrFile.columns:
            addrCol = addrFile[['city', 'geometry_name']].fillna('').apply(
                lambda x: kwargs['city']+', '+', '.join(x.astype(str)) if 'city' in kwargs
                else ', '.join(x.astype(str)), axis=1)
        else:
            addrCol = kwargs['city']+', '+addrFile['geometry_name'].fillna('') if 'city' in kwargs else addrFile['geometry_name']

        #filter by empty addresses in data
        addrCol = addrCol.filter(items=data.index, axis=0)
        cords = addrCol.apply(getCords, apikey=apikey)

        data = pd.concat([cords, data], axis=1)
        # Collect errors from data
        errors_con = data[data.iloc[:, 0].str.contains('Ошибка')]
        errors_con.columns = ['error', 'name', 'geometry_name']
        errors = pd.concat([errors, errors_con])
        errors.index += 1

        data = data[~(data.iloc[:, 0].str.contains('Ошибка'))]
        altlat = data.iloc[:, 0].str.split(' ', expand=True)
        data = pd.concat([altlat, data.iloc[:, [1, 2]]], axis=1)


        # get errors if exist
        if errors.empty:
            print('Ура! Ошибок нет.')
        else:
            print('Ошибки:')
            print(errors)

            error_file = 'error.xlsx'
            try:
                errors.to_excel(error_file)
            except Exception as write_error:
                print('Запись ошибок в файл '+error_file+' не удалась. '+str(write_error))
            else:
                print('Проверьте файл '+error_file)

        data = data[[1, 0, 'name', 'geometry_name']]

        if 'sign' not in kwargs or kwargs['sign'] == 'number':
            data['geometry_name'] = range(1, len(data.index) + 1)
        else:
            if kwargs['sign'] == 'name':
                data['geometry_name'] = data_names


        data['Номер метки'] = range(1, len(data.index)+1)
        data.columns = ['Широта', 'Долгота', 'Описание', 'Подпись', 'Номер метки']

        print('\nИтоговые данные:')
        print(data)

        if not data.empty:
            try:
                data2File = 'address_map.xlsx'
                data.to_excel(data2File, index=False)

                print(f'Готово. Проверьте файл {data2File}')
                print('Отобразить на карте? \n'
                      '1 - Да\n'
                      '2 - Нет')
                draw = input(': ')

                if draw == '1':
                    draw_map(data)
                else:
                    exit('Выход.')

            except Exception as write_error:
                exit('Ошибка записи в файл. '+str(write_error))
        else:
            print('Данных нет. Что-то пошло не так...')

    else:
        exit('Проверьте наличие столбца "geometry_name" в файле')


if __name__ == '__main__':
    #geoAddress(city='Санкт-Петербург', name='name_address', sign='number')
    geoAddress(apikey='829ef111-8b8c-4dd2-802b-f6dfd6b03327', name='name_address', sign='number')
