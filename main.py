import win32com.client as wcom
# Объявление глобальных переменных
app = wcom.Dispatch('CT.Application')
job = app.CreateJobObject()
sht = job.CreateSheetObject()
sym = job.CreateSymbolObject()
dev = job.CreateDeviceObject()
txt = job.CreateTextObject()

# Тип надписи позиционного обозначения, заданный в Е3
txt_poz_type = 212
# Шаг положения символов на листе, который учитывается при сортировке символов, в мм
step_placement = 10


# функции программы
def _GetDevicesOnJob() -> dict:
    '''Получение словаря id устройств

        Returns:
            dict: {
                {id_device_1: 0},
                {id_device_2: 0},
                {id_assembly_3: [
                    id_device_3_1,
                    id_device_3_2,
                    ...]
                    },
                {id_device_4: 0}
        }
    '''
    # Получение списка блоков в проекте
    device_count, devices_ids = job.GetBlockIds()
    # print('Блоков в проекте: ', device_count)  # Debug info

    # если в проекте нет блоков
    if device_count == 0:
        print("Проект не содержит ни одного блока")
        return {}

    # Если есть хотя бы один блок в проекте
    else:
        devices = {}
        for id in devices_ids[1:]:
            dev.SetId(id)
            assembly_id = dev.GetRootAssemblyId()

            # если просто блок
            if assembly_id == 0:
                devices[id] = 0
            # если сборка уже существует в списке
            elif assembly_id in devices:
                devices[assembly_id].append(id)
            # если сборка ещё не добавлена в список
            else:
                devices[assembly_id] = []
                devices[assembly_id].append(id)

        return devices


def _GetList(dict_devices):
    dev = []
    for id, devices in dict_devices.items():
        if devices == 0:
            # print(id)  # Debug info
            dev.append(id)
        else:
            # print(id, devices)  # Debug info
            ass = (id, devices)
            dev.append(ass)
    # print(dev)  # Debug info
    return dev


def _GetKeyForSortSymbols(symbol_id: int) -> tuple:
    '''Функция получения ключа сортировки

        Args:
            symbol_id (int): id символа для получения ключа

        Returns:
            tuple: (Наименование_листа, координата_X, координата_Y)
    '''
    # проверяем что на входе и присваиваем id устройства
    if type(symbol_id) == int:
        sym.SetId(symbol_id)
    elif symbol_id is None:
        return ('!',)
    else:
        print('Прилетела какая-то непонятная херня!')
        return None

    # sym = job.CreateSymbolObject()

    # value = sym.GetName()
    sht_id, sym_X, sym_Y = sym.GetSchemaLocation()[0:3]
    sht.SetId(sht_id)
    value = (sht.GetName(), sym_X, sym_Y)
    # print(value)
    return value


def _GetKeyForSortDevices(device_item) -> tuple:
    '''Получение ключа для сортировки устройств

        Args:
            device_item (): id устройства для сортировки

        Returns:
            tuple: (Наименование_листа, координата_X, координата_Y)
    '''
    # проверяем что на входе и присваиваем id устройства
    # если сборка
    if type(device_item) == tuple:
        dev.SetId(device_item[0])
    # если устройство
    elif type(device_item) == int:
        dev.SetId(device_item)
    # если не понятно что
    else:
        print('Прилетела какая-то непонятная херня!')
        return None

    # если это сборка
    if dev.IsAssembly():
        # сортируем устройства в сборке
        device_item[1].sort(key=_GetKeyForSortDevices)
        # возвращаем элемент сортировки для первого устройства в
        #  отсортированной сборке
        return _GetKeyForSortDevices(device_item[1][0])

    # если это просто устройство
    else:
        symbols_count, symbols_ids = dev.GetSymbolIds()
        # если устройство не имеет символов
        if symbols_count == 0:
            # print(f'Устройство {dev.GetName()} не имеет символов')  # Debug info
            return ('z',)

        # если в устройстве всего 1 символ
        elif symbols_count == 1:
            return _GetKeyForSortSymbols(symbols_ids[1])

        # если в устройстве есть несколько символов
        else:
            symbols_ids = list(symbols_ids[1:])
            symbols_ids.sort(key=_GetKeyForSortSymbols)
            # print(_GetKeyForSortSymbols(symbols_ids[0]))  # Debug info
            return _GetKeyForSortSymbols(symbols_ids[0])


def _SortByPlacementWithStep(list_devices: list):
    '''Сортировка списка устройств с учётом шага на схеме (выполнять после сортировки с ключом _GetKeyForSortDevices)

        Args:
            list_devices (list): сортированный список устройств
    '''
    def _GetDev(id):
        '''Возвращает устройство в зависимости от переданного элемента'''
        # если сборка
        if type(id) == tuple:
            return dev.SetId(id[1][0])
        # если устройство
        elif type(id) == int:
            return dev.SetId(id)
        # если не понятно что
        else:
            print('Прилетела какая-то непонятная херня!')
            return None

    # dev1 = job.CreateDeviceObject()
    # dev2 = job.CreateDeviceObject()
    for index1 in range(len(list_devices[0:-1])):
        index2 = index1 + 1
        # print("===============")  # Debug info
        # print("Index1:", index1)  # Debug info
        # print("Index2:", index2)  # Debug info
        dev1_id = _GetDev(list_devices[index1])
        dev2_id = _GetDev(list_devices[index2])

        sort1 = _GetKeyForSortDevices(dev1_id)
        sort2 = _GetKeyForSortDevices(dev2_id)
        # print("Sort1:", sort1)  # Debug info
        # print("Sort2:", sort2)  # Debug info

        # если символы на одном листе
        if sort1[0] == sort2[0]:
            if sort1 == sort2:
                continue
            elif (sort1[1] + step_placement) > sort2[1]:
                ''' Варианты расстановки символов
                    Вариант 1
                    г==========¬
                    ¦   sym1   ¦
                    ¦          ¦
                    ¦          ¦
                    L==========˩
                    |<-sort1
                    |   |
                    |   |<-sort1+step
                    |   |
                    | |<-sort2
                    | г==========¬
                    | ¦   sym2   ¦
                    | ¦          ¦
                    | ¦          ¦
                    | L==========˩

                    Вариант 2
                    | г==========¬
                    | ¦   sym2   ¦
                    | ¦          ¦
                    | ¦          ¦
                    | L==========˩
                    | |<-sort2
                    |   |
                    |   |<-sort1+step
                    |   |
                    |<-sort1
                    г==========¬
                    ¦   sym1   ¦
                    ¦          ¦
                    ¦          ¦
                    L==========˩
                '''
                # если вариант 1, то оставляем как есть
                if sort1[2] > sort2[2]:
                    continue
                # если вариант 2, то меняем местами
                else:
                    list_devices[index1], list_devices[index2] = list_devices[index2], list_devices[index1]


def _RenameList(list_devices: list, designation_label="А", designation_position=1):
    '''Переименование сортированного списка

        Args:
            list_devices (list): список полученный после _GetList
            designation_label (str, optional): Буквенная часть позиционного обозначения
            designation_position (int, optional): Начальная позиция
    '''
    # проверяем что на входе и присваиваем id устройства
    for id in list_devices:
        # если сборка
        if type(id) == tuple:
            _RenameList(id[1], designation_label=f'{designation_label}{designation_position}-{designation_label}')
            dev.SetId(id[0])
        # если устройство
        elif type(id) == int:
            dev.SetId(id)
        # если не понятно что
        else:
            print('Прилетела какая-то непонятная херня!')
            continue

        dev.SetName(f'{designation_label}{designation_position}')
        designation_position += 1


def _TextPlaycementDev(list_devices: list):
    '''Получаем список id устройств переходим к символам и расставляем текст

        Args:
            list_devices: список полученный после _GetList
    '''
    # проверяем что на входе и присваиваем id устройства
    for dev_id in list_devices:
        # если сборка
        if type(dev_id) == tuple:
            _TextPlaycementDev(dev_id[1])
            continue
        # если устройство
        elif type(dev_id) == int:
            dev.SetId(dev_id)
        # если не понятно что
        else:
            print('Прилетела какая-то непонятная херня!')
            continue

        symbols_count, symbols_ids = dev.GetSymbolIds()
        # проверяем наличие символов
        if symbols_count == 0:
            # print('Нечего расставлять, Устройство не имеет изображений.')  # Debug info
            continue

        else:
            # Сортируем список символов
            symbols_ids = tuple(sorted(symbols_ids, key=_GetKeyForSortSymbols))

            # Перебираем отсортированный список
            for symbol_id in symbols_ids[1:]:
                sym.SetId(symbol_id)
                # Координаты верхней правой точки символа
                sym_X_max, sym_Y_max = sym.GetPlacedArea()[3:]
                # ID текста с типом Позиционного обозначения, заданного в Е3
                txt_id = sym.GetTextIds(None, txt_poz_type)[1][1]
                txt.SetId(txt_id)

                # Выравнивание по правому краю текста, чтобы не зависеть от длинны надписи
                txt.SetAlignment(3)
                # поворот позиционного обозначения всегда на 0
                txt.SetRotation(0.0)
                # цвет текста всегда черный
                txt.SetColour(0)
                # Координаты установки основного текста
                txt_X, txt_Y = sym_X_max, sym_Y_max + 2

                # если символов несколько, формируем доп текст и смещение основной надписи
                if symbols_count > 1:
                    att = job.CreateAttributeObject()
                    att_id = sym.SetAttributeValue('saberParam1', f'.{symbols_ids.index(symbol_id)}')
                    att.SetId(att_id)
                    # Создание дополнительного текста и настройка его вида
                    txt_att = job.CreateTextObject()
                    txt_att_id = att.DisplayAttribute()
                    txt_att.SetId(txt_att_id)
                    txt_att.SetRotation(0.0)                    # поворот
                    txt_att.SetStyle(txt.GetStyle())            # стиль
                    txt_att.SetAlignment(1)                     # выравнивание влево
                    txt_att.SetFontName(txt.GetFontName())      # наименования шрифта
                    txt_att.SetHeight(txt.GetHeight())          # высоты шрифта
                    txt_att.SetColour(txt.GetColour())          # цвета шрифта
                    txt_att.SetVisibility(txt.GetVisibility())  # видимость текста
                    # Задание координат Поз. обозначения с учётом доп.надписи
                    txt_X -= txt_att.GetWidth()
                    txt_att.SetSchemaLocation(txt_X, txt_Y)

                # Устанавливаем позиционное обозначение вверх вправо (с учётомтом наличия доп текста)
                txt.SetSchemaLocation(txt_X, txt_Y)
        # todo: надо получить символы соединителей и расставить текст у них
    # job.UpdateAllSymbols()


# # Основное тело программы
if __name__ == '__main__':
    # получаем список блоков
    block_ids = _GetDevicesOnJob()
    block_ids = _GetList(block_ids)

    # print(f'Список  до  сортивровки:    {block_ids}')  # Debug info
    block_ids.sort(key=_GetKeyForSortDevices)
    # print(f'Список после сортивровки:   {block_ids}')  # Debug info
    _SortByPlacementWithStep(block_ids)
    # print(f'Список после 2 сортивровки: {block_ids}')  # Debug info
    _RenameList(block_ids)
    _TextPlaycementDev(block_ids)

    # Это обязательный параметр для закрытия COM-обекта
    app.quit()
