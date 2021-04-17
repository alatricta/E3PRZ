import win32com.client as wcom
# Объявление глобальных переменных
app = wcom.Dispatch('CT.Application')
job = app.CreateJobObject()
sht = job.CreateSheetObject()
sym = job.CreateSymbolObject()
dev = job.CreateDeviceObject()
txt = job.CreateTextObject()


# функции программы
def _GetKeyForSortSymbols(symbol_id: int) -> tuple:
    '''Функция получения ключа сортировки

        Args:
            symbol_id (int): id символа для получения ключа

        Returns:
            tuple: (Наименование_листа, координата_X, координата_Y)
    '''
    # sym = job.CreateSymbolObject()
    sym.SetId(symbol_id)
    # value = sym.GetName()
    shm_location = sym.GetSchemaLocation()
    sht.SetId(shm_location[0])
    value = (sht.GetName(), shm_location[1], shm_location[2])
    # print(value)
    return value


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


def _GetKeyForSortDevices(device_item) -> tuple:
    '''Получение ключа для сортировки устройств

        Args:
            device_item (): id устройства для сортировки
        
        Returns:
            tuple: (Наименование_листа, координата_X, координата_Y)
    '''
    # проверяем что на входе и присваиваем id устройства
    if type(device_item) == tuple:
        dev.SetId(device_item[0])
    elif type(device_item) == int:
        dev.SetId(device_item)
    else:
        print('Прилетела какая-то непонятная фигня!')
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


def _RenameList(list_devices, designation_label="А", designation_position=1):
    '''Переименование сортированного списка

        Args:
            list_devices: список полученный после _GetList
            designation_label (str, optional): Буквенная часть позиционного обозначения
            designation_position (int, optional): Начальная позиция
    '''
    for id in list_devices:
        # проверяем что на входе и присваиваем id устройства
        if type(id) == tuple:
            dev.SetId(id[0])
            _RenameList(id[1], designation_label=f'{designation_label}{designation_position}-{designation_label}')
        elif type(id) == int:
            dev.SetId(id)
        else:
            print('Прилетела какая-то непонятная фигня!')

        dev.SetName(f'{designation_label}{designation_position}')
        designation_position += 1


def _TextPlaycementDev(device_id):
    '''Получаем id устройства переходим к символам и расставляем текст

    Args:
        device_id: Description
    '''
    pass


def _TextPlaycementSym(symbol_id):
    '''Получаем id символа и расставляем его тексты
    
    Args:
        symbol_id (): Description
    '''
    pass


# # Основное тело программы
if __name__ == '__main__':
    # получаем список блоков
    block_ids = _GetDevicesOnJob()
    block_ids = _GetList(block_ids)

    print(f'Список  до  сортивровки:  {block_ids}')
    block_ids.sort(key=_GetKeyForSortDevices)
    print(f'Список после сортивровки: {block_ids}')
    _RenameList(block_ids)
    # dev.SetId(22922)
    # print(dev.GetName())

    # Это обязательный параметр для закрытия COM-обекта
    app.quit()
