import win32com.client as wcom

# функции программы


def _GetSymOfDevOnSheets(job):
    '''
    Функция для получения списка id символов блоков на выделенных листах
    из передаваемого проекта Если нет ни одного выделенного листа в проекте,
    осуществляется выход из программы

    job: проект передаваемый функции
    return: {
        id_листа_1: [
            id_символа_1,
            id_символа_2,
            ...
            ],
        id_листа_2: [
            ...
            ]
        }
    '''

    # Получение количества и списка ID выделенных листов
    sht_count, sht_ids = job.GetTreeSelectedSheetIds()
    # Проверяем что есть выделенные листы
    if sht_count > 0:
        # Генерируем словарь
        devices = {k: [] for k in sht_ids[1:]}
        # Перебор по всем листам
        for sht_id in sht_ids[1:]:
            sht.SetId(sht_id)
            # Вначало добавляем имя листа, для сортировки
            devices[sht_id].append(sht.GetName())
            # Получаем количество и список id всех символов на листе
            sym_count, sym_ids = sht.GetSymbolIds()
            # проверяем что есть хотя бы один символ
            if sym_count > 0:
                # Перебор по всем символам
                for sym_id in sym_ids[1:]:
                    sym.SetId(sym_id)
                    # Если символ является блоком, вносим его в список
                    if sym.IsBlock():
                        devices[sht_id].append(sym_id)
                    else:
                        print('Выделите хотя бы один лист.')
                        quit()


def _GetKeyForSortSymbol(sym_id):
    '''
        Функция для получения ключа для сортировки символов.
        Ключом берём кортеж координат X и Y левого нижнего угла символа.
    '''
    sym.SetId(sym_id)
    _, Xmin, Ymin, _, _ = sym.GetPlacedArea()
    # Debug code
    # sym_texts = sym.GetTextIds()[1][1:]
    # txt.SetId(sym_texts[0])
    # print(txt.GetText(), 'ID:', sym_id)
    # print(Xmin, Ymin)
    return (Xmin, Ymin)


def _GetKeyFoSortSheet(sht_id):
    '''
        Функция для получения ключа для сортировки листов.
        Ключом берём наименование листа.
    '''
    sht.SetId(sht_id)
    return sht.GetName()


# Основное тело программы
if __name__ == '__main__':
    # Объявление переменных
    app = wcom.Dispatch('CT.Application')
    job = app.CreateJobObject()
    sht = job.CreateSheetObject()
    sym = job.CreateSymbolObject()
    dev = job.CreateDeviceObject()
    txt = job.CreateTextObject()

    # TODO: 1. Получить структуру данных
    # TODO: 2. Отсортировать структуру
    # TODO: 3. Пройтись по элементам структуры и переименовать по порядку
    # TODO: 4. Пройтись по элементам структуры и расставить надписи

    # Получаем список id всех символов на листах
    _GetSymOfDevOnSheets(job)

    # Это обязательный параметр для закрытия COM-обекта
    app.quit()
