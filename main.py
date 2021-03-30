import win32com.client as wcom
    
# функции программы
def _GetSymOfDevOnSheets(job): 
    '''
    Функция для получения списка id символов блоков на выделенных листах из передаваемого проекта
    Если нет ни одного выделенного листа в проекте, осуществляется выход из программы

    job: проект передаваемый функции
    return: { id_листа_1: ['наименование листа 1', id_символа_1, id_символа_2, ...], id_листа_2: [...] }
    '''

    # Получение количества и списка ID выделенных листов
    sht_count, sht_ids = job.GetTreeSelectedSheetIds()
    # Проверяем что есть выделенные листы
    if sht_count > 0:
        # Генерируем словарь 
        devices = {k:[] for k in sht_ids[1:]}

        # Перебор по всем листам 
        for sht_id in sht_ids[1:]:
            sht.SetId(sht_id)
            # Первой позицией добавляем имя листа, для использования его в сортировке
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


def _GetKeyForSort(sym_id):
    sym.SetId(sym_id)
    _, Xmin, Ymin, Xmax, Ymax = sym.GetPlacedArea()
    sym_texts = sym.GetTextIds()[1][1:]
    txt.SetId(sym_texts[0])
    print(txt.GetText(), 'ID:', sym_id)
    print(Xmin, Ymin, Xmax, Ymax)
    return (Xmin, Ymin)



# Основное тело программы
if __name__ == '__main__':
    # Объявление переменных
    app = wcom.Dispatch('CT.Application')
    job = app.CreateJobObject()
    sht = job.CreateSheetObject()
    sym = job.CreateSymbolObject()
    dev = job.CreateDeviceObject()
    txt = job.CreateTextObject()
    
    # Получаем список id всех символов на листах
    _GetSymOfDevOnSheets(job)

    


                    #     sym_texts = sym.GetTextIds()[1][1:]
                    #     for id_txt in sym_texts:
                    #         txt.SetId(id_txt)
                    #         print('ID текста:', id_txt, ' Текст: ', txt.GetText())
                    #     print('========================')

                    #     for id_txt in sym_texts:
                    #         txt.SetId(id_txt)
                    #         txt.SetText(f'{txt.GetText()}__')
                    #         print('ID текста:', id_txt, ' Новый текст: ', txt.GetText())
                    #     print('^^^^^^^^^^^^^^^^^^^^^^^^')
                        # print(sym_texts[0])
                        # _, Xmin, Ymin, Xmax, Ymax = sym.GetPlacedArea()
                        # print(Xmin, Ymin, Xmax, Ymax)
                        # devices[sym_id] = (coord[1:4], 0)
                        

    # Вывод текущего списка 
    # print(devices)
    # Сортировка списка листов по названию
    devices = {k: devices[k] for k in sorted(devices, key=lambda x: devices[x][0])}
    # print(devices)
    
    # TODO: можно попробовать отсортировать списки символов на листах по положению Х
    # А что это мне даст ?
    # А как потом мне отсортировать по Y? 

    for sym_ids in devices.values():
        # for sym_id in sym_ids[1:]:
        sort_dev = sorted(sym_ids[1:], key=lambda x: GetKeyForSort(x))
        print('++++++++++++++')
        print(sort_dev)
        print('==============')
            
    
    # Это обязательный параметр для закрытия COM-обекта
    app.quit()