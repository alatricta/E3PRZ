import win32com.client as wcom
# Объявление глобальных переменных
app = wcom.Dispatch('CT.Application')
job = app.CreateJobObject()
sht = job.CreateSheetObject()
sym = job.CreateSymbolObject()
dev = job.CreateDeviceObject()
txt = job.CreateTextObject()


# функции программы
def _key(symbol_id: int) -> list:
    ''' Функция получения ключа сортировки

        Args:
            symbol_id (int): id символа для получения ключа

        Returns:
            list: [Наименование_листа, координата_X, координата_Y]
    '''
    # sym = job.CreateSymbolObject()
    sym.SetId(symbol_id)
    # value = sym.GetName()
    shm_location = sym.GetSchemaLocation()
    sht.SetId(shm_location[0])
    value = [sht.GetName(), shm_location[1], shm_location[2]]
    # print(value)
    return value


def _GetIdBlocksOnJob() -> list:
    ''' Получение списка id блоков из проекта.
        Список сортируется по положению на листе
    '''
    # Получаем список символов блоков в проекте
    symbols_ids = job.GetSymbolIds()
    block_ids = []
    # делаем новый список только с блоками
    for id in symbols_ids[1][1:]:
        sym.SetId(id)
        if sym.IsBlock():
            block_ids.append(id)

    # сортируем по наименованию листа и положению на нём
    block_ids.sort(key=_key)
    return block_ids


# Основное тело программы
block_ids = _GetIdBlocksOnJob()
name_ids = []
txts_ids = []
for id in block_ids:
    sym.SetId(id)
    txts = sym.GetTextIds()
    txt.SetId(txts[1][1])
    name_ids.append(txt.GetText())
    txts_ids.append(txts[1][1])

print(name_ids)
print(txts_ids)
print(job.GetDeviceNameSeparator())

# txt.SetId(25544)
# txt.SetText('erq')
# Это обязательный параметр для закрытия COM-обекта
app.quit()
