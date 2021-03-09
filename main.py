# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import win32com.client as wcom

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app = wcom.Dispatch('CT.Application')
    job = app.CreateJobObject()
    sht = job.CreateSheetObject()
    sym = job.CreateSymbolObject()

    shts = job.GetTreeSelectedSheetIds()    #Получение списка ID выделенных листов
#    shts = job.GetActiveSheetId()    #Получение ID активного окна

    if shts[0] > 0:
        for id in shts[1][1:]:
            #Здесь надо получить список символов блоков со всех листов и скинуть всё в один список
            sht.SetId(id)
            sym_ids = sht.GetSymbolIds()
            if sym_ids[0] > 0:
                block_ids = []  # Создал пустой список ID блоков
                for sym_id in sym_ids[1][1:]:
                    sym.SetId(sym_id)
                    if sym.IsBlock():
                        block_ids.append(sym_id)

    print(block_ids)
    print('Количество блоков на листе:', len(block_ids))

