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

    # Получение количества и списка ID выделенных листов
    shts = job.GetTreeSelectedSheetIds()

    if shts[0] > 0:
        # TODO: получить словарь {id_символа : ( (Х_мин, Y_мин, Х_макс, Y_макс), id_сборки)}
        '''
        id_символа - id символа блока на листе
        Х_мин - координата левого края блока
        Y_мин - координата нижнего края блока
        Х_макс - координата правого края блока
        Y_макс - координата верхнего края блока   
        id_сборки - id сборки, в которую входит блок. Если блок в сборку не входит, то 0
        devices - описанный словарь
        coord - список координат, который будет внутри словаря
        '''
        devices = {}

        for sht_id in shts[1][1:]:
            sht.SetId(sht_id)
            sym_ids = sht.GetSymbolIds()    # Получаем количество и список id всех символов на листе
            if sym_ids[0] > 0:
                for sym_id in sym_ids[1][1:]:
                    sym.SetId(sym_id)
                    # вносим в словарь только символы блоков 
                    if sym.IsBlock():
                        coord = sym.GetArea()
                        devices[sym_id] = (coord[1:4], 0)
    else:
        print('Выделите хотя бы один лист.')

    #print(block_ids)
    #print('Количество блоков на листе:', len(block_ids))
