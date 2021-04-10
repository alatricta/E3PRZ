import win32com.client as wcom


def _GetDevicesOnJob(job) -> dict:
    '''Получает список устройств (блоков и сборок) в проекте

        Args:
            job (): Текущий проект

        Returns: (dict)
            {0: [
                id_device_1,
                id_device_2,
                ...]
            id_assembly_1: [
                id_dev_in_assembly_1,
                id_dev_in_assembly_2,
                ...],
            id_assembly_2: [
                ...]
            }
    '''
    dev = job.CreateDeviceObject()
    # Получение списка блоков в проекте
    device_count, devices_ids = job.GetBlockIds()
    # print('Блоков в проекте: ', device_count)  # develop

    # Если есть хотя бы один блок в проекте, тогда едем дальше
    if device_count != 0:
        # Список для устройств и сборок
        devices = {0: []}

        for id in devices_ids[1:]:
            dev.SetId(id)
            ass = dev.GetAssemblyId()
            # print('Блок ID:', id, ' Обозначение:', dev.GetName())  # develop
            # print('Имеет', sym_count, 'символов: ', sym_ids[1:])  # develop

            if ass in devices:
                devices[ass].append(id)
            else:
                devices[ass] = [id]

            # print('Входит в сборку ', ass)  # develop

        # print('Устройства в проекте: ', devices)  # develop
        return devices

    else:  # device_count == 0
        print("Нет блоков в проекте")
        return {}


def _GetSymbolsOnSheets(job, devices: dict):
    '''Получает словарь id символов блоков на выделенных листах

        Args:
            job: текущий проект
            Devices (list): список id устройств, полученный от _GetDevicesOnJob
        Return:
            { id_sheet_1: [
                id_symbol_1,
        !       {id_assembly_1: [
                    id_symbol_in_assembly_1,
                    id_symbol_in_assembly_2,
                    ...
                    ]
                },
                id_symbol_2,
                ...
                ],
            id_sheet_2: [
                ...
                ]
            }
    '''
    dev = job.CreateDeviceObject()
    sym = job.CreateSymbolObject()

    dev.SetId(devices[0][0])
    sym_count, sym_ids = dev.GetSymbolIds()
    sym.SetId(sym_ids[1])

    print(sym.GetSchemaLocation())
    print(job.GetSheetIds())


if __name__ == '__main__':
    app = wcom.Dispatch('CT.Application')
    job = app.CreateJobObject()

    sht = job.CreateSheetObject()
    sym = job.CreateSymbolObject()
    dev = job.CreateDeviceObject()

    devices = _GetDevicesOnJob(job)
    symbols = _GetSymbolsOnSheets(job, devices)

    # Это обязательный параметр для закрытия COM-обекта
    app.quit()


