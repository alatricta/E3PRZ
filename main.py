import win32com.client as wcom

app = wcom.Dispatch('CT.Application')
job = app.CreateJobObject()
sht = job.CreateSheetObject()
sym = job.CreateSymbolObject()
dev = job.CreateDeviceObject()

device_count, devices_ids = job.GetBlockIds()
print('Блоков в проекте: ', device_count)
# print('Список ID блоков ', devices_ids[1:])

asss = []
for id in devices_ids[1:]:
	dev.SetId(id)
	ass = dev.GetAssemblyId()
	sym_count, sym_ids = dev.GetSymbolIds()
	print('Блок ID:', id, ' Обозначение:', dev.GetName())
	print('Имеет', sym_count, 'символов: ', sym_ids[1:])
	if ass != 0:

		if asss.count(ass) == 0:
			asss.append(ass)
		print('Входит в сборку ', ass)
	print('===================')

print('Сборки в проекте: ', asss)

# TODO: 




# Это обязательный параметр для закрытия COM-обекта
app.quit()