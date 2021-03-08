# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import win32com.client as wcom

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app = wcom.Dispatch('CT.Application')
    job = app.CreateJobObject()
    sht = job.CreateSheetObject()

    shts=job.GetAllSheetIds()
#    print(shts[1][1])
    for id in shts[1][1:]:
        print(id)
#    app.PutMessage("Привет от Python!")

