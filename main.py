# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import win32com.client as wcom

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app = wcom.Dispatch('CT.Application')
    job = app.CreateJobObject()
    sht = job.CreateSheetObject()

    app.PutMessage("Привет от Python!")
