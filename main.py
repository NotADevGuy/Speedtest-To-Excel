import os
import xlsxwriter
import PySimpleGUI as sg
from datetime import datetime


def mainGUI():
    sg.theme("LightGray1")
    layout = [
        [sg.Text("Data Folder Path", font="Any 20", justification='center')],
        [sg.Input(key="dataPath", size=(50, 0), pad=((0, 0), (0, 0))), sg.FolderBrowse("Browse", font="Any 10", pad=((5, 0), (0, 0)))],

        [sg.Text("Excel File Name", font="Any 20", justification='center')],
        [sg.InputText(key="excelName", size=(60, 2))],

        [sg.Text("Save Location", font="Any 20", justification='center')],
        [sg.Input(str(os.getcwd()), key="savePath", pad=((5, 0), (0, 25))), sg.FolderBrowse("Browse", font="Any 10", pad=((5, 0), (0, 25)))],

        [sg.Button("RUN", size=(6, 2), pad=(10, 0)), sg.Button("Exit", key="Exit", size=(6, 2), pad=(10, 0))],
        [sg.Text(key="done", font="Any 20", justification='center')]
    ]

    jusCol = [
        [sg.Column(layout, element_justification='center')]
    ]
    mainWindow = sg.Window("SpeedTest Data Processor", jusCol)#, size=(800, 300), grab_anywhere=False)

    while True:
        event, values = mainWindow.read()
        print(values["excelName"][-5:])
        if event == "RUN":
            passed = True

            if os.path.exists(values["dataPath"]) == False:
                mainWindow["dataPath"].update("Path Doesn't Exist")
                passed = False

            if values["excelName"] == "":
                mainWindow["excelName"].update("defaultName.xlsx")
                passed = False  # Because this is still a technically allowed name.
            elif values["excelName"][-5:] != ".xlsx":
                values["excelName"] += ".xlsx"
            # elif ".xlsx" not in values["excelName"]: # This is old way haha
            #     values["excelName"] += ".xlsx"

            if os.path.exists(values["savePath"]) == False:
                mainWindow["savePath"].update(os.getcwd())
                print("WAS HERE")
                passed = False  # Yet again this won't stop the program, since it has a default

            if passed:
                main(values["dataPath"], values["excelName"], values["savePath"])
                mainWindow["done"].update("Done! " + values["excelName"] + " generated!")
                # mainWindow["done"].update("Done! " + values["excelName"] + ".xlsx generated!")  # OG one lol (double .xlsx extension on default.

        if event in (sg.WIN_CLOSED, "Exit"):
            mainWindow.close()
            break


def main(dataPath, excelName, savePath):
    # print(dataPath + "\n" + excelName + "\n" + savePath + "\n\n")
    renameFiles(dataPath)
    # if ".xlsx" not in excelName:
    #     excelName += ".xlsx"

    wb = xlsxwriter.Workbook(savePath + "\\" + excelName)
    ws = wb.add_worksheet()

    times = getTimesList(dataPath)

    printData(dataPath, times, wb, ws)

    ws.set_column(0, 0, 14)
    ws.set_column(1, 1, 16)
    ws.set_column(2, len(times) + 2, 11)
    wb.close()
    return 1


def renameFiles(dataPath):
    for root, dirs, files in os.walk(dataPath):
        if len(files) > 0:
            for f in files:
                hours, minutes = f[15:17], int(f[17:19])
                if 0 <= minutes <= 14:
                    minutes = "00"
                elif 15 <= minutes <= 44:
                    minutes = "30"
                elif 45 <= minutes <= 59:
                    minutes, hours = "00", int(hours)
                    hours += 1
                    if hours < 23:
                        hours = ("0" + str(hours)) if hours < 10 else str(hours)
                    elif hours == 24:
                        hours = "00"
                tmpS = f[:15] + hours + minutes + f[-4:]
                if not os.path.exists(root + "\\" + tmpS):
                    os.rename((root + "\\" + f), (root + "\\" + tmpS))
    # print("\nFiles should be sorted!\n")


def getTimesList(dataPath):
    times = []
    for files in os.walk(dataPath):
        if len(files[2]) > 0:
            for f in files[2]:
                dateTime = stripDateTime(f)
                if dateTime not in times:
                    times.append(dateTime)
    times.sort()
    return times


def stripDateTime(fileName):
    t = fileName.split("_")[1].split('.')[0]
    return datetime.strptime(t[0:2] + " " + t[2:4] + " " + t[5:7] + " " + t[-2:] + " 2021", "%m %d %H %M %Y")


def printData(dataPath, times, wb, ws):
    dateFormat = wb.add_format({'num_format': 'mm/dd hh:mm'})
    ws.write(0, 1, "Date / Time:")
    row, col = 1, 2
    for files in os.walk(dataPath):
        # print(files)
        if len(files[1]) > 0:
            for system in files[1]:
                ws.write(row, 0, system)
                row += 3
            row = 1

        # elif len(files[2]) > 0:
        else:
            # print(files[0].split("\\")[1])
            ws.write(row, 1, "Download Speed:")
            ws.write(row + 1, 1, "Latency:")
            i = 0
            while i < len(files[2]):
                dateTime = stripDateTime(files[2][i])
                if dateTime == times[col - 2]:
                    dlSpeed, latencySpeed = trim(files[0], files[2][i])
                    ws.write(row, col, dlSpeed)
                    ws.write(row + 1, col, latencySpeed)
                    i += 1
                ws.write_datetime(0, col, times[col - 2], dateFormat)
                col += 1
            col = 2
            row += 3


def trim(dataPath, fileName):
    dlSpeed, latencySpeed = '', ''
    if ".txt" not in fileName:
        return '', ''
    with open(str(dataPath + "\\" + fileName), 'r', encoding='utf-16-le') as file:
        for line in file:
            if "Latency:" in line and "ms" in line:
                latencySpeed = float((line.replace(" ", "")).split(":")[1].split("ms")[0])
            if "Download:" in line and "Mbps" in line:
                dlSpeed = float((line.replace(" ", "")).split(":")[1].split("M")[0])
            if latencySpeed != '' and dlSpeed != '':
                break
    return dlSpeed, latencySpeed


if __name__ == '__main__':
    mainGUI()
#
# What to do if folder is empty?
#
