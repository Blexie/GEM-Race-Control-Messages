import sys
import time
import pygsheets
from obswebsocket import obsws, requests
import logging

logging.basicConfig(level=logging.ERROR)

sys.path.append('../')

sheeturl = input("Google sheet URL: ")
worksheet = input("Name of worksheet with messages: ")
obsport = input("OBS Websocket plugin port: (default: 4444)")
obspassword = input("OBS Websocket plugin password: ")
gc = pygsheets.authorize(local=True)
sh = gc.open_by_url(str(sheeturl))
gws = sh.worksheet_by_title(str(worksheet))


host = "localhost"
if obsport == '':
    port = 4444
else:
    port = str(obsport)
password = str(obspassword)

ws = obsws(host, port, password)

while True:
    ws.connect()
    scene = ws.call(requests.GetCurrentScene()).getName()
    if scene != "In-Race":
        continue
    entries = gws.range('A:A')
    try:
        for column in entries:
            for cell in column:
                if cell.value == '' or cell.value == 'DECISION' or gws.cell("B"+str(cell.row)).value == "Broadcasted":
                    continue
                text = cell.value
                ws.call(requests.SetTextGDIPlusProperties(source="Race Control Message", text=text))
                ws.call(requests.SetSceneItemProperties(scene_name="In-Race", item="Race Control Message", visible=True))
                time.sleep(10)
                ws.call(requests.SetSceneItemProperties(item="Race Control Message", visible=False))
                gws.cell("B"+str(cell.row)).value = "Broadcasted"
                time.sleep(3)
            #info=ws.call(requests.GetTextGDIPlusProperties(source="Race Control Message"))
            #print(info)

    except KeyboardInterrupt:
        pass

    ws.disconnect()
    time.sleep(30)
