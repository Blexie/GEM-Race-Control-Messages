import sys
import time
import pygsheets
from obswebsocket import obsws, requests
import logging
import config

logging.basicConfig(level=logging.ERROR)

sys.path.append('../')

sheeturl = config.sheeturl
worksheet = config.worksheet
obsport = config.obsport
obspassword = config.obspassword
gc = pygsheets.authorize(local=True)
sh = gc.open_by_url(str(sheeturl))
gws = sh.worksheet_by_title(str(worksheet))
print("Race Control Alerts running...")

host = "localhost"

ws = obsws(host, obsport, obspassword)

while True:
    ws.connect()
    scene = ws.call(requests.GetCurrentScene()).getName()
    if scene == "In-Race":
        entries = gws.range('A:A')
        try:
            for column in entries:
                for cell in column:
                    if cell.value == '' or cell.value == 'DECISION' or gws.cell("B"+str(cell.row)).value == "Broadcasted":
                        continue
                    text = cell.value
                    ws.call(requests.SetTextGDIPlusProperties(source="Race Control Message", text=text))
                    ws.call(requests.SetSceneItemProperties(scene_name="In-Race", item="Race Control Message", visible=True))
                    time.sleep(60)
                    ws.call(requests.SetSceneItemProperties(item="Race Control Message", visible=False))
                    gws.cell("B"+str(cell.row)).value = "Broadcasted"
                    time.sleep(10)

        except KeyboardInterrupt:
            pass
    else:
        pass

    ws.disconnect()
    time.sleep(30)
