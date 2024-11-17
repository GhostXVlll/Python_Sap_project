import subprocess
import win32com.client
import time
import win32con
import win32gui
from datetime import datetime
import sys

# Путь к SAP GUI
sap_gui_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"

# Параметры для подключения к системе
sap_logon_parameters = [
    '-system=CR1',
    '-client=001',
    '-user=DDIC',
    '-pw=neMzcjWPp_6',
    '-guirunmode=1',  # Этот параметр отключает графический интерфейс
]

# Запуск SAP GUI
subprocess.Popen([sap_gui_path] + sap_logon_parameters)

# Подождать некоторое время, чтобы SAP GUI успел запуститься
time.sleep(15)

# Подключаемся к объекту SAP GUI
sap_gui = win32com.client.Dispatch("SAPGUI.ScriptingCtrl.1")
application = sap_gui.GetScriptingEngine()

# Получаем активное соединение
connection = application.OpenConnection('CR1', True)
session = connection.Children(0)

# Выполняем действия в SAP
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "sbook"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtI3-LOW").text = startdate
session.findById("wnd[0]/usr/ctxtI3-HIGH").text = enddate
session.findById("wnd[0]/usr/ctxtI3-HIGH").setFocus()
session.findById("wnd[0]/usr/ctxtI3-HIGH").caretPosition = 10
session.findById("wnd[0]").sendVKey(8)
session.findById("wnd[0]/tbar[1]/btn[43]").press()
session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderdir
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Завершаем работу
session.findById("wnd[0]").close()

# Завершить работу
sys.exit(0)