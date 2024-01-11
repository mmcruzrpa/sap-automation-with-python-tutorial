"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

import win32com.client
import win32gui
import win32con

import pandas as pd

import re

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = DesktopBot()

    credential = maestro.get_credential(label="SAP", key="SAP")

    sap_logon_file_path = r"path to sap logon executable"
    excel_file_path = r"path to the excel file"

    # Implement here your logic...
    bot.execute(sap_logon_file_path)
    bot.wait(10000)

    sapGui = win32com.client.GetObject("SAPGUI")
    application = sapGui.GetScriptingEngine
    connection = application.OpenConnection("SAP", True)
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize

    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = 800
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "MCRUZ"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = credential
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[0]/okcd").text = "me21n"
    session.findById("wnd[0]").sendVKey(0)

    df = pd.read_excel(excel_file_path)

    for index,row in df.iterrows():

        if index == 0 :
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = row['Vendor']
        
        session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,{index}]").text = row['Material']
        session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,{index}]").text = row['PO Quantity']
        session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,{index}]").text = row['Net Price']

    session.findById("wnd[0]").sendVKey(11)
    poCode = session.findById("wnd[0]/sbar").text
    poCode = re.search("\d+", poCode).group()

    df['Purchase Order Code'] = poCode
    df.to_excel(excel_file_path, index = False)

    #Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
         task_id=execution.task_id,
         status=AutomationTaskFinishStatus.SUCCESS,
         message="Task Finished OK."
    )

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()