def sap_login(environment):
    import win32com.client
    import pythoncom
    import subprocess
    import ctypes
    import sys
    import time
    try:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(10)

        # Necesario par correr win32com.client en Threading
        pythoncom.CoInitialize()

        sapmin = ctypes.windll.user32.FindWindowW(None, "SAP Logon 760")
        ctypes.windll.user32.ShowWindow(sapmin, 6)

        sapmin = ctypes.windll.user32.FindWindowW(None, "SAP Logon 740")
        ctypes.windll.user32.ShowWindow(sapmin, 6)

        sap_gui_auto = win32com.client.GetObject('SAPGUI')
        if not type(sap_gui_auto) == win32com.client.CDispatch:
            return

        application = sap_gui_auto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui_auto = None
            return

        if environment == "Q02":
            connection = application.OpenConnection("Q02 - Quality", True)
        if environment == "P02":
            connection = application.OpenConnection("P02 - Production", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui_auto = None
            return

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "5210almacen1"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "5210almacen01"
        session.findById("wnd[0]").sendVKey(0)

    except:
        print(sys.exc_info())
    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None


def terminate():
    import win32com.client
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    if not type(sap_gui_auto) == win32com.client.CDispatch:
        return

    application = sap_gui_auto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sap_gui_auto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        sap_gui_auto = None
        return

    if connection.DisabledByServer:
        application = None
        sap_gui_auto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        sap_gui_auto = None
        return

    if session.Info.IsLowSpeedConnection:
        connection = None
        application = None
        sap_gui_auto = None
        return
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


def lt01_query_withdraw(storage_location, material, quantity, storage_type, storage_bin, storage_unit, children):
    import win32com.client
    import re
    import json
    import time
    import pythoncom
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not type(sap_gui_auto) == win32com.client.CDispatch:
            return

        application = sap_gui_auto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui_auto = None
            return

        connection = application.Children(children)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            sap_gui_auto = None
            return

        if connection.DisabledByServer:
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui_auto = None
            return

        if session.Info.IsLowSpeedConnection:
            connection = None
            application = None
            sap_gui_auto = None
            return

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nLT01"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").text = "521"
        session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "998"
        session.findById("wnd[0]/usr/ctxtLTAP-MATNR").text = material
        session.findById("wnd[0]/usr/txtRL03T-ANFME").text = quantity
        session.findById("wnd[0]/usr/ctxtLTAP-WERKS").text = "5210"
        session.findById("wnd[0]/usr/ctxtLTAP-LGORT").text = f'00{storage_location}'
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLTAP-VLENR").text = storage_unit
        session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").text = storage_type
        session.findById("wnd[0]/usr/ctxtLTAP-NLBER").text = "001"
        session.findById("wnd[0]/usr/txtLTAP-NLPLA").text = storage_bin
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        sap_message = session.findById("wnd[0]/sbar/pane[0]").Text
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(.005)
        try:
            # Verify if Transfer order in message if not stk.end error state
            int(re.sub(r"\D", "", sap_message, 0))
            # capture(sap_message)
            response = {"result": sap_message, "error": "N/A"}
            return json.dumps(response)
        except:
            response = {"result": "N/A", "error": sap_message}
            return json.dumps(response)

    except Exception as e:
        print(e)
        time.sleep(.005)
        sap_error = session.findById("wnd[0]/sbar/pane[0]").Text
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        response = {"result": "N/A", "error": sap_error}
        return json.dumps(response)
    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None
        time.sleep(.005)


def lt01_query_insert(storage_location, material, quantity, storage_type, storage_bin, storage_unit, destination_storage_type, destination_storage_bin, children):
    import win32com.client
    import re
    import json
    import time
    import pythoncom
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not type(sap_gui_auto) == win32com.client.CDispatch:
            return

        application = sap_gui_auto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui_auto = None
            return

        connection = application.Children(children)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            sap_gui_auto = None
            return

        if connection.DisabledByServer:
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui_auto = None
            return

        if session.Info.IsLowSpeedConnection:
            connection = None
            application = None
            sap_gui_auto = None
            return

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nLT01"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").text = "521"
        session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "998"
        session.findById("wnd[0]/usr/ctxtLTAP-MATNR").text = material
        session.findById("wnd[0]/usr/txtRL03T-ANFME").text = quantity
        session.findById("wnd[0]/usr/ctxtLTAP-ALTME").text = ""
        session.findById("wnd[0]/usr/ctxtLTAP-WERKS").text = "5210"
        session.findById("wnd[0]/usr/ctxtLTAP-LGORT").text = f'00{storage_location}'
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLTAP-LETYP").text = "001"
        session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").text = storage_type
        session.findById("wnd[0]/usr/ctxtLTAP-VLBER").text = "001"
        session.findById("wnd[0]/usr/txtLTAP-VLPLA").text = storage_bin
        session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").text = destination_storage_type
        session.findById("wnd[0]/usr/ctxtLTAP-NLBER").text = "001"
        session.findById("wnd[0]/usr/txtLTAP-NLPLA").text = destination_storage_bin
        session.findById("wnd[0]/usr/ctxtLTAP-NLENR").text = storage_unit
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)

        sap_message = session.findById("wnd[0]/sbar/pane[0]").Text
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(.005)
        try:
            # Verify if Transfer order in message if not stk.end error state
            int(re.sub(r"\D", "", sap_message, 0))
            response = {"result": sap_message, "error": "N/A"}
            return json.dumps(response)
        except:
            response = {"result": "N/A", "error": sap_message}
            return json.dumps(response)

    except Exception as e:

        # print(sys.exc_info()[0])
        time.sleep(.005)
        # sap_error = session.findById("wnd[0]/sbar/pane[0]").Text
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        # err(sap_error)
        response = {"result": "N/A", "error": e}
        return json.dumps(response)
    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None
        time.sleep(.005)


def lt09_query(serial_num, children):
    """
    Function used to get the Material Number corresponding the Serial Number
    """

    import json
    import win32com.client
    import time
    import pythoncom
    # serial_num = sys.argv[1]
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not type(sap_gui_auto) == win32com.client.CDispatch:
            return

        application = sap_gui_auto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui_auto = None
            return

        connection = application.Children(children)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            sap_gui_auto = None
            return

        if connection.DisabledByServer:
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui_auto = None
            return

        if session.Info.IsLowSpeedConnection:
            connection = None
            application = None
            sap_gui_auto = None
            return

        # session.findById("wnd[0]").resizeWorkingPane(90, 24, 0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nLT09"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtLEIN-LENUM").text = serial_num
        session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "999"
        session.findById("wnd[0]").sendVKey(0)
        material_number = session.findById("wnd[0]/usr/subD0171_S:SAPML03T:1711/tblSAPML03TD1711/ctxtLTAP-MATNR[0,0]").Text
        quant = session.findById("wnd[0]/usr/subD0171_S:SAPML03T:1711/tblSAPML03TD1711/txtRL03T-ANFME[1,0]").Text
        material_description = session.findById("wnd[0]/usr/subD0171_S:SAPML03T:1711/tblSAPML03TD1711/txtLTAP-MAKTX[12,0]").Text
        storage_type = session.findById("wnd[0]/usr/ctxt*LTAP-VLTYP").Text
        storage_bin = session.findById("wnd[0]/usr/txt*LTAP-VLPLA").Text

        # session.findById("wnd[0]/tbar[0]/btn[12]").press()
        # session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        # session.findById("wnd[0]/tbar[0]/btn[12]").press()
        # # Se crea respuesta y se carga en un Json con dumps

        response = {"material": material_number, "quantity": quant, "error": "N/A", "storage_type": storage_type, "storage_bin": storage_bin}
        time.sleep(.005)
        return json.dumps(response)
    except Exception as e:
        print(e)
        if session.Children.Count == 2:
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()

        error = session.findById("wnd[0]/sbar/pane[0]").Text
        response = {"material": "N/A", "quantity": "N/A", "error": error,"storage_type": "N/A", "storage_bin": "N/A"}
        # session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        # session.findById("wnd[0]").sendVKey(0)

        return json.dumps(response)

    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None


def ls24_query(sap_num, storage_type, storage_bin, children):
    import win32com.client
    import json
    import pythoncom
    import re
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")

        application = sap_gui_auto.GetScriptingEngine

        connection = application.Children(children)

        if connection.DisabledByServer:
            print("Scripting is disabled by server")
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)

        if session.Info.IsLowSpeedConnection:
            print("Connection is low speed")
            connection = None
            application = None
            sap_gui_auto = None
            return

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nLS24"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtRL01S-LGNUM").text = "521"
        session.findById("wnd[0]/usr/ctxtRL01S-MATNR").text = sap_num
        session.findById("wnd[0]/usr/ctxtRL01S-WERKS").text = "5210"
        session.findById("wnd[0]/usr/ctxtRL01S-BESTQ").text = "*"
        session.findById("wnd[0]/usr/ctxtRL01S-SOBKZ").text = "*"
        session.findById("wnd[0]/usr/ctxtRL01S-LGTYP").text = storage_type
        session.findById("wnd[0]/usr/ctxtRL01S-LGPLA").text = storage_bin
        session.findById("wnd[0]/usr/ctxtRL01S-LISTV").text = "/DEL"
        session.findById("wnd[0]").sendVKey(0)

        try:
            error = session.findById("wnd[0]/sbar/pane[0]").Text
            if error != "":
                session.findById("wnd[0]/tbar[0]/btn[15]").press()
                session.findById("wnd[0]/tbar[0]/btn[15]").press()
                response = {"quantity": "N/A", "error": error}
                return json.dumps(response)
            else:
                raise Exception('I know Python!')
        except:
            try:
                quantity = session.findById("wnd[0]/usr/lbl[35,8]").Text

            except:
                pass
            session.findById("wnd[0]/tbar[0]/btn[15]").press()
            session.findById("wnd[0]/tbar[0]/btn[15]").press()

            response = {"quantity": int(float(re.sub(r",", "", quantity).strip())), "error": "N/A"}

            return json.dumps(response)

    except Exception as e:
        error = session.findById("wnd[0]/sbar/pane[0]").Text
        response = {"quantity": "N/A", "error": error}

        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        return json.dumps(response)

    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None
