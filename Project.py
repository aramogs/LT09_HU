
#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------


# import tkinter module
from tkinter import *
from tkinter import Tk
from tkinter.ttk import *
import center_tk_window
import threading

# Check to see if all checkboxex are checked

def check_status():
    if chkv1.get() & chkv2.get() & chkv3.get():
        #print("True")
        print(variable.get())
        opt.config(state=NORMAL)
    else:
        #print("False")
        opt.config(state=DISABLED)

def check_opt(self):
    if variable.get():
        #print("True")
        b1.config(state=NORMAL)
        return variable.get()
    else:
        #print("False")
        b1.config(state=DISABLED)


#####################################

def process_sap():

    import csv
    import sys
    import time


    with open('LT10_Transfers.csv', mode='r') as csv_file:

        total_rows = 0
        for row in open("LT10_Transfers.csv"):
            total_rows += 1

        csv_reader = csv.DictReader(csv_file)
        line_count = 0

        top = Toplevel(width=300, height=400, cursor="watch")
        top.title("Transferencia de Inventario")
        master.withdraw()
        center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = Label(top, image=img1)
        label.image = img1
        label.grid(row=0, column=0, padx=20, pady=20)
        label_1 = Label(top, text="En Proceso")
        label_1.grid(row=4, column=0)


        progress_bar = Progressbar(top, orient="horizontal", mode="determinate", maximum=100, value=0)
        progress_bar.grid(row=7, column=0)
        progress_bar.start()
        progress_bar.step(10)

        textbox = Text(top, height=9, width=30)
        textbox.grid(row=10, column=0)


        def saplogin():
            import win32com.client
            import pythoncom
            import subprocess
            import ctypes
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

                SapGuiAuto = win32com.client.GetObject('SAPGUI')
                if not type(SapGuiAuto) == win32com.client.CDispatch:
                    return

                application = SapGuiAuto.GetScriptingEngine
                if not type(application) == win32com.client.CDispatch:
                    SapGuiAuto = None
                    return
                connection = application.OpenConnection("P02 - Production", True)

                if not type(connection) == win32com.client.CDispatch:
                    application = None
                    SapGuiAuto = None
                    return

                session = connection.Children(0)
                if not type(session) == win32com.client.CDispatch:
                    connection = None
                    application = None
                    SapGuiAuto = None
                    return

                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "etiquetado2"
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "etiquetado02"
                session.findById("wnd[0]").sendVKey(0)

            except:
                print(sys.exc_info())
            finally:
                session = None
                connection = None
                application = None
                SapGuiAuto = None

        saplogin()

        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            storage_unit = row["Stor.Unit"]
            if len(storage_unit)<10:
                storage_unit = "0"+storage_unit
            line_count += 1

            current_line = line_count-1
            total_lines = total_rows-1
            label_2 = Label(top, text=f"{current_line} de {total_lines}")
            label_2.grid(row=5, column=0)
            progress_bar.destroy()

            def capture(message):
                # print(f"{line_count - 1} Numero de Storage_Unit Procesado: {storage_unit} Status: OK")
                from datetime import date
                today = str(date.today())

                with open(r'Resultado.csv', 'a', newline='') as csvfile:
                    fieldnames = ['Storage Unit', 'Date', 'Error']
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                    writer.writerow({'Storage Unit': storage_unit, 'Date': today, 'Error': message})
                    textbox.insert(END, "Serial: " + storage_unit + ' Status: ' + "OK\n")
                    textbox.see(END)

            def err(error):
                # print(f"{line_count - 1} Numero de Storage_Unit Procesado: {storage_unit} Status: Error")
                from datetime import date
                today = str(date.today())

                with open(r'Resultado.csv', 'a', newline='') as csvfile:
                    fieldnames = ['Storage Unit', 'Date', 'Error']
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                    writer.writerow({'Storage Unit': storage_unit, 'Date': today, 'Error': error})
                    textbox.insert(END, "Serial: " + storage_unit + ' Status: ' + "Err\n")
                    textbox.see(END)

            # -Sub Main--------------------------------------------------------------
            def Main():
                import win32com.client
                import pywintypes
                import ctypes
                try:
                    # Necesario par correr win32com.client en Threading
                    #pythoncom.CoInitialize()


                    sapmin = ctypes.windll.user32.FindWindowW(None, "P02(1)/200 SAP Easy Access")
                    ctypes.windll.user32.ShowWindow(sapmin, 6)

                    SapGuiAuto = win32com.client.GetObject("SAPGUI")
                    if not type(SapGuiAuto) == win32com.client.CDispatch:
                        return

                    application = SapGuiAuto.GetScriptingEngine
                    if not type(application) == win32com.client.CDispatch:
                        SapGuiAuto = None
                        return

                    connection = application.Children(0)
                    if not type(connection) == win32com.client.CDispatch:
                        application = None
                        SapGuiAuto = None
                        return

                    if connection.DisabledByServer == True:
                        application = None
                        SapGuiAuto = None
                        return

                    session = connection.Children(0)
                    if not type(session) == win32com.client.CDispatch:
                        connection = None
                        application = None
                        SapGuiAuto = None
                        return

                    if session.Info.IsLowSpeedConnection == True:
                        connection = None
                        application = None
                        SapGuiAuto = None
                        return


                    session.findById("wnd[0]/tbar[0]/okcd").text = "LT09"
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/txtLEIN-LENUM").text = storage_unit
                    session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "998"
                    session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").setFocus()
                    session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").caretPosition = 3
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/ctxt*LTAP-NLTYP").text = "VUL"
                    session.findById("wnd[0]/usr/ctxt*LTAP-NLBER").text = "001"
                    session.findById("wnd[0]/usr/txt*LTAP-NLPLA").text = variable.get()
                    session.findById("wnd[0]/usr/txt*LTAP-NLPLA").setFocus()
                    session.findById("wnd[0]/usr/txt*LTAP-NLPLA").caretPosition = 6
                    session.findById("wnd[0]/tbar[0]/btn[11]").press()

                    SAPmessage = session.findById("wnd[0]/sbar/pane[0]").Text
                    time.sleep(.005)
                    capture(SAPmessage)

                    session.findById("wnd[0]/tbar[0]/btn[15]").press()



                except:
                   # print(sys.exc_info()[0])
                    time.sleep(.005)
                    SAPerror = session.findById("wnd[0]/sbar/pane[0]").Text
                    session.findById("wnd[0]/tbar[0]/btn[15]").press()
                    err(SAPerror)
                finally:
                    session = None
                    connection = None
                    application = None
                    SapGuiAuto = None
                    time.sleep(.005)

            # -Main------------------------------------------------------------------
            Main()

            # -End-------------------------------------------------------------------

    time.sleep(2)
    top.destroy()
    new_window()


def startsap():
    t1 = threading.Thread(target=process_sap, daemon=True)
    t1.start()
def terminate():
    import win32com.client
    import ctypes
    sapmin = ctypes.windll.user32.FindWindowW(None, "P02(1)/200 SAP Easy Access")
    ctypes.windll.user32.ShowWindow(sapmin, 6)

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return

    if connection.DisabledByServer == True:
        application = None
        SapGuiAuto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return

    if session.Info.IsLowSpeedConnection == True:
        connection = None
        application = None
        SapGuiAuto = None
        return
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    master.quit()


def new_window():
    top = Toplevel(width=300, height=300)
    top.title("Proceso terminado")
    master.withdraw()
    center_tk_window.center_on_screen(top)
    top.resizable(False, False)
    top.lift()
    label = Label(top, image=img1)
    label.image = img1
    label.grid(row=0, column=0, padx=20, pady=20)
    label_1 = Label(top, text="Procesado")
    label_1.grid(row=5, column=0)
    progress_bar = Progressbar(top, orient="horizontal", mode="determinate", maximum=100, value=100)
    progress_bar.grid(row=6, column=0)

    button = Button(top, text="Terminar", width=50, command=terminate)
    button.grid(row=8, column=0, columnspan=2, sticky=W, pady=10)
    #####################



# creating main tkinter window/toplevel
master: Tk = Tk()
master.title("Movimientos de material en SLOCK:VUL")

style = Style()
style.configure("TLabelframe.Label", font=("TkDefaultFont",10, "bold"),foreground ='blue')

# adding image (remember image should be PNG and not JPG)
img = PhotoImage(file=r"./img/tristone.png")
img1 = img.subsample(2, 2)

# setting image with the help of label
Label(master, image=img1).grid(row=0, column=2,columnspan=5, rowspan=2, padx=5, pady=5)

# this will create CheckBoxes
chkv1 = BooleanVar()
chkv2 = BooleanVar()
chkv3 = BooleanVar()
chkv1.set(False)
chkv2.set(False)
chkv3.set(False)

lfw = LabelFrame(master, text="Instrucciones")
lw1 = Label(lfw, text="Para poder continuar cerciorarse de tener los siguientes puntos                                                ")
chkb1 = Checkbutton(lfw, text="1.-Cerrar todas las ventanas de SAP abiertas",var=chkv1,command=check_status,)
chkb2 = Checkbutton(lfw, text="2.-Insertar los Storage Unit a ser cambiados en el archivo LT10_Transfers.csv",var=chkv2, command=check_status)
chkb3 = Checkbutton(lfw, text="3.-Una vez insertados los Storage Unit, guardar y cerrar el archivo.",var=chkv3, command= check_status)


# grid method to arrange labels in respective
# rows and columns as specified
lfw.grid(row=6, column=2,columnspan=5, sticky=W, pady=4, padx=10)
lw1.grid(row=8, column=2,columnspan=5, sticky=W, pady=4)
chkb1.grid(row=12, column=3, columnspan=5, sticky=W, pady=0)
chkb2.grid(row=13, column=3, columnspan=5, sticky=W, pady=0)
chkb3.grid(row=14, column=3, columnspan=5, sticky=W, pady=0)


OptionList = ["","V01", "V02", "V03", "V04", "V05"]

variable = StringVar(master)
variable.set(OptionList[0])
opt = OptionMenu(lfw, variable, *OptionList,command=check_opt)
opt.config(state=DISABLED)
opt.grid(row=16, column=3, columnspan=1, sticky=W, pady=0)

lw1 = Label(lfw, text="4.-Seleccionar ubicacion")
lw1.grid(row=16, column=3, sticky=W, padx=55)

# button widget
b1 = Button(master, text="Transferir", width=50, state = DISABLED,command=lambda: startsap())

# arranging button widgets
b1.grid(row=18, column=3,columnspan=5, sticky=W, pady=10)
# infinite loop which can be terminated
# by keyboard or mouse interrupt

center_tk_window.center_on_screen(master)
master.resizable(False, False)
mainloop()
