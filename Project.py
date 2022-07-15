# -Begin-----------------------------------------------------------------
import datetime
import inspect
import json
import math
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from tkinter import messagebox
from tkinter import ttk

import center_tk_window
import openpyxl
import pandas
from openpyxl.utils.dataframe import dataframe_to_rows

# - Functions------------------------------------------------------------
import SAP_Functions


# Check to see if all checkboxes are checked
def check_status():
    if check_v_1.get() & check_v_2.get() & check_v_3.get():
        b1.config(state=tk.NORMAL)
    else:
        b1.config(state=tk.DISABLED)


def excel_save(in_queue):
    workbook = openpyxl.load_workbook(project_file)
    sheet = workbook['Result']
    while True:
        item = in_queue.get()
        # process
        query = json.loads(item)
        from datetime import date
        today = str(date.today())
        current_time = str(datetime.datetime.now().strftime("%H:%M:%S"))
        df = pandas.DataFrame([[query["serial"], today, current_time, query["status"], query["result"], query["process"]]],
                              columns=["Serial", "Date", "Time", "Status", "SAP_Result", "Process"])
        for row_df in dataframe_to_rows(df, header=False, index=False):
            if not row_df[0] is None:
                sheet.append(row_df)
        workbook.save(project_file)
        workbook.close()


def open_sap(in_queue, ):
    while True:
        item = in_queue.get()
        # process
        SAP_Functions.sap_login(environment)
        in_queue.task_done()


def check_status_2():
    if check_v_1_2.get() & check_v_2_2.get() & check_v_3_2.get():
        b2.config(state=tk.NORMAL)
    else:
        b2.config(state=tk.DISABLED)


#####################################

def process_sap_withdraw():
    try:
        global label_2
        global textbox
        global sap_instances
        global all_threads
        global stop_threads

        excel_file = pandas.read_excel(project_file, sheet_name="LT01_Serials_Withdraw")
        total_rows = excel_file.shape[0]

        top = tk.Toplevel(width=275, height=400, cursor="watch")
        top.iconbitmap('./img/image.ico')
        top.title("Transferencia Seriales")
        root.withdraw()
        # center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = ttk.Label(top, image=img1)
        label.image = img1
        label.grid(row=0, column=0, padx=20, pady=20)
        label_1 = ttk.Label(top, text="En Proceso")
        label_1.grid(row=4, column=0)

        progress_bar = ttk.Progressbar(top, orient="horizontal", mode="determinate", maximum=100, value=0)
        progress_bar.grid(row=7, column=0)
        progress_bar.start()
        progress_bar.step(10)

        textbox = tk.Text(top, height=9, width=30)
        textbox.grid(row=10, column=0)

        label_2 = ttk.Label(top, text=f"0 de {total_rows}")
        label_2.grid(row=5, column=0)

        thread_excel = threading.Thread(target=excel_save, args=(excel_queue,))
        thread_excel.daemon = True
        thread_excel.start()

        def capture(serial_, message_, sst_, ssb_, quantity_, process_):
            excel_queue.put(json.dumps(
                {"serial": f"{serial_}", "quantity": f"{quantity_}", "result": f"{message_}", "status": "OK", "sst": f"{sst_}", "ssb": f"{ssb_}", "process": f"{process_}"}))
            textbox.insert('1.0', "SAP: " + serial_ + ' Status: ' + "OK\n")

        def err(serial_, error_, sst_, ssb_, quantity_, process_):
            excel_queue.put(json.dumps(
                {"serial": f"{serial_}", "quantity": f"{quantity_}", "result": f"{error_}", "status": "ERROR", "sst": f"{sst_}", "ssb": f"{ssb_}", "process": f"{process_}"}))
            textbox.insert('1.0', "SAP: " + serial_ + ' Status: ' + "ERR\n")

        def do_work(in_queue):
            global current_process
            global label_2
            current_process = 0
            current_t = threading.current_thread()

            while True:
                item = in_queue.get()
                # process
                query = json.loads(item)
                if current_process == 0:
                    current_process += 1
                total_lines = total_rows
                progress_bar.destroy()
                label_2.config(text=f"{current_process} de {total_lines}")
                label_2.grid(row=5, column=0)

                response_lt09 = json.loads(SAP_Functions.lt09_query(query["serial"], int(current_t.getName())))
                lt09_error = response_lt09["error"]
                lt09_material = response_lt09["material"]
                try:
                    lt09_quantity = re.sub(r",", "", response_lt09["quantity"])
                except:
                    lt09_quantity = response_lt09["quantity"]

                if lt09_error != "N/A":
                    err(query["serial"], lt09_error, query["sst"], query["ssb"], query["quantity"], query["process"])
                else:
                    if int(float(lt09_quantity)) >= int(query["quantity"]):
                        response = json.loads(SAP_Functions.lt01_query_withdraw(storage_location, lt09_material, query["quantity"], query["sst"], query["ssb"], query["serial"],
                                                                                int(current_t.getName())))
                        result = response["result"]
                        error = response["error"]

                        if error != "N/A":
                            err(query["serial"], error, query["sst"], query["ssb"], query["quantity"], query["process"])
                        else:
                            capture(query["serial"], result, query["sst"], query["ssb"], query["quantity"], query["process"])
                    else:
                        error_quantity = f'Requested amount exceeded by {math.ceil((int(query["quantity"]) - int(float(lt09_quantity))) / int(float(lt09_quantity)) * 100)}%'
                        err(query["serial"], error_quantity, query["sst"], query["ssb"], query["quantity"], query["process"])

                current_process += 1
                in_queue.task_done()

        for x in range(int(sap_instances)):
            thread_sap = threading.Thread(target=open_sap, name=str(x), args=(sap_queue,))
            thread_sap.daemon = True
            thread_sap.start()

        for z in range(int(sap_instances)):
            sap_queue.put(z)
        sap_queue.join()

        for index, row in excel_file.iterrows():
            serial = str(row["Serial"]).replace(".0", "")
            quantity = row["Quantity"]
            sst = str(row["StorageType"]).replace(".0", "")
            ssb = str(row["StorageBin"]).replace(".0", "")

            work.put(json.dumps({"serial": f"{int(serial)}", "quantity": f"{quantity}", "sst": f"{sst}", "ssb": f"{ssb}", "process": "withdraw"}))

        for y in range(int(sap_instances)):
            thread_work = threading.Thread(target=do_work, name=str(y), args=(work,))
            thread_work.daemon = True
            all_threads.append(thread_work)
            thread_work.start()
        work.join()
        top.destroy()
        new_window()
    except Exception as e:
        error_window(e, traceback)


def process_sap_insert():
    try:
        global label_2
        global textbox
        global sap_instances
        global all_threads
        global stop_threads

        excel_file = pandas.read_excel(project_file, sheet_name="LT01_Serials_Insert")
        total_rows = excel_file.shape[0]
        line_count = 0

        top = tk.Toplevel(width=275, height=400, cursor="watch")
        top.iconbitmap('./img/image.ico')
        top.title("Transferencia Seriales")
        root.withdraw()
        # center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = ttk.Label(top, image=img1)
        label.image = img1
        label.grid(row=0, column=0, padx=20, pady=20)
        label_1 = ttk.Label(top, text="En Proceso")
        label_1.grid(row=4, column=0)

        progress_bar = ttk.Progressbar(top, orient="horizontal", mode="determinate", maximum=100, value=0)
        progress_bar.grid(row=7, column=0)
        progress_bar.start()
        progress_bar.step(10)

        textbox = tk.Text(top, height=9, width=30)
        textbox.grid(row=10, column=0)

        label_2 = ttk.Label(top, text=f"0 de {total_rows}")
        label_2.grid(row=5, column=0)

        thread_excel = threading.Thread(target=excel_save, args=(excel_queue,))
        thread_excel.daemon = True
        thread_excel.start()

        def capture(serial_, message_, sst_, ssb_, quantity_, process_):
            excel_queue.put(json.dumps(
                {"serial": f"{serial_}", "quantity": f"{quantity_}", "result": f"{message_}", "status": "OK", "sst": f"{sst_}", "ssb": f"{ssb_}", "process": f"{process_}"}))
            textbox.insert('1.0', "SAP: " + serial_ + ' Status: ' + "OK\n")

        def err(serial_, error_, sst_, ssb_, quantity_, process_):
            excel_queue.put(json.dumps(
                {"serial": f"{serial_}", "quantity": f"{quantity_}", "result": f"{error_}", "status": "ERROR", "sst": f"{sst_}", "ssb": f"{ssb_}", "process": f"{process_}"}))
            textbox.insert('1.0', "SAP: " + serial_ + ' Status: ' + "ERR\n")

        def do_work(in_queue):
            global current_process
            global label_2
            current_process = 0
            current_t = threading.current_thread()

            while True:
                item = in_queue.get()
                # process
                query = json.loads(item)
                if current_process == 0:
                    current_process += 1
                total_lines = total_rows
                progress_bar.destroy()
                label_2.config(text=f"{current_process} de {total_lines}")
                label_2.grid(row=5, column=0)

                response_lt09 = json.loads(SAP_Functions.lt09_query(query["serial"], int(current_t.getName())))
                lt09_error = response_lt09["error"]
                lt09_material = response_lt09["material"]
                lt09_quantity = response_lt09["quantity"]
                lt09_storage_type = response_lt09["storage_type"]
                lt09_sorage_bin = response_lt09["storage_bin"]

                if lt09_error != "N/A":
                    err(query["serial"], lt09_error, query["sst"], query["ssb"], query["quantity"], query["process"])
                else:
                    ls24_response = json.loads(SAP_Functions.ls24_query(lt09_material, query["sst"], query["ssb"], int(current_t.getName())))
                    ls24_quantity = ls24_response["quantity"]
                    ls24_error = ls24_response["error"]

                    if ls24_error != "N/A":
                        err(query["serial"], ls24_error, query["sst"], query["ssb"], query[quantity], query["process"])
                    else:
                        if lt09_error != "N/A":
                            err(query["serial"], lt09_error, query["sst"], query["ssb"], query[quantity], query["process"])
                        else:
                            if int(float(ls24_quantity)) >= int(query["quantity"]):
                                response = json.loads(SAP_Functions.lt01_query_insert(storage_location, lt09_material, query["quantity"], query["sst"], query["ssb"], query["serial"],
                                                                                      lt09_storage_type, lt09_sorage_bin, int(current_t.getName())))
                                result = response["result"]
                                error = response["error"]

                                if error != "N/A":
                                    err(query["serial"], error, query["sst"], query["ssb"], query[quantity], query["process"])
                                else:
                                    capture(query["serial"], result, query["sst"], query["ssb"], query["quantity"], query["process"])
                            else:
                                error_quantity = f'Requested amount exceeded by {math.ceil((int(query["quantity"]) - int(float(ls24_quantity))) / int(float(ls24_quantity)) * 100)}%'
                                err(query["serial"], error_quantity, query["sst"], query["ssb"], query["quantity"], query["process"])

                current_process += 1
                in_queue.task_done()

        for x in range(int(sap_instances)):
            thread_sap = threading.Thread(target=open_sap, name=str(x), args=(sap_queue,))
            thread_sap.daemon = True
            thread_sap.start()

        for z in range(int(sap_instances)):
            sap_queue.put(z)
        sap_queue.join()

        for index, row in excel_file.iterrows():
            serial = str(int(row["Serial"]))
            sst = str(row["StorageType"]).replace(".0", "")
            ssb = str(row["StorageBin"]).replace(".0", "")
            quantity = row["Quantity"]

            if len(serial) < 10:
                serial = "0" + serial

            work.put(json.dumps(
                {"serial": f"{int(serial)}", "quantity": f"{quantity}", "sst": f"{sst}", "ssb": f"{ssb}", "process": "Insert"}))

        for y in range(int(sap_instances)):
            thread_work = threading.Thread(target=do_work, name=str(y), args=(work,))
            thread_work.daemon = True
            all_threads.append(thread_work)
            thread_work.start()
        work.join()
        top.destroy()
        new_window()
    except Exception as e:
        error_window(e, traceback)


def lt01_withdraw():
    t1 = threading.Thread(target=process_sap_withdraw, daemon=True)
    t1.start()


def lt01_insert():
    t1 = threading.Thread(target=process_sap_insert, daemon=True)
    t1.start()


def terminate(root_, top):
    check_v_1.set(False)
    check_v_2.set(False)
    check_v_3.set(False)
    check_v_1_2.set(False)
    check_v_2_2.set(False)
    check_v_3_2.set(False)
    top.update()
    top.destroy()
    root_.deiconify()
    os.system(f'taskkill /im saplogon.exe /T /F')
    root_.destroy()
    # SAP_Functions.terminate()


def refresh():
    global file
    global environment
    global storage_location
    global sap_instances

    file = pandas.read_excel(project_file, sheet_name="CONFIG")
    environment = str(file["ENVIRONMENT"][0])
    storage_location = str(file["Storage_Location"][0])
    if int(float(file["SAP_Instances"][0])) >= int(sap_instances):
        sap_instances = int(float(file["SAP_Instances"][0]))
    root.title(f'MST:       Env: {environment}     |       St.Loc: {storage_location} | SAP.Instances: {sap_instances}')


def help_file():
    try:
        if getattr(sys, 'frozen', False):
            test = re.sub(r"(.*\\).*", "\\1", sys.executable)
            help_f = f'{test}\\help\\{project_name}_HELP.docx'

        else:
            help_f = f'{os.path.dirname(__file__)}/help/{project_name}_HELP.docx'
        subprocess.Popen([help_f], shell=True)
    except Exception as e:
        error_window(e, traceback)


def about():
    try:
        top = tk.Toplevel(width=305, height=200)
        top.iconbitmap('./img/image.ico')
        top.title("About MTS")

        center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = ttk.Label(top, image=img3)
        label.image = img3
        label.grid(row=0, column=0, padx=40, pady=0, sticky=tk.W)
        label_1 = ttk.Label(top, text="Version 1.1.0")
        label_1.grid(row=5, column=0)

        #####################
    except Exception as e:
        error_window(e, traceback)


def new_window():
    try:
        top = tk.Toplevel(width=305, height=300)
        top.iconbitmap('./img/image.ico')
        top.title("Proceso terminado")
        root.withdraw()
        center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = ttk.Label(top, image=img1)
        label.image = img1
        label.grid(row=0, column=0, padx=40, pady=0, sticky=tk.W)
        label_1 = ttk.Label(top, text="Procesado")
        label_1.grid(row=5, column=0)
        progress_bar = ttk.Progressbar(top, orient="horizontal", mode="determinate", maximum=100, value=100)
        progress_bar.grid(row=6, column=0)

        button = ttk.Button(top, text="Terminar", width=50, command=lambda: terminate(root, top))
        button.grid(row=8, column=0, sticky=tk.W, pady=10)
        #####################
    except Exception as e:
        error_window(e, traceback)


def error_window(e_, traceback_):
    root.withdraw()
    res = messagebox.showerror(f'Error: {inspect.stack()[1].function}', f'{e_} \n\n {traceback_.format_exc()}')
    if res:
        root.quit()


try:
    # if getattr(sys, 'frozen', False):
    #     project_file = re.sub(r".exe$", "", re.sub(r".*\\", "", sys.executable))
    #     project_file = f'{project_file}.xlsx'
    # else:
    #     project_file = f'{re.sub(r".py", "", os.path.basename(__file__))}.xlsx'
    work = queue.Queue()
    results = queue.Queue()
    sap_queue = queue.Queue()
    excel_queue = queue.Queue()
    current_process = 0
    label_2 = None
    textbox = None
    all_threads = []

    if getattr(sys, 'frozen', False):
        project_file = re.sub(r".exe$", "", re.sub(r".*\\", "", sys.executable))
        project_name = re.sub(r".exe$", "", re.sub(r".*\\", "", sys.executable))
        project_file = f'{project_file}.xlsx'
    else:
        project_file = f'{re.sub(r".py", "", os.path.basename(__file__))}.xlsx'
        project_name = f'{re.sub(r".py", "", os.path.basename(__file__))}'

    file = pandas.read_excel(project_file, sheet_name="CONFIG")
    environment = str(file["ENVIRONMENT"][0])
    storage_location = str(file["Storage_Location"][0])
    sap_instances = int(float(file["SAP_Instances"][0]))
    root = tk.Tk()
    root.title(f'MST:       Environment: {environment} | Storage Location: {storage_location}')
    root.iconbitmap('./img/image.ico')
    tabControl = ttk.Notebook(root)

    menu_bar = tk.Menu(root)
    file_menu = tk.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Refresh", command=refresh)
    menu_bar.add_cascade(label="File", menu=file_menu)

    help_menu = tk.Menu(menu_bar, tearoff=0)
    help_menu.add_command(label="MST Help", command=help_file)
    help_menu.add_command(label="About...", command=about)
    menu_bar.add_cascade(label="Help", menu=help_menu)

    root.config(menu=menu_bar)

    tab1 = ttk.Frame(tabControl)
    tab2 = ttk.Frame(tabControl)
    tab1_image = tk.PhotoImage(file=r"./img/withdrawal_32.png")
    tab2_image = tk.PhotoImage(file=r"./img/insert_32.png")

    tabControl.add(tab1, text='Reducir Cantidad a Seriales', image=tab1_image, compound=tk.LEFT)
    tabControl.add(tab2, text='Agregar Cantidad a Seriales', image=tab2_image, compound=tk.LEFT)
    tabControl.grid(column=0, row=0, sticky=tk.E + tk.W + tk.N + tk.S)

    style = ttk.Style()
    style.configure("TLabelframe.Label", font=("TkDefaultFont", 10, "bold"), foreground='blue')
    # Notebook tab size
    current_theme = style.theme_use()
    style.theme_settings(current_theme, {"TNotebook.Tab": {"configure": {"padding": [20, 3]}}})

    # adding image (remember image should be PNG and not JPG)
    img1 = tk.PhotoImage(file=r"./img/Tristone_Withdrawal.png").subsample(2, 2)
    img2 = tk.PhotoImage(file=r"./img/Tristone_Insert.png").subsample(2, 2)
    img3 = tk.PhotoImage(file=r"./img/Tristone.png").subsample(2, 2)

    # setting image with the help of label
    ttk.Label(tab1, image=img1).grid(row=0, column=2, columnspan=5, rowspan=2, padx=5, pady=5)
    ttk.Label(tab2, image=img2).grid(row=0, column=2, columnspan=5, rowspan=2, padx=5, pady=5)

    # this will create CheckBoxes
    check_v_1 = tk.BooleanVar()
    check_v_2 = tk.BooleanVar()
    check_v_3 = tk.BooleanVar()
    check_v_1.set(False)
    check_v_2.set(False)
    check_v_3.set(False)

    check_v_1_2 = tk.BooleanVar()
    check_v_2_2 = tk.BooleanVar()
    check_v_3_2 = tk.BooleanVar()
    check_v_1_2.set(False)
    check_v_2_2.set(False)
    check_v_3_2.set(False)

    lfw = ttk.LabelFrame(tab1, text="Instrucciones")
    lw1 = ttk.Label(lfw, text="Para poder continuar cerciorarse de tener los siguientes puntos                                                ")
    check_1 = ttk.Checkbutton(lfw, text="1.-Cerrar todas las ventanas de SAP abiertas", var=check_v_1, command=check_status, )
    check_2 = ttk.Checkbutton(lfw, text="2.-Insertar los Seriales a ser cambiados en el archivo .xlsx", var=check_v_2, command=check_status)
    check_3 = ttk.Checkbutton(lfw, text="3.-Una vez insertados los Seriales, guardar y cerrar el archivo.", var=check_v_3, command=check_status)

    lfw_2 = ttk.LabelFrame(tab2, text="Instrucciones")
    lw1_2 = ttk.Label(lfw_2, text="Para poder continuar cerciorarse de tener los siguientes puntos                                                ")
    check_1_2 = ttk.Checkbutton(lfw_2, text="1.-Cerrar todas las ventanas de SAP abiertas", var=check_v_1_2, command=check_status_2, )
    check_2_2 = ttk.Checkbutton(lfw_2, text="2.-Insertar los Seriales a ser cambiados en el archivo .xlsx", var=check_v_2_2, command=check_status_2)
    check_3_2 = ttk.Checkbutton(lfw_2, text="3.-Una vez insertados los Seriales, guardar y cerrar el archivo.", var=check_v_3_2, command=check_status_2)

    # grid method to arrange labels in respective
    # rows and columns as specified
    lfw.grid(row=6, column=2, columnspan=5, sticky=tk.W, pady=4, padx=10)
    lw1.grid(row=8, column=2, columnspan=5, sticky=tk.W, pady=4)
    check_1.grid(row=12, column=3, columnspan=5, sticky=tk.W, pady=0)
    check_2.grid(row=13, column=3, columnspan=5, sticky=tk.W, pady=0)
    check_3.grid(row=14, column=3, columnspan=5, sticky=tk.W, pady=0)

    lfw_2.grid(row=6, column=2, columnspan=5, sticky=tk.W, pady=4, padx=10)
    lw1_2.grid(row=8, column=2, columnspan=5, sticky=tk.W, pady=4)
    check_1_2.grid(row=12, column=3, columnspan=5, sticky=tk.W, pady=0)
    check_2_2.grid(row=13, column=3, columnspan=5, sticky=tk.W, pady=0)
    check_3_2.grid(row=14, column=3, columnspan=5, sticky=tk.W, pady=0)

    # button widget
    b1 = ttk.Button(tab1, text="Transferir", width=50, state=tk.DISABLED, command=lambda: lt01_withdraw())
    b2 = ttk.Button(tab2, text="Transferir", width=50, state=tk.DISABLED, command=lambda: lt01_insert())

    # arranging button widgets
    b1.grid(row=18, column=3, columnspan=5, sticky=tk.W, pady=10)
    b2.grid(row=18, column=3, columnspan=5, sticky=tk.W, pady=10)

    # infinite loop which can be terminated
    # by keyboard or mouse interrupt
    center_tk_window.center_on_screen(root)
    root.resizable(False, False)
    root.mainloop()
except Exception as E:
    error_window(E, traceback)
