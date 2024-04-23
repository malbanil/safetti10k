import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import openpyxl
import pandas as pd
import datetime
import json
import os

EX_FILE    = 'subscribers_10k.xlsx'
WOMPY_FILE = 'wompi.xlsx'
MATCH_FILE = 'consolidate_subscribers_10k.xlsx'
ERROR_FILE = 'error_log.txt'
LASTM_FILE = 'last_match.txt'

def read_last_match():
    """
    Read txt file with info
    Returns:
        string: txt with the file info
    """
    with open(LASTM_FILE, "r") as archivo:
        contenido = archivo.read()
    return contenido

def write_last_match():
    """
    Write txt file with info
    Returns:
        (empty)
    """
    date_now = datetime.datetime.now()
    fecha_str = date_now.strftime("%Y-%m-%d %H:%M:%S") 
    with open(LASTM_FILE, "w") as archivo:
         archivo.write(fecha_str)
        
def write_error_log(str_error):
    """
    Create txt file with error log messages
    Args:
        str_error (string): With the error message
    Returns:
        file: txt with the ERROR_FILE global file value
    """
    date_now = datetime.datetime.now()
    with open(ERROR_FILE, 'a') as archivo:
        archivo.write(str(date_now) + " : " + str_error + ".\n")

def load_json():
    """
    Create excel from Json Url to Online Form Safetti 10k
    Args:
        (Empty)
    Returns:
        file: Excel with the EX_FILE global file value
    Raises:
        TypeError: If the endpoint doesn't exist or the function can't create a excel file
    """
    try:
        url = "https://eventos.safetti.com/api/subscriptions?token=0d7d6c29-9abd-4b78-af28-bf570fe07276"
        df = pd.read_json(url)
        df.to_excel(EX_FILE, index=False)
    except FileNotFoundError:
        write_error_log("Error - El archivo JSON no fue encontrado.")
    except ValueError:
        write_error_log("Error - El archivo JSON es inválido o está mal formateado.")
    except Exception as e:
        write_error_log("Error - Ocurrió un error inesperado, " + str(e))    

def load_data():
    try:
        # Intenta abrir el archivo
        with open(MATCH_FILE, 'r'):
            dft = pd.read_excel(MATCH_FILE) # create DataFrame
            df = dft[['id', 'idNumber','firstName','lastName','email','mobile','monto','medio de pago','fecha']]
            l1 = list(df)  
            r_set = df.to_numpy().tolist() 
            treeview["height"] = 10  # Number of rows to display, default is 10
            treeview["show"] = "headings"
            # column identifiers
            treeview["columns"] = l1
            for i in l1:
                treeview.column(i, width=90)
                # Headings of respective columns
                treeview.heading(i, text=i)
            for dt in r_set:
                v = [r for r in dt]  # creating a list from each row
                treeview.insert("", "end", iid=v[0], values=v)  # adding row   
    except FileNotFoundError:
        # Si el archivo no se encuentra, se maneja la excepción
         write_error_log("Error - El archivo CONSOLIDADO no existe.")  

def user_search():
    dft = pd.read_excel(MATCH_FILE) # create DataFrame
    df = dft[['id', 'idNumber','firstName','lastName','email','mobile','monto','medio de pago','fecha']]
    treeview["height"] = 10  # Number of rows to display, default is 10
    treeview["show"] = "headings"
    
    treeview.delete(*treeview.get_children())
    query = name_entry.get().strip() # get user entered strißng
    if query.isdigit():  # if query is number
        str1 = df["idNumber"] == query #
    else:
        str1 = df.firstName.str.contains(query, case=False)  # name column value matching
   
    df2 = df[(str1)]  # combine all conditions using | operator
    l1 = list(df)  # List of column names as headers
    treeview["columns"] = l1
    r_set = df2.to_numpy().tolist()  # Create list of list using rows
    #trv = ttk.Treeview(root, selectmode="browse")  # selectmode="browse" or "extended"
    #trv.grid(row=2, column=1, columnspan=3, padx=10, pady=20)  #

    # column identifiers
    for i in l1:
        treeview.column(i,)
        # Headings of respective columns
        treeview.heading(i, text=i)
    for dt in r_set:
        v = [r for r in dt]  # creating a list from each row
        treeview.insert("", "end", iid=v[0], values=v)  # adding row 

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

def match_report():
    dfform = pd.read_excel(EX_FILE) # create DataFrame
    dfform['idNumber'] = dfform['idNumber'].str.replace('.', '')
    dfform['idNumber'] = dfform['idNumber'].str.replace(',', '')
    #dfwompy = pd.read_csv(WOMPY_FILE,sep=';')
    dfwompy = pd.read_excel(WOMPY_FILE) # create DataFrame
    dfmatch_headers = dfform.columns.to_list() + dfwompy.columns.to_list()
    dfmatch = pd.DataFrame(columns = dfmatch_headers)
    check_file = os.path.isfile(MATCH_FILE)
    if check_file:
        os.remove(MATCH_FILE)

    #loop excel 10k
    for index, row in dfform.iterrows():
        t_cc = row['idNumber']
        t_email = row['email']
      
        #find by int(cc) in wompy
        if t_cc.isdigit():
            resultw = dfwompy[dfwompy['documento del pagador'] == int(t_cc)]
        else:
            resultw = dfwompy[dfwompy['documento del pagador'] == str(t_cc)]
        
        if not resultw.empty :
            serie = pd.Series(resultw.iloc[0])
            c_rst = pd.concat([row,serie])
            dfmatch.loc[len(dfmatch)] = c_rst
        else:
            #find by str(cc) in wompy
            if not resultw.empty :
                serie = pd.Series(resultw.iloc[0])
                c_rst = pd.concat([row,serie])
                dfmatch.loc[len(dfmatch)] = c_rst        
            else:
                #find by email in wompy
                resultw = dfwompy[dfwompy['email del pagador'] == t_email]    
                if not resultw.empty :
                    serie = pd.Series(resultw.iloc[0])
                    c_rst = pd.concat([row,serie])
                    dfmatch.loc[len(dfmatch)] = c_rst
                else:
                    dfmatch.loc[len(dfmatch)] = row
          

    dfmatch.to_excel(MATCH_FILE, index=False)
    write_last_match()
    write_error_log("Notice - Match de reportes realizado!!.")
    mostrar_ventana_emergente()
    
def mostrar_ventana_emergente():
    ventana_emergente = tk.Toplevel(root)
    ventana_emergente.title("Proeceso de Match")
    etiqueta = tk.Label(ventana_emergente, text="El proceso fue ejecutado exitosamente. Ver el archivo consolidate_subscribers_10k.xlsx")
    etiqueta.pack(padx=20, pady=20)
    boton_cerrar = tk.Button(ventana_emergente, text="Cerrar", command=ventana_emergente.destroy)
    boton_cerrar.pack(padx=10, pady=10)

### ----- **** MAIN **** ------
    
#0. test process
#match_report()
#exit()
#1. Get de Json from endpoint
load_json()
#Screen
root = tk.Tk()
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
root.title('Safetti 10k Monitoring V1.0')
style.theme_use("forest-dark")
frame = ttk.Frame(root)
frame.pack()

# -Widget Left
widgets_frame = ttk.LabelFrame(frame, text="Buscar usuario:")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widgets_frame)
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

button = ttk.Button(widgets_frame, text="Buscar", command=user_search)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

button_match = ttk.Button(widgets_frame, text="Consolidar", command=match_report)
button_match.grid(row=6, column=0, padx=5, pady=5)
str_lastm = read_last_match()
texto_widget = tk.Label(widgets_frame, text="Ultimo: " + str_lastm)
texto_widget.grid(row=7, column=0, padx=5, pady=5)

separator = ttk.Separator(widgets_frame)
separator.grid(row=8, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(
    widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=9, column=0, padx=5, pady=10, sticky="nsew")

# -Widget Right
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("#0", "#1", "#2", "#3")
treeview = ttk.Treeview(treeFrame, show="headings", columns=cols, height=13)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()
root.mainloop()
