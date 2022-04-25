#!/usr/bin/env python
# pflege scan tasks
import PySimpleGUI as sg
import operator

import os
import io
from datetime import datetime
import glob
import subprocess
import re
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileMerger
from pdf2image import convert_from_path
import pytesseract
import colored  
from colored import stylize
import config

def analyze_command_parameters(cmd_string) :

    split_strings = []
    for i in config_parms["metadata_parameter"].keys() :
        split_strings.append(i + "=")

    result = re.split("("+ "|".join(split_strings) + ")", cmd_string)

    selector = {}    

    for i in range(len(result)-1) :
        for j in config_parms["metadata_parameter"].keys() :
            if result[i] == j + "=" :
                try :
                    values_dict = config_parms["metadata_value_" + j]
                    selector[j] = ""
                    for ii,jj in values_dict.items() :
                        if result[i+1].strip()[0:1] == ii :
                            selector[j] = jj
                except IndexError:
                    selector[j] = result[i+1].strip()

    return(selector)

sg.theme('Light blue 2')
# ------ Some functions to help generate data for the table ------


def make_table(num_rows, num_cols, selector):
    data = []
    data_line = []

    taskcount = 0
    task_s = {}
    k = "l "

    count = 0
    for filename in sorted(glob.glob(source_dir + '*.pdf')):
        if os.path.isfile(filename) and ".pdf" in filename :

            data_line = []
            metadata = {}
            count += 1
            taskname = filename.replace(source_dir,"").replace(".pdf","")
            metadata["TASK"] = taskname
            data_line.append(taskname)

            reader = PdfFileReader(filename)

            for i,j in config_parms["metadata_parameter"].items() :
                try :
                    metadata[i] = reader.documentInfo['/' + j]
                except :
                    metadata[i] = ""
                data_line.append(metadata[i])

            try :
                sid = reader.documentInfo['/SID']
                metadata["SID"] = sid
                data_line.append(sid)
            except :
                data_line.append("")

            try :
                erledigt = reader.documentInfo['/Erledigt']
            except :
                erledigt = ""
            data_line.append(erledigt)
            metadata["FINISHED"] = erledigt

            #if True :
            #if taskname == selected["taskname"] :
            #    current = True
            #else :
            #    current = False
            #data_line.append(current)

            matched = True
            for i,j in selector.items() :
                if selector[i] not in metadata[i].lower() :
                    matched = False
            
            if matched :
                data.append(data_line)

    return data

# ------ Make the Table Data ------

config_parms = config.config()

pytesseract.pytesseract.tesseract_cmd = config_parms["tasks_tesseract_cmd"]
TESSDATA_PREFIX = config_parms["tasks_tesseract_prefix"]
tessdata_dir_config = config_parms["tasks_tesseract_config"]
source_dir = config_parms["tasks_source_dir"]

selector ={}
data = make_table(num_rows=5, num_cols=5, selector=selector)
# headings = [str(data[0][x])+'     ..' for x in range(len(data[0]))]

headings = []
headings.append("TASK")
for i,j in config_parms["metadata_parameter"].items() :
    headings.append(j)
headings.append("SID")
headings.append("FINISHED")

row_x = 0
col_x = 0

def sort_table(table, cols):
    """ sort a table by multiple columns
        table: a list of lists (or tuple of tuples) where each inner list
               represents a row
        cols:  a list (or tuple) specifying the column numbers to sort by
               e.g. (1,0) would sort by column 1, then by column 0
    """
    for col in reversed(cols):
        try:
            table = sorted(table, key=operator.itemgetter(col))
        except Exception as e:
            sg.popup_error('Error in sort_table', 'Exception in sort_table', e)
    return table

layout = [[sg.Table(values=data[:][:], headings=headings, max_col_width=25,
                    auto_size_columns=True,
                    display_row_numbers=False,
                    justification='left',
                    num_rows=20,
                    alternating_row_color='lightblue',
                    key='-TABLE-',
                    selected_row_colors='red on yellow',
                    enable_events=True,
                    expand_x=True,
                    expand_y=True,
                    enable_click_events=True,           # Comment out to not enable header and other clicks
                    right_click_menu=['&Right', ['&Delete', '&Finish', '&Modify', '&PDF', '&Text','&Register', 'T&o Temp', '&Invoice']],
                    tooltip='This is a table')],
          [sg.Button('New'), sg.Button('Filter'), sg.Button('Help'), sg.Button('Exit')],
          [sg.Text('Cell clicked:'), sg.T(k='-CLICKED-')]]

window = sg.Window('FLUSS Scan Tasks V01', layout,
                   ttk_theme='clam',
                   resizable=True)

# ------ Event Loop ------
while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == "Exit" :
        break
    if event == 'Double':
        for i in range(1, len(data)):
            data.append(data[i])
        window['-TABLE-'].update(values=data[:][:])
    elif event == 'Change Colors':
        window['-TABLE-'].update(row_colors=((8, 'white', 'red'), (9, 'green')))
    elif event == 'Filter':

        print("Filter requested", row_x, col_x, values['-TABLE-'])
        input_layout =  []

        for h in headings : 
            input_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key=f'-{h}-')])
        
        input_layout.append([sg.Button('Apply'), sg.Button('Exit')])

        window_input = sg.Window('Filter by Column', input_layout,
                   ttk_theme='clam',
                   resizable=True, finalize=True)
        window_input['-SID-'].update("0000")
        while True :
            event_input, values_input = window_input.read()
            print(event_input, values_input)
            if event_input == sg.WIN_CLOSED:
                window_input.close()
                break
            if event_input == "Exit":
                window_input.close()
                break
            if event_input == "Apply":
                selector = {}
                try :
                    selector["m"] = values_input["-Mandant-"]
                except IndexError :
                    pass
                window_input.close()
                data = make_table(num_rows=5, num_cols=5, selector=selector)
                window['-TABLE-'].update(data)
                break

    if isinstance(event, tuple):
        # TABLE CLICKED Event has value in format ('-TABLE=', '+CLICKED+', (row,col))
        if event[0] == '-TABLE-':
            row_x = event[2][0]
            col_x = event[2][1]
            if event[2][0] == -1 and event[2][1] != -1:           # Header was clicked and wasn't the "row" column
                col_num_clicked = event[2][1]
                new_table = sort_table(data[:][:],(col_num_clicked, 0))
                window['-TABLE-'].update(new_table)
                data = [data[0]] + new_table
            window['-CLICKED-'].update(f'{event[2][0]},{event[2][1]}')
window.close()

# -------------------------------- mainline -------------------------------------------------------------------------

config_parms = config.config()

pytesseract.pytesseract.tesseract_cmd = config_parms["tasks_tesseract_cmd"]
TESSDATA_PREFIX = config_parms["tasks_tesseract_prefix"]
tessdata_dir_config = config_parms["tasks_tesseract_config"]
source_dir = config_parms["tasks_source_dir"]

print("--- Manage scan tasks V01 ---")
print("Source directory: ", source_dir)
print("Press h for help")

normal = colored.fg("grey_74")
finished = colored.fg("aquamarine_1a") 

selected = {}
selected["taskname"] = ""

no_selected = 0

while False:
    if selected["taskname"] != "": 
        k = input(str(no_selected) + '> ') + " "
    else :
        k = input('> ') + " "

    if len(k) == 0 :
        k = "  "
    # Anzeigen aller offenen Tasks
    if k[0:2] == 'l ' or k[0:2] == "la":
        taskcount = 0
        task_s = {}

        header_line = k[0:2]

        selector = {}
        for i in config_parms["metadata_parameter"].keys() :
            match = re.search(i + r'=(\S+)',k)
            if match :
                selector[i] = match.group(1).lower()
                header_line = header_line + " " + i + "=" + selector[i]  
  
        header_line = 9*"-" + " " + header_line + " "+ (80-10-1-len(header_line))*"-"
        print(header_line)
        for filename in sorted(glob.glob(source_dir + '*.pdf')):
            if os.path.isfile(filename) and ".pdf" in filename :

                taskname = filename.replace(source_dir,"").replace(".pdf","")

                reader = PdfFileReader(filename)

                metadata = {}
                metadata_string = ""
                for i,j in config_parms["metadata_parameter"].items() :
                    try :
                        metadata[i] = reader.documentInfo['/' + j]
                    except :
                        metadata[i] = ""
                    metadata_string += " " + i + "=" + metadata[i]

                try :
                    sid = reader.documentInfo['/SID']
                    metadata_string += " " + "SID" + "=" + sid
                except :
                    pass

                try :
                    erledigt = reader.documentInfo['/Erledigt']
                except :
                    erledigt = ""

                if taskname == selected["taskname"] :
                    current = True
                else :
                    current = False

                matched = True
                for i,j in selector.items() :
                    if selector[i] not in metadata[i].lower() :
                        matched = False

                if (k[0:2] == "l " and erledigt == "" and matched) :
                    taskcount += 1
                    string1 = str(taskcount) + ": " + taskname + metadata_string
                    if erledigt != "" : 
                        string1 += " ERL:" + erledigt

                    if not current :
                        print(stylize(string1, normal))
                    else :                
                        print(stylize(string1, normal + colored.attr("bold")))   
                        no_selected = taskcount          
                    
                    metadata["taskname"] = taskname
                    metadata["erledigt"] = erledigt
                    task_s[taskcount] = metadata
                
                if (k[0:2] == "la" and matched) :
                    taskcount += 1
                    string1 = str(taskcount) + ": " + taskname + metadata_string
                    if erledigt != "" : 
                        string1 += " ERL:" + erledigt
                    
                    if not current :
                        if erledigt : 
                            print(stylize(string1, finished))
                        else :
                            print(stylize(string1, normal))
                    else :                 
                        if erledigt : 
                            print(stylize(string1, finished + colored.attr("bold")))
                        else :
                            print(stylize(string1, normal + colored.attr("bold")))
                        no_selected = taskcount          
                    
                    metadata["erledigt"] = erledigt
                    metadata["taskname"] = taskname
                    task_s[taskcount] = metadata

        print(80*"-")

    # Selektiere Datei zur Bearbeitung
    elif k.replace(" ","").isdigit():
        try : 
            no_selected             = int(k)
            selected                = task_s[int(k)]

            metadata_string = ""
            for i,j in config_parms["metadata_parameter"].items() :
                metadata_string += " " + i + "=" + selected[i]
            print("... Task selected",": ", selected["taskname"], metadata_string)
        except KeyError :
            selected["taskname"] = ""         
            no_selected = 0  

    # Dateien kopieren
    elif k[0] == 'c':
        if selected["taskname"] != "" :

            name_string = ""
            for i,j in config_parms["metadata_parameter"].items() :
                try :
                    name_string += "_" + selected[i]
                except :
                    pass

            os.system("cp " + source_dir + selected["taskname"] + ".pdf" + " " + source_dir + "temp/" + selected["taskname"] + name_string.replace(" ","_") + ".pdf")
            print("... PDF-file of task ", selected["taskname"], " copied ")

    # Dateien löschen
    elif k[0] == 'd':
        if selected["taskname"] != "" :
            l = input("Dateien zu " + selected["taskname"] + " wirklich löschen? (j/n)  >")
            if l.lower() == "j" :
                os.system("rm " + source_dir + selected["taskname"] + ".pdf")
                os.system("rm " + source_dir + selected["taskname"] + ".txt")
                print("... Task ", selected["taskname"], " deleted")
                selected["taskname"] = ""  
                no_selected = 0            

    # Metadata ändern
    elif k[0] == 'm':
        if selected["taskname"] != "" :
            selector = analyze_command_parameters(k[2:])   
            fin = open(source_dir + selected["taskname"] + ".pdf", 'rb')
            reader = PdfFileReader(fin)
            writer = PdfFileMerger()
            writer.append(reader)
            metadata = reader.getDocumentInfo()
            writer.addMetadata(metadata)
            for i,j in selector.items() :
                writer.addMetadata({"/" + config_parms["metadata_parameter"][i]: j})
            fin.close()
            fout = open(source_dir + selected["taskname"] + ".pdf", 'wb')
            writer.write(fout)
            fout.close()
            print("... Metadata of task ", selected["taskname"], " changed ")

    # Task abschließen
    elif k[0] == 'f':
        if selected["taskname"] != "" :
            selector = analyze_command_parameters(k[2:])   
            fin = open(source_dir + selected["taskname"] + ".pdf", 'rb')
            reader = PdfFileReader(fin)
            writer = PdfFileMerger()
            writer.append(reader)
            metadata = reader.getDocumentInfo()
            writer.addMetadata(metadata)

            for i,j in selector.items() :
                writer.addMetadata({"/" + config_parms["metadata_parameter"][i]: j})

            writer.addMetadata({'/Erledigt': datetime.now().strftime("%Y-%m-%d_%H:%M:%S")})
            fin.close()
            fout = open(source_dir + selected["taskname"] + ".pdf", 'wb')
            writer.write(fout)
            fout.close()

            print("... task ", selected["taskname"], " finished")

    # neue Task anlegen
    elif k[0] == 'n':
        selector = analyze_command_parameters(k[2:]) 
        fin = open("dummy.pdf", 'rb')
        reader = PdfFileReader(fin)
        writer = PdfFileMerger()

        writer.append(reader)

        metadata = reader.getDocumentInfo()
        writer.addMetadata(metadata)

        for i,j in selector.items() :
            writer.addMetadata({"/" + config_parms["metadata_parameter"][i]: j})

        fin.close()

        taskname_temp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        fout = open(source_dir + taskname_temp + ".pdf", 'wb')
        writer.write(fout)
        fout.close()
             
        text_f = open(source_dir + taskname_temp + ".txt", "wt")
        text_all = ""

        for i,j in selector.items() :
            text_all += config_parms["metadata_parameter"][i] + ": " + j + "\n"

        text_f.write(text_all)
        text_f.close()

        print("... task ", taskname_temp, " manually created")

    # PDF-Viewer aufrufen
    elif k[0] == 'p':
        if selected["taskname"] != "" :
            subprocess.call("xreader " + source_dir + selected["taskname"] + ".pdf",shell=True)

    # Txt anzeigen
    elif k[0] == 't':
        if selected["taskname"] != "" :
            subprocess.call("xed " + source_dir + selected["taskname"] + ".txt",shell=True)

    # Rechnung analysieren
    elif k[0] == 'r':
        if selected["taskname"] != "" :

            file = open(source_dir + selected["taskname"] + ".txt")
            data = file.readlines()
            file.close()

            print("... Rechnungsdaten zu ", selected["taskname"])
            print("--- Absender ---")
            for d in data[:20] :
                if d.strip() != "" and "Meyer-Martin" not in d and "Martin" not in d and "22395" not in d and "Kirchenheide" not in d and "Herr" not in d  and "Frau" not in d:
                    print(d.replace("\n",""))

            print("--- IBAN ---")
            for d in data :
                reg = re.findall(r"iban", d.lower())
                if len(reg) > 0 :
                    print(d.replace("\n","")) 

            print("--- Betrag ---")
            for d in data :
                reg = re.findall(r"betrag|summe|forderung|gesamt", d.lower())
                if len(reg) > 0 :
                    print(d.replace("\n","")) 

            print("--- Rechnung ---")
            for d in data :
                reg = re.findall(r"rechnung|nr", d.lower())
                if len(reg) > 0 :
                    print(d.replace("\n","")) 

    # register pdf
    elif k[0] == 'y':
        if selected["taskname"] != "" :
            selector = analyze_command_parameters(k[2:]) 
            merger = PdfFileMerger()
            text_all = ""
            taskname_temp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
            for ss in convert_from_path(source_dir + selected["taskname"] + ".pdf") :
                result =  pytesseract.image_to_pdf_or_hocr(ss, lang="deu", config=tessdata_dir_config)
                pdf_file_in_memory = io.BytesIO(result)        
                merger.append(pdf_file_in_memory)   
                text = pytesseract.image_to_string(ss, lang="deu")
                text_all += text

            for i,j in selector.items() :
                writer.addMetadata({"/" + config_parms["metadata_parameter"][i]: j})

            merger.write(source_dir + taskname_temp + ".pdf")
            merger.close()
        
            text_f = open(source_dir + taskname_temp + ".txt", "wt")
            text_f.write(text_all)
            text_f.close()

            os.system("rm " + source_dir + selected["taskname"] + ".pdf")
            os.system("rm " + source_dir + selected["taskname"] + ".txt")

    # closest document
    elif k[0] == 'k':
        if selected["taskname"] != "" :
            selector = analyze_command_parameters(k[2:]) 
        
            text_f = open(source_dir + selected["taskname"] + ".txt", mode="r")
            text_all = text_f.read()
            text_f.close()
            searches = re.findall(r'\S\S\S\S\S+',text_all)[:100]
            level = 0
            filename_l = ""
            text_l = ""
            matches_l = []
            for filename in sorted(glob.glob(source_dir + '*.txt')):
                if os.path.isfile(filename) and ".txt" in filename and source_dir + selected["taskname"] + ".txt" != filename :

                    text_c = open(filename, mode="r")
                    text_all_c = text_c.read()
                    text_c.close()
                    count_c = 0
                    matches = []
                    for s in searches :
                        if s in text_all_c :
                            count_c += 1
                            matches.append(s)
                    if count_c > level :
                        filename_l = filename
                        level = count_c
                        text_l = text_all
                        matches_l = []
                        matches_l = matches
            print(filename_l, level, "\n", matches_l)          

            # merger = PdfFileMerger()
            # text_all = ""
            # taskname_temp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
            # for ss in convert_from_path(source_dir + selected["taskname"] + ".pdf") :
            #     result =  pytesseract.image_to_pdf_or_hocr(ss, lang="deu", config=tessdata_dir_config)
            #     pdf_file_in_memory = io.BytesIO(result)        
            #     merger.append(pdf_file_in_memory)   
            #     text = pytesseract.image_to_string(ss, lang="deu")
            #     text_all += text

            # for i,j in selector.items() :
            #     writer.addMetadata({"/" + config_parms["metadata_parameter"][i]: j})

            # merger.write(source_dir + taskname_temp + ".pdf")
            # merger.close()

            # os.system("rm " + source_dir + selected["taskname"] + ".pdf")
            # os.system("rm " + source_dir + selected["taskname"] + ".txt")

    elif k[0] == 'h':
        print("List of actions ...")
        print("... h=help")
        print("... k=closest document")
        print("... l=list open tasks")
        print("... la=list all tasks")
        print("... p=view pdf")
        print("... t=view txt")
        print("... 99=select task from list")
        print("... m=change metadata")
        print("... f=finish")
        print("... n=new task")
        print("... y=register")
        print("... d=delete task")
        print("... c=copy pdf to /temp dir")
        print("... r=rechnung")
        print("... x=exit")
        print("List of parameters ...")
        for i,j in config_parms["metadata_parameter"].items() :
            print("... " + i + "=" + j)

    # Exit
    elif k[0] == 'x':
        break
    else : 
        print("unknown command")
