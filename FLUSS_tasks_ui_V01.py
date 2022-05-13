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
import FLUSS_pdfviewer

# ------ Some functions to help generate data for the table ------


def make_dict():

    index = "TASK"
    
    headings = []
    headings.append("TASK")
    for i,j in config_parms["metadata_parameter"].items() :
        headings.append(j)
    headings.append("SID")
    headings.append("OPEN")
    headings.append("FINISHED")

    data_dict = {}
    count = 0
    for filename in sorted(glob.glob(source_dir + '*.pdf')):
        if os.path.isfile(filename) and ".pdf" in filename :

            data_dict_line = {}
            count += 1
            taskname = filename.replace(source_dir,"").replace(".pdf","")
            data_dict_line["TASK"] = taskname

            reader = PdfFileReader(filename)

            for i,j in config_parms["metadata_parameter"].items() :
                try :
                    metadata_temp = reader.documentInfo['/' + j]
                except :
                    metadata_temp = ""
                data_dict_line[j] = metadata_temp

            try :
                metadata_temp = reader.documentInfo['/SID']
            except :
                metadata_temp = ""
            data_dict_line["SID"] = metadata_temp

            try :
                metadata_temp = reader.documentInfo['/Erledigt']
            except :
                metadata_temp = ""
            data_dict_line["FINISHED"] = metadata_temp
            if metadata_temp == "" :
                data_dict_line["OPEN"] = "X"
            else :
                data_dict_line["OPEN"] = ""
            data_dict[taskname] = data_dict_line

    return index, headings, data_dict

def make_table(index, headings, data_dict, selector):
    data_table = []
    data_line = []

    for i,j in data_dict.items() :

        match = True
        for k,l in selector.items() :
            if l.lower() not in j[k].lower() :
                match = False
        if match : 
            data_line = []
            for h in headings :
                try :
                    data_line.append(j[h])
                except (KeyError, IndexError) :
                    data_line.append("")
            data_table.append(data_line)

    return data_table

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

# ------ Make the Table Data ------

config_parms = config.config()

pytesseract.pytesseract.tesseract_cmd = config_parms["tasks_tesseract_cmd"]
TESSDATA_PREFIX = config_parms["tasks_tesseract_prefix"]
tessdata_dir_config = config_parms["tasks_tesseract_config"]
source_dir = config_parms["tasks_source_dir"]

sg.theme('Light blue 2')
selector = {"OPEN":"X"} 

index, headings, data_dict = make_dict()

data = make_table(index, headings, data_dict, selector)


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
                    right_click_menu=['&Right', ['&Delete', '&Finish', '&Modify', '&PDF', '&Text','&Register', 'T&o Temp', '&Invoice', '&View']],
                    tooltip='This is a table')],
          [sg.Button('New'), sg.Button('Filter'), sg.Button('Help'), sg.Button('Exit')],
          [sg.Text('Cell clicked:'), sg.T(k='-CLICKED-')]]

window = sg.Window('FLUSS Tasks', layout,
                   ttk_theme='clam',
                   resizable=True)

# ------ Event Loop ------
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == "Exit" :
        break

    elif event == 'Filter':

        input_layout =  []

        for h in headings : 

            try :
                dict_values = list(config_parms["metadata_value_" + h].values())
            except (KeyError,IndexError) :
                dict_values = []
            if len(dict_values) > 0 :
                input_layout.append([sg.Text(h,size=(15, None)) , sg.Combo(dict_values,default_value=dict_values[0:0],key=f'-{h}-')])
            else :
                input_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key=f'-{h}-')])
        
        input_layout.append([sg.Button('Apply'), sg.Button('Exit')])

        window_input = sg.Window('Filter by Column', input_layout,
                   ttk_theme='clam',
                   resizable=True, finalize=True)
        for i,j in selector.items() :
            window_input["-" + i + "-"].update(j)
        while True :
            event_input, values_input = window_input.read()
            if event_input == sg.WIN_CLOSED:
                window_input.close()
                break
            if event_input == "Exit":
                window_input.close()
                break
            if event_input == "Apply":
                selector = {}
                for i,j in values_input.items() :
                    if j != "" :
                        selector[i.replace("-","")] = j
                window_input.close()
                data = make_table(index, headings, data_dict, selector)
                window['-TABLE-'].update(data)
                break

    elif event == 'New':
        index_key = headings.index(index)

        key_data = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")

        new_layout =  []

        for h in headings : 
            if h == index :
                pass
            elif h == "OPEN" :
                pass
            elif h == "FINISHED" :
                pass
            else :
                try :
                    dict_values = list(config_parms["metadata_value_" + h].values())
                except (KeyError,IndexError) :
                    dict_values = []
                if len(dict_values) > 0 :
                    new_layout.append([sg.Text(h,size=(15, None)) , sg.Combo(dict_values,default_value=dict_values[0:0],key=f'-{h}-')])
                else :
                    new_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key=f'-{h}-')])
        
        new_layout.append([sg.Button('Apply'), sg.Button('Exit')])

        window_new = sg.Window('New', new_layout,
                ttk_theme='clam',
                resizable=True, finalize=True)
        for i in headings :
            if i != index and i != "OPEN" and i != "FINISHED":
                window_new["-" + i + "-"].update("")
        while True :
            event_new, values_new = window_new.read()
            if event_new == sg.WIN_CLOSED:
                window_new.close()
                break
            if event_new == "Exit":
                window_new.close()
                break
            if event_new == "Apply": 

                for i,j in values_new.items() :
                    if i.replace("-","") in list(config_parms["metadata_parameter"].values()) :
                        writer.addMetadata({"/" + i.replace("-",""): j})
                    elif i.replace("-","") == "SID" :
                        writer.addMetadata({"/" + "SID": j})
                    elif i.replace("-","") == "FINISHED" :
                        writer.addMetadata({"/" + "Erledigt": j})
                fin.close()

                taskname_temp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
                fout = open(source_dir + taskname_temp + ".pdf", 'wb')
                writer.write(fout)
                fout.close()
                    
                text_f = open(source_dir + taskname_temp + ".txt", "wt")
                text_all = ""

                for i,j in values_new.items() :
                    if i.replace("-","") in list(config_parms["metadata_parameter"].values()) :
                        text_all += "/" + i.replace("-","") + ": " + j + "\n"
                    elif i.replace("-","") == "SID" :
                        text_all += "/" + "SID" + ": " + j + "\n"
                    elif i.replace("-","") == "FINISHED" :
                        text_all += "/" + "Erledigt" + ": " + j + "\n"

                text_f.write(text_all)
                text_f.close()

                window_new.close()
                break

        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'Modify':

        index_key = headings.index(index)

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]

            modify_layout =  []

            for h in headings : 
                if h == index :
                    modify_layout.append([sg.Text(h,size=(15, None)) , sg.Text(key_data)])
                elif h == "OPEN" :
                    modify_layout.append([sg.Text(h,size=(15, None)) , sg.Text(data_dict[key_data]["OPEN"])])
                else :
                    try :
                        dict_values = list(config_parms["metadata_value_" + h].values())
                    except (KeyError,IndexError) :
                        dict_values = []
                    if len(dict_values) > 0 :
                        modify_layout.append([sg.Text(h,size=(15, None)) , sg.Combo(dict_values,default_value=dict_values[0:0],key=f'-{h}-')])
                    else :
                        modify_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key=f'-{h}-')])
            
            modify_layout.append([sg.Button('Apply'), sg.Button('Exit')])

            window_modify = sg.Window('Modify', modify_layout,
                    ttk_theme='clam',
                    resizable=True, finalize=True)
            for i,j in data_dict[key_data].items() :
                if i != index and i != "OPEN":
                    window_modify["-" + i + "-"].update(j)
            while True :
                event_modify, values_modify = window_modify.read()
                if event_modify == sg.WIN_CLOSED:
                    window_modify.close()
                    break
                if event_modify == "Exit":
                    window_input.close()
                    break
                if event_modify == "Apply": 
                    fin = open(source_dir + key_data + ".pdf", 'rb')
                    reader = PdfFileReader(fin)
                    writer = PdfFileMerger()
                    writer.append(reader)
                    metadata = reader.getDocumentInfo()
                    writer.addMetadata(metadata)
                    for i,j in values_modify.items() :
                        if i.replace("-","") in list(config_parms["metadata_parameter"].values()) :
                            writer.addMetadata({"/" + i.replace("-",""): j})
                        elif i.replace("-","") == "SID" :
                            writer.addMetadata({"/" + "SID": j})
                        elif i.replace("-","") == "FINISHED" :
                            writer.addMetadata({"/" + "Erledigt": j})
                    fin.close()
                    fout = open(source_dir + key_data + ".pdf", 'wb')
                    writer.write(fout)
                    fout.close()

                    window_modify.close()
                    break

        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'Register':

        index_key = headings.index(index)

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]

            register_layout = []

            for h in headings : 
                if h == index :
                    register_layout.append([sg.Text(h,size=(15, None)) , sg.Text(key_data)])
                elif h == "OPEN" :
                    register_layout.append([sg.Text(h,size=(15, None)) , sg.Text(data_dict[key_data]["OPEN"])])
                else :
                    try :
                        dict_values = list(config_parms["metadata_value_" + h].values())
                    except (KeyError,IndexError) :
                        dict_values = []
                    if len(dict_values) > 0 :
                        register_layout.append([sg.Text(h,size=(15, None)) , sg.Combo(dict_values,default_value=dict_values[0:0],key=f'-{h}-')])
                    else :
                        register_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key=f'-{h}-')])
            
            register_layout.append([sg.Button('Apply'), sg.Button('Exit')])

            window_register = sg.Window('Register', register_layout,
                    ttk_theme='clam',
                    resizable=True, finalize=True)
            for i,j in data_dict[key_data].items() :
                if i != index and i != "OPEN":
                    window_register["-" + i + "-"].update(j)
            while True :
                event_register, values_register = window_register.read()

                if event_register == sg.WIN_CLOSED:
                    window_register.close()
                    break
                if event_register == "Exit":
                    window_input.close()
                    break
                if event_register == "Apply": 
                    merger = PdfFileMerger()
                    text_all = ""
                    taskname_temp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
                    for ss in convert_from_path(source_dir + key_data + ".pdf") :
                        result =  pytesseract.image_to_pdf_or_hocr(ss, lang="deu", config=tessdata_dir_config)
                        pdf_file_in_memory = io.BytesIO(result)        
                        merger.append(pdf_file_in_memory)   
                        text = pytesseract.image_to_string(ss, lang="deu")
                        text_all += text

                    for i,j in values_register.items() :
                        if i.replace("-","") in list(config_parms["metadata_parameter"].values()) :
                            merger.addMetadata({"/" + i.replace("-",""): j})
                        elif i.replace("-","") == "SID" :
                            merger.addMetadata({"/" + "SID": j})
                        elif i.replace("-","") == "FINISHED" :
                            merger.addMetadata({"/" + "Erledigt": j})

                    merger.write(source_dir + taskname_temp + ".pdf")
                    merger.close()
                
                    text_f = open(source_dir + taskname_temp + ".txt", "wt")
                    text_f.write(text_all)
                    text_f.close()

                    os.system("rm " + source_dir + key_data + ".pdf")
                    os.system("rm " + source_dir + key_data + ".txt")

                    window_register.close()
                    break

        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'Delete':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]

            result = sg.popup_ok_cancel("Task " + key_data + ' löschen?')  # Shows OK and Cancel buttons

            if result == "OK" :
                os.system("rm " + source_dir + key_data + ".pdf")
                os.system("rm " + source_dir + key_data + ".txt")
                confirm_message += "... Task " + key_data + " deleted\n"

        sg.popup(confirm_message)
        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'To Temp':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            
            name_string = ""
            for i,j in config_parms["metadata_parameter"].items() :
                try :
                    name_string += "_" + data_dict[key_data][j]
                except :
                    pass

            os.system("cp " + source_dir + key_data + ".pdf" + " " + source_dir + "temp/" + key_data + name_string.replace(" ","_") + ".pdf")
            confirm_message += "... PDF-file of task " + key_data + " copied \n"    

        sg.popup(confirm_message)

    elif event == 'Finish':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            fin = open(source_dir + key_data + ".pdf", 'rb')
            reader = PdfFileReader(fin)
            writer = PdfFileMerger()
            writer.append(reader)
            metadata = reader.getDocumentInfo()
            writer.addMetadata(metadata)
            writer.addMetadata({'/Erledigt': datetime.now().strftime("%Y-%m-%d_%H:%M:%S")})
            fin.close()
            fout = open(source_dir + key_data + ".pdf", 'wb')
            writer.write(fout)
            fout.close()

        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'PDF':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            subprocess.call(config_parms["PDF viewer"] + " " + source_dir + key_data + ".pdf",shell=True)
 
    elif event == 'View':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            FLUSS_pdfviewer.pdfviewer(source_dir + key_data + ".pdf")
 
    elif event == 'Text':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            subprocess.call("xed " + source_dir + key_data + ".txt",shell=True)

    elif event == 'Invoice':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]

            file = open(source_dir + key_data + ".txt")
            data = file.readlines()
            file.close()

            invoice_layout = [[sg.Multiline(key='-MULTILINE KEY-', disabled=True, size=(70,40))]]
            window_invoice = sg.Window('Rechnung', invoice_layout,
                    ttk_theme='clam',
                    resizable=True, finalize=True)

            window_invoice['-MULTILINE KEY-'].print("Rechnungsdaten zu ", key_data)
            window_invoice['-MULTILINE KEY-'].print("----- Absender -----")
            for d in data[:20] :
                if d.strip() != "" and "Meyer-Martin" not in d and "Martin" not in d and "22395" not in d and "Kirchenheide" not in d and "Herr" not in d  and "Frau" not in d:
                    window_invoice['-MULTILINE KEY-'].print(d.replace("\n",""))

            window_invoice['-MULTILINE KEY-'].print("----- IBAN -----")
            for d in data :
                reg = re.findall(r"iban", d.lower())
                if len(reg) > 0 :
                    window_invoice['-MULTILINE KEY-'].print(d.replace("\n","")) 

            window_invoice['-MULTILINE KEY-'].print("----- Betrag -----")
            for d in data :
                reg = re.findall(r"betrag|summe|forderung|gesamt", d.lower())
                if len(reg) > 0 :
                    window_invoice['-MULTILINE KEY-'].print(d.replace("\n","")) 

            window_invoice['-MULTILINE KEY-'].print("----- Rechnung -----")
            for d in data :
                reg = re.findall(r"rechnung|nr", d.lower())
                if len(reg) > 0 :
                    window_invoice['-MULTILINE KEY-'].print(d.replace("\n","")) 

            while True :
                event_invoice, values_invoice = window_invoice.read()
                if event_invoice == sg.WIN_CLOSED:
                    window_invoice.close()
                    break
                if event_invoice == "Exit":
                    window_invoice.close()
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
                data = new_table
            window['-CLICKED-'].update(f'{event[2][0]},{event[2][1]}')
window.close()

