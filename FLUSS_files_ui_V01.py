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
import pikepdf
from PIL import Image
import img2pdf

# ------ Some functions to help generate data for the table ------


def make_dict():

    index = "FILE"
    
    headings = []
    headings.append("FILE")

    data_dict = {}
    count = 0
    for filename in sorted(glob.glob(source_dir + '*.*')):
        if os.path.isfile(filename):

            data_dict_line = {}
            count += 1
            ff = filename.replace(source_dir,"")
            data_dict_line["FILE"] = ff
            data_dict[ff] = data_dict_line

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
source_dir = config_parms["files_source_dir"]
target_dir = config_parms["files_target_dir"]

sg.theme(config_parms["theme"])
selector = {} 

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
                    right_click_menu=['&Right', ['&Convert', '&Delete', '&Merge', '&PDF', 'T&o Temp']],
                    tooltip='FLUSS files')],
          [sg.Button('Filter'), sg.Button('Refresh'), sg.Button('Help'), sg.Button('Exit')],
          [sg.Text('Cell clicked:'), sg.T(k='-CLICKED-')]]

window = sg.Window('FLUSS Files', layout,
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

    elif event == 'Convert':

        index_key = headings.index(index)

        log_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))]]
        window_log = sg.Window('Log', log_layout, ttk_theme='clam', resizable=True, finalize=True)

        merger = PdfFileMerger(strict=False)

        for l in values['-TABLE-'] :
            filename = data[int(l)][index_key]
            path_with_file = source_dir + filename

            image = Image.open(path_with_file)
            pdf_bytes = img2pdf.convert(image.filename)
            file = open(path_with_file + ".pdf", "wb")
            file.write(pdf_bytes)
            image.close()
            file.close()
            window_log['-MULTILINE KEY-'].print("... image-file ", filename, " converted")

        while True :
            event_log, values_log = window_log.read()
            if event_log == sg.WIN_CLOSED or event_log == "Exit":
                window_log.close()
                break

    elif event == 'Merge':

        index_key = headings.index(index)

        log_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))]]
        window_log = sg.Window('Log', log_layout, ttk_theme='clam', resizable=True, finalize=True)

        merger = PdfFileMerger(strict=False)

        for l in values['-TABLE-'] :
            filename = data[int(l)][index_key]
            if filename.endswith(".pdf"):
                path_with_file = source_dir + filename

                my_pdf = pikepdf.Pdf.open(path_with_file, allow_overwriting_input=True)
                my_pdf.save(path_with_file)
                merger.append(path_with_file, bookmark=filename, pages=None, import_bookmarks=False )

                window_log['-MULTILINE KEY-'].print("... PDF-file ", filename, " merged")

        target_file = datetime.now().strftime("%Y-%m-%d_%H:%M:%S") + "_merged.pdf"
        merger.write(target_dir + target_file)
        window_log['-MULTILINE KEY-'].print("... target file ", target_file)

        while True :
            event_log, values_log = window_log.read()
            if event_log == sg.WIN_CLOSED or event_log == "Exit":
                window_log.close()
                break

    elif event == 'Delete':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]

            result = sg.popup_ok_cancel("File " + key_data + ' löschen?')  # Shows OK and Cancel buttons

            if result == "OK" :
                os.system("rm " + source_dir + key_data)
                confirm_message += "... Task " + key_data + " deleted\n"

        index, headings, data_dict = make_dict()
        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'Refresh':

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

    elif event == 'PDF':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            subprocess.call(config_parms["PDF viewer"] + " " + source_dir + key_data,shell=True)
 
    elif event == 'Text':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            subprocess.call("xed " + source_dir + key_data,shell=True)

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
                    window_modify.close()
                    break
                if event_invoice == "Exit":
                    window_input.close()
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

