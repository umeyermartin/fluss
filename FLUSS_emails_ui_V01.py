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
import config
from datetime import date
from imap_tools import MailBox, AND
import pdfkit
import string

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

def make_headings():

    index = "DATE"
    
    headings = []
    headings.append("DATE")
    headings.append("SUBJECT")
    headings.append("SENDER")
    headings.append("PDFS")

    return index, headings

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
selector = {} 
data_dict = {} 

index, headings = make_headings()

data = make_table(index, headings, data_dict, selector)

layout = [[sg.Table(values=data[:][:], headings=headings, max_col_width=25,
                    auto_size_columns=False,
                    col_widths = [26,40,40,6],
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
                    right_click_menu=['&Right', ['export &PDF', 'export &Text',]],

                    tooltip='Scans')],
          [sg.Button('Filter'), sg.Button('Save')],
          [sg.Text('Cell clicked:'), sg.T(k='-CLICKED-')]]

window = sg.Window('FLUSS Emails', layout,
                   ttk_theme='clam',
                   resizable=True)

mb = MailBox(config_parms["emails_server"]).login(config_parms["emails_user"], config_parms["emails_password"]) 

# ------ Event Loop ------
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == "Exit" :
        break

    elif event == 'Filter':

        input_layout =  []

        for h in headings : 

            if h != "DATE" :
                input_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key=f'-{h}-')])
            else :
                input_layout.append([sg.Text(h,size=(15, None)) , sg.Input(key='-DATEFROM-', size=(20,1)), sg.CalendarButton('from date',  target='-DATEFROM-', locale='de_DE', begin_at_sunday_plus=1, format=('%Y-%m-%d') ), sg.Input(key='-DATETO-', size=(20,1)), sg.CalendarButton('to date',  target='-DATETO-',  locale='de_DE', begin_at_sunday_plus=1, format=('%Y-%m-%d') )])
        
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
                print("APPLY")
                try : 
                    selector_m = values_input["-SENDER-"]
                except (KeyError, IndexError) :
                    selector_m = ""
                try : 
                    selector_s = values_input["-SUBJECT-"]
                except (KeyError, IndexError) :
                    selector_s = ""
                try : 
                    selector_d = values_input["-DATEFROM-"] if values_input["-DATEFROM-"] != "" else "1900-01-01"
                except (KeyError, IndexError) :
                    selector_d = "1900-01-01"
                try : 
                    selector_e = values_input["-DATETO-"] if values_input["-DATETO-"] != "" else "9999-12-31"
                except (KeyError, IndexError) :
                    selector_e = "9999-12-31"
                selector_l = 50

                print(date.fromisoformat(selector_d), date.fromisoformat(selector_e), selector_s)

                messages = mb.fetch(criteria=AND(from_= selector_m, subject=selector_s, date_gte = date.fromisoformat(selector_d), date_lt = date.fromisoformat(selector_e)),limit = selector_l,reverse = True, bulk=True)
                data_dict = {}
                print(messages)
                for msg in messages:

                    data_dict_line = {}
                    data_dict_line["DATE"] = msg.date.strftime("%Y-%m-%d_%H:%M:%S")
                    print(msg.subject)
                    data_dict_line["SUBJECT"] = msg.subject.encode("ascii", "ignore").decode()
                    data_dict_line["SUBJECT"] = ''.join(s for s in msg.subject if s in string.printable or s in list("äöüÄÖÜß"))         
                    data_dict_line["SENDER"] = msg.from_
                    data_dict_line["msg"]    = msg

                    data_dict_line["PDFS"] = ""
                    for att in msg.attachments:  # list: imap_tools.MailAttachment  
                        if att.content_type == "application/pdf" :
                            data_dict_line["PDFS"] = "X"

                    data_dict[data_dict_line["DATE"]] = data_dict_line

                window_input.close()
                selector = {}
                data = make_table(index, headings, data_dict, selector)

                print(data_dict)
                window['-TABLE-'].update(data)
                break

    elif event == 'export PDF':

        index_key = headings.index(index)
        confirm_message = ""
        log_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))]]
        window_log = sg.Window('Log', log_layout, ttk_theme='clam', resizable=True, finalize=True)

        for l in values['-TABLE-'] :
            ll = "01"
            key_data = data[int(l)][index_key]
            j = data_dict[key_data]
            for att in j["msg"].attachments:  # list: imap_tools.MailAttachment
                if att.content_type == "application/pdf" :
                    filename = config_parms["emails_target_dir"] + ll.strip() + "_" + att.filename
                    with open(filename, 'wb') as f:
                        f.write(att.payload) 

                    window_log['-MULTILINE KEY-'].print("... PDF attachment ", filename, " copied ") 

        while True :
            event_log, values_log = window_log.read()
            if event_log == sg.WIN_CLOSED or event_log == "Exit":
                window_log.close()
                break
        break

    elif event == 'export Text':

        index_key = headings.index(index)
        confirm_message = ""
        log_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))]]
        window_log = sg.Window('Log', log_layout, ttk_theme='clam', resizable=True, finalize=True)

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            ll = "01"
            j = data_dict[key_data] 

            filename = config_parms["emails_target_dir"] + ll.strip() + "_" + j["DATE"] + "_" + j["msg"].subject.replace(" ","_") + ".pdf"
            #pdfkit.from_string(msg.html, filename)
            print("TEXT", msg.text, "\nHTML", msg.html)
            if msg.html != "" :
                pdfkit.from_string(msg.html, filename)
            else :
                options = {'encoding': "utf8"}
                prefix = '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
                pdfkit.from_string(prefix + msg.text.replace("\n","<br>"), filename, options=options)

            window_log['-MULTILINE KEY-'].print("... PDF Text ", filename, " copied ") 

        while True :
            event_log, values_log = window_log.read()
            if event_log == sg.WIN_CLOSED or event_log == "Exit":
                window_log.close()
                break

    elif event == 'View':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            subprocess.call(config_parms["picture viewer"] + " " + config_parms["scan_src_dir"] + key_data + "_100.tif",shell=True)
 

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

