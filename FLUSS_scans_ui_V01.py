#!/usr/bin/env python
# pflege scan tasks
import PySimpleGUI as sg
import operator

import os
import io
from datetime import datetime
import glob
import subprocess
import sys
import re
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileMerger
from pdf2image import convert_from_path
import pytesseract
import config

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

    index = "TASK"
    
    headings = []
    headings.append("TASK")
    for i,j in config_parms["metadata_parameter"].items() :
        headings.append(j)

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

def runCommand(cmd, timeout=None, window=None):
    p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    output = ''
    for line in p.stdout:
        line = line.decode(errors='replace' if (sys.version_info) < (3, 5) else 'backslashreplace').rstrip()
        output += line
        window.Refresh() if window else None        # yes, a 1-line if, so shoot me
    retval = p.wait(timeout)
    return (retval, output)                         # also return the output just for fun

# ------ Make the Table Data ------

config_parms = config.config()

pytesseract.pytesseract.tesseract_cmd = config_parms["tasks_tesseract_cmd"]
TESSDATA_PREFIX = config_parms["tasks_tesseract_prefix"]
tessdata_dir_config = config_parms["tasks_tesseract_config"]
source_dir = config_parms["tasks_source_dir"]

sg.theme(config_parms["theme"])
selector = {} 
data_dict = {} 

index, headings = make_headings()

data = make_table(index, headings, data_dict, selector)

layout = [[sg.Table(values=data[:][:], headings=headings, max_col_width=25,
                    #auto_size_columns=True,
                    auto_size_columns=False,
                    col_widths = [16,16,40,40,40,40],
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
                    right_click_menu=['&Right', ['&View']],
                    tooltip='Scans')],
          [sg.Button('Scan'), sg.Button('Save')],
          [sg.Text('Cell clicked:'), sg.T(k='-CLICKED-')]]

window = sg.Window('FLUSS Scans', layout,
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
                input_layout.append([sg.Text(h,size=(15, None)) , sg.Combo(dict_values,default_value=dict_values[0],key=f'-{h}-')])
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

    elif event == 'Scan':
        index_key = headings.index(index)

        key_data = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")

        new_layout =  []

        for h in headings : 
            if h == index :
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

        new_layout.append([sg.Radio('Double', group_id="RADIO1", default=True), sg.Radio('Single', group_id="RADIO1", default=False)])      
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

                data_dict_line = {}
                data_dict_line["TASK"] = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
                for i,j in values_new.items() :
                    if str(i).replace("-","") in list(config_parms["metadata_parameter"].values()) :
                        data_dict_line[i.replace("-","")] = j
                    elif i == 0 and j == True : 
                        source = '"ADF Duplex"'
                    elif i == 1 and j == True : 
                        source = '"ADF Front"'
      
                data_dict[data_dict_line["TASK"]] = data_dict_line

                #sg.popup_notify("... now scanning " + " ".join(list(data_dict_line.values())))

                multiline_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))], 
                                    [sg.Button('Exit')]]
                window_multiline = sg.Window('Scan-Result', multiline_layout,
                        ttk_theme='clam',
                        resizable=True, finalize=True)

                scan_command = "scanimage --format=tiff --batch=" + config_parms["scan_src_dir"] + data_dict_line["TASK"] + "_%d.tif --batch-start=100 --resolution 300 --source " + source + " -p -x 210.01 -y 297.364"

                retval, output = runCommand(scan_command)

                window_multiline['-MULTILINE KEY-'].print(scan_command)
                window_multiline['-MULTILINE KEY-'].print(re.sub(r"Progress: \d{1,3}.\d%", "", output))

                window_new.close()

                while True :
                    event_multiline, values_multiline = window_multiline.read()
                    if event_multiline == sg.WIN_CLOSED or event_multiline == "Exit":
                        window_multiline.close()
                        break

                break

        data = make_table(index, headings, data_dict, selector)
        window['-TABLE-'].update(data)

    elif event == 'Save':

        index_key = headings.index(index)

        log_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))],
                     [sg.Button('Exit')]]
        window_log = sg.Window('Log', log_layout, ttk_theme='clam', resizable=True, finalize=True)

        count = 0
        for i,j in data_dict.items() :
            count += 1
            sg.popup_notify("*** now merging pages file " + str(count) + " ***")
            window_log['-MULTILINE KEY-'].print("*** now merging pages file " + str(count) + " ***")

            metadata = {}
            for k,l in j.items() :
                if k in list(config_parms["metadata_parameter"].values()) :
                    metadata["/" + k]      = l

            src_dir = config_parms["scan_src_dir"]
            out_dir = config_parms["scan_out_dir"]
            pdf_file = out_dir + j["TASK"] + ".pdf"
            txt_file = out_dir + j["TASK"] + ".txt"
            
            merger = PdfFileMerger()
            
            text_all = ""
            files = False
            
            for filename in sorted(glob.glob(src_dir + '*.tif')):
                if j["TASK"] in filename :
                    f = filename
                    if os.path.isfile(f) and ".tif" in filename and "_deleted.tif" not in filename:

                        window_log['-MULTILINE KEY-'].print("... merging file " +  f)
                        files = True
                        result =  pytesseract.image_to_pdf_or_hocr(f, lang="deu", config=tessdata_dir_config)
                        pdf_file_in_memory = io.BytesIO(result)        
                        merger.append(pdf_file_in_memory)
            
                        text = pytesseract.image_to_string(f, lang="deu")
                        text_all += text
            
                        os.rename(f, f.replace(".tif","_deleted.tif"))
            
            if files :

                z = re.findall('SID:(\d\d\d\d\d\d)', text_all)
                if len(z) > 0 :
                    metadata["/SID"]      = z[0]
                    window_log['-MULTILINE KEY-'].print("... SID found " +  z[0])

                merger.addMetadata(metadata)    
                merger.write(pdf_file)
                merger.close()
            
                text_f = open(txt_file, "wt")
                text_f.write(text_all)
                text_f.close()
                window_log['-MULTILINE KEY-'].print("... files created " +  j["TASK"] + ".pdf" + " und " + j["TASK"] + ".txt")
            else :
                window_log['-MULTILINE KEY-'].print("... no files found")
        window_log['-MULTILINE KEY-'].print(count, " input files processed")
        while True :
            event_log, values_log = window_log.read()
            if event_log == sg.WIN_CLOSED or event_log == "Exit":
                window_log.close()
                data_dict = {}
                selector  = {}
                data = make_table(index, headings, data_dict, selector)
                window['-TABLE-'].update(data)
                break

    elif event == 'View':

        index_key = headings.index(index)
        confirm_message = ""

        for l in values['-TABLE-'] :
            key_data = data[int(l)][index_key]
            runCommand(config_parms["picture viewer"] + " " + config_parms["scan_src_dir"] + key_data + "_100.tif")
 

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

