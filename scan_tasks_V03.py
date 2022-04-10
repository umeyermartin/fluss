# -*- coding: utf-8 -*-
"""
Created on Tue Oct 30 17:19:20 2018

@author: umeyermartin
"""
# V01 11.02.2022 UMM Neuerstellung

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

while True:
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

    elif k[0] == 'h':
        print("List of actions ...")
        print("... h=help")
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