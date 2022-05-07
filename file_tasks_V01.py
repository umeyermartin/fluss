# -*- coding: utf-8 -*-
"""
Created on Tue Oct 30 17:19:20 2018

@author: umeyermartin
"""
# V01 11.02.2022 UMM Neuerstellung

import os
from datetime import datetime
import glob
import subprocess
import re
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileMerger
import colored  
from colored import stylize
import config
import pikepdf
from PIL import Image

def analyze_command_parameters(cmd_string) :

    split_strings = []
    for i in config_parms["files_parameter"].keys() :
        split_strings.append(i + "=")

    result = re.split("("+ "|".join(split_strings) + ")", cmd_string)

    selector = {}    

    for i in range(len(result)-1) :
        for j in config_parms["files_parameter"].keys() :
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
source_dir = config_parms["files_source_dir"]
target_dir = config_parms["files_target_dir"]

print("--- Manage file tasks V01 ---")
print("Source Directory: ", source_dir)
print("Press h for help")

normal = colored.fg("grey_74")
finished = colored.fg("aquamarine_1a") 

selected = {}
selected["filename"] = ""

no_selected = 0

while True:
    if selected["filename"] != "": 
        k = input(str(no_selected) + '> ') + " "
    else :
        k = input('> ') + " "

    if len(k) == 0 :
        k = "  "
    # Anzeigen aller offenen Tasks
    if k[0:2] == 'l ' or k[0:2] == "la":
        filecount = 0
        file_s = {}

        header_line = k[0:2]

        selector = {}
        for i in config_parms["files_parameter"].keys() :
            match = re.search(i + r'=(\S+)',k)
            if match :
                selector[i] = match.group(1).lower()
                header_line = header_line + " " + i + "=" + selector[i]  

        header_line = 9*"-" + " " + header_line + " "+ (80-10-1-len(header_line))*"-"
        for f in sorted(glob.glob(source_dir+"*.*")):
            if os.path.isfile(f) :

                filename = f.replace(source_dir,"")

                metadata = {}
                metadata_string = ""
                for i,j in config_parms["files_parameter"].items() :
                    if i == "f" :
                        metadata[i] = filename

                if filename == selected["filename"] :
                    current = True
                else :
                    current = False

                matched = True
                for i,j in selector.items() :
                    if selector[i] not in metadata[i].lower() :
                        matched = False

                if (k[0:2] == "l " and matched) :
                    filecount += 1
                    string1 = str(filecount) + ": " + filename + metadata_string

                    if not current :
                        print(stylize(string1, normal))
                    else :                
                        print(stylize(string1, normal + colored.attr("bold")))   
                        no_selected = filecount          
                    
                    metadata["filename"] = filename
                    file_s[filecount] = metadata
                
                if (k[0:2] == "la" and matched) :
                    filecount += 1
                    string1 = str(filecount) + ": " + filename + metadata_string
                    
                    if not current :
                        print(stylize(string1, normal))
                    else :                 
                        print(stylize(string1, normal + colored.attr("bold")))
                        no_selected = filecount          
                    
                    metadata["filename"] = filename
                    file_s[filecount] = metadata

        print(80*"-")

    # Selektiere Datei zur Bearbeitung
    elif k.replace(" ","").isdigit():
        try : 
            no_selected             = int(k)
            selected                = file_s[int(k)]

            metadata_string = ""
            print("... Task selected",": ", selected["filename"], metadata_string)
        except KeyError :
            selected["filename"] = ""         
            no_selected = 0  

    # Dateien kopieren
    elif k[0] == 'c':
        if selected["filename"] != "" :

            name_string = ""
            for i,j in config_parms["files_parameter"].items() :
                try :
                    name_string += "_" + selected[i]
                except :
                    pass

            os.system("cp " + source_dir + selected["filename"] + " " + target_dir + selected["filename"])
            print("... PDF file ", selected["filename"], " copied ")

    # Dateien löschen
    elif k[0] == 'd':
        if selected["filename"] != "" :
            l = input("Datei " + selected["filename"] + " wirklich löschen? (j/n)  >")
            if l.lower() == "j" :
                os.system("rm " + source_dir + selected["filename"])
                print("... File ", selected["filename"], " deleted")
                selected["filename"] = ""  
                no_selected = 0            

    # PDF-Viewer aufrufen
    elif k[0] == 'p':
        if selected["filename"] != "" and selected["filename"].endswith(".pdf") :
            subprocess.call("xreader " + source_dir + selected["filename"],shell=True)

    # Txt anzeigen
    elif k[0] == 't':
        if selected["filename"] != "" and selected["filename"].endswith(".txt") :
            subprocess.call("xed " + source_dir + selected["filename"],shell=True)

    # PDFS mergen
    elif k[0] == 'm':

        count_files = len(file_s)
        l = input(str(count_files) + " Dateien mergen? (j/n)  >")
        if l.lower() == "j" :

            merger = PdfFileMerger(strict=False)

            for i,j in file_s.items() :
                if not j["filename"].endswith(".pdf"):
                    path_with_file = os.path.join(source_dir, j["filename"])
                    image_temp = Image.open(path_with_file)
                    image_temp.save(path_with_file.replace(".png",".pdf"))

            for f in sorted(glob.glob(source_dir+"*.*")):
                if os.path.isfile(f) :

                    filename = f.replace(source_dir,"")

                    if filename.endswith(".pdf"):
                        path_with_file = f

                        my_pdf = pikepdf.Pdf.open(path_with_file, allow_overwriting_input=True)
                        my_pdf.save(path_with_file)
                        merger.append(path_with_file, bookmark=filename, pages=None, import_bookmarks=False )

                        print("... PDF-file ", filename, " merged")

            target_file = datetime.now().strftime("%Y-%m-%d_%H:%M:%S") + "_merged.pdf"
            merger.write(target_dir + target_file)
            print("... target file ", target_file)

    elif k[0] == 'h':
        print("List of actions ...")
        print("... h=help")
        print("... l=list files")
        print("... m=merge pdfs")
        print("... p=view pdf")
        print("... t=view txt")
        print("... 99=select file from list")
        print("... d=delete file")
        print("... c=copy file to target")
        print("... x=exit")
        print("List of parameters ...")
        for i,j in config_parms["files_parameter"].items() :
            print("... " + i + "=" + j)

    # Exit
    elif k[0] == 'x':
        break
    else : 
        print("unknown command")