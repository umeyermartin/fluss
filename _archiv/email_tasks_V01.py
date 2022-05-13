# -*- coding: utf-8 -*-
"""
Created on Tue Oct 30 17:19:20 2018

@author: umeyermartin
"""
# V01 11.02.2022 UMM Neuerstellung

import os
import io
import sys
from datetime import datetime
from datetime import date
import glob
import subprocess
import re
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileMerger
from pdf2image import convert_from_path
import pytesseract
import colored  
from colored import stylize
import readline
import config
from imap_tools import MailBox, AND
import pdfkit


def analyze_command_parameters(cmd_string) :

    split_strings = []
    for i in config_parms["emails_parameter"].keys() :
        split_strings.append(i + "=")

    result = re.split("("+ "|".join(split_strings) + ")", cmd_string)

    selector = {}    

    for i in range(len(result)-1) :
        for j in config_parms["emails_parameter"].keys() :
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

def fill_value(dictx, keyx, default) :

    try : 
        result = dictx[keyx]
    except :
        result = default

    return(result)

# -------------------------------- mainline -------------------------------------------------------------------------

print("Manage email tasks V01", "\n", "Press h for help")

config_parms = config.config()

print("--- Manage email tasks V01 ---")
print("eMail User: ", config_parms["emails_user"])
print("Press h for help")

normal = colored.fg("grey_74")
finished = colored.fg("aquamarine_1a") 

selected = {}
selected["email"] = ""

no_selected = 0

mb = MailBox(config_parms["emails_server"]).login(config_parms["emails_user"], config_parms["emails_password"]) 

while True:
    if selected["email"] != "": 
        k = input(str(no_selected) + '> ') + " "
    else :
        k = input('> ') + " "

    if len(k) == 0 :
        k = "  "
    # Anzeigen aller eMails
    if k[0:2] == 'l ' or k[0:2] == "la":
        emailcount = 0
        email_s = {}

        header_line = k[0:2]

        selector = {}
        for i in config_parms["emails_parameter"].keys() :
            match = re.search(i + r'=(\S+)',k)
            if match :
                selector[i] = match.group(1).lower()
                header_line = header_line + " " + i + "=" + selector[i]  
  
        header_line = 9*"-" + " " + header_line + " "+ (80-10-1-len(header_line))*"-"
        print(header_line)

        selector_m = fill_value(selector, "m", "")
        selector_s = fill_value(selector, "s", "")
        selector_d = fill_value(selector, "d", "1900-01-01")
        selector_e = fill_value(selector, "e", "9999-12-31")
        selector_l = int(fill_value(selector, "l", "50"))

        messages = mb.fetch(criteria=AND(from_= selector_m, subject=selector_s, date_gte = date.fromisoformat(selector_d), date_lt = date.fromisoformat(selector_e)),limit = selector_l,reverse = True, bulk=True)

        for msg in messages:

            emailname = msg.date.strftime("%Y-%m-%d_%H:%M:%S")

            metadata = {}
            metadata_string = ""
            for i,j in config_parms["emails_parameter"].items() :
                if i == "s" :
                    metadata[i] = msg.subject
                    metadata_string += "  " + metadata[i]
                if i == "d" :
                    metadata[i] = msg.date.strftime("%Y-%m-%d_%H:%M:%S")
                    metadata_string += "  " + metadata[i]
                if i == "m" :
                    metadata[i] = msg.from_
                if i == "p" :
                    metadata[i] = ""
                    for att in msg.attachments:  # list: imap_tools.MailAttachment  
                        if att.content_type == "application/pdf" :
                            metadata[i] = "X"  
                    if metadata[i] != "" :                         
                        metadata_string += "  " + "PDF"

            if emailname == selected["email"] :
                current = True
            else :
                current = False

            emailcount += 1
            string1 = str(emailcount) + ": " + emailname + metadata_string

            if not current :
                print(stylize(string1, normal))
            else :                
                print(stylize(string1, normal + colored.attr("bold")))   
                no_selected = emailcount              
                
            metadata["emailname"] = emailname
            metadata["msg"]       = msg

            email_s[emailcount] = metadata

        print(80*"-")

    # Selektiere Datei zur Bearbeitung
    elif k.replace(" ","").isdigit():
        try : 
            no_selected             = int(k)
            selected                = email_s[int(k)]

            metadata_string = ""

            count = 0
            for i,j in config_parms["emails_parameter"].items() :
                count += 1
                if count <= 3 :
                    metadata_string += " " + i + "=" + metadata[i]

            print("... eMail selected",": ", selected["email"], metadata_string)
        except KeyError :
            selected["email"] = ""         
            no_selected = 0  

    # PDF-Attachments exportieren
    elif k[0] == 'p':
        l = input("Dateien exportieren mit Prefix? (leer = abbrechen)  >")
        if l.strip() != "" :
            for i,j in email_s.items() :
                for att in j["msg"].attachments:  # list: imap_tools.MailAttachment
                    if att.content_type == "application/pdf" :
                        filename = config_parms["emails_target_dir"] + l.strip() + "_" + att.filename
                        with open(filename, 'wb') as f:
                            f.write(att.payload)  
                            print("... PDF attachment ", filename, " copied ")     

    # eMail-Texte exportieren
    elif k[0] == 't':
        l = input("eMail-Texte als PDF exportieren mit Prefix? (leer = abbrechen)  >")
        if l.strip() != "" :
            for i,j in email_s.items() :
                filename = config_parms["emails_target_dir"] + l.strip() + "_" + j["d"] + "_" + j["msg"].subject.replace(" ","_") + ".pdf"
                pdfkit.from_string(msg.html, filename)
                print("... PDF Text ", filename, " copied ")     


    elif k[0] == 'h':
        print("List of actions ...")
        print("... h=help")
        print("... l=list open tasks")
        print("... p=export pdf")
        print("... t=export txt")
        print("....99=select task from list")
        print("... x=exit")
        print("List of parameters ...")
        print("... m=email sender")
        print("... s=subject")
        print("... d=from date")
        print("... e=to date")
        print("... l=limit number of mails")

    # Exit
    elif k[0] == 'x':
        break
    else : 
        print("unknown command")