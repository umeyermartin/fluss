# importiere alle tifs eines vorgegebenen Verzeichnisses in ein PDF und eine Textdatei

from PyPDF2 import PdfFileMerger
import io
import pytesseract
import xml.dom.minidom
from datetime import datetime
import os
import glob

pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
TESSDATA_PREFIX = '/usr/share/tesseract-ocr/4.00/tessdata/'
tessdata_dir_config = '--tessdata-dir /usr/share/tesseract-ocr/4.00/tessdata/'

print("*** ask for metadata and scan ***")

request_list = []
while True:
    request_dict = {}
    now_tupel = datetime.now() 
    current_date_time = now_tupel.strftime('%Y-%m-%d_%H:%M:%S')
    request_dict["/Zeit"] = current_date_time
    k = input('mandant (m=m-m-cs,g=gerda,p=privat, x=exit)> ')
    if k == "x" :
        break
    if k != "" :
        if k[0] == "m":
            request_dict["/Mandant"] = "m-m-cs"
        elif k[0] == "g":
            request_dict["/Mandant"] = "gerda"
        elif k[0] == "p":
            request_dict["/Mandant"] = "privat"
    k = input('Beschreibung > ')
    if k != "" :
        request_dict["/Beschreibung"] = k
    else :
        request_dict["/Beschreibung"] = ""
    k = input('Aufgabe > ')
    if k != "" :
        request_dict["/Aufgabe"] = k
    else : 
        request_dict["/Aufgabe"] = ""

    k = input('Duplex scan? (J/N) and INSERT DOCUMENT and START > ')
    if k.lower() != "n" :
        source = '"ADF Duplex"'
    else : 
        source = '"ADF Front"'

    request_list.append(request_dict)
    
    print("   *** now scanning ***")

    scan_command = "scanimage --format=tiff --batch=/home/pi/scan_server/temp/" + request_dict["/Zeit"] + "_%d.tif --batch-start=100 --resolution 300 --source " + source + " -p" #  -x 210.01 -y 297.364"
    print(scan_command)
    os.system(scan_command)

    print("*** next document ***")

count = 0
for r in request_list :
    count += 1
    metadata = {}
    metadata["/Mandant"]      = r["/Mandant"]
    metadata["/Beschreibung"] = r["/Beschreibung"]
    metadata["/Aufgabe"]      = r["/Aufgabe"]

    print("*** now merging pages file", count, "***")
    
    src_dir = "/home/pi/scan_server/temp/"
    out_dir = "/home/pi/scan_server/temp_out/"
    pdf_file = out_dir + r["/Zeit"] + ".pdf"
    txt_file = out_dir + r["/Zeit"] + ".txt"
    
    merger = PdfFileMerger()
    
    text_all = ""
    files = False
    
    for filename in sorted(glob.glob(src_dir + '*.tif')):
        if r["/Zeit"] in filename :
            f = filename
            if os.path.isfile(f) and ".tif" in filename and "_deleted.tif" not in filename:
    
                print("   processing ...", filename)
                files = True
                result =  pytesseract.image_to_pdf_or_hocr(f, lang="deu", config=tessdata_dir_config)
                pdf_file_in_memory = io.BytesIO(result)        
                merger.append(pdf_file_in_memory)
    
                text = pytesseract.image_to_string(f, lang="deu")
                text_all += text
    
                os.rename(f, f.replace(".tif","_deleted.tif"))
    
    if files :
        merger.addMetadata(metadata)
    
        merger.write(pdf_file)
        merger.close()
    
        text_f = open(txt_file, "wt")
        text_f.write(text_all)
        text_f.close()
        print("   *** Dateien erstellt: ",  r["/Zeit"] + ".pdf", "und", r["/Zeit"] + ".txt", "***")
    else :
        print("   *** keine Dateien gefunden ***")

print("*** Insgesamt", count, "Dateien erzeugt ***")
