# importiere alle tifs eines vorgegebenen Verzeichnisses in ein PDF und eine Textdatei

from PyPDF2 import PdfFileMerger
import io
import pytesseract
from datetime import datetime
import os
import glob
import config

config_parms = config.config()

print("--- Scan to PDF V06 ---")
print("Nextcloud directory: ", config_parms["scan_nextcloud_dir"])

pytesseract.pytesseract.tesseract_cmd = config_parms["scan_tesseract_cmd"]
TESSDATA_PREFIX = config_parms["scan_tesseract_prefix"]
tessdata_dir_config = config_parms["scan_tesseract_config"]

print("*** ask for metadata and scan ***")

request_list = []
while True:
    request_dict = {}
    now_tupel = datetime.now() 
    current_date_time = now_tupel.strftime('%Y-%m-%d_%H:%M:%S')
    request_dict["/Zeit"] = current_date_time
    exit = False

    for i,j in config_parms["metadata_parameter"].items() :
        input_string = j + " ("
        try :
            values_dict = config_parms["metadata_value_" + i]
            for ii,jj in values_dict.items() :
                input_string += jj + ","
        except (IndexError, KeyError) :
            pass
        input_string += " x=exit)> "

        k = input(input_string)

        if k == "x" :
            exit = True
            break

        value = k
        try :
            values_dict = config_parms["metadata_value_" + i]
            value_temp = ""
            for ii,jj in values_dict.items() :
                if k.strip()[0:1] == ii :
                    value_temp = jj
            if value_temp != "" :
                value = value_temp
            else :
                value = ""
        except (IndexError, KeyError) :
            pass

        if k != "" :
            request_dict["/" + j] = value
        else :
            request_dict["/" + j] = ""

    if exit :
        break    

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
    for i,j in r.items() :
        metadata[i]      = r[i]

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

        z = re.findall('SID:(\d\d\d\d\d\d)', text_all)
        if len(z) > 0
            metadata["/SID"]      = z[0]

        merger.addMetadata(metadata)    
        merger.write(pdf_file)
        merger.close()
    
        text_f = open(txt_file, "wt")
        text_f.write(text_all)
        text_f.close()
        print("   *** Files created: ",  r["/Zeit"] + ".pdf", "und", r["/Zeit"] + ".txt", "***")
    else :
        print("   *** No files found ***")

print("*** Total", count, "files created ***")

move_command = "sudo mv " + config_parms["scan_out_dir"] + "*.* " + config_parms["scan_nextcloud_dir"]

print("*** Files moved to Nextcloud ***")
