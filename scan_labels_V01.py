# Erzeuge Paginierlabel

from fpdf import FPDF
from datetime import datetime
import config

class PDF(FPDF):
    def header(self):
        pass

# -------------------------------- mainline -------------------------------------------------------------------------

config_parms = config.config()
target_dir = config_parms["labels_target_dir"]

print("--- Print scan labels V01 ---")
print("Target Directory: ", target_dir)
print("")

selected = {}
selected["filename"] = ""

no_selected = 0

k = input("number of first label > ") 
first_label = int(k)
k = input("number of last label  > ") 
last_label = int(k)

pdf = PDF()

pdf.alias_nb_pages()
pdf.add_page("p")
pdf.set_auto_page_break(True, margin = 0.0)
pdf.set_margins(0,0)

page_top_offset = 0             # zusätzlicher versatz der Seite horizontal
page_top_boarder= 4             # oberer Rand des Etikettenblattes
page_left_offset = 0            # zusätzlicher versatz am linken Rand
page_height     = 297           # Seitenhöhe
page_width      = 210           # Seitenbreite
boarder         = 5             # Rand um das Etikett herum
inner_boarder   = 3             # innerer Abstand zwischen verschiedenen Elementen
no_label_height = 17            # Anzahl der Label in der Höhe
no_label_width  = 4             # Anzahl der Label in der Breite
logo_height     = 0             # Höhe des Quellengalerie-Logos
logo_width      = 0             # Höhe des Quellengalerie-Logos
header_text_heigth = 5

label_heigth    = (page_height - 2*page_top_boarder) / no_label_height
label_width     = page_width / no_label_width        

header_text_width  = label_width - 2*boarder - inner_boarder - logo_width        

h = 0
w = 0

for v in range(first_label,last_label+1) :

    # Bestimme linke obere Ecke der ersten Kopf-Überschrift (Name)
    h_left_upper = page_top_offset + page_top_boarder + (h * label_heigth) + boarder
    w_left_upper = page_left_offset +                   (w * label_width)  + boarder
                        
    # setze Position auf die linke obere Ecke und erzeuge Namens-zeile
    pdf.set_xy(w_left_upper, h_left_upper)
    pdf.set_font('arial', '', 10)
    pdf.multi_cell(header_text_width, header_text_heigth, "SID:" + str(v).zfill(6), 0, 0, '')  

    if w + 1 >= no_label_width :
        w = 0
        if h + 1 >= no_label_height :
            h = 0
            pdf.add_page("p")
        else :
            h = h + 1
    else :        
            w = w + 1

file_name_out = target_dir + datetime.now().strftime("%Y-%m-%d_%H:%M:%S") + "_labels_" + str(first_label).zfill(6) + "_" + str(last_label).zfill(6) + ".pdf"                
pdf.output(file_name_out, '')
print("... file created ", file_name_out)