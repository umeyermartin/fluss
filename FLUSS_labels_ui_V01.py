# Erzeuge Paginierlabel

from fpdf import FPDF
from datetime import datetime
import config
import PySimpleGUI as sg

class PDF(FPDF):
    def header(self):
        pass

# -------------------------------- mainline -------------------------------------------------------------------------

config_parms = config.config()
target_dir = config_parms["labels_target_dir"]

select_layout =  []

select_layout.append([sg.Text("Number of first label:",size=(20, None)) , sg.Input(key=f'-FROM-')])
select_layout.append([sg.Text("Number of last label:",size=(20, None)) , sg.Input(key=f'-TO-')])  
select_layout.append([sg.Button('Apply'), sg.Button('Exit')])

window_select = sg.Window('Print Label', select_layout,
        ttk_theme='clam',
        resizable=True, finalize=True)
window_select["-FROM-"].update("")
window_select["-TO-"].update("")
while True :
    event_select, values_select = window_select.read()
    if event_select == sg.WIN_CLOSED:
        window_select.close()
        break
    if event_select == "Exit":
        window_select.close()
        break
    if event_select == "Apply": 
        first_label = int(values_select["-FROM-"])
        last_label = int(values_select["-TO-"])

        pdf = PDF()

        pdf.alias_nb_pages()
        pdf.add_page("p")
        pdf.set_auto_page_break(True, margin = 0.0)
        pdf.set_margins(0,0)

        page_top_offset = 3             # zusätzlicher versatz der Seite horizontal
        page_left_offset = 5            # zusätzlicher versatz am linken Rand
        page_top_boarder= 9             # oberer Rand des Etikettenblattes
        page_left_boarder= 9             # oberer Rand des Etikettenblattes
        page_height     = 295           # Seitenhöhe
        page_width      = 210           # Seitenbreite
        boarder         = 3             # Rand um das Etikett herum
        inner_boarder   = 3             # innerer Abstand zwischen verschiedenen Elementen
        no_label_height = 13            # Anzahl der Label in der Höhe
        no_label_width  = 5             # Anzahl der Label in der Breite
        logo_height     = 0             # Höhe des Quellengalerie-Logos
        logo_width      = 0             # Höhe des Quellengalerie-Logos
        header_text_heigth = 5

        label_heigth    = 1.04 * (page_height - 2*page_top_boarder) / no_label_height
        label_width     = 1.03 * (page_width-2*page_left_boarder) / no_label_width        

        header_text_width  = label_width - 2*boarder - inner_boarder - logo_width        

        h = 0
        w = 0

        for v in range(first_label,last_label+1) :

            # Bestimme linke obere Ecke der ersten Kopf-Überschrift (Name)
            h_left_upper = page_top_offset + page_top_boarder + (h * label_heigth) + boarder
            w_left_upper = page_left_offset + page_left_boarder + (w * label_width)  + boarder
                                
            # setze Position auf die linke obere Ecke und erzeuge Namens-zeile
            pdf.set_xy(w_left_upper, h_left_upper)
            pdf.set_font('arial', '', 10)
            pdf.multi_cell(header_text_width, header_text_heigth, "SID:" + str(v).zfill(6), 0, 0, '')  

            if w + 1 >= no_label_width :
                w = 0
                if h + 1 >= no_label_height :
                    h = 0
                    if v != last_label :
                        pdf.add_page("p")
                else :
                    h = h + 1
            else :        
                    w = w + 1

        file_name_out = target_dir + datetime.now().strftime("%Y-%m-%d_%H:%M:%S") + "_labels_" + str(first_label).zfill(6) + "_" + str(last_label).zfill(6) + ".pdf"                
        pdf.output(file_name_out, '')
        print("... file created ", file_name_out)

        multiline_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))], 
                            [sg.Button('Exit')]]
        window_multiline = sg.Window('Label-Result', multiline_layout,
                ttk_theme='clam',
                resizable=True, finalize=True)

        window_multiline['-MULTILINE KEY-'].print("... file created ", file_name_out)

        window_select.close()

        while True :
            event_multiline, values_multiline = window_multiline.read()
            if event_multiline == sg.WIN_CLOSED or event_multiline == "Exit":
                window_multiline.close()
                break

        break