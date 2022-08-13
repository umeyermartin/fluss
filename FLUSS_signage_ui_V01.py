# FLUSS_signage_UI_V01
import io
import os
import pdf2image
import PySimpleGUI as sg
from PIL import Image
import datetime
import config

file_types = [("JPEG (*.jpg)", "*.jpg"),
              ("All files (*.*)", "*.*")]

#DECLARE CONSTANTS
DPI = 200
OUTPUT_FOLDER = None
FIRST_PAGE = None
LAST_PAGE = None
FORMAT = 'jpg'
THREAD_COUNT = 1
USERPWD = None
USE_CROPBOX = False
STRICT = False

config_parms = config.config()

def main():

    while True :

        # PDF-Datei von Nextcloud holen und PDF in Bilder zerlegen
        os.system("sudo cp " + config_parms["signage_source_file"] + " " + config_parms["signage_temp_file"])
        stat_old = os.stat(config_parms["signage_source_file"])

        pil_images = pdf2image.convert_from_path(config_parms["signage_temp_file"], dpi=DPI, output_folder=OUTPUT_FOLDER, first_page=FIRST_PAGE, last_page=LAST_PAGE, fmt=FORMAT, thread_count=THREAD_COUNT, userpw=USERPWD, use_cropbox=USE_CROPBOX, strict=STRICT)
        dt_last = datetime.datetime.now()

        image_no = 0
        pil_l = len(pil_images) 

        pil_images[image_no].thumbnail(config_parms["signage_thumbnail_size"])
        bio = io.BytesIO()
        pil_images[image_no].save(bio, format="PNG")

        text_signage_left = "FLUSS_signage_V01" + 150 * " "
        text_signage_right = len(text_signage_left) * " "

        layout_column1 = [
            [sg.Image(data=bio.getvalue(),key="-IMAGE-")],
            [sg.Text(text_signage_left), sg.Text(datetime.datetime.now().strftime("%H:%M"),key = '-datetime-'), sg.Button("Exit"), sg.Text(str(image_no+1) + "/" + str(len(pil_images)),key = '-page-'),sg.Text(text_signage_right)],
        ]

        layout = [[sg.Column(layout_column1,element_justification="center")]]
        window = sg.Window("Image Viewer", layout, no_titlebar=True, location=(0,0), size=config_parms["signage_screen_size"], keep_on_top=True, finalize = False)

        while True:
            event, values = window.read(timeout=config_parms["signage_time_slide"])
            if event == "Exit" or event == sg.WIN_CLOSED:
                exit()
            dt_now = datetime.datetime.now()
            if dt_now - dt_last > datetime.timedelta(seconds=config_parms["signage_time_slide"]) :
                dt_last = dt_now
                image_no += 1
                if image_no == pil_l :
                    if stat_old != os.stat(config_parms["signage_source_file"]) :
                        print("*** file changed ***")
                        break
                    else :
                        image_no = 0
                pil_images[image_no].thumbnail(config_parms["signage_thumbnail_size"])
                bio = io.BytesIO()
                pil_images[image_no].save(bio, format="PNG")
                window["-IMAGE-"].update(data=bio.getvalue())
                window["-datetime-"].update(datetime.datetime.now().strftime("%H:%M"))
                window["-page-"].update(str(image_no+1) + "/" + str(len(pil_images)))
        window.close()

if __name__ == "__main__":
    main()
