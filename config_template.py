# -*- coding: utf-8 -*-
# set configuration parameters 

from collections import namedtuple

def config():

    config_parms = {}

    # --- Metadaten ---
    config_parms["metadata_parameter"]              = dict(m="Mandant", b="Beschreibung", a="Aufgabe")
    config_parms["metadata_value_m"]                = dict(g="gerda", m="m-m-cs", p="privat")
    
    # --- Scan-Konfiguration ---
    config_parms["scan_tesseract_cmd"]              = '/usr/bin/tesseract'
    config_parms["scan_tesseract_prefix"]           = '/usr/share/tesseract-ocr/4.00/tessdata/'
    config_parms["scan_tesseract_config"]           = '--tessdata-dir /usr/share/tesseract-ocr/4.00/tessdata/'
    config_parms["scan_src_dir"]                    = "/home/pi/scan_server/temp/"
    config_parms["scan_out_dir"]                    = "/home/pi/scan_server/temp_out/"
    config_parms["scan_nextcloud_mount"]            = "/home/pi/nextcloud/"
    config_parms["scan_nextcloud_dir"]              = "/home/ulrich/Nextcloud/scan/"
    
    # --- Scan-Tasks-Konfiguration ---
    config_parms["tasks_tesseract_cmd"]             =  '/usr/bin/tesseract'
    config_parms["tasks_tesseract_prefix"]          = '/usr/share/tesseract-ocr/4.00/tessdata/'
    config_parms["tasks_tesseract_config"]          = '--tessdata-dir /usr/share/tesseract-ocr/4.00/tessdata/'
    config_parms["tasks_source_dir"]                = "/home/ulrich/Nextcloud/scan/"
    config_parms["tasks_temp_dir"]                  = "/home/ulrich/Nextcloud/scan/temp/"
    
    # --- eMail-Tasks-Konfiguration ---
    config_parms["emails_server"]                   = "exchange.ionos.eu"
    config_parms["emails_user"]                     = "yyy@xxx.com" 
    config_parms["emails_password"]                 = "ppp"
    config_parms["emails_target_dir"]               = "/home/ulrich/Nextcloud/UFD/m-m-cs/09-Abrechnung/Monatsabschluss/2022.XX/"
    config_parms["emails_parameter"]                = dict(m="mail sender",s="subject", d="date_from", e="date_to", l="limit", p="pdf-attachment")
    
    # --- File-Tasks-Konfiguration ---
    #config_parms["files_source_dir"]                = "/home/ulrich/Nextcloud/scan/"
    config_parms["files_source_dir"]                = "/home/ulrich/Nextcloud/UFD/m-m-cs/09-Abrechnung/Monatsabschluss/2022.XX/"
    config_parms["files_target_dir"]                = "/home/ulrich/Nextcloud/UFD/m-m-cs/09-Abrechnung/Monatsabschluss/2022.XX/"
    config_parms["files_parameter"]                 = dict(f="file name")
    
    # --- Labels-Konfiguration ---
    config_parms["labels_target_dir"]                = "/home/ulrich/Nextcloud/UFD/m-m-cs/09-Abrechnung/Monatsabschluss/2022.XX/"

    return(config_parms)

# -------  mainline ----------------------------------------------------------
if __name__ == "__main__":

    print(config())
