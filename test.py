# import re
# import subprocess
# import PySimpleGUI as sg

# text_all = " jg jh gjg jg jg jh   SID:004711 ilkh kh k hkj hk hkj hkj hkj hk h"
# #re.match(r'\d', '3')
# z = re.findall('SID:(\d\d\d\d\d\d)', text_all)
# print(z[0])
# # -------------------
# import time
  





# invoice_layout = [[sg.Multiline(key='-MULTILINE KEY-',  disabled=True, size=(70,40))],

#      [sg.Combo(['New York','Chicago','Washington', 'Colorado','Ohio','San Jose','Fresno','San Fransisco'],default_value='Utah',key='-combo-')]]

# window_invoice = sg.Window('Rechnung', invoice_layout,
#         ttk_theme='clam',
#         resizable=True, finalize=True)
# for i in range(5) :
#     command = "ls"
#     sg.popup_notify("cammand executing ..." + command + str(i))
#     window_invoice['-MULTILINE KEY-'].print("---",command,"---",i)
#     proc = subprocess.Popen(command, stdout=subprocess.PIPE)
#     tmp = proc.stdout.read()
#     #print(tmp.decode())
#     window_invoice['-MULTILINE KEY-'].print(tmp.decode())
#     time.sleep(3)

# while True :
#     event_invoice, values_invoice = window_invoice.read()
#     if event_invoice == sg.WIN_CLOSED:
#         window_invoice.close()
#         break
#     if event_invoice == "Exit":
#         window_invoice.close()
#         break

# import subprocess
# import sys
# import PySimpleGUI as sg

# def main():
#     layout = [  [sg.Text('Enter a command to execute (e.g. dir or ls)')],
#             [sg.Input(key='_IN_')],             # input field where you'll type command
#             [sg.Output(size=(60,15))],          # an output area where all print output will go
#             [sg.Button('Run'), sg.Button('Exit')] ]     # a couple of buttons

#     window = sg.Window('Realtime Shell Command Output', layout)
#     while True:             # Event Loop
#         event, values = window.Read()
#         if event in (None, 'Exit'):         # checks if user wants to 
#             exit
#             break

#         if event == 'Run':                  # the two lines of code needed to get button and run command
#             runCommand(cmd=values['_IN_'], window=window)

#     window.Close()

# # This function does the actual "running" of the command.  Also watches for any output. If found output is printed
# def runCommand(cmd, timeout=None, window=None):
#     p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
#     output = ''
#     for line in p.stdout:
#         line = line.decode(errors='replace' if (sys.version_info) < (3, 5) else 'backslashreplace').rstrip()
#         output += line
#         print(line)
#         window.Refresh() if window else None        # yes, a 1-line if, so shoot me
#     retval = p.wait(timeout)
#     return (retval, output)                         # also return the output just for fun

# if __name__ == '__main__':
#     main()

#!/usr/bin/env python
import PySimpleGUI as sg

"""
    Simple test harness to demonstate how to use the CalendarButton and the get date popup
"""
# sg.theme('Dark Red')
layout = [[sg.Text('Date Chooser Test Harness', key='-TXT-')],
      [sg.Input(key='-IN-', size=(20,1)), sg.CalendarButton('Cal US No Buttons Location (0,0)', close_when_date_chosen=True,  target='-IN-', location=(0,0), no_titlebar=False, )],
      [sg.Input(key='-IN3-', size=(20,1)), sg.CalendarButton('Cal Monday', title='Pick a date any date', no_titlebar=True, close_when_date_chosen=False,  target='-IN3-', begin_at_sunday_plus=1, month_names=('студзень', 'люты', 'сакавік', 'красавік', 'май', 'чэрвень', 'ліпень', 'жнівень', 'верасень', 'кастрычнік', 'лістапад', 'снежань'), day_abbreviations=('Дш', 'Шш', 'Шр', 'Бш', 'Жм', 'Иш', 'Жш'))],
      [sg.Input(key='-IN2-', size=(20,1)), sg.CalendarButton('Cal German Feb 2020',  target='-IN2-',  default_date_m_d_y=(2,None,2020), locale='de_DE', begin_at_sunday_plus=1 )],
      [sg.Input(key='-IN4-', size=(20,1)), sg.CalendarButton('Cal Format %m-%d Jan 2020',  target='-IN4-', format='%m-%d', default_date_m_d_y=(1,None,2020), )],
      [sg.Button('Read'), sg.Button('Date Popup'), sg.Exit()]]

window = sg.Window('window', layout)

while True:
    event, values = window.read()
    print(event, values)
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    elif event == 'Date Popup':
        sg.popup('You chose:', sg.popup_get_date())
window.close()