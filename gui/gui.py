import os
import re
import PySimpleGUI as sg
from pptx import Presentation

# pyinstaller -F --noconsole -n PrivatePowerPointNotes gui.py  

RED = "#f02828"
ORANGE = "#e0741f"
GREEN = "#08c65b"


layout = [
    [sg.Text('Select a .pptx file:')],

    # Select file and start
    [sg.Input(key='pptx_input'), sg.FileBrowse(key="pptx_input_browse")],
    [sg.Button('Hide', key='hide')],

    # Info for user
    [sg.Text('_' * 50)],
    [sg.Text('Waiting...', key='process')],
]
window = sg.Window('App').Layout(layout)


# Functions
def hide(PATH):
    prs = Presentation(PATH)
    for slide in prs.slides:
        if slide.has_notes_slide:
            slide.notes_slide.notes_placeholder.text = re.sub(r"\n?<hide>[\s\S]*?</hide>", "", slide.notes_slide.notes_placeholder.text)
    
    outputname = os.path.split(PATH)[0] + "/" + os.path.split(PATH)[1].split(".")[0] + "_OUTPUT.pptx"
    print(outputname)
    prs.save(outputname)

# Actions during running
while True:

    event, values = window.Read()

    # hide event started
    if event == "hide":
        PATH = values["pptx_input_browse"]

        if os.path.exists(PATH):
            # if not correct file format
            if PATH.split(".")[-1] != "pptx":
                print(window.FindElement("process").Size)
                window.FindElement("process").Update(value="Not .pptx", text_color=RED)
                window.Refresh()
            else:
                window.FindElement("process").Update(value="Working...", text_color=ORANGE)
                window.Refresh()
                try:
                    hide(PATH)
                    window.FindElement("process").Update(value="Done!", text_color=GREEN)
                    window.Refresh()
                except:
                    window.FindElement("process").Update(value="Error!", text_color=RED)
                    window.Refresh()

    # closing program
    if event is None or event == 'Exit':
        break

window.Close()