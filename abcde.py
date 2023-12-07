import sys,os
from docxtpl import DocxTemplate
import PySimpleGUI as sg
from pathlib import Path
import datetime

os.chdir(sys.path[0])
doc = DocxTemplate("C:\\Users\\USER\\Desktop\\USG.docx")


context =[
    [sg.Text("Liver"),sg.Input(key="LIVER",do_not_clear = False)],
    [sg.Text("Common Bile Duct"),sg.Input(key="CBD",do_not_clear = False)],
    [sg.Text("Pancreas"),sg.Input(key="PAN",do_not_clear = False)],
    [sg.Text("Spleen"),sg.Input(key="SPL",do_not_clear = False)],
    [sg.Text("prostate"),sg.Input(key="PRT",do_not_clear = False)],
    [sg.Text("Urinary Bladder"),sg.Input(key="UB",do_not_clear = False)],
    [sg.Text("Ureter"),sg.Input(key="UT",do_not_clear = False)],
    [sg.Button("UPLOAD REPORT"),sg.Exit()],
    
]
window = sg.Window("USG REPORT",context,element_justification="Right")
while True:
    event,values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "UPLOAD REPORT":
       
       doc.render(values)   
       document_path = (Path(__file__).parent/"C:\\Users\\USER\\Desktop\\USG5.docx") 
       doc.save(document_path)
       sg.popup("File saved",f"File has been saved here:{document_path}") 
           
      
     

window.close()