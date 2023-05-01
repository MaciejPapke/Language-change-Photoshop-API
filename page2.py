import win32com.client
import os
from photoshop import Session

psApp = win32com.client.Dispatch('Photoshop.Application')
mainPsd = psApp.Open(r"D:\Files\Professional\CV-Portfolio-Cover-letters\CV2023p2.psd")
doc = psApp.Application.ActiveDocument


#Go through all layers and groups - make NOK invisible, ENG visible
for i in doc.Layers:
    if i.name.startswith("NOK"):
        i.visible = False
    if i.name.startswith("ENG"):
        i.visible = True
#LayerSets are groups
for i in doc.LayerSets:
    if i.name.startswith("NOK"):
        i.visible = False
    if i.name.startswith("ENG"):
        i.visible = True
    for x in i.Layers:
        if x.name.startswith("NOK"):
            x.visible = False
        if x.name.startswith("ENG"):
            x.visible = True
    for y in i.LayerSets:
        for z in y.Layers:
            if z.name.startswith("NOK"):
                z.visible = False
            if z.name.startswith("ENG"):
                z.visible = True


#Exporting
pdfFileNOK = "D:\Files\Professional\CV-Portfolio-Cover-letters\CV-versions\Auto-generated\MACIEJ-PAPKE-CV-NORSK-p2.pdf"
pdfFileENG = "D:\Files\Professional\CV-Portfolio-Cover-letters\CV-versions\Auto-generated\MACIEJ-PAPKE-CV-ENG-p2.pdf"

#Save current active document as a PDF file.
with Session() as ps:
    option = ps.PDFSaveOptions(presetfile = "CV")
    pdf = os.path.join(pdfFileENG)
    ps.active_document.saveAs(pdf, option)


#"""
#Go through all layers and groups - make NOK invisible, ENG visible
for i in doc.Layers:
    if i.name.startswith("ENG"):
        i.visible = False
    if i.name.startswith("NOK"):
        i.visible = True
#LayerSets are groups
for i in doc.LayerSets:
    if i.name.startswith("ENG"):
        i.visible = False
    if i.name.startswith("NOK"):
        i.visible = True
    for x in i.Layers:
        if x.name.startswith("ENG"):
            x.visible = False
        if x.name.startswith("NOK"):
            x.visible = True
    for y in i.LayerSets:
        for z in y.Layers:
            if z.name.startswith("ENG"):
                z.visible = False
            if z.name.startswith("NOK"):
                z.visible = True

#"""





#Save current active document as a PDF file.
with Session() as ps:
    option = ps.PDFSaveOptions(presetfile = "CV")
    pdf = os.path.join(pdfFileNOK)
    ps.active_document.saveAs(pdf, option)
