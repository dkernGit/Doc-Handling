import win32com.client as win32
import os

from tkinter import *
from tkinter import filedialog
from tkinter import ttk


def document_prop():
    filename = os.path.normpath(filedialog.askopenfilename())
    
    word = win32.Dispatch("Word.Application")
    word.Visible = 0
    # file = "C:\\Users\\dkern\\Desktop\\Cylindricity Measurement on ESX.doc"
    doc = word.Documents.Open(filename)

    doc_level.set(doc.CustomDocumentProperties('prqdocsubtype').value)
    doc_number.set(doc.CustomDocumentProperties('prqdocnumber').value)
    doc_title.set(doc.CustomDocumentProperties('prqdoctitle').value)
    proc_area.set(doc.CustomDocumentProperties('prquserfield1').value)
    itar.set(doc.CustomDocumentProperties('prquserfield2').value)
    doc_status.set(doc.CustomDocumentProperties('prqdocdraft').value)
    doc_date.set(doc.CustomDocumentProperties('prqdocdate').value)
    doc_rev.set(doc.CustomDocumentProperties('prqdocissue').value)
    doc_author.set(doc.CustomDocumentProperties('prqdocauthor').value)
    doc_author_name.set(doc.BuiltInDocumentProperties('Author').value)

    file_name.set(filename)

    word.Application.Quit(-1)
    return True

def update_document():
    word = win32.Dispatch("Word.Application")
    word.Visible = 0
    doc = word.Documents.Open(file_name.get())
    try:
        # Set Custom Property Values from Variables
        doc.CustomDocumentProperties('prqdocsubtype').value = doc_level.get()
        doc.CustomDocumentProperties('prqdocnumber').value = doc_number.get()
        doc.CustomDocumentProperties('prqdoctitle').value = doc_title.get()
        doc.CustomDocumentProperties('prquserfield1').value = proc_area.get()
        doc.CustomDocumentProperties('prquserfield2').value = itar.get()
        doc.CustomDocumentProperties('prqdocdraft').value = doc_status.get()
        doc.CustomDocumentProperties('prqdocdate').value = doc_date.get()
        doc.CustomDocumentProperties('prqdocissue').value = doc_rev.get()
        doc.CustomDocumentProperties('prqdocauthor').value = doc_author.get()
        doc.BuiltInDocumentProperties('Author').value = doc_author_name.get()


        file_status.set(doc.CustomDocumentProperties('prqdocnumber').value + " updated")
        # os.getcwd()+'\\'+doc_number.get()+'-'+doc_title.get() + '.docx'
        doc.SaveAs(file_name.get(), FileFormat = 12)
        word.Application.Quit(-1)
        return True
    except:
        file_status.set(file_name.get() + " not updated.")
        word.Application.Quit(-1)
        return False

root = Tk()
root.title("Document Control Property Editor")

mainframe = Frame(root)
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)
mainframe.pack(pady=10, padx=10)

doc_level = StringVar()
doc_number = StringVar()
doc_title = StringVar()
proc_area = StringVar()
itar = StringVar()
doc_status = StringVar()
doc_date = StringVar()
doc_rev = StringVar()
doc_author = StringVar()
doc_author_name = StringVar()
file_name = StringVar()
file_status = StringVar()


itar_info = ['ITAR Controlled Information', '']
document_status = ['DRAFT - NOT FOR OPERATIONAL USE',
                   'PROCESS DOCUMENT - FOR REFERENCE USE ONLY']

# Row 1
ttk.Label(mainframe, text="Adjust Document Properties Below").grid(row=1, column=2, sticky=(E))
get_document = ttk.Button(mainframe, text="Get Document", command=document_prop).grid(row=1, column=3)
# Row 2
ttk.Label(mainframe, text="Document Level: ").grid(row=2, column=2, sticky=(E))
document_level = ttk.Entry(mainframe, text=doc_level, width=40).grid(row=2, column=3)
# Row 3
ttk.Label(mainframe, text="Document Number: ").grid(row=3, column=2, sticky=(E))
document_number = ttk.Entry(mainframe, text=doc_number, width=40).grid(row=3, column=3)
# Row 4
ttk.Label(mainframe, text="Document Title: ").grid(row=4, column=2, sticky=(E))
document_title = ttk.Entry(mainframe, text=doc_title, width=40).grid(row=4, column=3)
# Row 5
ttk.Label(mainframe, text="Process Area: ").grid(row=5, column=2, sticky=(E))
process_area = ttk.Entry(mainframe, text=proc_area, width=40).grid(row=5, column=3)
# Row 6
ttk.Label(mainframe, text="ITAR Controlled?: ").grid(row=6, column=2, sticky=(E))
itar_info = ttk.Entry(mainframe, text=itar, width=40).grid(row=6, column=3)
# Row 7 ttk.OptionMenu(mainframe, week_var, *choices)
ttk.Label(mainframe, text="Document Status: ").grid(row=7, column=2, sticky=(E))
document_status = ttk.OptionMenu(mainframe, doc_status, *document_status).grid(row=7, column=3)
# Row 8
ttk.Label(mainframe, text="Document Date: ").grid(row=8, column=2, sticky=(E))
document_date = ttk.Entry(mainframe, text=doc_date, width=40).grid(row=8, column=3)
# Row 9
ttk.Label(mainframe, text="Document Rev: ").grid(row=9, column=2, sticky=(E))
document_rev = ttk.Entry(mainframe, text=doc_rev, width=40).grid(row=9, column=3)
# Row 10
ttk.Label(mainframe, text="Document Owner (Position): ").grid(row=10, column=2, sticky=(E))
document_author = ttk.Entry(mainframe, text=doc_author, width=40).grid(row=10, column=3)
# Row 11
ttk.Label(mainframe, text="Document Owner Email: ").grid(row=11, column=2, sticky=(E))
document_author_name = ttk.Entry(mainframe, text=doc_author_name, width=40).grid(row=11, column=3)
# Row 12
set_document = ttk.Button(mainframe, text="Update Document", command=update_document).grid(row=12, column=3)
ttk.Label(mainframe, textvariable=file_status, background="yellow").grid(row=12, column=2)
root.mainloop()
