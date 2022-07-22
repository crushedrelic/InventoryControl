
"""
Excel files should be in the filename regex as follows
For
    Windows UCMDB = windows_ucmdb
    Windows SCCM = windows_sccm
    Windows Antivirus = windows_antivirus
    Unix Antivirus = unix_antivirus
    Unix UCMDB = unix_ucmdb
    Unix Satellite = unix_satellite
    Unix Puppet = unix_puppet

"""
# Import module
import glob
from tkinter import *
from tkinter import filedialog
import pandas as pd
from array import *

# Create object
root = Tk()
root.title("Inventory and Integrity Control Application")

# Adjust size
root.geometry("800x600")
globalDirectoryName = StringVar


# Browse Files and extract all the excel files.
def browseFiles():
    directoryName = filedialog.askdirectory(initialdir="/", title="Select excel directory")
    # filename = filedialog.askopenfilename(initialdir="/",title="Select excel directory", filetypes=(("Excel files","*.xlsx"),("Excel Files","*.xls")))
    labelFileBrowser.config(text="File opened: " + directoryName)
    global globalDirectoryName
    globalDirectoryName = directoryName


excelArray = []


# Compare files extracted Excel files according to selected control type
def compareFiles():
    print("Compare")
    filenames = glob.glob(globalDirectoryName + "/*.xlsx")
    print(filenames)

    for file in filenames:
        # TODO: Put ui debug in here !1
        print(file)
        excelFile = pd.read_excel(file)
        excelArray.append(excelFile)

    print(excelArray[0])

    # excelFile = pd.read_excel(globalDirectoryName)
    # print(excelFile)
    # print(excelFile[excelFile.columns.values[0]])



# Change the label text


def showSelected():
    label3.config(text="Selected Controls: " + clickedOS.get() + " and " + clickedControl.get() + "Control")


# Dropdown menu options

optionsOS = [

    "Windows",
    "Linux"
]

optionControl = [
    "Inventory",
    "Patch",
    "Antivirus"

]

# datatype of menu text
clickedOS = StringVar()
clickedControl = StringVar()

# initial menu text
clickedOS.set("Windows")
clickedControl.set("Envanter")

# Create Dropdown menu
dropOS = OptionMenu(root, clickedOS, *optionsOS)
dropOS.place(x=150, y=25)
dropType = OptionMenu(root, clickedControl, *optionControl)
dropType.place(x=255, y=25)

# Create button, it will change label text
buttonOS = Button(root, text="Select", command=showSelected).place(x=150, y=50)
buttonFileExp = Button(root, text="BrowseFile", command=browseFiles).place(x=150, y=150)
buttonCompare = Button(root, text="Compare", command=compareFiles).place(x=250, y=300)

# Create Label

label3 = Label(root, text=" ")
label3.place(x=10, y=75)

labelFileBrowser = Label(root, text="File Browser", width=50, height=2)
labelFileBrowser.place(x=10, y=100)
# Execute tkinter
root.mainloop()
