
"""
Excel files should be in the filename pattern as follows
For
    Windows UCMDB = windows_ucmdb
    Windows SCCM = windows_sccm
    Windows Antivirus = windows_antivirus
    Unix Antivirus = unix_antivirus
    Unix UCMDB = unix_ucmdb
    Unix Satellite = unix_satellite
    Unix Puppet = unix_puppet

!!! Do not forget to add desktop central for non domain windows machines. 
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
    if clickedOS.get() == "Windows":
        if clickedControl.get() == "Inventory":
            # windows inventory controls will be done in here
            print("Windows inventory control will be processed.")
            # files array should be converted from directory to filename
            i = 0
            # Take only the filename not the entire path.
            for file in filenames:
                # TO DO: change this according to operating system; in windows "\" in OSX "/"
                splittedFileNames = filenames[i].split("/")
                print(splittedFileNames[(len(splittedFileNames)-1)])
                filenames[i] = splittedFileNames[len(splittedFileNames)-1]
                i = i+1
                splittedFileNames.clear()
            print("Filenames: ") # Debug
            print(filenames) # Debug
            # check if all required files are uploaded
            if "windows_ucmdb.xlsx" in filenames and "windows_sccm.xlsx" in filenames and "windows_antivirus.xlsx" in filenames:
                print("Files are in the list... Processing comparision operation...")

                # TO DO: Do comparision operations in here


            else:
                print()
                print(filenames)
                print("Files are missing in the list, for windows inventory control put at least windows_ucmdb, windows_sccm, windows_antivirus lists")


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
clickedControl.set("Inventory")

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
