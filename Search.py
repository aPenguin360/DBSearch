import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog

## Initialize Global Variables
global widthEnt
global lengthEnt
global thickEnt
global boreEnt
global keyEnt
global cbEnt
global shapeEnt
global tolEnt
global repVar
widthEnt = 0
lengthEnt = 0
thickEnt = 0
boreEnt = 100
keyEnt = 100
cbEnt = 1000
shapeEnt = 100
tolEnt = 0
repVar = 0
matchRes = ''

## Initialize GUI
root = tk.Tk()
root.title('Import')
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'gray')
canvas1.pack()

## Define Spreadsheet Import Method
def getFile():
    global ef

    filepath = tk.filedialog.askopenfilename()
    ef = pd.read_excel(filepath)
    print("Reading From %s", filepath)
    root.destroy()

## Create Import GUI Button
browseButton_Excel = tk.Button(text = 'Import Excel File', command = getFile, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)
root.mainloop()

## -----Initialize Spreadsheet Columns by Parameter-----
partNumber = pd.DataFrame(ef, columns = ['Part #'], dtype = object)
plateSize = pd.DataFrame(ef, columns = ['Plate Size'], dtype = float)
xLength = pd.DataFrame(ef, columns = ['Width'], dtype = float)
yLength = pd.DataFrame(ef, columns = ['Length'], dtype = float)
zLength = pd.DataFrame(ef, columns = ['Thickness'], dtype = float)
bores = pd.DataFrame(ef, columns = ['Bores'], dtype = float)
keys = pd.DataFrame(ef, columns = ['Keyed'], dtype = float)
cbAng = pd.DataFrame(ef, columns = ['Counterbore Angle'], dtype = float)
bShape = pd.DataFrame(ef, columns = ['Trap/Para/Rect'], dtype = float)


## Initialize Secoond GUI Screen
root = tk.Tk()
root.title('Sorting Parameters')

## Program Repeat Loop
while repVar != 1:
    ## Reinitialize Entered Parameters
    widthEnt = 0
    lengthEnt = 0
    thickEnt = 0
    boreEnt = 100
    keyEnt = 100
    cbEnt = 1000
    shapeEnt = 100
    tolEnt = 0
    repVar = 0

    var = tk.IntVar()
    var2 = tk.IntVar()
    var3 = tk.StringVar()

    var.set(100)
    var2.set(100)

    ## -----Define Methods to Repeat or End the Program-----
    def repButton():
        global widthEnt
        global lengthEnt
        global thickEnt
        global boreEnt
        global keyEnt
        global cbEnt
        global shapeEnt
        global tolEnt
        global repVar

        repVar = 0
        widthEnt = 0
        lengthEnt = 0
        thickEnt = 0
        boreEnt = 100
        keyEnt = 100
        cbEnt = 1000
        shapeEnt = 100
        tolEnt = 0
        repVar = 0

        matchLab.delete(0, tk.END)

        startSearch()

    def closeButton():
        global repVar
        repVar = 1
        sys.exit()

    ## -----Define Methods to Recieve Input Parameters-----
    def getWidth():
        global widthEnt
        if widthEntry.get() == '':
            widthEnt = 0
        else:
            widthEnt = float(widthEntry.get())
        print('Width: ', widthEnt, 'in.')

    def getLength():
        global lengthEnt
        if lengthEntry.get() == '':
            lengthEnt = 0
        else:
            lengthEnt = float(lengthEntry.get())
        print('Length: ', lengthEnt, 'in.')

    def getThick():
        global thickEnt
        if thickEntry.get() == '':
            thickEnt = 0
        else:
            thickEnt = float(thickEntry.get())
        print('Thickness: ', thickEnt, 'in.')

    def getBores():
        global boreEnt
        if boreEntry.get() == '':
            boreEnt = 100
        else:
            boreEnt = float(boreEntry.get())
        print('Number of Bores: ', boreEnt)

    def getTol():
        global tolEnt
        tolEnt = float(tolSpin.get()) / 100
        print('Parameter Tolerence: ', tolEnt * 100, '%')

    def getKey():
        global keyEnt
        keyEnt = var.get()
        if(keyEnt):
            print('Keyed')

        else:
            print('Not Keyed')

    def getShape():
        global shapeEnt
        shapeEnt = var2.get()
        if(shapeEnt == 1):
            print('Shape: Trapezoidal')

        elif(shapeEnt == 2):
            print('Shape: Para')

        elif(shapeEnt == 3):
            print('Shape: Rectangular')

        else:
            shapeEnt = 100

    def getCB():
        global cbEnt
        if cbEntry.get() == '':
            cbEnt = 1000
        else:
            cbEnt = float(cbEntry.get())
        print('Counterbore Angle: ', cbEnt, 'Degrees')

    ## -----Define Search Method-----
    def startSearch():
        ## Initialize Lists
        global outMatch
        global matchRes

        widthMatch = [1] * (len(xLength) + 1)
        lengthMatch = [1] * (len(xLength) + 1)
        thickMatch = [1] * (len(xLength) + 1)
        boreMatch = [1] * (len(xLength) + 1)
        cbMatch = [1] * (len(xLength) + 1)
        keyMatch = [1] * (len(xLength) + 1)
        shapeMatch = [1] * (len(xLength) + 1)
        outMatch = [0] * (len(xLength) + 1)
        matchRes = ''

        ## Recheck Input Parameters
        getWidth()
        getLength()
        getThick()
        getBores()
        getTol()
        getKey()
        getShape()
        getCB()

        var.set(100)
        var2.set(100)

        ## -------Change Start Button to Reset-------
        execButton.grid_remove()
        reptButton = tk.Button(root, text = 'Run Again', command = repButton, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
        reptButton.grid(row = 3, column = 1, pady = 2, padx = 2)

        ## -------No Parameters Entered Case-----
        if ((widthEnt == 0) and (lengthEnt == 0) and (thickEnt == 0) and (boreEnt == 100) and (cbEnt == 1000) and (keyEnt == 100) and (shapeEnt == 100)):
            matchRes = 'No Input Parameters Set'
            var3.set(matchRes)

        ## -------Parameters Entered Case-----
        else:
            ## -------Input Parameter Matching with Imported Data-------
            if widthEnt != 0:
                for i in range(len(xLength)):
                    if ((widthEnt <= (xLength.at[i, 'Width'] + tolEnt * xLength.at[i, 'Width'])) and (widthEnt >= (xLength.at[i, 'Width'] - tolEnt * xLength.at[i, 'Width']))):
                        widthMatch[i] = xLength.at[i, 'Width']
                    else:
                        widthMatch[i] = 0

            if lengthEnt != 0:
                for i in range(len(yLength)):
                    if ((lengthEnt <= (yLength.at[i, 'Length'] + tolEnt * yLength.at[i, 'Length'])) and (lengthEnt >= (yLength.at[i, 'Length'] - tolEnt * yLength.at[i, 'Length']))):
                        lengthMatch[i] = yLength.at[i, 'Length']
                    else:
                        lengthMatch[i] = 0

            if thickEnt != 0:
                for i in range(len(zLength)):
                    if ((thickEnt <= (zLength.at[i, 'Thickness'] + tolEnt * zLength.at[i, 'Thickness'])) and (thickEnt >= (zLength.at[i, 'Thickness'] - tolEnt * zLength.at[i, 'Thickness']))):
                        thickMatch[i] = zLength.at[i, 'Thickness']
                    else:
                        thickMatch[i] = 0

            if boreEnt != 100:
                for i in range(len(bores)):
                    if ((boreEnt <= (bores.at[i, 'Bores'] + tolEnt * bores.at[i, 'Bores'])) and (boreEnt >= (bores.at[i, 'Bores'] - tolEnt * bores.at[i, 'Bores']))):
                        boreMatch[i] = bores.at[i, 'Bores']
                    else:
                        boreMatch[i] = 0

            if cbEnt != 1000:
                for i in range(len(cbAng)):
                    if ((cbEnt <= (cbAng.at[i, 'Counterbore Angle'] + tolEnt * cbAng.at[i, 'Counterbore Angle'])) and (cbEnt >= (cbAng.at[i, 'Counterbore Angle'] - tolEnt * cbAng.at[i, 'Counterbore Angle']))):
                        cbMatch[i] = 1
                    else:
                        cbMatch[i] = 0

            if keyEnt != 100:
                for i in range(len(keys)):
                    if (keyEnt == keys.at[i, 'Keyed'] ):
                        keyMatch[i] = 1
                    else:
                        keyMatch[i] = 0

            if shapeEnt != 100:
                for i in range(len(bShape)):
                    if (shapeEnt == bShape.at[i, 'Trap/Para/Rect']):
                        shapeMatch[i] = 1
                    else:
                        shapeMatch[i] = 0

            ## -------Compare Match Arrays-------
            j = 0
            for i in range(len(xLength)):
                if (widthMatch[i] and lengthMatch[i] and thickMatch[i] and boreMatch[i] and keyMatch[i] and cbMatch[i] and shapeMatch[i]):
                    outMatch[j] = partNumber.at[i, 'Part #']
                    j += 1

            ## -------Initialize and De-Zero Output DataFrame-------
            matchRes = pd.DataFrame({'Matches' : outMatch})
            indexZ = matchRes[matchRes['Matches'] == 0].index
            matchRes.drop(indexZ, inplace = True)

            ## -------Add DataFrame Elements to Listbox-------
            for i in range(len(matchRes)):
                matchLab.insert(tk.END, str(matchRes.at[i, 'Matches']))

    ## -------GUI Handling-------
    widthTitle = tk.LabelFrame(root, text = 'Enter Width:')
    widthTitle.grid(row = 0, column = 0, pady = 2, padx = 2)
    widthEntry = tk.Entry(widthTitle, text = 'Width', bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'))
    widthEntry.grid(row = 0, column = 0, pady = 2, padx = 2)
    widthButton = tk.Button(widthTitle, text = 'Enter', command = getWidth, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    widthButton.grid(row = 0, column = 1, pady = 2, padx = 2)

    lengthTitle = tk.LabelFrame(root, text = 'Enter Length:')
    lengthTitle.grid(row = 1, column = 0, pady = 2, padx = 2)
    lengthEntry = tk.Entry(lengthTitle, text = 'Length', bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'))
    lengthEntry.grid(row = 0, column = 0, pady = 2, padx = 2)
    lengthButton = tk.Button(lengthTitle, text = 'Enter', command = getLength, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    lengthButton.grid(row = 0, column = 1, pady = 2, padx = 2)

    thickTitle = tk.LabelFrame(root, text = 'Enter Height:')
    thickTitle.grid(row = 2, column = 0, pady = 2, padx = 2)
    thickEntry = tk.Entry(thickTitle, text = 'Thickness', bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'))
    thickEntry.grid(row = 0, column = 0, pady = 2, padx = 2)
    thickButton = tk.Button(thickTitle, text = 'Enter', command = getThick, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    thickButton.grid(row = 0, column = 1, pady = 2, padx = 2)


    boreTitle = tk.LabelFrame(root, text = 'Enter Number of Holes:')
    boreTitle.grid(row = 0, column = 1, pady = 2, padx = 2)
    boreEntry = tk.Entry(boreTitle, text = 'Bores', bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'))
    boreEntry.grid(row = 0, column = 0, pady = 2, padx = 2)
    boreButton = tk.Button(boreTitle, text = 'Enter', command = getBores, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    boreButton.grid(row = 0, column = 1, pady = 2, padx = 2)

    cbTitle = tk.LabelFrame(root, text = 'Enter Counterbore Angle:')
    cbTitle.grid(row = 1, column = 1, pady = 2, padx = 2)
    cbEntry = tk.Entry(cbTitle, text = 'Counterbore Angle', bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'))
    cbEntry.grid(row = 0, column = 0, pady = 2, padx = 2)
    cbButton = tk.Button(cbTitle, text = 'Enter', command = getCB, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    cbButton.grid(row = 0, column = 1, pady = 2, padx = 2)

    tolTitle = tk.LabelFrame(root, text = 'Parameter Tolerance:')
    tolTitle.grid(row = 2, column = 1, pady = 2, padx = 2)
    tolSpin = tk.Spinbox(tolTitle, from_ = 0, to = 20,wrap = True, bg ='gold', fg ='black', activebackground = 'black', font = ('helvetica', 12, 'bold'))
    tolSpin.grid(row = 0, column = 0, pady = 2, padx = 2)
    tolButton = tk.Button(tolTitle, text = 'Enter', command = getTol, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    tolButton.grid(row = 0, column = 1, pady = 2, padx = 2)


    keyTitle = tk.LabelFrame(root, text = 'Keys:')
    keyTitle.grid(row = 0, column = 2, pady = 2, padx = 2)
    keyRad1 = tk.Radiobutton(keyTitle, text = 'Keyed Holes', command = getKey, variable = var, value = 1, bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'), indicatoron = 0, activebackground = 'black')
    keyRad1.grid(row = 0, column = 0, pady = 2, padx = 2)
    keyRad2 = tk.Radiobutton(keyTitle, text = 'No Keyed Holes', command = getKey, variable = var, value = 0, bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'), indicatoron = 0, activebackground = 'black')
    keyRad2.grid(row = 1, column = 0, pady = 2, padx = 2)

    shapeTitle = tk.LabelFrame(root, text = 'Blade Shape:')
    shapeTitle.grid(row = 1, column = 2, pady = 2, padx = 2)
    shapeRad1 = tk.Radiobutton(shapeTitle, text = 'Trapezoid', command = getShape, variable = var2, value = 1, bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'), indicatoron = 0, activebackground = 'black')
    shapeRad1.grid(row = 0, column = 0, pady = 2, padx = 2)
    shapeRad2 = tk.Radiobutton(shapeTitle, text = 'Para', command = getShape, variable = var2, value = 2, bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'), indicatoron = 0, activebackground = 'black')
    shapeRad2.grid(row = 1, column = 0, pady = 2, padx = 2)
    shapeRad3 = tk.Radiobutton(shapeTitle, text = 'Rectangular', command = getShape, variable = var2, value = 3, bg ='gold', fg ='black', font = ('helvetica', 12, 'bold'), indicatoron = 0, activebackground = 'black')
    shapeRad3.grid(row = 2, column = 0, pady = 2, padx = 2)


    matchTitle = tk.LabelFrame(root, text = 'Matching Part Numbers:')
    matchTitle.grid(row = 0, column = 3, rowspan = 3, columnspan = 2, padx = 5, pady = 5)
    matchLab = tk.Listbox(matchTitle, fg ='black', font = ('helvetica', 12, 'bold'))
    matchLab.grid(row = 0, column = 0, padx = 2, pady = 2)
    scrollbar = tk.Scrollbar(matchTitle, orient = 'vertical', command = matchLab.yview)
    scrollbar.grid(row = 0, column = 1)
    matchLab.config(yscrollcommand = scrollbar.set)

    execButton = tk.Button(root, text = 'Enter', command = startSearch, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    execButton.grid(row = 3, column = 1, pady = 2, padx = 2)
    endButton1 = tk.Button(root, text = 'Close Program', command = closeButton, bg = 'gold', fg = 'black', font = ('helvetica', 12, 'bold'), activebackground = 'black')
    endButton1.grid(row = 3, column = 2, pady = 2, padx = 2)

    root.mainloop()
