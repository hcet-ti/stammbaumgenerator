import os
import json
import math
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
#import configparser
import numpy as np
import xlsxwriter
import webbrowser


class Stammbaum:
    generations = {}

def CheckFloat(x):
    if x - int(x) == 0:
        return True
    else:
        return False
def get_all_elements_in_list_of_lists(list):
    count = 0
    for element in list:
        count += len(element)
    return count
def checkListForSymbol(list, symbol):
    for i in range(len(list)):
        if symbol in list[i]:
            return i  # list.index(list[i])
def findLastSymbolInMatrix(matrix, symbol, positionOfTarget):
    for i in range(len(matrix)):
        if symbol == matrix[i][positionOfTarget]:
            return i
def replaceEmptyWithSpaceInList(matrix):
    max = 0
    for y in range(len(matrix)):
        for x in range(len(matrix[y])):
            if max < len(matrix[y][x]):
                max = len(matrix[y][x])

    for y in range(len(matrix)):
        for x in range(len(matrix[y])):
            matrix[y][x] = str(matrix[y][x]) + ' ' * (max -len(matrix[y][x]))

    matrix = np.ndarray.tolist(np.asarray(matrix))
    return matrix
def replaceEmpty(list):
    for n in range(len(list)):
        for i in range(len(list[n])):
            if list[n][i] == '':
                list[n][i] = ' '
    return list
def roundDown(float):
    if not CheckFloat(float):
        num = int(float) 
        return num
    else:
        return int(float)
def roundUp(float):
    if not CheckFloat(float):
        num = int(float) + 1
        return num
    else:
        return int(float)
def create_tree(personen):
    #print(personen)
    stammbaum = Stammbaum
    stammbaum.generations = {}
    #print(math.log(int(len(personen.keys()) + 1), 2))
    amountGenerations = int(math.log(int(len(personen.keys()) + 1), 2))  # int(sorted(personen.keys())[-1])
    alteredPpl = ''
    for i in range(len(personen.keys())):
        if CheckFloat(math.log(int(sorted(personen.keys())[i]), 2)):  # if the Person ID is a power of 2
            alteredPpl = alteredPpl + '|'
            if alteredPpl[-1] == '|':
                alteredPpl = alteredPpl + str(sorted(personen.keys())[i])
            else:
                alteredPpl = alteredPpl + ', ' + str(sorted(personen.keys())[i])
        else:
            if alteredPpl[-1] == '|':
                alteredPpl = alteredPpl + str(sorted(personen.keys())[i])
            else:
                alteredPpl = alteredPpl + ', ' + str(sorted(personen.keys())[i])
    alteredPpl = alteredPpl[1:]
    alteredPpl = alteredPpl.split('|')
    #print(alteredPpl)

    for i in range(amountGenerations):
        stammbaum.generations.update({i:alteredPpl[i].replace(' ', '').split(',')})
        #print(stammbaum.generations)
    
    IDgens = stammbaum.generations
    #print('IDGENS:',IDgens)
    generations = len(IDgens.keys())
    IDs = get_all_elements_in_list_of_lists(IDgens.values())
    MARGIN = int(list(os.get_terminal_size())[0] / 2)
    CELLCOLUMNS = 11
    CELLROWS = 3
    HEIGHT = CELLROWS * generations
    WIDTH = CELLCOLUMNS * IDs #+ MARGIN * 2
    grid = []
    for i in range(generations):
        List = []
        for n in range(IDs):
            List.append('')
        grid.append(List)

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # Create the grid with IDs replaced by 'X'
    CELLSYMBOL = 'X'
    LINESYMBOL = '-'
    INTERSECTIONSYMBOL = '+'

    # Create the X's at the bottom
    for x in range(IDs):
        if (x % 2) == 0:
            # x is even
            grid[generations - 1][x] = CELLSYMBOL

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # Create the rest of the X's
    for i in range(generations - 1, 0, -1):
        #print('i:', i)  # debug
        start = -1
        for n in range(IDs):
            #print('n:', n)  # debug
            if grid[i][n] == CELLSYMBOL:
                if start == -1:
                    start = n
                else:
                    middle = int(start + (n - start) / 2)  # get middle of start and n
                    grid[i - 1][middle] = CELLSYMBOL  # create the X in the middle of start and n one generation above the two of them
                    start = -1  # reset start

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # # Create the lines, that connect the 'X's
    # for i in range(1, generations, 1):
    #     start = -1
    #     for n in range(IDs):
    #         if grid[i][n] == CELLSYMBOL:
    #             if start == -1:
    #                 start = n
    #             else:
    #                 line = range(start + 1, n)
    #                 for k in line:
    #                     grid[i][k] = LINESYMBOL
    #                 start = -1  # reset start

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # # Create the intersections
    # for i in range(1, generations, 1):
    #     start = -1
    #     for n in range(IDs):
    #         if grid[i][n] == CELLSYMBOL:
    #             if start == -1:
    #                 start = n
    #             else:
    #                 middle = int(start + (n - start) / 2)  # get middle of start and n
    #                 grid[i][middle] = INTERSECTIONSYMBOL  # create the X in the middle of start and n one generation above the two of them
    #                 start = -1  # reset start

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # Replace the X's with the IDs
    ids = list(range(1, IDs + 1, 1))
    for i in range(generations):
        for n in range(IDs):
            if grid[i][n] == CELLSYMBOL:
                grid[i][n] = str(ids[0])
                ids.pop(0)

    replaceEmptyWithSpaceInList(grid)
    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # If the IDs go past 1 digit, create new characters to fill the space
    highest = []
    for i in range(generations):
        highest.append(max(grid[i], key=len))
    highest = max(highest, key=len)

    grid = grid[::-1]  # this inverts the list
    return grid
def createTree(personen):
    # Create the tree
    grid = create_tree(personen)

    # Create lines between the nums
    line = []
    for i in range(len(grid[0])):
        line.append('')
    n = 0
    while n in range(len(grid)):
        if (n % 2) != 0:
            grid.insert(n, line)
        n += 1

    HORIZONTAL = '─'
    LEFTCORNER = '└'
    RIGHTCORNER = '┘'
    INTERSECTION = '┬'

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

    # Convert the list so multiple list entrys aren't referenced the same...
    grid = np.ndarray.tolist(np.asarray(grid))  # THIS TOOK ONE WHOLE DAY TO FIGURE OUT!!!

    for y in range(len(grid)):
        if (y % 2) != 0:
            LEFT = True
            for x in range(len(grid[0])):
                if grid[y - 1][x].replace(' ', '').isnumeric():
                    if LEFT:
                        grid[y][x] = LEFTCORNER
                        LEFT = False
                    elif not LEFT:
                        grid[y][x] = RIGHTCORNER
                        LEFT = True

    # Create the lines, that connect the 'X's
    for i in range(1, len(grid), 1):
        start = -1
        for n in range(len(grid[0])):
            if grid[i][n].replace(' ', '') in [LEFTCORNER, RIGHTCORNER]:
                if start == -1:
                    start = n
                else:
                    line = range(start + 1, n)
                    for k in line:
                        grid[i][k] = HORIZONTAL
                    start = -1  # reset start

    # Create the intersections
    for i in range(1, len(grid), 1):
        start = -1
        for n in range(len(grid[0])):
            if grid[i][n] in [LEFTCORNER, RIGHTCORNER]:
                if start == -1:
                    start = n
                else:
                    middle = int(start + (n - start) / 2)  # get middle of start and n
                    grid[i][middle] = INTERSECTION  # create the X in the middle of start and n one generation above the two of them
                    start = -1  # reset start

    np.asarray(replaceEmptyWithSpaceInList(grid))
    #print(np.asarray(replaceEmptyWithSpaceInList(grid)))  # debug
    return grid


def def_factory(arg):
    def edit_person():
        def on_closing():
            for i in range(len(textBoxes)): 
                if i == 0:
                    person["Name"] = " ".join([textBoxes[0].get(), textBoxes[1].get()])
                elif i == 1 or i > 10:
                    pass
                else:
                    person[list(person.keys())[i - 1]] = textBoxes[i].get()
            personen.update({id:person})
            window.destroy()

        id = str(arg)
        personName = personen.get(id).get("Name")

        window = Toplevel(root)
        window.title("Edit " + personName)

        frame = ttk.Frame(window, padding="3 3 12 12")
        frame.grid(column=0, row=0, sticky=(N, W, E, S))
        window.columnconfigure(0, weight=1)  # frame should expand to fill any extra space if the window is resized.
        window.rowconfigure(0, weight=1)  # frame should expand to fill any extra space if the window is resized.

        person = personen.get(id)

        textBoxes = {}
        for i in range(len(list(person))):  
            if i == 0:
                textBoxes.update({i:list(person.values())[i].split(" ")[0]})
                textBoxes.update({i + 1:list(person.values())[i].split(" ")[1]})
            else:
                textBoxes.update({i + 1:list(person.values())[i]})  # the +1 is bcs we have split Name into 2 boxes
            
        for i in range(len(list(person))):
            if list(person)[i] == "Name":
                ttk.Label(frame, text=list(person)[i] + ":").grid(column=0, row=i, sticky=E)
            
                textBoxes[i] = StringVar(value=textBoxes[i])
                textBoxes[i + 1] = StringVar(value=textBoxes[i + 1])
                ttk.Entry(frame, width=15, textvariable=textBoxes[i]).grid(column=1, row=i, sticky=(W, E))
                ttk.Entry(frame, width=15, textvariable=textBoxes[i + 1]).grid(column=2, row=i, sticky=(W, E))
            else:
                ttk.Label(frame, text=list(person)[i] + ":").grid(column=0, row=i, sticky=E)
                
                textBoxes[i + 1] = StringVar(value=textBoxes[i + 1])
                ttk.Entry(frame, width=30, textvariable=textBoxes[i + 1]).grid(column=1, row=i, sticky=(W, E), columnspan=2)
        
        ttk.Button(frame, text="Abbrechen", command=window.destroy).grid(column=1, row=len(list(person)), sticky=(N, S))
        ttk.Button(frame, text="Fertig", command=on_closing).grid(column=2, row=len(list(person)), sticky=(N, S))
        
        # Polish
        for child in frame.winfo_children(): 
            child.grid_configure(padx=1, pady=1)
                

        window.protocol("WM_DELETE_WINDOW", on_closing)

    return edit_person

personen = {"1": {"Name": "G\u00e4rtner Gabi", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}, "2": {"Name": "- -", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}, "3": {"Name": "- -", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}, "4": {"Name": "- -", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}, "5": {"Name": "- -", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}, "6": {"Name": "- -", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}, "7": {"Name": "- -", "Heiratsdatum": "", "Heiratsort": "", "weitereHeiraten": "", "Geschwister": "", "Geburtsort": "", "Sterbeort": "", "Geburtsdatum": "", "Sterbedatum": "", "Anmerkung": ""}}
root = Tk()
root.title("Stammbaumgenerator")
root.iconbitmap('./img/icon.ico')
currentfile = ''


def showTutorial():
    webbrowser.open('http://stammbaumgenerator.great-site.net/tutorial.html', new=0)
def showFAQ():
    webbrowser.open('http://stammbaumgenerator.great-site.net/faq.html#', new=0)
def removeGeneration():
    global personen
    last = int((len(personen) + 1) / 2 - 1)
    #print(last)
    a = list(personen.keys())
    b = list(personen.values())
    del a[last:len(personen)]
    del b[last:len(personen)]
    personen = dict(zip(a, b))
    # while len(personen) > last:
    #     personen.pop(list(personen.keys())[-1])
    #print(personen)
    # refresh the tree
    refreshTree()
def addGeneration():
    global personen
    for i in range((len(personen) + 1) * 2 - 1):
        if not str(i+1) in list(personen.keys()):
            personen.update({str(i+1):{"Name": "- -","Heiratsdatum": "","Heiratsort": "","weitereHeiraten": "","Geschwister": "","Geburtsort": "","Sterbeort": "","Geburtsdatum": "","Sterbedatum": "","Anmerkung": ""}})
    # refresh the tree
    refreshTree()
def refreshTree():
    #print(personen)
    for widget in secondframe.winfo_children():
        widget.destroy()
    stammbaum = createTree(personen)

    commands = []
    for i in range(len(stammbaum[0])):
        for y in range(len(stammbaum)):
            if stammbaum[y][i].replace(" ", "").isdecimal():
                id = int(stammbaum[y][i])
                #print('id:',id)
                commands.append(def_factory(id))
    #print(len(commands))

    for y in range(len(stammbaum)):
        for x in range(len(stammbaum[y])):
            value = stammbaum[y][x].replace(" ", "")
            if value.isdecimal():
                ttk.Button(secondframe, text=value, command=commands[x], width=4).grid(column=x, row=y)
            else:
                if value == "└":
                    value = " └─"
                    ttk.Label(secondframe, text=value, font="Ariel 10").grid(column=x, row=y, sticky=E)
                elif value == "┬":
                    value = "─┬─"
                    ttk.Label(secondframe, text=value, font="Ariel 10").grid(column=x, row=y, sticky=(N, E, S, W))
                elif value == "┘":
                    value = "─┘ "
                    ttk.Label(secondframe, text=value, font="Ariel 10").grid(column=x, row=y, sticky=W)
                elif value == "─":
                    value = "───"
                    ttk.Label(secondframe, text=value, font="Ariel 10").grid(column=x, row=y, sticky=(N, E, S, W))
                else:
                    pass
                    #ttk.Label(secondframe, text=value, font="Ariel 10").grid(column=x, row=y, sticky=(N, E, S, W))
def onOpen():
    input = filedialog.askopenfilename(initialdir = "/",title = "Datei öffnen",filetypes = (("JSON Dateien","*.json"),("Alle Dateien","*.*")))
    try: 
        with open(os.path.join(os.path.dirname(__file__), input), 'r') as input_file:
            personen = json.loads(input_file.read())
        global currentfile
        currentfile = input

    except:
        if input:
            messagebox.showerror("Fehler", "Datei konnte nicht geöffnet werden.")
def onSave():
    input = filedialog.asksaveasfilename(initialdir = "/",title = "Datei speichern",filetypes = (("JSON Dateien","*.json"),("Alle Dateien","*.*")))
    if not '.json' in input:
        input = input + '.json'
    try: 
        with open(os.path.join(os.path.dirname(__file__), input), 'w') as output_file:
            output_file.write(json.dumps(personen))
        global currentfile
        currentfile = input

    except:
        if input:
            messagebox.showerror("Fehler", "Datei konnte nicht gespeichert werden.")
def quickSave():
    try: 
        with open(os.path.join(os.path.dirname(__file__), currentfile), 'w') as output_file:
            output_file.write(json.dumps(personen))

    except:
        if input:
            messagebox.showerror("Fehler", "Datei konnte nicht gespeichert werden.")
def bindSave(a):
    try: 
        with open(os.path.join(os.path.dirname(__file__), currentfile), 'w') as output_file:
            output_file.write(json.dumps(personen))

    except:
        if input:
            messagebox.showerror("Fehler", "Datei konnte nicht gespeichert werden.")
def exportToExcel():
    # window = Toplevel(root)
    # window.title("Exporting to excel...")

    # scroll = Scrollbar(window)
    # Tex = Text(window, height=4, width=50)
    # scroll.pack(side=RIGHT, fill=tkinter.Y)
    # Tex.pack(side=LEFT, fill=tkinter.Y)
    # scroll.config(command=Tex.yview)
    # Tex.config(yscrollcommand=scroll.set)
    # quote = """HAMLET: To be, or not to be--that is the question:
    # Whether 'tis nobler in the mind to suffer
    # The slings and arrows of outrageous fortune
    # Or to take arms against a sea of troubles
    # And by opposing end them. To die, to sleep--
    # No more--and by a sleep to say we end
    # The heartache, and the thousand natural shocks
    # That flesh is heir to. 'Tis a consummation
    # Devoutly to be wished."""
    # Tex.insert(END, quote)

    #with console.status("[bold green]Exporting...") as status:
    workbook = xlsxwriter.Workbook('Stammbaum von ' + personen.get('1').get('Name') + '.xlsm')
    workbook.set_vba_name('DieseArbeitsMappe')
    #console.log(f"[green]Created workbook[/green]")
    worksheet = workbook.add_worksheet()
    worksheet.set_vba_name('Tabelle1')
    #console.log(f"[green]Created worksheet[/green]")

    bold = workbook.add_format({'bold':True})
    bold.set_left()
    bold.set_right()
    bold.set_font_size(10)
    bold.set_font_name('Bahnschrift Light SemiCondensed')

    anmerkungen = workbook.add_format({'bold':True})
    anmerkungen.set_font_size(10)
    anmerkungen.set_font_name('Bahnschrift Light SemiCondensed')

    normal = workbook.add_format({'font_size': 10, 'font_name': 'Bahnschrift Light SemiCondensed'})

    green = workbook.add_format({'font_color':'green'})

    title = workbook.add_format({'align': 'center', 'bold':  True, 'valign': 'vcenter', 'font_size': 16, 'font_name': 'Bahnschrift Light SemiCondensed'})

    center = workbook.add_format({'font_size': 10, 'font_name': 'Bahnschrift Light SemiCondensed', 'align': 'center'})

    topCellBorder = workbook.add_format()
    topCellBorder.set_top()
    topCellBorder.set_left()
    topCellBorder.set_right()
    topCellBorder.set_align('right')
    topCellBorder.set_align('vcenter')
    topCellBorder.set_bg_color('#C6EFCE')
    topCellBorder.set_font_size(10)
    topCellBorder.set_font_name('Bahnschrift Light SemiCondensed')

    grey = workbook.add_format()
    grey.set_left()
    grey.set_right()
    grey.set_font_size(10)
    grey.set_bg_color('#D6D6D6')
    grey.set_font_name('Bahnschrift Light SemiCondensed')

    betweenCellBorder = workbook.add_format()
    betweenCellBorder.set_left()
    betweenCellBorder.set_right()
    betweenCellBorder.set_font_size(10)
    betweenCellBorder.set_font_name('Bahnschrift Light SemiCondensed')

    bottomCellBorder = workbook.add_format()
    bottomCellBorder.set_bottom()
    bottomCellBorder.set_left()
    bottomCellBorder.set_right()
    bottomCellBorder.set_font_size(10)
    bottomCellBorder.set_font_name('Bahnschrift Light SemiCondensed')

    grid = createTree(personen)
    #print(grid)

    HORIZONTAL = '─'
    LEFTCORNER = '└'
    RIGHTCORNER = '┘'
    INTERSECTION = '┬'

    amountOfAttributes = 9

    amountOfBrosAndSis = 0
    for i in range(len(personen.keys())):
        #print(len(list(filter(None, personen.get(list(personen.keys())[i]).get('Geschwister').replace('; ', ',').replace(';', ',').replace(', ', ',').split(',')))))
        if len(list(filter(None, personen.get(list(personen.keys())[i]).get('Geschwister').replace('; ', ',').replace(';', ',').replace(', ', ',').split(',')))) > amountOfBrosAndSis:
            amountOfBrosAndSis = len(list(filter(None, personen.get(list(personen.keys())[i]).get('Geschwister').replace('; ', ',').replace(';', ',').replace(', ', ',').split(','))))

    if amountOfAttributes >= amountOfBrosAndSis:
        CELLHEIGHT = amountOfAttributes
    else:
        CELLHEIGHT = amountOfBrosAndSis

    IMAGEHEIGHT = 4
    CELLWIDTH = 1
    #CELLHEIGHT = IMAGEHEIGHT + additive  # TODO: W.I.P.
    #CELLHEIGHT = 9

    SPACE = ' '

    SP = 0
    LC = -1
    RC = -2
    IS = -3
    HL = -4

    sizedGrid = np.ndarray.tolist(np.asarray(grid))
    for y in range(len(sizedGrid)):
        for x in range(len(sizedGrid[y])):
            if sizedGrid[y][x] == ' '*len(sizedGrid[y][x]):
                SPACE = sizedGrid[y][x]
                sizedGrid[y][x] = SP
            elif sizedGrid[y][x].replace(' ', '') == LEFTCORNER:
                sizedGrid[y][x] = LC
            elif sizedGrid[y][x].replace(' ', '') == RIGHTCORNER:
                sizedGrid[y][x] = RC
            elif sizedGrid[y][x].replace(' ', '') == INTERSECTION:
                sizedGrid[y][x] = IS
            elif sizedGrid[y][x].replace(' ', '') == HORIZONTAL:
                sizedGrid[y][x] = HL
            elif type(sizedGrid[y][x]) == str:
                sizedGrid[y][x] = int(sizedGrid[y][x])

    #print(sizedGrid)
    sizedGrid = np.kron(np.asarray(sizedGrid), np.ones((CELLHEIGHT, CELLWIDTH)))
    #print(sizedGrid)
    sizedGrid = np.ndarray.tolist(sizedGrid)

    for y in range(len(sizedGrid)):
        for x in range(len(sizedGrid[y])):
            if sizedGrid[y][x] == SP:
                sizedGrid[y][x] = SPACE
            elif sizedGrid[y][x] == LC:
                sizedGrid[y][x] = LEFTCORNER
            elif sizedGrid[y][x] == RC:
                sizedGrid[y][x] = RIGHTCORNER
            elif sizedGrid[y][x] == IS:
                sizedGrid[y][x] = INTERSECTION
            elif sizedGrid[y][x] == HL:
                sizedGrid[y][x] = HORIZONTAL
            elif type(sizedGrid[y][x]) == float:
                sizedGrid[y][x] = str(int(sizedGrid[y][x]))

    np.asarray(replaceEmptyWithSpaceInList(sizedGrid))
    #print(np.asarray(replaceEmptyWithSpaceInList(sizedGrid)), '\n')  # debug
    #print(grid)

    for i in range(len(grid)):
        grid[i].append(' ')
    for i in range(len(sizedGrid)):
        sizedGrid[i].append(' ')

    for Y in range(len(grid)):
        for X in range(len(grid[Y])):
            if grid[Y][X] in ['', ' '*len(str(grid[Y][X]))]:
                if str(grid[Y][X-1]).replace(' ', '').isnumeric and not str(grid[Y][X-1]).replace(' ', '') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y][X-1]))]:
                    iteration = 0
                    for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                        if not personen.get(str(grid[Y][X-1]).replace(' ', '')).get('Geschwister') in ['', ' '*len(personen.get(str(grid[Y][X-1]).replace(' ', '')).get('Geschwister'))]:
                            for k in range(len(list(str(personen.get(str(grid[Y][X-1].replace(' ', ''))).get('Geschwister')).replace(';',',').replace(', ',',').split(',')))):
                                try:
                                    sizedGrid[i][X] = ' ' + list(str(personen.get(str(grid[Y][X-1].replace(' ', ''))).get('Geschwister')).replace(';',',').replace(', ',',').split(','))[iteration]
                                except IndexError:
                                    pass
                        iteration += 1
            elif grid[Y][X].replace(' ','') == LEFTCORNER:
                iteration = 0
                for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                    if iteration == 0:
                        sizedGrid[i][X] = '[Leftcorner.png]'
                    else:
                        sizedGrid[i][X] = ' '
                    iteration += 1
            elif grid[Y][X].replace(' ','') == RIGHTCORNER:
                iteration = 0
                for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                    if iteration == 0:
                        sizedGrid[i][X] = '[Rightcorner.png]'
                    else:
                        sizedGrid[i][X] = ' '
                    iteration += 1
            elif grid[Y][X].replace(' ','') == INTERSECTION:
                iteration = 0
                for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                    if iteration == 0:
                        sizedGrid[i][X] = '[Intersection.png]'
                    elif iteration == 1:
                        id1 = str(int(grid[Y + 1][X]) * 2)
                        #get person up then left
                        # for m in range(X + 1, 0, -1):
                        #     if not str(grid[Y - 1][m]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y - 1][m]).replace(' ',''))]:
                        #         id1 = grid[Y - 1][m].replace(' ','')
                        id2 = str(int(grid[Y + 1][X]) * 2 + 1)
                        #get person up then right
                        # for m in range(X + 1, len(grid[Y - 1]), 1):
                        #     if not str(grid[Y - 1][m]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y - 1][m]).replace(' ',''))]:
                        #         id2 = grid[Y - 1][m].replace(' ','')

                        if not personen.get(str(id1)).get('weitereHeiraten') in ['', ' ' * len(personen.get(str(id1)).get('weitereHeiraten'))] or not personen.get(str(id2)).get('weitereHeiraten') in ['', ' ' * len(personen.get(str(id2)).get('weitereHeiraten'))]:  # weitere Ehe
                            sizedGrid[i][X] = '[center]' + str(personen.get(str(id1)).get('weitereHeiraten')) + str(personen.get(str(id2)).get('weitereHeiraten'))
                        else:
                            sizedGrid[i][X] = ' '
                    elif iteration == 2:
                        id = str(int(grid[Y + 1][X]) * 2)
                        #get person up then left
                        # for m in range(X + 1, 0, -1):
                        #     if not str(grid[Y - 1][m]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y - 1][m]).replace(' ',''))]:
                        #         id = grid[Y - 1][m].replace(' ','')

                        sizedGrid[i][X] = '[center]oo ' + str(personen.get(str(id)).get('Heiratsdatum')) + ' | ' + str(personen.get(str(id)).get('Heiratsort'))
                    else:
                        sizedGrid[i][X] = ' '
                    iteration += 1
            elif grid[Y][X].replace(' ','') == HORIZONTAL:
                iteration = 0
                for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                    if iteration == 0:
                        sizedGrid[i][X] = '[Horizontal.png]'
                    else:
                        sizedGrid[i][X] = ' '
                    iteration += 1
            elif not str(grid[Y][X]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y][X]).replace(' ',''))]:
                #print(grid[Y][X])
                ID = int(grid[Y][X])
                #print(ID)
                attribute = personen.get(str(ID))
                #print(attribute)
                iteration = 0
                for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                    if iteration == 0:
                        sizedGrid[i][X] = '[topCellBorder][person.jpg]      ' + str(ID)
                    elif iteration in [1, 2, 3]:
                        sizedGrid[i][X] = '[grey]'
                    elif iteration == 4:
                        sizedGrid[i][X] = '[betweenCellBorder][bold]' + str(attribute.get('Name').split(' ')[1])
                    elif iteration == 5:
                        sizedGrid[i][X] = '[betweenCellBorder]' + str(attribute.get('Name').split(' ')[0])
                    elif iteration == 6:
                        sizedGrid[i][X] = '[betweenCellBorder]' + str(attribute.get('Geburtsdatum') + ' - ' + attribute.get('Sterbedatum'))
                    elif iteration == 7:
                        sizedGrid[i][X] = '[betweenCellBorder]' + str('* ' + attribute.get('Geburtsort'))
                    elif iteration == 8:
                        sizedGrid[i][X] = '[bottomCellBorder]' + str('† ' + attribute.get('Sterbeort'))
                    iteration += 1

    np.asarray(replaceEmptyWithSpaceInList(sizedGrid))
    #print(np.asarray(replaceEmptyWithSpaceInList(sizedGrid)), '\n')  # debug


    # config = configparser.ConfigParser()
    # config.read('settings.ini')
    # MTOP = config['EXCEL']['topmargin']
    # MLEFT = config['EXCEL']['leftmargin']
    MTOP = 1
    MLEFT = 1

    MARGIN_COL = int(MLEFT)
    MARGIN_ROW = 2 + int(MTOP)

    # Write the values
    for row in range(len(sizedGrid)):
        for col in range(len(sizedGrid[row])):
            if '[Horizontal.png]' in sizedGrid[row][col]:
                worksheet.insert_image(row + MARGIN_ROW, col + MARGIN_COL, './img/Horizontal.png', {'object_position': 1, 'x_scale': 0.501})
            elif '[Intersection.png]' in sizedGrid[row][col]:
                worksheet.insert_image(row + MARGIN_ROW, col + MARGIN_COL, './img/Intersection.png', {'object_position': 1, 'x_scale': 0.501})
            elif '[Rightcorner.png]' in sizedGrid[row][col]:
                worksheet.insert_image(row + MARGIN_ROW, col + MARGIN_COL, './img/Rightcorner.png', {'object_position': 1, 'x_scale': 0.501})
            elif '[Leftcorner.png]' in sizedGrid[row][col]:
                worksheet.insert_image(row + MARGIN_ROW, col + MARGIN_COL, './img/Leftcorner.png', {'object_position': 1, 'x_scale': 0.501})
            elif '[person.jpg]' in sizedGrid[row][col]:
                if '[topCellBorder]' in sizedGrid[row][col]:
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[person.jpg]', '').replace('[topCellBorder]', '').replace(' ', '') + ' ', topCellBorder)
                else:
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[person.jpg]', ''))
                worksheet.insert_image(row + MARGIN_ROW, col + MARGIN_COL, './img/person.jpg', {'x_offset': 1, 'y_offset': 1, 'x_scale':1.05, 'y_scale':1.075})
            else:
                if '[topCellBorder]' in sizedGrid[row][col]:
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[topCellBorder]', ''), topCellBorder)
                elif '[center]' in sizedGrid[row][col]:
                    while sizedGrid[row][col][-1] == ' ':
                        sizedGrid[row][col] = sizedGrid[row][col][:-1]
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[center]', ''), center)
                elif '[grey]' in sizedGrid[row][col]:
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[grey]', ''), grey)
                elif '[betweenCellBorder]' in sizedGrid[row][col]:
                    if '[bold]' in sizedGrid[row][col]:
                        worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[betweenCellBorder]', '').replace('[bold]', ''), bold)
                    else:
                        worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[betweenCellBorder]', ''), betweenCellBorder)
                elif '[bottomCellBorder]' in sizedGrid[row][col]:
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col].replace('[bottomCellBorder]', ''), bottomCellBorder)
                else:
                    worksheet.write(row + MARGIN_ROW, col + MARGIN_COL, sizedGrid[row][col], normal)
    #console.log(f"[green]Created tree[/green]")

    # Create the title
    worksheet.merge_range(1, 1, 1, len(sizedGrid[0]) - 1, 'Stammbaum von ' + personen.get('1').get('Name'), title)
    worksheet.set_row(1, 48)
    #console.log(f"[green]Created title[/green]")

    # Create the infos (Anmerkungen)
    Infos = []
    for i in range(len(personen)):
        if not personen.get(str(i + 1)).get('Anmerkung') in ['', ' ' * len(personen.get(str(i + 1)).get('Anmerkung'))]:
            Infos.append(f'{i + 1}: ' + personen.get(str(i + 1)).get('Anmerkung'))

    MARGIN_TOP = 2

    worksheet.write(MARGIN_TOP, len(sizedGrid[0]) + MARGIN_COL, 'Anmerkungen:', anmerkungen)

    n = 1
    for i in Infos:
        worksheet.write(MARGIN_TOP + n, len(sizedGrid[0]) + MARGIN_COL, i, normal)
        n += 1
    #console.log(f"[green]Created infos[/green]")

    # hide the gridlines.
    worksheet.hide_gridlines(2)
    #console.log(f"[green]Hid gridlines[/green]")

    # Create VBA-AutoFit-Makro, that runs automatically when opening the workbook
    workbook.add_vba_project('./vbaProject.bin')
    #console.log(f"[green]Added vbaProject binary[/green]")


    try:
        workbook.close()
        #console.log('[green]File successfully saved.[/green]')
    except:
        #console.log('[b][red]File couldn\'t be saved. Is the file already open?[/b][/red]')
        return

    #console.log('[yellow]Opening file...[/yellow]')
    os.startfile('Stammbaum von ' + personen.get('1').get('Name') + '.xlsm')
    #console.log('[b][green]File successfully exported.[/b][/green]')
def showVersion():
    messagebox.showinfo("Version", "Stammbaumgenerator Version 1.0.0.0")  # (Major version).(Minor version).(Revision number).(Build number)

menubar = Menu(root)

filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Öffnen...", command=onOpen)
filemenu.add_command(label="Speichern", command=quickSave)
filemenu.add_command(label="Speichern als...", command=onSave)
filemenu.add_command(label="Exportieren", command=exportToExcel)
filemenu.add_command(label="Beenden", command=root.quit)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Häufig gestellete Fragen", command=showFAQ)
helpmenu.add_command(label="Tutorial", command=showTutorial)
helpmenu.add_command(label="Version", command=showVersion)

editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Generation hinzufügen", command=addGeneration)
editmenu.add_command(label="Generation löschen", command=removeGeneration)

menubar.add_cascade(label="Datei", menu=filemenu)
menubar.add_cascade(label="Bearbeiten", menu=editmenu)
menubar.add_cascade(label="Hilfe", menu=helpmenu)

root.config(menu=menubar)
root.bind('<Control-s>', bindSave)

firstframe = Frame(root)
firstframe.pack(fill=BOTH,expand=1)
mainframe = ttk.Frame(firstframe, padding="3 3 12 12")
mainframe.pack(fill=X,side=BOTTOM)
# mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1) 
root.rowconfigure(0, weight=1)  
canvas = Canvas(firstframe)
canvas.pack(side=LEFT,fill=BOTH,expand=1)
secondframe = Frame(canvas)
canvas.create_window((0,0),window= secondframe, anchor="nw")
x_scrollbar = ttk.Scrollbar(mainframe,orient=HORIZONTAL,command=canvas.xview)
x_scrollbar.pack(side=BOTTOM,fill=X)

y_scrollbar = ttk.Scrollbar(firstframe,orient=VERTICAL,command=canvas.yview)
y_scrollbar.pack(side=RIGHT,fill=Y)

# canvas.configure(xscrollcommand=x_scrollbar.set)

# canvas.configure(yscrollcommand=y_scrollbar.set)

# canvas.bind("<Configure>",lambda e: canvas.config(scrollregion= canvas.bbox(ALL)))

refreshTree()

# Polish
# for child in mainframe.winfo_children(): 
#             child.grid_configure(padx=0, pady=0)

mainloop()