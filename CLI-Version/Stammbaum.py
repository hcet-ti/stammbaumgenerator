import math
import json
from time import sleep
import os
import configparser


config = configparser.ConfigParser()
config.read('settings.ini')
if config['FIRSTLOAD']['firstexecution'] == 'yes':
    os.system('setup.py')
    config['FIRSTLOAD']['firstexecution'] = 'no'  # TODO: Uncomment after done with app
    with open('settings.ini', 'w') as configfile:
        config.write(configfile)


import numpy as np
from rich.highlighter import RegexHighlighter
from rich.theme import Theme
from rich.console import Console
import xlsxwriter


class TreeHighlighter(RegexHighlighter):
    """Apply style to any number."""

    base_style = "num"
    highlights = [r"(?P<numbers>[\d])"]

console = Console()

menu_options = {
    1: 'Create new Person',
    2: 'Inspect/Edit Person',
    3: 'Delete Person',
    4: 'View full tree',
    5: 'Save project',
    6: 'Export to excel file',
    7: 'Exit',
}

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

def print_tree(IDgens):
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

    np.asarray(replaceEmpty(grid))
    #print(np.asarray(replaceEmpty(grid)), '\n')  # debug

    # Create the grid with IDs replaced by 'X'
    config = configparser.ConfigParser()
    config.read('settings.ini')
    CELLSYMBOL = 'X'
    LINESYMBOL = config['TREE']['LINESYMBOL']
    INTERSECTIONSYMBOL = config['TREE']['INTERSECTIONSYMBOL']

    # Create the X's at the bottom
    for x in range(IDs):
        if (x % 2) == 0:
            # x is even
            grid[generations - 1][x] = CELLSYMBOL

    np.asarray(replaceEmpty(grid))
    #print(np.asarray(replaceEmpty(grid)), '\n')  # debug

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

    np.asarray(replaceEmpty(grid))
    #print(np.asarray(replaceEmpty(grid)), '\n')  # debug

    # Create the lines, that connect the 'X's
    for i in range(1, generations, 1):
        start = -1
        for n in range(IDs):
            if grid[i][n] == CELLSYMBOL:
                if start == -1:
                    start = n
                else:
                    line = range(start + 1, n)
                    for k in line:
                        grid[i][k] = LINESYMBOL
                    start = -1  # reset start

    np.asarray(replaceEmpty(grid))
    #print(np.asarray(replaceEmpty(grid)), '\n')  # debug

    # Create the intersections
    for i in range(1, generations, 1):
        start = -1
        for n in range(IDs):
            if grid[i][n] == CELLSYMBOL:
                if start == -1:
                    start = n
                else:
                    middle = int(start + (n - start) / 2)  # get middle of start and n
                    grid[i][middle] = INTERSECTIONSYMBOL  # create the X in the middle of start and n one generation above the two of them
                    start = -1  # reset start

    np.asarray(replaceEmpty(grid))
    #print(np.asarray(replaceEmpty(grid)), '\n')  # debug

    # Replace the X's with the IDs
    ids = list(range(1, IDs + 1, 1))
    for i in range(generations):
        for n in range(IDs):
            if grid[i][n] == CELLSYMBOL:
                grid[i][n] = str(ids[0])
                ids.pop(0)

    replaceEmpty(grid)
    np.asarray(replaceEmpty(grid))
    #print(np.asarray(replaceEmpty(grid)), '\n')  # debug

    # If the IDs go past 1 digit, create new characters to fill the space
    highest = []
    for i in range(generations):
        highest.append(max(grid[i], key=len))
    highest = max(highest, key=len)
    #print("highest digits in a number were:", len(highest))  # debug

    for i in range(generations):
        for n in range(IDs):
            if len(grid[i][n]) < len(highest):
                if grid[i][n] == ' ':
                    grid[i][n] = ' ' * (len(highest))
                elif grid[i][n] == LINESYMBOL:
                    grid[i][n] = LINESYMBOL * (len(highest))
                elif grid[i][n] == INTERSECTIONSYMBOL:
                    if (len(highest) % 2) != 0:
                        grid[i][n] = roundDown(len(highest)/2) * LINESYMBOL + INTERSECTIONSYMBOL + LINESYMBOL * roundUp(len(highest)/2)
                    else:
                        grid[i][n] = (int(len(highest)/2) - 1) * LINESYMBOL + INTERSECTIONSYMBOL * 2 + LINESYMBOL * (int(len(highest)/2) - 1)
                elif grid[i][n] == '1':
                    grid[i][n] = roundDown(len(highest)/2) * ' ' + '1' + ' ' * roundUp(len(highest)/2)
                elif (int(grid[i][n]) % 2) == 0:
                    grid[i][n] = grid[i][n] + LINESYMBOL * (len(highest) - len(max(grid[i][n], key=len)))
                elif (int(grid[i][n]) % 2) != 0:
                    grid[i][n] = (len(highest) - len(max(grid[i][n],key=len))) * LINESYMBOL + grid[i][n]

    # print the tree
    theme = Theme({"numnumbers": "bold cyan"})
    treeConsole = Console(highlighter=TreeHighlighter(), theme=theme)
    Out = ''
    for i in range(generations - 1, -1, -1):
        for n in range(IDs):
            Out = Out + grid[i][n]
        treeConsole.print((MARGIN - int(len(Out) / 2)) * ' ' + Out)
        Out = ''

def print_menu():
    for key in menu_options.keys():
        console.print (key, '--', '[b]' + menu_options[key] + '[/b]')

def option1(Personen):
    os.system('cls')
    print('IDs, die es schon gibt: ', end='')
    for i in range(len(Personen.keys())):
        if i != len(Personen.keys()) - 1:
            print(str(list([int(x) for x in Personen.keys()])[i]) + ', ', end='')
        else:
            print(str(list(Personen.keys())[i]))
    ID = 0
    while ID < 2:
        ID = input("ID: ")
        if ID in ['' ' '*len(str(ID))]:
            os.system('cls')
            return
        try:
            ID = int(ID)
            if str(ID) in Personen.keys():
                print("ID already exists!")
                ID = 0
            elif ID < 2:
                print('ID is invalid')
        except:
            print("Not a number.")
            ID = 0
    Vorname = input("Vorname: ").replace(' ', '')  # Heinz
    if Vorname == '':
        Vorname = '-'
    Nachname = input("Nachname: ").replace(' ', '')  # Heinz
    if Nachname == '':
        Nachname = '-'
    #MännlichWeiblich = input("Männlich oder weiblich: ")
    #Eltern = input("Eltern: ")  # 5, 4
    #Geheiratet = input("relevanter Ehemann/relevante Ehefrau: ")  # 6
    Heiratsdatum = input("Heiratsdatum: ").replace(' ', '')  # DD-MM-YYYY
    Heiratsort = input("Heiratsort: ")  # Pfarrkirche Eisenstadt 1657
    weitereHeiraten = input("weitere Ehemänner/Ehefrauen: ")  # 6, 7, 4
    Geschwister = input("Geschwister: ")  # hier gehen keine IDs, sondern nur Namen´
    Geburtsort = input("Geburtsort: ")  # Eisenbach 7893
    Sterbeort = input("Sterbeort: ")  # Eisenstadt 1657
    Geburtsdatum = input("Geburtsdatum: ").replace(' ', '')  # DD-MM-YYYY
    Sterbedatum = input("Sterbedatum: ").replace(' ', '')  # DD-MM-YYYY
    Anmerkung = input("Anmerkung: ")
    Personen.update({f'{ID}':{'Name':Vorname + ' ' + Nachname, 'Heiratsdatum':Heiratsdatum, 'Heiratsort':Heiratsort, 'weitereHeiraten':weitereHeiraten, 'Geschwister':Geschwister, 'Geburtsort':Geburtsort, 'Sterbeort':Sterbeort, 'Geburtsdatum':Geburtsdatum, 'Sterbedatum':Sterbedatum, 'Anmerkung':Anmerkung}})  # int(sorted(Personen.keys())[-1])+mw2Num(MännlichWeiblich)
    os.system('cls')
    print(f"{Vorname + ' ' + Nachname} wurde kreiert.\n")
    #print(Personen)  # debug

def option2(Personen):
    os.system('cls')
    print('IDs: ', end='')
    for i in range(len(Personen.keys())):
        if i != len(Personen.keys()) - 1:
            print(str(list([int(x) for x in Personen.keys()])[i]) + ', ', end='')
        else:
            print(str(list(Personen.keys())[i]))
    proceed = True
    while proceed:
        person = input('ID: ')
        payload = Personen.get(person)
        if payload == None:
            print('This ID does not exist.')
        else:
            proceed = False
    print()
    #console.print(payload)  # debug
    while True:
        for i in range(len(payload)):
            key = list(payload.keys())[i]
            val = payload.get(key)
            console.print(str(i + 1) + ' ' + '-' * ((len(str(len(payload))) + 2) - len(str(i + 1))) + ' [b]' + key + ':[/b] ' + val)
        print('\nChoose an attribute to change. Type nothing to exit.')
        choice = input('Edit: ')
        try:
            if choice in ['', ' ', '  ', '   ', '    ', '     ']:
                os.system('cls')
                return
            elif int(choice) < (len(payload) + 1):
                choice = int(choice)
                if list(payload)[choice - 1] == 'Name':
                    Vorname = input("Vorname: ").replace(' ', '')  # Heinz
                    if Vorname == '':
                        Vorname = '-'
                    Nachname = input("Nachname: ").replace(' ', '')  # Heinz
                    if Nachname == '':
                        Nachname = '-'
                    arg = Vorname + ' ' + Nachname
                else:
                    arg = input(str(list(payload)[choice - 1]) + ': ')
                payload[list(payload)[choice - 1]] = arg
        except:
            print('Invalid option.')

def option3(Personen):
    os.system('cls')
    print('IDs: ', end='')
    for i in range(len(Personen.keys())):
        if i != len(Personen.keys()) - 1:
            print(str(list([int(x) for x in Personen.keys()])[i]) + ', ', end='')
        else:
            print(str(list(Personen.keys())[i]))
    Proceed = True
    while Proceed:
        try:
            print('Delete which person?')
            person = input('ID: ')
            if person in ['', ' '*len(str(person))]:
                os.system('cls')
                return
            name = Personen.get(person).get('Name')
            del Personen[person]
            Proceed = False
        except:
            console.print('[b][red]That ID doesn\'t exist[/b][/red]')
    os.system('cls')
    console.print('[green][b]Succesfully deleted [white]{}[/white][green][/b]'.format(name))

def option4(Personen):
    os.system('cls')
    # print(math.log(len(Personen.keys()), 2))  # debug
    # print(int(math.log(len(Personen.keys()), 2)))  # debug
    # print(math.log(len(Personen.keys()), 2) - int(math.log(len(Personen.keys()), 2)))  # debug
    # print(CheckFloat(math.log(len(Personen.keys()), 2)))  # debug
    # if CheckFloat(math.log(len(Personen.keys()), 2)):  # TODO: this checker is broken and I don't know why
    #     print('ERROR: Amount of People is not a power of 2!')
    #     return
    stammbaum = Stammbaum
    #print(Personen)  # debug
    #print('len(Personen.keys(): ', len(Personen.keys()))  # debug
    amountGenerations = int(math.log(int(len(Personen.keys())), 2))  # int(sorted(Personen.keys())[-1])
    #print('amountGenerations: ', amountGenerations)  # debug
    alteredPpl = ''
    for i in range(len(Personen.keys())):
        if CheckFloat(math.log(int(sorted(Personen.keys())[i]), 2)):  # if the Person ID is a power of 2
            alteredPpl = alteredPpl + '|'
            if alteredPpl[-1] == '|':
                alteredPpl = alteredPpl + str(sorted(Personen.keys())[i])
            else:
                alteredPpl = alteredPpl + ', ' + str(sorted(Personen.keys())[i])
        else:
            if alteredPpl[-1] == '|':
                alteredPpl = alteredPpl + str(sorted(Personen.keys())[i])
            else:
                alteredPpl = alteredPpl + ', ' + str(sorted(Personen.keys())[i])
    alteredPpl = alteredPpl[1:]
    alteredPpl = alteredPpl.split('|')
    #print(alteredPpl)  # debug

    """additiveList = []
    splitList = []
    for i in range(len(Personen.keys())):
        if CheckFloat(math.log(sorted(Personen.keys())[i], 2)):  # if the Person ID is a power of 2
            additiveList.append(sorted(Personen.keys())[i])
            splitList.append(additiveList)
            additiveList = []
        else:
            additiveList.append(sorted(Personen.keys())[i])"""

    for i in range(amountGenerations+1):
        # 0, 1
        # i is the current generation
        stammbaum.generations.update({i:alteredPpl[i].replace(' ', '').split(',')})
    """for i in range(amountGenerations+1):
        for n in range(i, 2**i):
            n += 1
            print('n:',  n)
            stammbaum.generations.update({i:{n:Personen.get(str(n))}})"""
    #print(stammbaum.generations)  # debug
    console.print('[u][b]Stammbaum von {}[/b][/u]\n'.format(Personen.get('1').get('Name')), justify='center')
    # underline = ''
    # for i in range(len('Stammbaum von {}'.format(Personen.get('1').get('Name')))):
    #     underline = underline + '‾'
    # console.print(underline, justify='center')
    print_tree(stammbaum.generations)
    print()
    input('Type any key to continue...')
    os.system('cls')

def option5(Personen):
    os.system('cls')
    ProceedSave = True
    while ProceedSave:
        try:
            with open(os.path.join(os.path.dirname(__file__), input('Speicherort: ')), 'w') as output_file:
                output_file.write(json.dumps(Personen))
                os.system('cls')
                console.log('[b][green]File successfully saved.[/b][/green]')
                ProceedSave = False
        except:
            console.print('[b][red]ERROR: Couldn\'t save file.[/b][/red]')

def create_tree(Personen):
    stammbaum = Stammbaum
    amountGenerations = int(math.log(int(len(Personen.keys())), 2))  # int(sorted(Personen.keys())[-1])
    alteredPpl = ''
    for i in range(len(Personen.keys())):
        if CheckFloat(math.log(int(sorted(Personen.keys())[i]), 2)):  # if the Person ID is a power of 2
            alteredPpl = alteredPpl + '|'
            if alteredPpl[-1] == '|':
                alteredPpl = alteredPpl + str(sorted(Personen.keys())[i])
            else:
                alteredPpl = alteredPpl + ', ' + str(sorted(Personen.keys())[i])
        else:
            if alteredPpl[-1] == '|':
                alteredPpl = alteredPpl + str(sorted(Personen.keys())[i])
            else:
                alteredPpl = alteredPpl + ', ' + str(sorted(Personen.keys())[i])
    alteredPpl = alteredPpl[1:]
    alteredPpl = alteredPpl.split('|')

    for i in range(amountGenerations+1):
        stammbaum.generations.update({i:alteredPpl[i].replace(' ', '').split(',')})
    
    IDgens = stammbaum.generations
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
    #print("highest digits in a number were:", len(highest))  # debug

    # for i in range(generations):
    #     for n in range(IDs):
    #         if len(grid[i][n]) < len(highest):
    #             if grid[i][n] == ' ':
    #                 grid[i][n] = ' ' * (len(highest))
    #             elif grid[i][n] == LINESYMBOL:
    #                 grid[i][n] = LINESYMBOL * (len(highest))
    #             elif grid[i][n] == INTERSECTIONSYMBOL:
    #                 if (len(highest) % 2) != 0:
    #                     grid[i][n] = roundDown(len(highest)/2) * LINESYMBOL + INTERSECTIONSYMBOL + LINESYMBOL * roundUp(len(highest)/2)
    #                 else:
    #                     grid[i][n] = (int(len(highest)/2) - 1) * LINESYMBOL + INTERSECTIONSYMBOL * 2 + LINESYMBOL * (int(len(highest)/2) - 1)
    #             elif grid[i][n] == '1':
    #                 grid[i][n] = roundDown(len(highest)/2) * ' ' + '1' + ' ' * roundUp(len(highest)/2)
    #             elif (int(grid[i][n]) % 2) == 0:
    #                 grid[i][n] = grid[i][n] + LINESYMBOL * (len(highest) - len(max(grid[i][n], key=len)))
    #             elif (int(grid[i][n]) % 2) != 0:
    #                 grid[i][n] = (len(highest) - len(max(grid[i][n],key=len))) * LINESYMBOL + grid[i][n]

    # print the tree
    grid = grid[::-1]  # this inverts the list
    return grid

def option6(Personen):
    os.system('cls')
    with console.status("[bold green]Exporting...") as status:
        workbook = xlsxwriter.Workbook('Stammbaum von ' + Personen.get('1').get('Name') + '.xlsm')
        workbook.set_vba_name('DieseArbeitsMappe')
        console.log(f"[green]Created workbook[/green]")
        worksheet = workbook.add_worksheet()
        worksheet.set_vba_name('Tabelle1')
        console.log(f"[green]Created worksheet[/green]")

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

        # Create the tree
        grid = create_tree(Personen)

        # Create lines between the nums
        line = []
        for i in range(len(grid[0])):
            line.append('')
        n = 0
        while n in range(len(grid)):
            if (n % 2) != 0:
                grid.insert(n, line)
            n += 1


        config = configparser.ConfigParser()
        config.read('settings.ini')
        HORIZONTAL = config['TREE']['HORIZONTAL']
        LEFTCORNER = config['TREE']['LEFTCORNER']
        RIGHTCORNER = config['TREE']['RIGHTCORNER']
        INTERSECTION = config['TREE']['INTERSECTION']

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
        #print(np.asarray(replaceEmptyWithSpaceInList(grid)), '\n')  # debug

        # Scale grid according to Cellwidth and Cellheight.
        # amountOfAttributes = 0
        # for i in range(len(Personen.keys())):
        #     print(len(Personen.get(list(Personen.keys())[i]).keys()))
        #     if len(Personen.get(list(Personen.keys())[i]).keys()) > amountOfAttributes:
        #         amountOfAttributes = len(Personen.get(list(Personen.keys())[i]).keys())
        amountOfAttributes = 9

        amountOfBrosAndSis = 0
        for i in range(len(Personen.keys())):
            #print(len(list(filter(None, Personen.get(list(Personen.keys())[i]).get('Geschwister').replace('; ', ',').replace(';', ',').replace(', ', ',').split(',')))))
            if len(list(filter(None, Personen.get(list(Personen.keys())[i]).get('Geschwister').replace('; ', ',').replace(';', ',').replace(', ', ',').split(',')))) > amountOfBrosAndSis:
                amountOfBrosAndSis = len(list(filter(None, Personen.get(list(Personen.keys())[i]).get('Geschwister').replace('; ', ',').replace(';', ',').replace(', ', ',').split(','))))

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
                    if str(grid[Y][X-1]).isnumeric and not str(grid[Y][X-1]).replace(' ', '') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y][X-1]))]:
                        iteration = 0
                        for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                            if not Personen.get(str(grid[Y][X-1]).replace(' ', '')).get('Geschwister') in ['', ' '*len(Personen.get(str(grid[Y][X-1]).replace(' ', '')).get('Geschwister'))]:
                                for k in range(len(list(str(Personen.get(str(grid[Y][X-1].replace(' ', ''))).get('Geschwister')).replace(';',',').replace(', ',',').split(',')))):
                                    try:
                                        sizedGrid[i][X] = list(str(Personen.get(str(grid[Y][X-1].replace(' ', ''))).get('Geschwister')).replace(';',',').replace(', ',',').split(','))[iteration]
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
                            #get person up then left
                            for m in range(X + 1, 0, -1):
                                if not str(grid[Y - 1][m]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y - 1][m]).replace(' ',''))]:
                                    id1 = grid[Y - 1][m].replace(' ','')
                            #get person up then right
                            for m in range(X + 1, len(grid[Y - 1]), 1):
                                if not str(grid[Y - 1][m]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y - 1][m]).replace(' ',''))]:
                                    id2 = grid[Y - 1][m].replace(' ','')

                            if not Personen.get(str(id1)).get('weitereHeiraten') in ['', ' ' * len(Personen.get(str(id1)).get('weitereHeiraten'))] or not Personen.get(str(id2)).get('weitereHeiraten') in ['', ' ' * len(Personen.get(str(id2)).get('weitereHeiraten'))]:  # weitere Ehe
                                sizedGrid[i][X] = '[center]' + str(Personen.get(str(id1)).get('weitereHeiraten')) + str(Personen.get(str(id2)).get('weitereHeiraten'))
                            else:
                                sizedGrid[i][X] = ' '
                        elif iteration == 2:
                            #get person up then left
                            for m in range(X + 1, 0, -1):
                                if not str(grid[Y - 1][m]).replace(' ','') in [LEFTCORNER, INTERSECTION, HORIZONTAL, RIGHTCORNER, '', ' '*len(str(grid[Y - 1][m]).replace(' ',''))]:
                                    id = grid[Y - 1][m].replace(' ','')

                            sizedGrid[i][X] = '[center]oo ' + str(Personen.get(str(id)).get('Heiratsdatum')) + ' | ' + str(Personen.get(str(id)).get('Heiratsort'))
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
                    attribute = Personen.get(str(ID))
                    #print(attribute)
                    iteration = 0
                    for i in range(Y * CELLHEIGHT, (Y + 1) * CELLHEIGHT, 1):
                        if iteration == 0:
                            sizedGrid[i][X] = '[topCellBorder][person.jpg]' + str(ID)
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

        MTOP = config['EXCEL']['topmargin']
        MLEFT = config['EXCEL']['leftmargin']

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
                    worksheet.insert_image(row + MARGIN_ROW, col + MARGIN_COL, './img/person.jpg', {'x_offset': 1, 'y_offset': 1, 'x_scale':1.05, 'y_scale':1.05})
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
        console.log(f"[green]Created tree[/green]")

        # Create the title
        worksheet.merge_range(1, 1, 1, len(sizedGrid[0]) - 1, 'Stammbaum von ' + Personen.get('1').get('Name'), title)
        worksheet.set_row(1, 48)
        console.log(f"[green]Created title[/green]")

        # Create the infos (Anmerkungen)
        Infos = []
        for i in range(len(Personen)):
            if not Personen.get(str(i + 1)).get('Anmerkung') in ['', ' ' * len(Personen.get(str(i + 1)).get('Anmerkung'))]:
                Infos.append(f'{i + 1}: ' + Personen.get(str(i + 1)).get('Anmerkung'))

        MARGIN_TOP = 2

        worksheet.write(MARGIN_TOP, len(sizedGrid[0]) + MARGIN_COL, 'Anmerkungen:', anmerkungen)

        n = 1
        for i in Infos:
            worksheet.write(MARGIN_TOP + n, len(sizedGrid[0]) + MARGIN_COL, i, normal)
            n += 1
        console.log(f"[green]Created infos[/green]")

        # hide the gridlines.
        worksheet.hide_gridlines(2)
        console.log(f"[green]Hid gridlines[/green]")

        # Create VBA-AutoFit-Makro, that runs automatically when opening the workbook
        workbook.add_vba_project('./vbaProject.bin')
        console.log(f"[green]Added vbaProject binary[/green]")


        try:
            workbook.close()
            console.log('[green]File successfully saved.[/green]')
        except:
            console.log('[b][red]File couldn\'t be saved. Is the file already open?[/b][/red]')
            return

        console.log('[yellow]Opening file...[/yellow]')
        os.startfile('Stammbaum von ' + Personen.get('1').get('Name') + '.xlsm')
        sleep(5)
        os.system('cls')
        console.log('[b][green]File successfully exported.[/b][/green]')

class Person:  # NOTE: This isn't used anywhere in the Code.
    ID = 0                # 2
    MännlichWeiblich = "" # "m" / "w"
    Geheiratet = 0        # 3
    weitereHeiraten = []  # [10, 11, 12]
    Eltern = []           # [8, 7]
    Geschwister = []      # [9, 6, 4]                
    Name = ""             # Heinrich Josef
    Geburtsort = ""       # Eisenstadt 1657
    Sterbeort = ""        # Eisenstadt 1657
    Geburstdatum = ""     # DD-MM-YYYY
    Sterbedatum = ""      # DD-MM-YYYY
    Heiratsdatum = ""     # DD-MM-YYYY
    Heiratsort = ""       # Pfarrkirche Eisenstadt 1657
    Anmerkung = ""        #

class Stammbaum:
    generations = {} # {0:[1], 1:[2, 3], 2:[4, 5, 6, 7]}

Personen = {}  # dictionary
StartStop = True

os.system('cls')
console.print('[b][white]Stammbaumgenerator v20.6.2022[/white][/b]\n')
ProceedSave = True
while ProceedSave:
    console.print('Projekt laden?')
    if console.input('[b][green]Y[/green]/[red]N[/red][/b]: ') in ['y', 'Y', 'yes', 'Yes', 'YES', 'J', 'Ja', 'ja', 'JA', 'j']:
        try: 
            with open(os.path.join(os.path.dirname(__file__), input('Dateipfad: ')), 'r') as input_file:
                content = json.loads(input_file.read())
                Personen = content
                ProceedSave = False
        except:
            console.print('[b][red]ERROR: Couldn\'t read file.[/b][/red]')
    else:
        ProceedSave = False

    #print(type(Personen))  # debug
    #print(Personen)  # debug

os.system('cls')
while StartStop:
    if len(Personen) == 0:
        console.print('\n[b][u]Ich[/u][/b]', justify='center')
        Vorname = input("Vorname: ").replace(' ', '')  # Heinz
        if Vorname == '':
            Vorname = '-'
        Nachname = input("Nachname: ").replace(' ', '')  # Heinz
        if Nachname == '':
            Nachname = '-'
        #MännlichWeiblich = input("Männlich oder weiblich: ")
        #Eltern = input("Eltern: ")  # 5, 4
        #Geheiratet = input("relevanter Ehemann/relevante Ehefrau: ")  # 6
        Heiratsdatum = input("Heiratsdatum: ").replace(' ', '')  # DD-MM-YYYY
        Heiratsort = input("Heiratsort: ")  # Pfarrkirche Eisenstadt 1657
        weitereHeiraten = input("weitere Ehemänner/Ehefrauen: ")  # 6, 7, 4
        Geschwister = input("Geschwister: ")  # hier gehen keine IDs, sondern nur Namen´
        Geburtsort = input("Geburtsort: ")  # Eisenbach 7893
        Sterbeort = input("Sterbeort: ")  # Eisenstadt 1657
        Geburtsdatum = input("Geburtsdatum: ").replace(' ', '')  # DD-MM-YYYY
        Sterbedatum = input("Sterbedatum: ").replace(' ', '')  # DD-MM-YYYY
        Anmerkung = input("Anmerkung: ")
        Personen.update({'1':{'Name':Vorname + ' ' + Nachname, 'Heiratsdatum':Heiratsdatum, 'Heiratsort':Heiratsort, 'weitereHeiraten':weitereHeiraten, 'Geschwister':Geschwister, 'Geburtsort':Geburtsort, 'Sterbeort':Sterbeort, 'Geburtsdatum':Geburtsdatum, 'Sterbedatum':Sterbedatum, 'Anmerkung':Anmerkung}})
        console.print(f"[green][b]{Vorname + ' ' + Nachname} wurde kreiert.[/green][/b]")
        #print(Personen)  # debug
        os.system('cls')
    print()
    print_menu()
    print()
    option = ''
    try:
        option = int(input('Enter your choice: '))
    except:
        print('Wrong input. Please enter a number ...')
    #Check what choice was entered and act accordingly
    if option == 1:
        try:
            option1(Personen)
        except:
            os.system('cls')
            print()
            console.log('[b][red]An Error occurred during Execution.[/b][/red]')
    elif option == 2:
        try:
            option2(Personen)
        except:
            os.system('cls')
            print()
            console.log('[b][red]An Error occurred during Execution.[/b][/red]')
    elif option == 3:
        try:        
            option3(Personen)
        except:
            os.system('cls')
            print()
            console.log('[b][red]An Error occurred during Execution.[/b][/red]')
    elif option == 4:
        try:
            option4(Personen)
        except:
            os.system('cls')
            print()
            console.log('[b][red]An Error occurred during Execution.[/b][/red]')
    elif option == 5:
        try:
            option5(Personen)
        except:
            os.system('cls')
            print()
            console.log('[b][red]An Error occurred during Execution.[/b][/red]')
    elif option == 6:
        option6(Personen)
    elif option == 7:
        os.system('cls')
        print('Program wird beendet...')
        exit()
    else:
        print('Invalid option. Please enter a number between 1 and 7.')
