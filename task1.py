
import pandas as pd
from openpyxl.reader.excel import load_workbook

dataframe = pd.read_excel("C:/Users/vruta/OneDrive/Desktop/sutt.xlsx")
wb = load_workbook("C:/Users/vruta/OneDrive/Desktop/sutt.xlsx")
sheet = wb.worksheets[0]

loops = sheet.max_row

dictionary = []
i = 1
for i in range(1, loops):
    dict = {}
    dict['name'] = f"{sheet.cell(i + 1, 2).value}"
    dict['id'] = f"{sheet.cell(i + 1, 1).value}"
    dict['branch'] = ''
    id = list(dict['id'].strip(""))
    if id[6:7] == ['P']:
        if id[4:5] == ['A'] and id[5:6] == ['A']:
            dict['branch'] += 'ECE'
        elif id[4:5] == ['A'] and id[5:6] == ['B']:
            dict['branch'] += 'Manufacturing'
        elif id[4:5] == ['A'] and id[5:6] == ['1']:
            dict['branch'] += 'Chemical'
        elif id[4:5] == ['A'] and id[5:6] == ['2']:
            dict['branch'] += 'Civil'
        elif id[4:5] == ['A'] and id[5:6] == ['3']:
            dict['branch'] += 'EEE'
        elif id[4:5] == ['A'] and id[5:6] == ['4']:
            dict['branch'] += 'Mechanical'
        elif id[4:5] == ['A'] and id[5:6] == ['5']:
            dict['branch'] += 'B Pharma'
        elif id[4:5] == ['A'] and id[5:6] == ['7']:
            dict['branch'] += 'Computer Science'
        elif id[4:5] == ['A'] and id[5:6] == ['8']:
            dict['branch'] += 'ENI'
        elif id[4:5] == ['B'] and id[5:6] == ['1']:
            dict['branch'] += 'Msc Biology'
        elif id[4:5] == ['B'] and id[5:6] == ['2']:
            dict['branch'] += 'Msc Chemistry'
        elif id[4:5] == ['B'] and id[5:6] == ['3']:
            dict['branch'] += 'Msc Economics'
        elif id[4:5] == ['B'] and id[5:6] == ['4']:
            dict['branch'] += 'Msc Maths'
        elif id[4:5] == ['B'] and id[5:6] == ['5']:
            dict['branch'] += 'Msc Physics'
    else:
        if id[4:5] == ['A'] and id[5:6] == ['A']:
            dict['branch'] += 'ECE'
        elif id[4:5] == ['A'] and id[5:6] == ['B']:
            dict['branch'] += 'Manufacturing'
        elif id[4:5] == ['A'] and id[5:6] == ['1']:
            dict['branch'] += 'Chemical'
        elif id[4:5] == ['A'] and id[5:6] == ['2']:
            dict['branch'] += 'Civil'
        elif id[4:5] == ['A'] and id[5:6] == ['3']:
            dict['branch'] += 'EEE'
        elif id[4:5] == ['A'] and id[5:6] == ['4']:
            dict['branch'] += 'Mechanical'
        elif id[4:5] == ['A'] and id[5:6] == ['5']:
            dict['branch'] += 'B Pharma'
        elif id[4:5] == ['A'] and id[5:6] == ['7']:
            dict['branch'] += 'Computer Science'
        elif id[4:5] == ['A'] and id[5:6] == ['8']:
            dict['branch'] += 'ENI'
        elif id[4:5] == ['B'] and id[5:6] == ['1']:
            dict['branch'] += 'Msc Biology'
        elif id[4:5] == ['B'] and id[5:6] == ['2']:
            dict['branch'] += 'Msc Chemistry'
        elif id[4:5] == ['B'] and id[5:6] == ['3']:
            dict['branch'] += 'Msc Economics'
        elif id[4:5] == ['B'] and id[5:6] == ['4']:
            dict['branch'] += 'Msc Maths'
        elif id[4:5] == ['B'] and id[5:6] == ['5']:
            dict['branch'] += 'Msc Physics'

        dict['branch'] += ' + '

        if id[6:7] == ['A'] and id[7:8] == ['A']:
            dict['branch'] += 'ECE'
        elif id[6:7] == ['A'] and id[7:8] == ['B']:
            dict['branch'] += 'Manufacturing'
        elif id[6:7] == ['A'] and id[7:8] == ['1']:
            dict['branch'] += 'Chemical'
        elif id[6:7] == ['A'] and id[7:8] == ['2']:
            dict['branch'] += 'Civil'
        elif id[6:7] == ['A'] and id[7:8] == ['3']:
            dict['branch'] += 'EEE'
        elif id[6:7] == ['A'] and id[7:8] == ['4']:
            dict['branch'] += 'Mechanical'
        elif id[6:7] == ['A'] and id[7:8] == ['5']:
            dict['branch'] += 'B Pharma'
        elif id[6:7] == ['A'] and id[7:8] == ['7']:
            dict['branch'] += 'Computer Science'
        elif id[6:7] == ['A'] and id[7:8] == ['8']:
            dict['branch'] += 'ENI'
        elif id[6:7] == ['B'] and id[7:8] == ['1']:
            dict['branch'] += 'Msc Biology'
        elif id[6:7] == ['B'] and id[7:8] == ['2']:
            dict['branch'] += 'Msc Chemistry'
        elif id[6:7] == ['B'] and id[7:8] == ['3']:
            dict['branch'] += 'Msc Economics'
        elif id[6:7] == ['B'] and id[7:8] == ['4']:
            dict['branch'] += 'Msc Maths'
        elif id[6:7] == ['B'] and id[7:8] == ['5']:
            dict['branch'] += 'Msc Physics'
    str = ''
    for ele in id[8:12]:
        str += ele
    dict['email'] = f"f2022{str}@pilani.bits-pilani.ac.in"
    dictionary.append(dict)
print(dictionary)
