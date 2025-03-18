from tkinter import filedialog as fd
from dateutil.parser import parse
# excel file reader
import pandas as pd
# excel file writer
import xlsxwriter

# get a number from a text string
def text_2_float(text):
    import locale 

    textnocomma = text.replace(',','')
    textnoparen = textnocomma.replace('(','-')
    textnp = textnoparen.replace(')','')
    return locale.atof(textnp)

def write_excel(path, transactions):
    title = ['Fund', 'Organization','Account','Program','Activity','Location','Fund Type',
    'Organization','Fund Type','Organization Level 4','Account Level 2','Transaction Date',
    'Transaction Desc','Document Type Desc','Document','Vendor ID','Budget','Trans_Amount','Encumbered']
    
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet()
    for idx_col, col in enumerate(title):
        sheet.write(0, idx_col, col)
    for idx_row, row in enumerate(transactions):
        values = transactions[idx_row]
        for idx_col, col in enumerate(values):
            sheet.write(idx_row+1, idx_col, col)
        # clean up some emptly columns
        sheet.write(idx_row+1, len(values), '   ')
        sheet.write(idx_row+1, len(values)+1, '   ')
    workbook.close()

# Parsing out financial information
filetypes = (
    ('excel files', '*.xlsx'),
    ('csv files', '*.csv'),
    ('All files', '*.*')
)

getfile = fd.askopenfilename(
    title='Open File',
    initialdir='.',
    filetypes=filetypes)

if(getfile != ''):
    print(getfile)

    with pd.ExcelFile(getfile) as excl:
        sheets = excl.sheet_names
        df = excl.parse(sheets[0])
        matrix = df.to_numpy()

       # Asari
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("110000" in row[0] and "151500" in row[4]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Asari\\Asari.xlsx'
        write_excel(path, transactions)

       # Chodavarapu
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("110000" in row[0] and "151512" in row[4]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Chodavarapu\\Chodavarapu.xlsx'
        write_excel(path, transactions)

       # Doll
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("110036" in row[0]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Doll\\Doll.xlsx'
        write_excel(path, transactions)

        # Hardie
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("110000" in row[0] and "151502" in row[4]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Hardie\\Hardie.xlsx'
        write_excel(path, transactions)

        # Hirakawa
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if(("414005" in row[0]) or ("414055" in row[0]) or ("110000" in row[0] and "151503" in row[4])):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Hirakawa\\Hirakawa.xlsx'
        write_excel(path, transactions)

        # Ordonez
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("110000" in row[0] and "151508" in row[4]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Ordonez\\Ordonez.xlsx'
        write_excel(path, transactions)

        # Ratliff
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("410304" in row[0] or "110035" in row[0]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Ratliff\\Ratliff.xlsx'
        write_excel(path, transactions)

       # Rigling
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("220451" in row[1] or ("110000" in row[0] and "151506" in row[4])):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Rigling\\Rigling.xlsx'
        write_excel(path, transactions)

        # Subramanyam
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if(("115054" in row[0] and "10BU04" in row[5]) or ("110000" in row[0] and "151505" in row[4])):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Subramanyam\\Subramanyam.xlsx'
        write_excel(path, transactions)

        # Taha
        transactions = []
        for listrow in matrix:
            row = list(map(str,listrow))
            if(row[0] != '' and row[0][:1].isdigit()):
                # skip transactions of 0 dollars
                budget = text_2_float(row[14])
                trans = text_2_float(row[15])
                encumbered = text_2_float(row[16])
                if(budget != 0 or trans != 0 or encumbered != 0):
                    try:
                        date = parse(row[9], None)
                        row[9] = date.strftime("%m/%d/%Y")
                        # clean up some of the data
                        if(row[13] == 'nan'):
                            row[13] = ''
                        row[14] = str(round(budget,2))
                        row[15] = str(round(trans, 2))
                        row[16] = str(round(encumbered, 2))

                        if("110000" in row[0] and "151509" in row[4]):
                            transactions.append(row)
                    except:
                        continue

        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Taha\\Taha.xlsx'
        write_excel(path, transactions)
