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

def get_discretionary(matrix, path, activity):
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

                    if("110000" in row[0] and activity in row[4]):
                        transactions.append(row)
                except:
                    continue
    write_excel(path, transactions)

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

        ########################## Faculty Discretionary ##################################
        # Asari
        print("Asari")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Asari\\Asari.xlsx'
        get_discretionary(matrix, path, "151500")
 
        # Balster
        print("Balster")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Balster\\Balster.xlsx'
        get_discretionary(matrix, path, "151501")

        # Chodavarapu
        print("Chodavarapu")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Chodavarapu\\Chodavarapu.xlsx'
        get_discretionary(matrix, path, "151512")

        # Doll
        print("Doll")
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
        print("Hardie")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Hardie\\Hardie.xlsx'
        get_discretionary(matrix, path, "151502")

        # Hirakawa
        print("Hirakawa")
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
        print("Ordonez")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Ordonez\\Ordonez.xlsx'
        get_discretionary(matrix, path, "151508")

        # Ratliff
        print("Ratliff")
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
        print("Rigling")
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
        print("Subramanyam")
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
        print("Taha")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Taha\\Taha.xlsx'
        get_discretionary(matrix, path, "151509")
    
        ############################# Staff Discretionary ####################################
        # Yakopcic
        print("Yakopcic")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Yakopcic\\Yakopcic.xlsx'
        get_discretionary(matrix, path, "151511")
    
        # Aspiras
        print("Aspiras")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Aspiras\\Aspiras.xlsx'
        get_discretionary(matrix, path, "151514")
    
        # Shin
        print("Shin")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Shin\\Shin.xlsx'
        get_discretionary(matrix, path, "151515")
    
        # Kumar
        print("Kumar")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Kumar\\Kumar.xlsx'
        get_discretionary(matrix, path, "151516")
    
        # Atahary
        print("Atahary")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Atahary\\Atahary.xlsx'
        get_discretionary(matrix, path, "151517")
    
        # Liu
        print("Liu")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Liu\\Liu.xlsx'
        get_discretionary(matrix, path, "151518")
    
        # Batts
        print("Batts")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Batts\\Batts.xlsx'
        get_discretionary(matrix, path, "151519")
    
        # Nehrbass
        print("Nehrbass")
        path = 'C:\\Users\\ebalster1\\Box\\ebalster1 workspace\\Faculty\\Nehrbass\\Nehrbass.xlsx'
        get_discretionary(matrix, path, "151520")
    
        print("Done")
