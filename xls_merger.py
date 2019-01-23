import pandas as pd
import datetime
import os
#install pandas, openpyxl, xlrd
start = datetime.datetime.now()
print("Processo Iniciado")
print('ok')
def csv_merger():
    
    path = os.getcwd()
    files = os.listdir(path)
    print(files)

    files_xls = [f for f in files if f[-3:] == 'xls']
    print (files_xls)
    # read them in
    excels = [pd.ExcelFile(name) for name in files_xls]

    # turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    # concatenate them..
    combined = pd.concat(frames)

    # write it out 
    combined.to_excel("Envios - " + str(start.strftime("%Y-%m-%d - %H")) + "h.xlsx" , header=False, index=False)

csv_merger()

end = datetime.datetime.now()
tempo = end - start
print("CSV Gerado em " + str(tempo))
