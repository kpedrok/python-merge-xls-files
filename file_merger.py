import pandas as pd
import datetime
import os
# install pandas, openpyxl, xlrd


def xlsx_merger():
    path = os.getcwd()
    files = os.listdir(path)
    files_xls = [f for f in files if f[-4:] == 'xlsx']
    print(files_xls)
    # read them in
    excels = [pd.ExcelFile(name) for name in files_xls]

    # turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None)
              for x in excels]

    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    # concatenate them..
    combined = pd.concat(frames)

    # write it out
    combined.to_excel("Envios - " + str(start.strftime("%Y-%m-%d - %H")
                                        ) + "h.xlsx", header=False, index=False)


def xls_merger():
    path = os.getcwd()
    files = os.listdir(path)

    files_xls = [f for f in files if f[-3:] == 'xls']
    print(files_xls)
    # read them in
    excels = [pd.ExcelFile(name) for name in files_xls]

    # turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None)
              for x in excels]

    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    # concatenate them..
    combined = pd.concat(frames)

    # write it out
    combined.to_excel("Envios - " + str(start.strftime("%Y-%m-%d - %H")
                                        ) + "h.xlsx", header=False, index=False)


def csv_merger():
    path = os.getcwd()
    files = os.listdir(path)
    filenames = [f for f in files if f[-3:] == 'csv']
    print(filenames)
    combined_csv = pd.concat([pd.read_csv(f) for f in filenames])
    combined_csv.to_csv(
        "Raskon - " + str(start.strftime("%Y-%m-%d - %H")) + "h.csv", header=True, index=False)


try:
    print("Processo Iniciado")
    file_type = input("Enter file type (csv, xls, or xlsx):")
    start = datetime.datetime.now()

    if file_type == "csv":
        csv_merger()
    elif file_type == "xls":
        xls_merger()
    elif file_type == "xlsx":
        xlsx_merger()
    else:
        print("File type is incompatible")

    # Calcular tempo de execução
    end = datetime.datetime.now()
    tempo = end - start
    print("\nCSV Gerado em " + str(tempo))
except Exception as e:
    print(e)


input("\nPress enter to exit")
