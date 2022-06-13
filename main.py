import camelot
import pandas as pd
import os
import tkinter as tk


ALL_ROWS_FLAG = False
excel_file = 'output.xlsx'
data_sheet = 'Данные'
chart_sheet = 'График'


def get_files(extensions=['.pdf',]):
    path = os.getcwd()
    dir_list = os.listdir(path)
    files = list()
    for name in dir_list:
        for ext in extensions:
            if ext in name:
                files.append(name)
    return files


def is_there_data(row, all_rows=True):
    """A check that it is not a line from head or bottom of a table and not an empty line"""
    flag = False
    if ':' in row[0] and len(row[0]) < 6:
        flag = True
    return flag


def row_data_extraction(row):
    def look_for_wob(row):
        for line in row:
            if 'WOB' in line:
                return line

    def look_for_spp(row):
        for line in row:
            if 'SPP' in line or 'GPM,' in line:
                return line

    def get_wob_value(line):
        line = line.split(' ')
        line = clear_row(line)
        wob_min = None
        wob_max = None
        wob_name = ''
        wob_i = None
        for name in ['WOB', 'WOB.']:
            try:
                wob_i = line.index(name)
                wob_name = name
            except ValueError:
                pass
        if not wob_name:
            return [-10, -10]

        if wob_name == 'WOB.':
            if line[wob_i-2] == '-':
                wob_min, wob_max = val_to_float(line[wob_i-3]), val_to_float(line[wob_i-1])
            else:
                wob_min = wob_max = val_to_float(line[wob_i-1])
        elif wob_name == 'WOB':
            if line[wob_i+1] == '=':
                if line[wob_i + 3] == '-':
                    wob_min, wob_max = val_to_float(line[wob_i + 2]), val_to_float(line[wob_i + 4])
                else:
                    wob_min = wob_max = val_to_float(line[wob_i + 2])
            else:
                if line[wob_i+2] == '-':
                    wob_min, wob_max = val_to_float(line[wob_i+1]),  val_to_float(line[wob_i+3])
                else:
                    wob_min = wob_max = val_to_float(line[wob_i+1])
        return [wob_min, wob_max]

    def get_spp_value(line):
        line = line.split(' ')
        spp_val = None
        spp_i = None
        if 'SPP' in line:
            spp_i = line.index('SPP')
            spp_val = line[spp_i+1]
        elif 'GPM,' in line:
            spp_i = line.index('GPM,')
            spp_val = line[spp_i+1]
        elif 'GPM' in line:
            spp_i = line.index('GPM')
            spp_val = line[spp_i+2]
        else:
            spp_val = -1000
        spp_value = val_to_float(spp_val)
        return spp_value

    def get_zaboy_value(row):
        for i in range(len(row)-1, 0, -1):
            if not pd.isna(row[i]):
                return float(row[i])

    def clear_row(row):
        return [x for x in row if x != '' and not pd.isna(x)]

    def val_to_float(val):
        _val = val
        for suffix in [',', '.', 'K', '/']:
            if _val[-1] == suffix:
                _val = _val[:-1]
        try:
            float(_val)
        except ValueError:
            _val = 0
        return _val

    row = clear_row(row)
    l = list()
    row_description_part = row[3].split('\n')

    spp_value = None
    wob_value = None
    zaboy_value = get_zaboy_value(row)
    line_with_wob = look_for_wob(row_description_part)
    line_with_spp = look_for_spp(row_description_part)
    if line_with_wob:
        wob_value = get_wob_value(line_with_wob)
    if line_with_spp:
        spp_value = get_spp_value(line_with_spp)
    l.append(spp_value) if spp_value else l.append(0)
    l.extend(wob_value) if wob_value else l.extend([0, 0])
    l.append(zaboy_value)
    return l


def add_charts(writer, data_len, charts=None):
    if charts:
        worksheet = writer.book.add_worksheet(chart_sheet)
        if 'wob:zaboy' in charts:
            chart = writer.book.add_chart({'type': 'line'})
            chart.add_series({
                'categories': [data_sheet, 1, 1, data_len - 1, 1],
                'values': [data_sheet, 1, 3, data_len - 1, 3],
            })
            chart.set_x_axis({'name': 'WOB', 'position_axis': 'on_tick'})
            chart.set_y_axis({'name': 'Забой', 'major_gridlines': {'visible': False}})
            worksheet.insert_chart('I2', chart)
        if 'spp:zaboy' in charts:
            chart = writer.book.add_chart({'type': 'line'})
            chart.add_series({
                'categories': [data_sheet, 1, 0, data_len - 1, 0],
                'values': [data_sheet, 1, 3, data_len - 1, 3],
            })
            chart.set_x_axis({'name': 'SPP', 'position_axis': 'on_tick'})
            chart.set_y_axis({'name': 'Забой', 'major_gridlines': {'visible': False}})
            worksheet.insert_chart('A2', chart)


def collect_data(files):
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
    dfs = []
    for file in files:
        tables = camelot.read_pdf(file)
        df = tables[0].df
        dfs.append(df)
    df = pd.concat(dfs, axis=0)  # concatenate all dataframes in list

    records = df.to_records(index=False)
    result = [row_data_extraction(row) for row in records if is_there_data(row, all_rows=ALL_ROWS_FLAG)]

    if not ALL_ROWS_FLAG:
        result = [line for line in result if line[0] and line[1] and line[2]]

    data_len = len(result)
    df = pd.DataFrame(result, columns=['SPP', 'WOB_min', 'WOB_max', 'Забой'])
    df.to_excel(writer, sheet_name=data_sheet, index=False)
    add_charts(writer, data_len, charts=['wob:zaboy', 'spp:zaboy'])
    writer.save()


def main():
    root = tk.Tk()
    root.geometry("500x380")

    box = tk.Listbox(selectmode=tk.EXTENDED)
    box.pack()
    pdf_files = get_files()
    for name in pdf_files:
        box.insert(tk.END, name)

    f = tk.Frame()
    f.pack(padx=10)

    def select_all_names():
        box.select_set(0, tk.END)

    def start_with_curselection():
        names = []
        for i in box.curselection():
            names.append(box.get(i))
        print(names)
        collect_data(names)

    tk.Button(f, text="Выбрать все", command=select_all_names).pack(fill=tk.X)
    tk.Button(f, text="Старт", command=start_with_curselection).pack(fill=tk.X)

    root.mainloop()


if __name__ == '__main__':
    main()
