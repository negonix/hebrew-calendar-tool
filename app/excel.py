from os import path
import pandas as pd

def test_col_names_match(old_sheet, new_sheet):
    old_tab = pd.read_excel(old_sheet, 1)
    new_tab = pd.read_excel(new_sheet, 1)

    old_data = pd.DataFrame(old_tab)
    new_data = pd.DataFrame(new_tab)

    old_column_names = old_data.columns.values.tolist()
    new_column_names = new_data.columns.values.tolist()

    old_column_names = [col.strip() for col in old_column_names if col.startswith('Unnamed') == False]
    new_column_names = [col.strip() for col in new_column_names if col.startswith('Unnamed') == False]
    
    if old_column_names == new_column_names:
        return True
    else:
        return False

def test_date_exists(row, new_data, lang):
    if lang == 'he':
        results = [new_row for new_row in new_data if new_row[0].strip() == row[0].strip() and new_row[1] == row[1]]
    else:
        results = [new_row for new_row in new_data if new_row[3].split(' ')[0] == row[3].split(' ')[0] and new_row[3].split(' ')[1].replace('-', '') == row[3].split(' ')[1].replace('-', '')]
        
    if len(results):
        return True
    else:
        return False

def np_to_df(new_sheet, new_data):
    tab = pd.read_excel(new_sheet, 1)
    data = pd.DataFrame(tab)
    column_names = data.columns.values.tolist()
    df = pd.DataFrame(new_data, columns = column_names)
    return df

def available_lines(empty_lines, lang):
    lines_available = []
    for row in empty_lines:
        if lang == 'he':
            temp = []
            for idx, line in enumerate(lines_available):
                if f"{row[0]} {row[1]}" in line:
                    temp.append(idx)
            if len(temp):
                lines_available[temp[0]][1] +=1
            else:
                lines_available.append([f"{row[0]} {row[1]}", 1])
            temp = []
        else:
            temp = []
            for idx, line in enumerate(lines_available):
                if row[3].replace('-','').strip().split('Line')[0].strip() in line:
                    temp.append(idx)
            if len(temp):
                lines_available[temp[0]][1] +=1
            else:
                lines_available.append([row[3].replace('-','').strip().split('Line')[0].strip(), 1])
            temp = []
    return lines_available

def excel(old_sheet, new_sheet, lang, output_directory):
    old_sheet = pd.ExcelFile(old_sheet)
    new_sheet = pd.ExcelFile(new_sheet)

    extras = {'dates do not exist': [], 'line not available': []}

    old_tab = pd.read_excel(old_sheet, 1)
    new_tab = pd.read_excel(new_sheet, 1)

    old_data = pd.DataFrame(old_tab).to_numpy()
    new_data = pd.DataFrame(new_tab).to_numpy()

    if not test_col_names_match(old_sheet, new_sheet):
        return "Error: Column names do not match!"

#      0         1                            3                          4                      5
# ['Hebrew', ' Date', 'Unnamed: 2',    'Date and Line ',              'Text',          'Character Count']
# ['Elul',      14,       nan,       'September 3 - Line 2 (35)',   'Happy Birthday',          14]

    for row in old_data:
        if type(row[4]) == str:
            if not test_date_exists(row, new_data, lang):
                extras['dates do not exist'].append(row)
                continue
            
            if lang == 'he':
                matching_line = [idx for (idx, new_row) in enumerate(new_data) if new_row[0].strip() == row[0].strip() and new_row[1] == row[1] and row[3].split('Line')[1].split('(')[0].strip() == new_row[3].split('Line')[1].split('(')[0].strip()]
                if not len(matching_line):
                    extras['line not available'].append(row)
                else:
                    for line in matching_line:
                        new_data[line][4] = row[4]

            else:
                matching_line = [idx for (idx, new_row) in enumerate(new_data) if new_row[3].strip().replace('-', '') == row[3].strip().replace('-', '')]
                if not len(matching_line):
                    extras['line not available'].append(row)
                else:
                    for line in matching_line:
                        new_data[line][4] = row[4]

    new_df = np_to_df(new_sheet=new_sheet, new_data=new_data)

    with pd.ExcelWriter(path.join(output_directory, 'output.xlsx'), engine='xlsxwriter') as writer: 
        new_df.to_excel(writer, sheet_name=new_sheet.sheet_names[1], index=False)

    empty_lines = [row for row in new_data if type(row[4]) != str]
    lines_available = available_lines(empty_lines=empty_lines, lang=lang)
    extra1_df = pd.DataFrame(extras['dates do not exist'], columns=pd.DataFrame(new_tab).columns.values.tolist())
    extra2_df = pd.DataFrame(extras['line not available'], columns=pd.DataFrame(new_tab).columns.values.tolist())
    extra3_df = pd.DataFrame(lines_available, columns=["Date", "Lines Available"])
    
    with pd.ExcelWriter(path.join(output_directory, 'extras.xlsx'), engine='xlsxwriter') as writer: 
        extra1_df.to_excel(writer, sheet_name='date not found', index=False)
        extra2_df.to_excel(writer, sheet_name='original line taken', index=False)
        extra3_df.to_excel(writer, sheet_name='unused lines (available)', index=False)

    return True