from os import path
import pandas as pd

def remove_unnamed_columns(df):
    return df.loc[:, ~df.columns.str.contains('^Unnamed')]

def test_col_names_match(old_data, new_data):
    old_column_names = old_data.columns.values.tolist()
    new_column_names = new_data.columns.values.tolist()

    stripped_old_column_names = [col.strip().lower() for col in old_column_names]
    stripped_new_column_names = [col.strip().lower() for col in new_column_names]
    
    if stripped_old_column_names == stripped_new_column_names:
        return True
    else:
        print("\nError: column names do not match:\n")
        print("OLD:", old_column_names)
        print("NEW", new_column_names)
        return False

def test_date_exists(row, new_data, lang):
    if lang == 'he':
        results = [new_row for new_row in new_data if new_row[0].strip() == row[0].strip() and new_row[1] == row[1]]
    else:
        results = [new_row for new_row in new_data if new_row[2].split(' ')[0] == row[2].split(' ')[0] and new_row[2].split(' ')[1].replace('-', '') == row[2].split(' ')[1].replace('-', '')]
        
    if len(results):
        return True
    else:
        return False

def np_to_df(new_sheet, new_data):
    data = remove_unnamed_columns(pd.DataFrame(new_sheet))
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
                if row[2].replace('-','').strip().split('Line')[0].strip() in line:
                    temp.append(idx)
            if len(temp):
                lines_available[temp[0]][1] +=1
            else:
                lines_available.append([row[2].replace('-','').strip().split('Line')[0].strip(), 1])
            temp = []
    return lines_available

def excel(old_file, new_file, lang, output_directory):
    extras = {'dates do not exist': [], 'line not available': [], 'exceeds character limit': []}

    old_file = pd.ExcelFile(old_file)
    new_file = pd.ExcelFile(new_file)

    old_sheet = pd.read_excel(old_file, len(old_file.sheet_names) - 1)
    new_sheet = pd.read_excel(new_file, len(new_file.sheet_names) - 1)

    old_data = remove_unnamed_columns(pd.DataFrame(old_sheet))
    new_data = remove_unnamed_columns(pd.DataFrame(new_sheet))

    if not test_col_names_match(old_data, new_data):
        return "Error: Column names do not match!"

#      0         1                            3                          4                      5
# ['Hebrew', ' Date', 'Unnamed: 2',    'Date and Line ',              'Text',          'Character Count']
# ['Elul',      14,       nan,       'September 3 - Line 2 (35)',   'Happy Birthday',          14]

    old_data = old_data.to_numpy()
    new_data = new_data.to_numpy()

    for row in old_data:
        if type(row[3]) == str:
            if not test_date_exists(row, new_data, lang):
                extras['dates do not exist'].append(row)
            
            elif lang == 'he':
                matching_line = [idx for (idx, new_row) in enumerate(new_data) if new_row[0].strip() == row[0].strip() and new_row[1] == row[1] and row[2].split('Line')[1].split('(')[0].strip() == new_row[2].split('Line')[1].split('(')[0].strip()]
            else:
                matching_line = [idx for (idx, new_row) in enumerate(new_data) if new_row[2].strip().replace('-', '') == row[2].strip().replace('-', '')]
            
            if not len(matching_line):
                extras['line not available'].append(row)
            else:
                for line in matching_line:
                    allowed_character_count = int(new_data[line][2].rsplit("(")[1].rsplit(")")[0])
                    if len(row[3]) > allowed_character_count:
                        extras['exceeds character limit'].append(row)
                    else:
                        new_data[line][3] = row[3]
                        new_data[line][4] = len(row[3])

    new_df = np_to_df(new_sheet=new_sheet, new_data=new_data)

    with pd.ExcelWriter(path.join(output_directory, 'output.xlsx'), engine='xlsxwriter') as writer: 
        new_df.to_excel(writer, sheet_name=new_file.sheet_names[len(new_file.sheet_names) - 1], index=False)

    empty_lines = [row for row in new_data if type(row[3]) != str]
    lines_available = available_lines(empty_lines=empty_lines, lang=lang)
    extra1_df = pd.DataFrame(extras['dates do not exist'], columns=remove_unnamed_columns(pd.DataFrame(new_sheet)).columns.values.tolist())
    extra2_df = pd.DataFrame(extras['line not available'], columns=remove_unnamed_columns(pd.DataFrame(new_sheet)).columns.values.tolist())
    extra3_df = pd.DataFrame(extras['exceeds character limit'], columns=remove_unnamed_columns(pd.DataFrame(new_sheet)).columns.values.tolist())
    extra4_df = pd.DataFrame(lines_available, columns=["Date", "Lines Available"])
    
    with pd.ExcelWriter(path.join(output_directory, 'extras.xlsx'), engine='xlsxwriter') as writer: 
        extra1_df.to_excel(writer, sheet_name='date not found', index=False)
        extra2_df.to_excel(writer, sheet_name='original line taken', index=False)
        extra3_df.to_excel(writer, sheet_name='exceeds character limit', index=False)
        extra4_df.to_excel(writer, sheet_name='unused lines (available)', index=False)

    return True