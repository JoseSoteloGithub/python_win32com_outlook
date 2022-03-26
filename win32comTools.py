import sys


def handle_attribute_error_CLSIDToClassMap(attribute_error_str):
    # Use this on Attribute Error "has no attribute 'CLSIDToClassMap'"
    # This error happens when on this call win32com.client.Dispatch()
    from shutil import rmtree

    first_section_find_str = 'win32com.gen_py.'
    first_section_index = attribute_error_str.find(first_section_find_str)
    first_section_index += len(first_section_find_str)
    second_section_index = attribute_error_str.find('\'', first_section_index)
    folder_name = attribute_error_str[first_section_index:second_section_index]
    rmtree(f"{win32com.__gen_path__}\{folder_name}")
    sys.exit(f'AttributeError detected and path {win32com.__gen_path__}\{folder_name} has been removed.  Restart the program')

def get_range_whole(findString, currentSheet, withinRange=None, afterRange=None):

    if len(findString) <= 255:

        if (withinRange is None) and (afterRange is None):
            getPosition = currentSheet.Cells.Find(What=findString, After=currentSheet.Range("A1"), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
    
        if (not withinRange is None) and (afterRange is None):
            getPosition = currentSheet.Range(withinRange.Address).Find(What=findString, After=currentSheet.Range("A" + withinRange.Row), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
    
        elif (not withinRange is None) and (not afterRange is None):
            getPosition = currentSheet.Range(withinRange.Address).Find(What=findString, After=currentSheet.Range(afterRange.Address), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
        
        elif (withinRange is None) and (not afterRange is None):
            getPosition = currentSheet.Cells.Find(What=findString, After=currentSheet.Range(afterRange.Address), LookAt=constants.xlWhole, SearchDirection=constants.xlNext)
    
    get_range_whole = getPosition
    
    return get_range_whole

def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start

def get_column_letter(worksheet_obj, substring_str, header_row_int, worksheet_column_letters_dict={}):
    # Returns column letter or -1 if not found
    try:
        column_letter = worksheet_column_letters_dict[substring_str]
        print('Returned from try: on worksheet_column_letters_dict')
        return column_letter
    except:
        if header_row_int == 0:
            column_header = worksheet_obj.Cells.Find(What=substring_str, SearchOrder=constants.xlByRows, 
                                                            SearchDirection=constants.xlNext)        
        else:
            column_header = worksheet_obj.Rows(header_row_int).Find(What=substring_str, SearchOrder=constants.xlByRows,
                                                            SearchDirection=constants.xlNext)

        if not column_header is None:
            first_dolla = find_nth(column_header.Address, '$', 1)
            second_dolla = find_nth(column_header.Address, '$', 2)
            column_letter = column_header.Address[first_dolla + 1:second_dolla]
        else:
            column_letter = -1

        return column_letter

def get_column_number(worksheet_obj, substring_str, header_row_int, worksheet_column_numbers_dict={}):
    # Returns column number or -1 if not found
    try:
        column_number = worksheet_column_numbers_dict[substring_str]
        print('Returned from try: on worksheet_column_numbers_dict')
        return column_number
    except:
        if header_row_int == 0:
            column_header = worksheet_obj.Cells.Find(What=substring_str, SearchOrder=constants.xlByRows,
                                                            SearchDirection=constants.xlNext)
        else:
            column_header = worksheet_obj.Rows(header_row_int).Find(What=substring_str, SearchOrder=constants.xlByRows,
                                                            SearchDirection=constants.xlNext)

        if not column_header is None:
            column_number = column_header.Column
        else:
            column_number = -1

        return column_number    
    
def get_last_column_letter(worksheet_obj, header_row_int=0):
    # Returns last column letter
    if header_row_int != 0:
        last_column_letter_address = worksheet_obj.Rows(header_row_int).Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    else:
        last_column_letter_address = worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious).Address
    last_column_letter = last_column_letter_address.split('$')[1]
    return(last_column_letter)

def get_last_column_index(worksheet_obj, header_row_int=0):
    # Returns last column as index
    if header_row_int != 0:
        last_column_range = worksheet_obj.Rows(header_row_int).Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious)
    else:
        last_column_range = worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByColumns, SearchDirection=constants.xlPrevious)

    if last_column_range == None:
        return 0
    else:
        return(last_column_range.Column)

def get_header_row(worksheet_obj, search_string_str):
    # returns header row as int, -1 if not found
    header_range = worksheet_obj.Rows.Find(What=search_string_str, SearchOrder=constants.xlByRows, 
                                                        SearchDirection=constants.xlNext)
    if not header_range is None:
        header_row = header_range.Row
    else:
        header_row = -1
    header_row = str(header_row)
    return header_row

def get_last_row(worksheet_obj):
    # Returns last row as int
    return(worksheet_obj.Cells.Find(What='*', SearchOrder=constants.xlByRows, SearchDirection=constants.xlPrevious).Row)

def sheet_exist(workbook, string):
    # Returns True if exists
    for ws_target in workbook.Worksheets:
        if string in ws_target.Name:
            return True
    return False

def create_sheets(workbook_obj, worksheet_obj, name_li):
    # Loop through list of name, check if there's a worksheet for the name,
    # if not then copy the Active worksheet and paste it at the end, rename it, apply filter for that name
    for name in name_li:
        name_sheet_exist = sheet_exist(workbook_obj, name)
        if name_sheet_exist == False:
            worksheet_obj.Copy(After=workbook_obj.Sheets(len(workbook_obj.Worksheets)))
            workbook_obj.Worksheets(len(workbook_obj.Worksheets)).Name = name

def get_dictionary_column_letters(worksheet_obj, column_letters_dict, header_row_int):
    # Return dictionary of column letters dictionary_name['ColumnHeaderString'] = 'ColumnLetter'

    last_column_index = get_last_column_index(worksheet_obj, header_row_int)

    for i in range(1, last_column_index + 1):
        value = worksheet_obj.Range(worksheet_obj.Cells(header_row_int, i), worksheet_obj.Cells(header_row_int, i)).Value
        column_letters_dict[value] = get_column_letter(worksheet_obj, value, header_row_int, column_letters_dict)

    return column_letters_dict

import win32com.client
constants = win32com.client.constants
