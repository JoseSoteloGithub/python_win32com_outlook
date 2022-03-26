# This loops through items in Outlook and saves them to To Do workbook in worksheet Outlook
import win32com.client
import os
import win32comTools


def find_nth(haystack, needle, n):
    # Returns index of start of substring found in string by nth occurrence
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start + len(needle))
        n -= 1
    return start


def send_info_to_row(send_to_row):
    outlook_ws.Range(outlook_ws.Cells(send_to_row, outlook_column_index_dict['Source']),
                     outlook_ws.Cells(send_to_row,
                                      outlook_column_index_dict[
                                          'Source'])).Value = target_inbox_folder_source_str

    outlook_ws.Range(outlook_ws.Cells(send_to_row, outlook_column_index_dict['ReceivedTime']),
                     outlook_ws.Cells(send_to_row,
                                      outlook_column_index_dict[
                                          'ReceivedTime'])).Value = message.ReceivedTime
    outlook_ws.Range(outlook_ws.Cells(send_to_row, outlook_column_index_dict['Subject']),
                     outlook_ws.Cells(send_to_row,
                                      outlook_column_index_dict['Subject'])).Value = message.Subject
    outlook_ws.Range(outlook_ws.Cells(send_to_row, outlook_column_index_dict['EntryID']),
                     outlook_ws.Cells(send_to_row,
                                      outlook_column_index_dict['EntryID'])).Value = message.EntryID
    outlook_ws.Range(outlook_ws.Cells(send_to_row, outlook_column_index_dict['ConversationID']),
                     outlook_ws.Cells(send_to_row,
                                      outlook_column_index_dict[
                                          'ConversationID'])).Value = message.ConversationID
    outlook_ws.Range(outlook_ws.Cells(send_to_row, outlook_column_index_dict['To Do']),
                     outlook_ws.Cells(send_to_row,
                                      outlook_column_index_dict['To Do'])).Value = outlook_to_do

constants = win32com.client.constants

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
except AttributeError as attribute_error:
    if str(attribute_error).find("has no attribute 'CLSIDToClassMap'") > -1:
        win32comTools.handle_attribute_error_CLSIDToClassMap(str(attribute_error))

mapi = outlook.GetNamespace("MAPI")
your_folder = mapi.Folders

try:
    xl = win32com.client.Dispatch('Excel.Application')
except AttributeError as attribute_error:
    if str(attribute_error).find("has no attribute 'CLSIDToClassMap'") > -1:
        win32comTools.handle_attribute_error_CLSIDToClassMap(str(attribute_error))
xl.Visible = True

to_do_filename = 'To Do.xlsx'

if len(xl.Workbooks) == 0:
    to_do_wb = xl.Workbooks.Add()
    to_do_wb.SaveAs(Filename='To Do.xlsx')

for to_do_wb in xl.Workbooks:
    if to_do_wb.Name == to_do_filename:
        break

if to_do_wb.Name != to_do_filename:
    if os.path.exists(to_do_filename):
        to_do_wb = xl.Workbooks.Open(to_do_filename)
    else:
        to_do_wb = xl.Workbooks.Add()
        to_do_wb.SaveAs(Filename='To Do.xlsx')

if not win32comTools.sheet_exist(to_do_wb, 'Outlook'):
    outlook_ws = to_do_wb.Worksheets.Add()
    outlook_ws.Name = 'Outlook'
else:
    outlook_ws = to_do_wb.Worksheets('Outlook')

#outlook_ws.Cells.Clear()

outlook_header_row = 1

# This section looks for column headers and creates them if they're missing START
outlook_column_li = ['Source',
                     'ReceivedTime',
                     'Subject',
                     'EntryID',
                     'ConversationID',
                     'To Do']

outlook_column_index_dict = {}

for outlook_column in outlook_column_li:
    outlook_column_index_dict[outlook_column] = win32comTools.get_column_number(outlook_ws, outlook_column,
                                                                                outlook_header_row,
                                                                                outlook_column_index_dict)
    if outlook_column_index_dict[outlook_column] == -1:
        last_column_index = win32comTools.get_last_column_index(outlook_ws, outlook_header_row)
        outlook_ws.Range(outlook_ws.Cells(outlook_header_row, last_column_index + 1),
                         outlook_ws.Cells(outlook_header_row, last_column_index + 1)).Value = outlook_column
        outlook_column_index_dict[outlook_column] = last_column_index + 1

outlook_column_letter_dict = {}

for outlook_column in outlook_column_li:
    outlook_column_letter_dict[outlook_column] = win32comTools.get_column_letter(outlook_ws, outlook_column,
                                                                                 outlook_header_row,
                                                                                 outlook_column_letter_dict)

# This section looks for column headers and creates them if they're missing END

# This section gives option of which Outlook account to iterate through START
for i, folder in enumerate(mapi.Folders):
    print(i, folder.Name)

folder_select_input: int = int(input('Select folder to iterate and export: '))

if mapi.Folders[folder_select_input].Name == mapi.Folders[int(folder_select_input)].Name:
    target_folder = mapi.Folders[int(folder_select_input)]
else:
    target_folder = mapi.Folders[int(folder_select_input) + 1]

print(target_folder.Name)
# This section gives option of which Outlook account to iterate through END

target_inbox_folder = target_folder.Folders['Inbox']

target_inbox_folder_source = target_inbox_folder

target_inbox_folder_source_str = target_inbox_folder_source.Name

while str(target_inbox_folder_source) != 'Mapi':
    target_inbox_folder_source = target_inbox_folder_source.Parent
    if str(target_inbox_folder_source) != 'Mapi':
        target_inbox_folder_source_str = f'{str(target_inbox_folder_source)}/{target_inbox_folder_source_str}'

outlook_last_row = win32comTools.get_last_row(outlook_ws)

for i, message in enumerate(target_inbox_folder.Items):
    if message.MessageClass == 'IPM.Note':

        # To Do handling START
        outlook_to_do_count = message.Body.count('TO DO: ')

        outlook_to_do_list = []
        for i_outlook in range(outlook_to_do_count):
            outlook_to_do_start_location = find_nth(message.Body, 'TO DO: ', i_outlook + 1)
            outlook_to_do_end_location = message.Body.find('\r\n', outlook_to_do_start_location)
            outlook_to_do_slice = message.Body[outlook_to_do_start_location + len('TO DO: '):outlook_to_do_end_location]
            outlook_to_do_list.append(outlook_to_do_slice)
        # To Do handling END

        if outlook_to_do_count == 0:
            outlook_to_do_list.append('')

        for outlook_to_do in outlook_to_do_list:

            if outlook_to_do == '':
                # If outlook to do is blank
                entry_id_to_do_result = xl.Evaluate(
                    f'MATCH("{message.EntryID}"&"{outlook_to_do}",\'[{to_do_wb.Name}]{outlook_ws.Name}\'!${outlook_column_letter_dict["EntryID"]}${outlook_header_row}:${outlook_column_letter_dict["EntryID"]}${outlook_last_row + 1}&\'[{to_do_wb.Name}]{outlook_ws.Name}\'!${outlook_column_letter_dict["To Do"]}${outlook_header_row}:${outlook_column_letter_dict["To Do"]}${outlook_last_row + 1},0)')
                if entry_id_to_do_result < 0:
                    send_info_to_row(outlook_last_row + 1)

                    outlook_last_row += 1

            elif outlook_to_do != '':

                # If to do is not blank, check if entry ID and to do.  If they exists, do nothing
                entry_id_to_do_result = xl.Evaluate(
                    f'MATCH("{message.EntryID}"&"{outlook_to_do}",\'[{to_do_wb.Name}]{outlook_ws.Name}\'!${outlook_column_letter_dict["EntryID"]}${outlook_header_row}:${outlook_column_letter_dict["EntryID"]}${outlook_last_row + 1}&\'[{to_do_wb.Name}]{outlook_ws.Name}\'!${outlook_column_letter_dict["To Do"]}${outlook_header_row}:${outlook_column_letter_dict["To Do"]}${outlook_last_row + 1},0)')

                if entry_id_to_do_result < 0:

                    #
                    entry_id_result = xl.Evaluate(
                        f'MATCH("{message.EntryID}",\'[{to_do_wb.Name}]{outlook_ws.Name}\'!${outlook_column_letter_dict["EntryID"]}${outlook_header_row}:${outlook_column_letter_dict["EntryID"]}${outlook_last_row + 1},0)')

                    if entry_id_result < 0:
                        send_info_to_row(outlook_last_row + 1)

                        outlook_last_row += 1

                    else:
                        outlook_ws_to_do_value = outlook_ws.Range(
                            outlook_ws.Cells(int(entry_id_result), outlook_column_index_dict['To Do']),
                            outlook_ws.Cells(int(entry_id_result), outlook_column_index_dict['To Do'])).Value

                        if outlook_to_do == outlook_ws_to_do_value:
                            pass

                        elif outlook_ws_to_do_value == None:
                            send_info_to_row(int(entry_id_result))

                        else:
                            send_info_to_row(outlook_last_row + 1)

                            outlook_last_row += 1

            # Select row on every 50 new entries to Excel
            if i % 50 == 0:
                outlook_ws.Activate()
                outlook_ws.Range(outlook_ws.Cells((outlook_last_row + 1), outlook_column_index_dict['To Do']),
                                 outlook_ws.Cells((outlook_last_row + 1),
                                                  outlook_column_index_dict['To Do'])).Select()

outlook_ws.Columns(outlook_column_index_dict['ReceivedTime']).NumberFormat = r"[$-en-US]mm/dd/yyyy h:mm:ss AM/PM;@"

outlook_last_column_letter = win32comTools.get_last_column_letter(outlook_ws, outlook_header_row)

outlook_ws.Range(f"$A${outlook_header_row}:${outlook_last_column_letter}${outlook_last_row}").RemoveDuplicates(
    Columns=(outlook_column_index_dict['EntryID'], outlook_column_index_dict['To Do']), Header=constants.xlYes)

outlook_ws.Range(f"$A${outlook_header_row}:${outlook_last_column_letter}${outlook_last_row}").Sort(
    Key1=outlook_ws.Range(f"${outlook_column_letter_dict['ReceivedTime']}${outlook_header_row}"), Order1=constants.xlDescending,
    Orientation=constants.xlSortColumns)

outlook_ws.Range(outlook_ws.Cells(1, 1), outlook_ws.Cells(1, 1)).Select()