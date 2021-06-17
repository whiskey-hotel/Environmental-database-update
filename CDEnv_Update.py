# -------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      jusin
#
# Created:     07/06/2021
# Copyright:   (c) jusin 2021
# Licence:     <your licence>
# -------------------------------------------------------------------------------

from openpyxl import workbook, load_workbook


def main():
    pass


if __name__ == '__main__':
    main()


# open and define workbooks
print('...Loading workbooks...')
project_wkb = load_workbook(
    "G:\\Shared drives\\Database\\Tables for Merge\\Working\\CD Projects.xlsx")
client_wkb = load_workbook(
    "G:\\Shared drives\\Database\\Tables for Merge\\Working\\Clients.xlsx")
np_wkb = load_workbook(
    "G:\\Shared drives\\Database\\Tables for Merge\\Working\\Newspapers_6.18.2020.xlsx")
env_wkb = load_workbook(
    "G:\\Shared drives\\Database\\Tables for Merge\\Working\\CD Env Review_5.12.2020_NEP DO NOT OVERWRITE.xlsx")

# define worksheets
print('Loading worksheets')
project_wks = project_wkb["CD Projects"]
client_wks = client_wkb["Clients"]
np_wks = np_wkb["Newspapers"]
env_wks = env_wkb["CD Env Review"]


def read_in_values(wks):
    """Stores worksheet values into a nested list"""
    wks_array = [[value for value in row]for row in wks.values]
#   wks_array[row][column]
    return wks_array


def flaten_list(wks_array, item_index):
    """Creates a new flat list from a specific index (item_index) in a nested list
     (wks_array)"""
    return [item[item_index] for item in wks_array]


def find_index(list1, search):

    return list1.index(search)


def check(env_row_check, wks_array_project, client_value):

    # temp_list_env = [item for item in wks_array_env if item[0] == client_value]
    temp_list_project = [
        item for item in wks_array_project if item[0] == client_value]
    if bool(temp_list_project) is False:
        return
    for count, i in enumerate(temp_list_project):
        if env_row_check[15] == temp_list_project[count][6]:  # Grant Program
            if env_row_check[19] == temp_list_project[count][10]:  # Grant Amount
                if env_row_check[20] == temp_list_project[count][11]:  # Grant Match
                    if env_row_check[28] == temp_list_project[count][12]:  # Total Amount
                        if env_row_check[33] == temp_list_project[count][2]:  # Year
                            # contract number
                            return temp_list_project[count][1]
        # make a list of the clients and corresponding columns
    # if the checks return true, input the contract value


def from_projects_wks(starting_row, client_array, client_search, news_array, news_search):
    new_row = []

    if starting_row[0] in client_search:
        row_index = find_index(client_search, starting_row[0])
        client_row = client_array[row_index]

        new_row.extend(starting_row[0:2])  # Client and Contract
        from_client_wks(new_row, client_row)
        new_row.extend(starting_row[6:12])
        # blank spaces for posting info
        new_row.extend([" ", " ", " ", " ", " "])
        from_client_wks(new_row, client_row)
        new_row.append(starting_row[12])
        new_row.append(starting_row[9])
        from_client_wks(new_row, client_row)
        new_row.extend([" ", " "])  # blank spaces for posting info
        new_row.append(starting_row[2])
        from_client_wks(new_row, client_row)
        from_np_wks(new_row, news_array, news_search)
        return new_row

    new_row.extend(starting_row[0:2])  # Client and Contract
    new_row.extend([" ", " ", " ", " ", " ", " ", " ",
                   " ", " ", " ", " ", " ", " ", " "])  # 12 blank spaces
    new_row.extend(starting_row[6:12])
    # 5 blank spaces for posting info
    new_row.extend([" ", " ", " ", " ", " "])
    new_row.extend([" ", " "])  # 2 blank spaces
    new_row.append(starting_row[12])
    new_row.append(starting_row[9])
    new_row.append('')
    new_row.extend([" ", " "])  # blank spaces for posting info
    new_row.append(starting_row[2])
    new_row.extend([" ", " ", " ", " "])
    from_np_wks(new_row, news_array, news_search)
    return new_row


def from_client_wks(new_row, client_row):

    if len(new_row) == 26:
        new_row.append(client_row[93])
        new_row.append(client_row[15])
        return new_row
    elif len(new_row) == 30:
        new_row.append(client_row[30])
        return new_row
    elif len(new_row) == 34:
        new_row.extend(client_row[82:84])
        new_row.append(client_row[99])
        new_row.append(client_row[84])
        return new_row

    new_row.extend(client_row[3:8])
    new_row.append(client_row[24])
    new_row.extend(client_row[19:23])
    new_row.extend(client_row[27:30])

    return new_row


def from_np_wks(new_row, np_array, np_line):

    if new_row[37] not in np_line:
        return new_row.extend([" ", " "])
    row_index = find_index(np_line, new_row[37])
    np_row = np_array[row_index]

    new_row.extend(np_row[3:5])
    return new_row


print('...defining intial arrays...')
env_array = read_in_values(env_wks)
project_array = read_in_values(project_wks)
client_array = read_in_values(client_wks)
news_array = read_in_values(np_wks)

print('...flattening arrays...')
env_contract_list = flaten_list(env_array, 1)
project_contract_list = flaten_list(project_array, 1)
client_by_client = flaten_list(client_array, 0)
news_by_news = flaten_list(news_array, 0)

# define trackers
total_contract_updates = 0
total_projects_added = 0

print('...Checking for existing projects with missing contract numbers...')
for index, env_rows in enumerate(env_wks.iter_rows(min_row=2, values_only=True)):
    if bool(env_wks.cell(row=index+2, column=1).value) is True and bool(env_wks.cell(row=index+2, column=2).value) is False:
        for row in env_wks.iter_rows(min_row=index+2, max_row=index+2, values_only=True):
            env_row = row
        env_wks.cell(row=index+2, column=2).value = check(env_row,
                                                          project_array, env_wks.cell(row=index+2, column=1).value)
        env_contract_list.append(env_wks.cell(row=index+2, column=2).value)
        total_contract_updates += 1
        #print(env_wks.cell(row=index+2, column=2).value)
# add a check for empty client info here

print('...Checking for missing projects...')
for i in project_contract_list:
    # add acheck here for empty contracts in projects workbook
    if i not in env_contract_list:
        env_contract_list.append(i)
        row_index = find_index(project_contract_list, i)
        # env_wks.append([i])
        # print(i)
        env_wks.append(
            from_projects_wks(project_array[row_index], client_array, client_by_client, news_array, news_by_news))
        total_projects_added += 1

print('...Saving Environmental Workbook...')
env_wkb.save("CD Env Review_5.12.2020_NEP DO NOT OVERWRITE.xlsx")
print(f'{total_contract_updates} Contract numbers were added to existing projects.\n{total_projects_added} Projects were added to the environmental database.')
print('Update process complete.')
