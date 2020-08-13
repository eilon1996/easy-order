echo installing easy order...

cd \
cd users
subst x: %USERPROFILE%
x:
cd desktop
mkdir easy order
cd easy order

https://aka.ms/nugetclidl
nuget.exe install python -Version 3.7.7


python -m pip install xlrd
python -m pip install xlsxwriter
python -m pip install pyinstaller

echo 
'import os
import xlrd
import xlsxwriter
import glob

def get_info(file_name):
    """
    :param file_name: file_name: the path to the fie where the user keep the xlsx files
    :type file_name: str
    :return: a list of branches contain a list of dictionaries, one for each supplier, which contain ('branch_name', list of string - his order)
    :rtype: list
    """
    files = os.listdir(file_name)
    files_xls = [f for f in files if f[-4:] == 'xlsx' and f != "final_order.xlsx"]
    branches = [None]*(len(files_xls)+1) # the first place will store the items names
    for b, f in enumerate(files_xls):

        workbook = xlrd.open_workbook(file_name+'\\'+f)
        # ignore certain suppliers
        all_sheets = [workbook.sheet_by_index(i) for i in range(workbook.nsheets)\
             if not workbook.sheet_by_index(i).visibility and not "טמפו" in workbook.sheet_by_index(i).name and\
              not "החברה המרכזית" in workbook.sheet_by_index(i).name and not "מחסן" in workbook.sheet_by_index(i).name]

        info = {}
        items_names = {}
        for index, sheet in enumerate(all_sheets):
            col = [sheet.cell_value(row, 1) for row in range(sheet.nrows)]
            empty = True  # if the sheet was empty don't enter it to the info
            for row in col[2:]:
                if row != "":
                    empty = False
                    break
            if not empty:
                info[sheet.name] = col
            if b == 0:
                # in the first loop also enter the products names
                items_names[sheet.name] = [sheet.cell_value(row, 0) for row in range(sheet.nrows)]
        if b == 0:
            branches[0] = items_names
        branches[b+1] = info
    return branches


def create_error_massage(error, massage=""):
    """
    :param error: the error that raise
    :type error: ErrorType
    :param massage: the massage relevant to the user
    :type massage: str
    :return: None
    """
    file = open(file_name+"\\error.txt", "w")
    file.write(str(error)+"\n"+massage+"\nif it's not the case report to Eilon Toledano")
    print(str(error)+"\n"+massage)
    file.close()
    raise SystemExit(0)


def create_separate_xlsx_files(file_name, branches_info):
    # saving file with existing file name will run-over the old one
    if branches_info is None:
        create_error_massage("branch_info is None",
         "check that excels files exist in 'easy order' folder and that at list one of them is not empty")
        return

    try:
        # delete all previous orders
        files = glob.glob(file_name+'\\*')
        for f in files:
            os.remove(f)
    except FileNotFoundError as e:
        create_error_massage(e, "check that inside 'easy order' folder there is a folder named 'your_orders'")
    except PermissionError as e:
        create_error_massage(e, "make sure that all excel files are closed before running the app'")

    # this xlsx fill will gather all orders in to 1 fill
    total_workbook = xlsxwriter.Workbook(file_name+'\\total_order.xlsx')
    workbooks = {}
    indexes = {}
    for col, branch in enumerate(branches_info):
        for supplier in branch:
            # create a new xlsx file to each supplier
            if workbooks.get(supplier, None) is None:
                workbooks[supplier] = xlsxwriter.Workbook(file_name+'\\'+supplier+'.xlsx')
                workbooks[supplier].add_worksheet('Sheet1')
                w = workbooks[supplier].get_worksheet_by_name('Sheet1')
                w.right_to_left()
                w.set_column(0, 0, width=40)
                indexes[supplier] = 0

            worksheet = workbooks[supplier].get_worksheet_by_name("Sheet1")
            total_worksheet = total_workbook.get_worksheet_by_name(supplier)

            if total_worksheet is None:
                total_worksheet = total_workbook.add_worksheet(supplier)
                total_worksheet.set_column(0, 0, width=40)

            for row, amount in enumerate(branch[supplier]):
                worksheet.write(row, indexes[supplier], amount)
                total_worksheet.write(row, col, amount)
            indexes[supplier] = indexes[supplier] + 1
            total_worksheet.right_to_left()
    for i in workbooks:
        workbooks[i].close()
    total_workbook.close()
    pass


if __name__ == '__main__':

    user_dir = os.path.expanduser("~")  # get dir to user
    file_name = user_dir+"\\Desktop\\easy order"

    try:
        info = get_info(file_name)
    except TypeError as e:
            create_error_massage(e, "raised from get_info function")
    try:
        create_separate_xlsx_files(file_name+"\\your_orders", info)
    except TypeError as e:
            create_error_massage(e, "raised from create_separate_xlsx_files function")

    print("done")
    create_error_massage("", "worked correct")'
> easy_order.py

pyinstaller --hidden-import=xlsxwriter --hidden-import=xlrd -F easy_order.py

mkdir your_orders