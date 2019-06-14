#! python3
import os, sys
import openpyxl
from openpyxl import utils
from openpyxl.utils import exceptions


def main():
    change_cwd()
    name = get_filename()
    wb = get_workbook(name)
    a_sheet = wb.active
    # generating tuples for specific sections
    gen_app_sec = tuple(a_sheet['A3':'C5'])
    info_class = tuple(a_sheet['A7':'C9'])
    sys_arch = tuple(a_sheet['A11':'C18'])
    access_control = tuple(a_sheet['A20':'C30'])
    data_trans_control = tuple(a_sheet['A32':'C36'])
    db_control = tuple(a_sheet['A38':'C43'])
    code_control = tuple(a_sheet['A45':'C47'])
    confidentiality = tuple(a_sheet['A49':'C53'])
    pw_control = tuple(a_sheet['A55':'C59'])
    testing_control = tuple(a_sheet['A61':'C63'])
    stride = tuple(a_sheet['A65':'C70'])
    threat_analysis = tuple(a_sheet['A72':'C79'])
    disaster_recov = tuple(a_sheet['A81':'C82'])
    # list of section tuples
    section_list = [gen_app_sec,
                    info_class,
                    sys_arch,
                    access_control,
                    data_trans_control,
                    db_control,
                    code_control,
                    confidentiality,
                    pw_control,
                    testing_control,
                    stride,
                    threat_analysis,
                    disaster_recov
                    ]
    for section in section_list:
        input_data(section)


def input_data(sheet):
    r = 2
    col = 1
    print(sheet.cell(row=r, column=col).value)

    for row in gen_app_sec:
        for cell in row:
            print('    ' + cell.coordinate, cell.value)
        print('-- END OF ROW --')



def change_cwd():
    print(r"Changing CWD to C:\Users\chtudor\Documents\ArchitectureAssessments")
    try:
        os.chdir(r"C:\Users\chtudor\Documents\ArchitectureAssessments")
        print("Directory changed.")
    except OSError:
        print("Error: unable to change CWD.")


def get_filename():
    print("What is the name of the .xlsx document?")
    name = input('--')
    name = name + '.xlsx'
    return name


def get_workbook(fn):
    try:
        wb = openpyxl.load_workbook(fn)
        return wb
    except openpyxl.utils.exceptions.SheetTitleException:
        print("Error: cannot find workbook ", fn)
        sys.exit(-1)





if __name__ == '__main__':
    main()
