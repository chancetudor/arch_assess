#! python3
import os, sys
import directory
import openpyxl
from openpyxl.styles import Font
from openpyxl import utils
from openpyxl.utils import exceptions


def main():
    change_cwd()
    name = get_filename()
    wb = get_workbook(name)
    a_sheet = wb.active
    # generating tuples for specific sections
    gen_app_sec = list(a_sheet['A3':'C5'])
    info_class = list(a_sheet['A7':'C9'])
    sys_arch = list(a_sheet['A11':'C18'])
    access_control = list(a_sheet['A20':'C30'])
    data_trans_control = list(a_sheet['A32':'C36'])
    db_control = list(a_sheet['A38':'C43'])
    code_control = list(a_sheet['A45':'C47'])
    confidentiality = list(a_sheet['A49':'C53'])
    pw_control = list(a_sheet['A55':'C59'])
    testing_control = list(a_sheet['A61':'C63'])
    stride = list(a_sheet['A65':'C70'])
    threat_analysis = list(a_sheet['A72':'C79'])
    disaster_recov = list(a_sheet['A81':'C82'])
    # dict of section tuples
    # sectionName : sectionTuple
    section_list = {
        'General Application Security': gen_app_sec,
        'Information Classification': info_class,
        'System Architecture': sys_arch,
        'Access Control': access_control,
        'Data and Transaction Controls': data_trans_control,
        'Database Controls': db_control,
        'Software and Proprietary and Code Control': code_control,
        'Confidentiality': confidentiality,
        'User Accounts and Password Control': pw_control,
        'Testing Controls': testing_control,
        'STRIDE Adherence': stride,
        'Threat Analysis On Vulnerable Modules': threat_analysis,
        'Disaster Recovery': disaster_recov
    }
    section_rating = list()
    # inputting data
    for k, v in section_list.items():
        print('*** ' + k + ' ***')  # section name
        rating = input_data(v)  # section tuple
        section_rating.append(rating)  # keep track of total section rating for subsequent processing
    # calc total rating
    final_rating = calc_rating(section_rating)
    a_sheet['C84'] = final_rating
    a_sheet['C84'].font = Font(bold=True)
    # calc net score percentage
    net_score = (final_rating / 315) * 100
    a_sheet['C85'] = net_score
    a_sheet['C85'].font = Font(bold=True)
    # saving edited workbook as a copy
    print('Enter a new name for this document.')
    copy_name = get_filename()
    wb.save(copy_name)


def calc_rating(section_rating):
    total = 0
    for rating in section_rating:
        total += rating
    return total


def input_data(section):
    total = 0
    for row in section:
        for cell in row:
            if cell.coordinate[0] != 'C':
                if cell.coordinate[0] == 'B':
                    print('        Description: ' + str(cell.value))
                else:
                    if cell.value is not None:
                        print('    ' + str(cell.value))
                    else:
                        print('    ' + cell.coordinate + ' None')
            else:
                rating = get_rating()
                while float(rating) < 0.0 or float(rating) > 5.0:
                    print('ERROR: Please enter a decimal value between 0-5')
                    rating = get_rating()
                # assign rating to cell
                cell.value = float(rating)
                cell.font = Font(bold=True)
                total = total + float(rating)
        print('----------------------------------------------------------------------')
    return total


def get_rating():
    print('    ' + 'Enter your rating (0-5):')
    rating = input('      ' + '--')
    return rating


def change_cwd():
    print('Changing CWD to ' + directory.directory)
    try:
        os.chdir(directory.directory)
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
