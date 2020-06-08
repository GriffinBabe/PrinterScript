"""
    Windows printer script.

    Prints a file every few times to prevent cartridge drying on cheap ass printers (like my dad's printer)

    GriffinBabe
"""

import win32print

from subprocess import call
from datetime import datetime

PRINTER_NAME = "HP OfficeJet 4650 series [F1117F]"  # change this for other printer
PRINT_FILE_NAME = "print_test.pdf"
DATE_FILE_NAME = "dates.txt"
ACROBAT = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
EVERY_X_DAYS = 7  # prints every week


def list_jobs():
    p_handle = win32print.OpenPrinter(PRINTER_NAME)  # printer handler
    print_jobs = win32print.EnumJobs(p_handle, 0, -1, 1)
    return print_jobs


def print_file(file_name):
    call([ACROBAT, "/T", PRINT_FILE_NAME, PRINTER_NAME])


def write_date():
    with open(DATE_FILE_NAME, 'w') as f:
        f.write("%s" % datetime.now())


def read_date():
    file = open(DATE_FILE_NAME, 'r')
    date_str = file.readline()
    parsed_date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S.%f')
    return parsed_date


def check_print():
    # checks if there is no print job yet
    for job in list_jobs():
        doc_name = job['pDocument']
        if doc_name == PRINT_FILE_NAME:
            print("Document already in printing queue")
            return False

    date_now = datetime.now()
    date_last = read_date()

    # checks if the last time the document was printed was less than 7 days ago
    if (date_last - date_now).days < EVERY_X_DAYS:
        print("Document was already printed %s days ago.".format((date_last - date_now).days))
        return False

    return should_print


should_print = check_print()

if should_print:
    print_file()
    write_date()
