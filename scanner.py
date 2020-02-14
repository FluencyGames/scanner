from _ctypes import COMError
from pywinauto.application import Application
from pywinauto.timings import TimeoutError
import pywinauto.keyboard as Keyboard
import pywinauto.findwindows as Windows
import shutil
import os
import time
import threading
import warnings
import json

Settings = {}

# disable warnings on 32/64 bit incompatibility
warnings.simplefilter('ignore', category=UserWarning)

# our app
# TODO: can we connect to existing app ever?
app = Application(backend="uia").start('C:\\HR Scanning\\Cobb HR File Scanning Tool.exe')
window = 'aXs Info - Cobb HR Scanning Tool'

# these will keep track of all our copy processes
procs = []
employees = []


class Copier(threading.Thread):
    global Settings

    def __init__(self, employee_id):
        super().__init__()
        self.employee_id = employee_id

    def run(self):
        src = 'C:\\HR Base Scan\\' + self.employee_id
        dest = 'D:\\' + Settings['DataDest'] + '\\' + Settings['DataDate'] + '\\' + self.employee_id + '\\'

        try:
            os.mkdir(dest)
        except FileExistsError:
            pass

        # dest
        for file in os.listdir(src):
            s = os.path.join(src, file)
            d = os.path.join(dest, file)
            try:
                shutil.copy2(s, d)
            except FileExistsError:
                pass
            except FileNotFoundError:
                print('{} not found, skipping.'.format(file))
            finally:
                pass


class Scanner(threading.Thread):

    def __init__(self):
        super().__init__()

    def run(self):
        # print('Scan files...')
        scanner_class = 'WindowsForms10.Window.8.app.0.141b42a_r6_ad1'
        time.sleep(4)
        scanner = None

        try:
            scanner = Application().connect(class_name='WindowsForms10.Window.8.app.0.141b42a_r6_ad1')
        except Windows.ElementAmbiguousError:
            print('Failed to scan, multiple instancees of Scanner occured.')
            return
        except Windows.ElementNotFoundError:
            print('Failed to scan, did not find scanning dialog. Try re-scanning.')
            return

        if scanner is not None:
            scanner[scanner_class].Scan.click()
        else:
            print('Scan dialog Not found. Aborting')
            return

        try:
            scanner[scanner_class].wait_not(wait_for_not='ready enabled visible', timeout=900, retry_interval=1)
        except TimeoutError:
            scanner[scanner_class].wait_not(wait_for_not='ready enabled visible', timeout=900, retry_interval=1)

        # print('Scan Complete!')


def save_settings():
    global Settings
    with open('scanner.cfg', 'w') as f:
        json.dump(Settings, f)


def read_settings():
    global Settings
    try:
        with open('scanner.cfg', 'r') as f:
            Settings = json.load(f)
            # print(Settings)
    except FileNotFoundError:
        Settings['DataDest'] = ''
        Settings['DataDate'] = ''
        Settings['WindowName'] = 'aXs Info - Cobb HR Scanning Tool'
        Settings['ScanningToolPath'] = 'C:\\HR Scanning\\Cobb HR File Scanning Tool.exe'
        Settings['ScannerClassName'] = 'WindowsForms10.Window.8.app.0.141b42a_r6_ad1'


def prompt_for_folder(prompt, root='', default=''):
    folder = input(prompt + ' (Default=\'{}\'): '.format(default))
    if len(folder) == 0:
        folder = default
    else:
        # create the destination directory if it does not exist
        dest = os.path.join('D:', root, folder);
        if not os.path.isdir(dest):
            r = input('Directory named \'{}\' does not exist. Create (y/n)? '.format(dest))
            if r in ['Y', 'y', 'Yes', 'yes']:
                try:
                    os.mkdir(dest)
                except FileExistsError:
                    pass
                except OSError as error:
                    folder = ''
                    print(error)
                finally:
                    pass

    return folder


# prompt the user in case anything has changed...
def prompt_current_settings():
    global Settings

    datadest = prompt_for_folder(prompt='Enter current destination folder', default=Settings['DataDest'])
    if len(datadest) > 0:
        Settings['DataDest'] = datadest

    datadate = prompt_for_folder('Enter current date (mmddyyyy): ', root=Settings['DataDest'],
                                 default=Settings['DataDate'])
    if len(datadate) != 0:
        Settings['DataDate'] = datadate


def scan_id(employee_id: str) -> bool:
    global app
    global window

    app[window].child_window(title='Employee ID').draw_outline()
    app[window].child_window(title='Employee ID').click()
    app[window].type_keys(employee_id)

    scanner = Scanner()
    scanner.start()

    try:
        # app[window].child_window(title='Start Scanning').click()
        app[window].child_window(title='Scanner:').draw_outline()
        app[window].child_window(title='Scanner:').click()
        Keyboard.send_keys("{TAB}"
                           "{VK_SPACE}")
    except COMError:
        print('Handled exception')
        return True
    finally:
        # print('Scanner completing...')
        scanner.join()
        return True


def copy_files(employee_id: str):
    global procs

    if len(employee_id) > 0:
        copier = Copier(employee_id)
        copier.start()
        procs.append(copier)


def enter_employee_ids():
    done = False
    id_list = []
    while not done:
        id_string = input('Enter Ids to scan (<Enter> to end) : ')
        if id_string == '':
            done = True
        else:
            id_list.append(id_string)

    return id_list


# read the current settings for scanning (destination directories, etc)
read_settings()
prompt_current_settings()

# get the list of employee ids
employee_list = []
employee = None
not_done_scanning = True
last_employee = ''

while not_done_scanning:
    employee_list.clear()

    # get the list of ids to scan
    employee_list = enter_employee_ids()
    if len(employee_list) == 0:
        exit(-1)
    print(employee_list)

    # get the first id to scan
    employee = employee_list.pop(0)
    last_employee = ''

    # for employee in employee_list:
    while employee is not None:
        prompt = '<Enter> to scan {}'.format(employee, last_employee)
        if last_employee != '':
            prompt += ' or (R)escan {}'.format(last_employee)
        resp = input(prompt + ': ')

        # copy the files from the previously scanned files
        copy_files(last_employee)

        if resp == 'q':
            break
        elif resp in ['R', 'r']:
            # if we selected to rescan, put this employee back into the list
            employee_list.insert(0, employee)
            employee = last_employee
        else:
            last_employee = employee

        # scan this employee id
        scan_id(employee)

        if len(employee_list) > 0:
            employee = employee_list.pop(0)
        else:
            resp = input('(R)escan {} or <Enter> to quit: '.format(employee))
            if resp in ['R', 'r']:
                employee = last_employee
            else:
                employee = None

    resp = input('<Enter> continue scanning or (Q)uit: ')
    if resp == 'q' or resp == 'Q':
        not_done_scanning = False

# when we are done with our scanning
# be sure to copy the last employee over to our backup
copy_files(last_employee)

# wait until we have finished copying
print('Hold on, copying the last files...')
for p in procs:
    p.join()

save_settings()
