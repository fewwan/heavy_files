import commctrl
import win32com.client as win32
import win32process
import win32con
import win32gui
import win32api
import psutil
import ctypes
import os
import sys
from pathlib import Path
import urllib.parse
from time import sleep
import argparse
from functools import lru_cache
user32 = ctypes.windll.user32
clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}' #Valid for InternetExplorer as well!
shellwindows = win32.Dispatch(clsid)
fso = win32.Dispatch('Scripting.FileSystemObject')
####################################################################################################

# naturalsize function from humanize, https://pypi.org/project/humanize/
suffixes = {
    "decimal": ("kB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"),
    "binary": ("KiB", "MiB", "GiB", "TiB", "PiB", "EiB", "ZiB", "YiB"),
    "gnu": "KMGTPEZY",
}
def naturalsize(value, binary=False, gnu=False, format="%.1f"):
    if gnu:
        suffix = suffixes["gnu"]
    elif binary:
        suffix = suffixes["binary"]
    else:
        suffix = suffixes["decimal"]

    base = 1024 if (gnu or binary) else 1000
    bytes = float(value)
    abs_bytes = abs(bytes)

    if abs_bytes == 1 and not gnu:
        return "%d Byte" % bytes
    elif abs_bytes < base and not gnu:
        return "%d Bytes" % bytes
    elif abs_bytes < base and gnu:
        return "%dB" % bytes

    for i, s in enumerate(suffix):
        unit = base ** (i + 2)
        if abs_bytes < unit and not gnu:
            return (format + " %s") % ((base * bytes / unit), s)
        elif abs_bytes < unit and gnu:
            return (format + "%s") % ((base * bytes / unit), s)
    if gnu:
        return (format + "%s") % ((base * bytes / unit), s)
    return (format + " %s") % ((base * bytes / unit), s)

####################################################################################################

# https://stackoverflow.com/questions/21241708/python-get-a-list-of-selected-files-in-explorer-windows-7/52959617#52959617

def getEditText(hwnd):
    '''
    api returns 16 bit characters so buffer needs 1 more char for null and twice the num of chars.
    '''
    buf_size = (win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0) +1 ) * 2
    target_buff = ctypes.create_string_buffer(buf_size)
    win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, buf_size, ctypes.addressof(target_buff))
    return target_buff.raw.decode('utf16')[:-1]# remove the null char on the end

def _normaliseText(controlText):
    '''
    Remove '&' characters, and lower case.
    Useful for matching control text.
    '''
    return controlText.lower().replace('&', '')

def _windowEnumerationHandler(hwnd, resultList):
    '''
    Pass to win32gui.EnumWindows() to generate list of window handle,
    window text, window class tuples.
    '''
    resultList.append((hwnd, win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd)))

def searchChildWindows(currentHwnd,
                       wantedText=None,
                       wantedClass=None,
                       selectionFunction=None):
    results = []
    childWindows = []
    try:
        win32gui.EnumChildWindows(currentHwnd,
                      _windowEnumerationHandler,
                      childWindows)
    except win32gui.error:
        # This seems to mean that the control *cannot* have child windows,
        # i.e. not a container.
        return
    for childHwnd, windowText, windowClass in childWindows:
        descendentMatchingHwnds = searchChildWindows(childHwnd)
        if descendentMatchingHwnds:
            results += descendentMatchingHwnds

        if wantedText and \
            not _normaliseText(wantedText) in _normaliseText(windowText):
                continue
        if wantedClass and \
            not windowClass == wantedClass:
                continue
        if selectionFunction and \
            not selectionFunction(childHwnd):
                continue
        results.append(childHwnd)
    return results


def selected_files():
    address_1=""
    files = []
    result=[]
    w=win32gui
    window = w.GetForegroundWindow()
    #print("window: %s" % window)
    try:
        if (window != 0):
            if (w.GetClassName(window) == 'CabinetWClass'): # the main explorer window
                #print("class: %s" % w.GetClassName(window))
                #print("text: %s " %w.GetWindowText(window))
                children = list(set(searchChildWindows(window)))
                addr_edit = None
                file_view = None
                for child in children:
                    if (w.GetClassName(child) == 'WorkerW'): # the address bar
                        addr_children = list(set(searchChildWindows(child)))
                        for addr_child in addr_children:
                            if (w.GetClassName(addr_child) == 'ReBarWindow32'):
                                addr_edit = addr_child
                                addr_children = list(set(searchChildWindows(child)))
                                for addr_child in addr_children:
                                    if (w.GetClassName(addr_child) == 'Address Band Root'):
                                        addr_edit = addr_child
                                        addr_children = list(set(searchChildWindows(child)))
                                        for addr_child in addr_children:
                                            if (w.GetClassName(addr_child) == 'msctls_progress32'):
                                                addr_edit = addr_child
                                                addr_children = list(set(searchChildWindows(child)))
                                                for addr_child in addr_children:
                                                    if (w.GetClassName(addr_child) == 'Breadcrumb Parent'):
                                                        addr_edit = addr_child
                                                        addr_children = list(set(searchChildWindows(child)))
                                                        for addr_child in addr_children:
                                                            if (w.GetClassName(addr_child) == 'ToolbarWindow32'):
                                                                text=getEditText(addr_child)
                                                                if "\\" in text:
                                                                    address_1=getEditText(addr_child)[text.index(" ")+1:]
                                                                    # print("Address --> "+address_1)

            for window in range(shellwindows.Count):
                window_URL = urllib.parse.unquote(shellwindows[window].LocationURL,encoding='ISO 8859-1')
                window_dir = window_URL.split('///')[1].replace("/", "\\")
                # print("Directory --> "+window_dir)
                if window_dir==address_1:
                    selected_files = shellwindows[window].Document.SelectedItems()
                    for file in range(selected_files.Count):
                        files.append(selected_files.Item(file).Path)
                    # print("Files --> "+str(files))
                    return tuple(files)
    except Exception as exc:
        pass

####################################################################################################

def change_speed(speed):
    """
    1 - slow
    10 - standard
    20 - fast
    """
    speed=int(speed)
    if speed in range(1, 21):
        set_mouse_speed = 113   # 0x0071 for SPI_SETMOUSESPEED
        ctypes.windll.user32.SystemParametersInfoA(set_mouse_speed, 0, speed, 0)
    else:
        raise ValueError('speed must be between 1 and 20')


def get_current_speed():
    get_mouse_speed = 112   # 0x0070 for SPI_GETMOUSESPEED
    speed = ctypes.c_int()
    ctypes.windll.user32.SystemParametersInfoA(get_mouse_speed, 0, ctypes.byref(speed), 0)

    return speed.value

####################################################################################################

default_speed = get_current_speed()

MAX_SIZE = 500
MIN_SPEED = 1
MAX_SPEED = default_speed
REFRESH_RATE = 1

parser = argparse.ArgumentParser(description='Slows down the mouse based on the size of the selected files.',
                                 formatter_class=lambda prog: argparse.HelpFormatter(prog, max_help_position=32))

parser.add_argument('--max_size', required=False, type=float, default=MAX_SIZE,
                    help=f"""Size in MB. The maximum size, when the speed is the minimum. (default: {naturalsize(MAX_SIZE)})""")

parser.add_argument('--min_speed', required=False, type=int, default=MIN_SPEED,
                    help=f"""Speed between 1 and 20. The minimum speed. (default: {MIN_SPEED})""")

parser.add_argument('--max_speed', required=False, type=int, default=MAX_SPEED,
                    help=f"""Speed between 1 and 20. The maximum speed. {f'(default: {MAX_SPEED}{" (current speed)" if MAX_SPEED==default_speed else ""})'}""")

parser.add_argument('--refresh_rate', required=False, type=float, default=REFRESH_RATE,
                    help=f"""(default: {REFRESH_RATE} s)""")

args = parser.parse_args()

MAX_SIZE = int(args.max_size*1000000)
MIN_SPEED = args.min_speed
MAX_SPEED = args.max_speed
REFRESH_RATE = args.refresh_rate

####################################################################################################
from timer import Timer
cache = {}
@lru_cache
def getsize(path):
    if os.path.isfile(path):
        return fso.GetFile(path).Size
    elif os.path.isdir(path):
        return fso.GetFolder(path).Size

try:
    while True:
        files = selected_files()
        if files and win32api.GetKeyState(0x01) < 0:
            size=sum(map(getsize, files))
            speed=MAX_SPEED-size*MAX_SPEED//MAX_SIZE
            if speed<MIN_SPEED:
                speed=MIN_SPEED
            print(naturalsize(size), speed)
            change_speed(speed)
        if win32api.GetKeyState(0x01) >= 0:
            change_speed(default_speed)
        sleep(REFRESH_RATE)
except Exception as exc:
    print(sys.exc_info())
finally:
    change_speed(default_speed)
