import subprocess
import webbrowser
import datetime

from tkinter import messagebox

EXIST=0
NEW=1
TAB=2

# reference:https://kazusa-pg.com/python-open-folder/
def open_folder(path):
    subprocess.run('explorer {}'.format(path))

def open_excel(path):
    subprocess.Popen([r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE',path])

def open_chrome(url,new):
    browser = webbrowser.get(r'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" %s')
    browser.open(url,new)

def open_default_browser(url,new):
    webbrowser.open(url,new)

def do_weekly_event():
    week = datetime.date.today().weekday()
    if week == 0: #Monday
        do_Monday_event()
    elif week == 1: #Tuesday
        do_Tuesday_event()
    elif week == 2: #Wednesday
        do_Wednesday_event()
    elif week == 3: #Thursday
        do_Thursday_event()
    elif week == 4: #Friday
        do_Friday_event()
    elif week == 5: #Saturday
        do_Saturday_event()
    elif week == 6: #Sunday
        do_Sunday_event()
    else : 
        pass #do nothing

def do_Monday_event():
    url = "https://www.gizmodo.jp/"
    open_default_browser(url,TAB)

def do_Tuesday_event():
    url = "https://forbesjapan.com/"
    open_default_browser(url,TAB)

def do_Wednesday_event():
    url = "https://techable.jp/"
    open_default_browser(url,TAB)

def do_Thursday_event():
    url = "https://github.com/"
    open_default_browser(url,TAB)

def do_Friday_event():
    url = "https://makezine.jp/blog/category/makers"
    open_default_browser(url,TAB)

def do_Saturday_event():
    url = "https://codezine.jp/"
    open_default_browser(url,TAB)

def do_Sunday_event():
    url = "https://hatenablog.com/"
    open_default_browser(url,TAB)

def do_Monthly_event():
    dt_now = datetime.datetime.now()
    day = dt_now.day
    if day == 25:
        messagebox.showinfo('AutoLauncher', 'Today is Payday!')
    else:
        pass

def main():
    #messagebox.showinfo('AutoLauncher', 'processing start')

    open_folder('C:')

    open_excel(r'C:\Users\user\Documents\temp\test.xlsx')
 
    url  = "http://google.com/"
    open_default_browser(url,EXIST)
 
    do_weekly_event()

    #messagebox.showinfo('AutoLauncher', 'processing end')

#
# main
#
if __name__ == '__main__':
        main()