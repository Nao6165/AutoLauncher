import subprocess
import webbrowser
import datetime

from tkinter import messagebox

#webbrowser open option
EXIST=0
NEW=1
TAB=2

def open_folder(path):
    subprocess.run('explorer {}'.format(path))

def open_text_file(path):
    subprocess.Popen([r'C:\WINDOWS\system32\notepad.exe',path])

def open_outlook():
    subprocess.Popen([r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'])

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
        do_monday_event()
    elif week == 1: #Tuesday
        do_tuesday_event()
    elif week == 2: #Wednesday
        do_wednesday_event()
    elif week == 3: #Thursday
        do_thursday_event()
    elif week == 4: #Friday
        do_friday_event()
    elif week == 5: #Saturday
        do_saturday_event()
    elif week == 6: #Sunday
        do_sunday_event()
    else : 
        pass #do nothing

def do_monday_event():
    url = "https://www.gizmodo.jp/"
    open_default_browser(url,TAB)

def do_tuesday_event():
    url = "https://forbesjapan.com/"
    open_default_browser(url,TAB)

def do_wednesday_event():
    url = "https://techable.jp/"
    open_default_browser(url,TAB)

def do_thursday_event():
    url = "https://github.com/"
    open_default_browser(url,TAB)

def do_friday_event():
    url = "https://makezine.jp/blog/category/makers"
    open_default_browser(url,TAB)

def do_saturday_event():
    url = "https://codezine.jp/"
    open_default_browser(url,TAB)

def do_sunday_event():
    url = "https://hatenablog.com/"
    open_default_browser(url,TAB)

def do_monthly_event():
    dt_now = datetime.datetime.now()
    day = dt_now.day
    if day == 25:
        messagebox.showinfo('AutoLauncher', 'Today is Payday!')
    else:
        pass

def main():
    #messagebox.showinfo('AutoLauncher', 'processing start')

    open_folder('C:')

    open_text_file(r'C:\Users\user\Documents\temp\test.txt')

    open_outlook()

    open_excel(r'C:\Users\user\Documents\temp\test.xlsx')
 
    url  = "http://google.com/"
    open_default_browser(url,EXIST)
 
    do_weekly_event()

    do_monthly_event()

    #messagebox.showinfo('AutoLauncher', 'processing end')

#
# main
#
if __name__ == '__main__':
        main()