import subprocess
import webbrowser

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

def main():
    #messagebox.showinfo('AutoLauncher', 'processing start')

    open_folder('C:')

    open_excel(r'C:\Users\user\Documents\temp\test.xlsx')
 
    url1  = "https://www.youtube.com/"
    open_default_browser(url1,EXIST)
 
    url2 = "http://google.com/"
    open_default_browser(url2,TAB)

    #messagebox.showinfo('AutoLauncher', 'processing end')

#
# main
#
if __name__ == '__main__':
        main()