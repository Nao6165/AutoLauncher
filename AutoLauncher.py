import subprocess
import webbrowser

# reference:https://kazusa-pg.com/python-open-folder/
def open_folder(path):
    subprocess.run('explorer {}'.format(path))

def open_chrome(url):
    browser = webbrowser.get(r'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" %s')
    browser.open(url)

def main():
    open_folder('C:')

    subprocess.Popen(r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')

    url = "Amazon.co.jp"
    open_chrome(url)

#
# main
#
if __name__ == '__main__':
        main()