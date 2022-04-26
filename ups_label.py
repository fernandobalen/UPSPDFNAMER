"https://www.blog.pythonlibrary.org/2014/10/20/pywin32-how-to-bring-a-window-to-front/"

import win32gui
import pandas as pd 
import pyautogui
import time
import os
import argparse
import locale 

language = locale.getdefaultlocale()

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


parser = argparse.ArgumentParser()
parser.add_argument('-c', '--csv',
                    help='path or name of the excel file', required=True)

args = vars(parser.parse_args())

new_args = args['csv']

if __name__ == "__main__":

    df = pd.read_csv(new_args)
    for j in range(1,len(df)+1):
        auftragnummer = df.loc[j-1][0].split(';')[0]
        results = []
        top_windows = []
        win32gui.EnumWindows(windowEnumerationHandler, top_windows)
        for i in top_windows:
            if list(language)[0] == "de_DE":
                if "druckausgabe speichern unter" in i[1].lower():
                    if os.path.exists("{}.pdf".format(auftragnummer)):
                        auftragnummer = "{}_{}".format(df.loc[j-1][0].split(';')[0],j)
                    else:
                        auftragnummer = df.loc[j-1][0].split(';')[0]
                    win32gui.ShowWindow(i[0],5)
                    win32gui.SetForegroundWindow(i[0])
                    pyautogui.typewrite('{}'.format(auftragnummer)) #take pdf out make new exe
                    pyautogui.press('enter')
            elif list(language)[0] == "en_DE":
                if "druckausgabe speichern unter" in i[1].lower():
                    if os.path.exists("{}.pdf".format(auftragnummer)):
                        auftragnummer = "{}_{}".format(df.loc[j-1][0].split(';')[0],j)
                    else:
                        auftragnummer = df.loc[j-1][0].split(';')[0]
                    win32gui.ShowWindow(i[0],5)
                    win32gui.SetForegroundWindow(i[0])
                    pyautogui.typewrite('{}'.format(auftragnummer)) #take pdf out make new exe
                    pyautogui.press('enter')
            else:
                if "save print output as" in i[1].lower():
                    if os.path.exists("{}.pdf".format(auftragnummer)):
                        auftragnummer = "{}_{}".format(df.loc[j-1][0].split(';')[0],j)
                    else:
                        auftragnummer = df.loc[j-1][0].split(';')[0]
                    win32gui.ShowWindow(i[0],5)
                    win32gui.SetForegroundWindow(i[0])
                    pyautogui.typewrite('{}'.format(auftragnummer)) #take pdf out make new exe
                    pyautogui.press('enter')

        time.sleep(1.5)
        print("{}/{} saved".format(j,len(df)))




