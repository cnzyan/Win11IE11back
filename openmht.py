import win32com 
import win32com.client 
import sys
import os
import subprocess
import winreg
import tkinter as tk
from tkinter import messagebox

# .venv\Scripts\Activate.ps1
# pyinstaller -F -w openmht.py -i mht.ico
# https://www.aconvert.com/ ICO转换
def get_default_open_with(extension):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\'+extension+'\\UserChoice') as key:
        value, _ = winreg.QueryValueEx(key, "Progid")
        return value
def popup(message):
    messagebox.showinfo("提示", message)

if __name__ == '__main__':
    if len(sys.argv)>1:
        try:
            filename=sys.argv[1]
            # filename='file://'+os.path.abspath(filename)
            ie = win32com.client.DispatchEx('InternetExplorer.Application')
            ie.Visible = 1
            ie.Navigate(filename) 

            ie.Delete()
            ie.Quit() 
            ie.close()
            os.system("taskkill -im openmht.exe /F > nul")
            subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass
    else:
        root = tk.Tk()
        root.withdraw()
        exename=str(sys.argv[0]).split('\\')[-1]
        exename_ext=(exename.split('.')[-1]).lower()
        if exename_ext=='exe':
            exe_cmd="\""+sys.argv[0]+"\""
            pass
        else:
            exe_cmd="python \""+sys.argv[0]+"\""
        default_open_with = get_default_open_with(".mht")
        # print(exename)
        # print(default_open_with)
        openwith=0

        subkey = r'Applications\\openmht.exe\\shell\\open\\command'
        if exename in default_open_with:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey) as key:
                value, _ = winreg.QueryValueEx(key, "")
            if exe_cmd in value:
                openwith=1

        if openwith==1:
            popup("MHT文件的打开方式设置正确")
            pass
        else:
            popup("请设置MHT文件的打开方式为本EXE文件")

            Key = winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, subkey, 0, winreg.KEY_WRITE)

            # 字符串格式（REG_SZ）修改：
            winreg.SetValueEx(Key, "", 0, winreg.REG_SZ, ""+exe_cmd+" \"%1\"")  # 将`Key`下的`@`值的数据修改为
            command=(".\\setuserfta.exe .mht Applications\\openmht.exe > nul")
            subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
            command=(".\\setuserfta.exe .mhtm Applications\\openmht.exe > nul")
            subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
            command=(".\\setuserfta.exe .mhtml Applications\\openmht.exe > nul")
            subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
            command=(".\\setuserfta.exe mhtmlfile Applications\\openmht.exe > nul")
            subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
        
        command=("taskkill -im openmht.exe /F > nul")
        subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)

