import win32com 
import win32com.client 
import sys
import configparser
import winreg
import os
import subprocess
import tkinter as tk
from tkinter import messagebox
# .venv\Scripts\Activate.ps1
# pyinstaller -F -w openie.py -i iexplore_00001.ico
def popup(message):
    messagebox.showinfo("提示", message)
if __name__ == '__main__':
    if len(sys.argv)>1:
        url=sys.argv[1]
    else:
        url="about:blank"
    if url=="-install":
        root = tk.Tk()
        root.withdraw()
        exename=str(sys.argv[0]).split('\\')[-1]
        exename_ext=(exename.split('.')[-1]).lower()
        if exename_ext=='exe':
            exe_cmd="\""+sys.argv[0]+"\""
            pass
        else:
            exe_cmd="python \""+sys.argv[0]+"\""
        subkey = r'Applications\\openie.exe\\shell\\open\\command'
        Key = winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, subkey, 0, winreg.KEY_WRITE)

        # 字符串格式（REG_SZ）修改：
        winreg.SetValueEx(Key, "", 0, winreg.REG_SZ, ""+exe_cmd+" \"%1\"")  # 将`Key`下的`@`值的数据修改为
        command=(".\\setuserfta.exe .url Applications\\openie.exe > nul")
        subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
        command=("taskkill -im openmht.exe /F > nul")
        subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
        popup("URL文件的打开方式已设置正确")
        print()
        pass
    else:
        try:
            file_name=str(sys.argv[1]).split('\\')[-1]
            file_name_ext=(file_name.split('.')[-1]).lower()
            # print(file_name_ext)
            if file_name_ext=='url':
                config = configparser.ConfigParser()  # 类实例化
                # 定义文件路径
                configpath = url
                config.read(configpath)
                url = config['InternetShortcut']['URL']
                pass
            else:
                
                pass
        except:
            pass
        # print(url)
        try:
            ie = win32com.client.DispatchEx('InternetExplorer.Application')
            ie.Visible = 1
            ie.Navigate(url) 

            ie.Delete()
            ie.Quit() 
            ie.close()
        except:
            pass
