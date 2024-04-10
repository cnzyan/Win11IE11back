import os,win32api,sys
from pathlib import Path
import configparser
import win32com 
import win32com.client
import tkinter as tk
from tkinter import ttk,messagebox
import json
import winreg
# .venv\Scripts\Activate.ps1
# pyinstaller -F -w iefav.py -i iexplore_00001.ico

def is_admin():
    # 由于win32api中没有IsUserAnAdmin函数,所以用了这种方法
    try:
        # 在c:\windows目录下新建一个文件test01.txt
        testfile=os.path.join(os.getenv("windir"),"test01.txt")
        open(testfile,"w").close()
    except OSError: # 不成功
        return False
    else: # 成功
        os.remove(testfile) # 删除文件
        return True

def open_url(url,browser):
    if browser=='ie':
        try:
            if 1==1:
                ie = win32com.client.DispatchEx('InternetExplorer.Application')
                ie.Visible = 1
                ie.Navigate(url) 
                ie.Delete()
                ie.Quit() 
                ie.close()
            else:
                os.system('start \"C:\\Program Files (x86)\\Internet Explorer\\Internet Explorer\\iexplore.exe\" '+url+' -Embedding')
        except:
            pass
    else:
        os.system('start '+url)
        pass
# 确保收藏夹路径存在
def list_files(favorites_path,depth):
    # print(favorites_path.name+f'(层级{depth})包含:')
    global url_list
    inner_dirs=[]
    if favorites_path.exists() and favorites_path.is_dir():
        # 列举收藏夹中的文件
        for filename in favorites_path.iterdir():
            if filename.is_dir():
                
                if depth==0:
                    temp_favorites_path='Favorites'
                else:
                    temp_favorites_path=favorites_path.name
                if temp_favorites_path in url_list:
                    # url_list[temp_favorites_path].update({file_name_main:{}})
                    pass
                else:
                    # url_list.update({temp_favorites_path:{file_name_main:{}}})
                    pass
                inner_dirs.append(filename)
                pass
            else:
                file_name=filename.name
                file_name_main=''.join(file_name.split('.')[0:-1]).lower()
                file_name_ext=(file_name.split('.')[-1]).lower()
                if file_name_ext=='url':
                    try:
                        config = configparser.ConfigParser()  # 类实例化
                        # 定义文件路径
                        configpath = filename
                        config.read(configpath)
                        url = config['InternetShortcut']['URL']
                        if depth==0:
                            temp_favorites_path='Favorites'
                        else:
                            temp_favorites_path=favorites_path.name
                        if temp_favorites_path in url_list:
                            url_list[temp_favorites_path].update({file_name_main:url})
                        else:
                            url_list.update({temp_favorites_path:{file_name_main:url}})

                        # print(file_name_main+':'+url)
                        
                    except Exception as e:
                        # print(e)
                        pass
        for inner_dir in inner_dirs:
            list_files(inner_dir,depth+1)
    else:
        # print(f"Favorites folder does not exist at: {favorites_path.name}")
        pass
def slct(evt):
    for item in tree.selection():
        # print(tree.item(item, "values")[0])
        try:
            open_url(tree.item(item, "values")[0],'ie')
        except:
            pass


def open_(evt):  # 
    for item in tree.selection():
        print(f"{item} has opened")
def close(evt):
    for item in tree.selection():
        print(f"{item} has closed")
def show_about():
    result = messagebox.showinfo(
        title='关于IE收藏夹', message='已开发Windows11系统收藏夹IE浏览器打开工具。\n用于在Windows11系统中调用IE浏览器打开系统收藏夹内的网址。\n Author：CrazYan 2024')
def show_help():
    popup('帮助','如果点击链接没有反应，可能是后台IE挂起，\n请从“操作”菜单执行“重置IE”指令')
def popup(title,message):
    messagebox.showinfo(title, message)
def reset_ie():
    import subprocess

    command = 'taskkill -im iexplore.exe /F'
    subprocess.call(command, creationflags=subprocess.CREATE_NO_WINDOW)
    #os.system("start /B taskkill -im iexplore.exe /F")
if __name__ == '__main__':
    if is_admin():
        # 获取当前用户的用户名
        user_name = os.getlogin()
        # 构建收藏夹的路径
        favorites_path = Path(os.path.expanduser('~')) / 'Favorites'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders') as key:
            favorites_path, _ = winreg.QueryValueEx(key, "Favorites")
        favorites_path=Path(favorites_path)    
        url_list={}
        url_json=json.dumps(url_list)
        url_list=json.loads(url_json) 
        list_files(favorites_path,0)
        # print(url_list)
        
        win=tk.Tk()
        win.title("IE收藏夹")
        # 获取窗口大小
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        Width = 400  # 窗口大小值
        Hight = 300
        # 计算中心坐标
        cen_x = (sw - Width) / 2
        cen_y = (sh - Hight) / 2
        # 设置窗口大小并居中
        win.geometry('%dx%d+%d+%d' % (Width, Hight, cen_x, cen_y)) 
        menubar = tk.Menu(win)
        filemenu =tk.Menu(menubar, tearoff=0)
        filemenu.add_command(
            label='重置IE', accelerator='Ctrl+R', command=reset_ie)
        menubar.add_cascade(label='操作', menu=filemenu)
        aboutmenu = tk.Menu(menubar, tearoff=0)
        aboutmenu.add_command(
            label='帮助', accelerator='Ctrl+H', command=show_help)
        aboutmenu.add_command(
            label='关于', accelerator='Ctrl+A', command=show_about)
        menubar.add_cascade(label='关于', menu=aboutmenu)
        win.config(menu=menubar)
        # 此为根节点
        tree = ttk.Treeview(win, show = "tree")
        for key in url_list:
            father = tree.insert("", 1, key, text=key)
            for v in url_list[key]:
                try:
                    tree.insert(father, 1, v, 
                        text=v, values=url_list[key][v])
                except:
                    pass
        tree.pack(side=tk.LEFT, expand = True, fill = tk.BOTH)
        tree.bind('<<TreeviewSelect>>', slct)
        tree.bind('<<TreeviewOpen>>', open_)
        tree.bind('<<TreeviewClose>>', close)
        scroll = ttk.Scrollbar(win)
        scroll.config(command=tree.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        # 给treeview添加配置
        tree.configure(yscrollcommand=scroll.set)
        win.mainloop()
    else:
        # 以管理员权限重新运行程序
        win32api.ShellExecute(None,"runas", sys.executable, __file__, None, 1)