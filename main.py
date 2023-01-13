import string
import os
import webbrowser
import winreg
import ctypes
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as msg
from tkinter import filedialog as filemsg
import base64
import datetime
import win32api
import win32con
from PIL import ImageTk, Image
from shutil import rmtree
from files import regfix, imgicon, imgbg, usehelptext

#使用方法
def usehelp():
    msgbox = Tk()
    helpwidth = 840
    helpheight = 880
    helpx = int((screenwidth - helpwidth) / 2)
    helpy = int((screenheight - helpheight) / 2)
    msgbox.geometry("%dx%d+%d+%d" % (helpwidth, helpheight, helpx, helpy))
    msgbox.resizable(0, 0)

    icontmp = open("icontmp.ico", "wb+")
    icontmp.write(base64.b64decode(imgicon))
    icontmp.close()
    msgbox.iconbitmap('icontmp.ico')
    os.remove('icontmp.ico')

    msgbox.title('使用方法')
    Label(msgbox, text=usehelptext, justify=tk.LEFT, font=("微软雅黑", 15)).pack(fill=X)
    msgbox.mainloop()

#关于作者
def about():
    if msg.askokcancel('关于作者', '这个程序是Xizo编写的。\n版本 V1.1-2023.01.13\n作者QQ群：912164832\n点击确定跳转至我的bilibili主页。'):
        webbrowser.open('https://space.bilibili.com/407420026', new=0, autoraise=True)

#选择图标
def choose_img():
    global imgpath
    imgpath = filemsg.askopenfile(filetypes=[("图片文件", (".ico", ".png", ".jpg", ".bmp", ".tif", ".tiff", ".jpeg", ".jpe"))])
    imgpath = imgpath.name
    choose_img_bar.configure(state='normal')
    choose_img_bar.delete("1.0","end")
    choose_img_bar.insert("insert", imgpath)
    choose_img_bar.configure(state='disabled')
    #预览小图标更新图片
    showimg_open = Image.open(imgpath)
    showimg = ImageTk.PhotoImage(showimg_open.resize((48,48)))
    ShowImage.config(image=showimg)
    ShowImage.image = showimg

#重置图标
def rest():
    errreg = 0
    errinf1 = 0
    errinf2 = 0
    errinf = 1
    try:
        queryimgpath = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + choose_disk_combobox.get()[0] + '\DefaultIcon')
        os.remove(queryimgpath)
        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + choose_disk_combobox.get()[0] + '\DefaultIcon')
        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + choose_disk_combobox.get()[0])
    except:
        print()
        errreg = 1
    try:
        os.remove(choose_disk_combobox.get()[0] + r":\icon.ico")
    except:
        errinf1 = 1
    try:
        os.remove(choose_disk_combobox.get()[0] + r":\autorun.inf")
    except:
        errinf2 = 1
    if errinf1 == 0 or errinf2 == 0:
        errinf = 0
    if errreg == 1 and errinf == 1:
        msg.showerror('重置图标', '错误：没有选择磁盘\n或图标已经重置了')
    if errreg == 1 and errinf == 0:
        msg.showinfo('重置图标', '图标重置成功(autorun.inf)')
    if errreg == 0 and errinf == 1:
        msg.showinfo('重置图标', '图标重置成功(注册表)')
    if errreg == 0 and errinf == 0:
        msg.showinfo('重置图标', '图标重置成功(autorun.inf + 注册表)')

#安装图标
def install():
    errimg = 0
    errdisk = 0
    if os.path.exists(r'C:\Windows\_XizoDiskIconChanger\icon') == False:
        os.makedirs(r'C:\Windows\_XizoDiskIconChanger\icon')
    try:
        originimgpath = str(choose_img_bar.get(1.0, END).replace('/','\\'))[:-1]
        imgname = originimgpath[originimgpath.rindex('\\')+1:][:-4]
        xizopath = r'C:\Windows\_XizoDiskIconChanger\icon\%s_%s.ico' % \
        (imgname[:-4], datetime.datetime.now().strftime('[%Y-%m-%d](%H.%M.%S)'))
    except:
        errimg = 1
    try:
        inputimg = Image.open(originimgpath)
        imgh, imgw = inputimg.size
        if imgh >= 48 and imgw >= 48:
            inputimg.save(xizopath, sizes=[(48, 48)])
        else:
            inputimg.save(xizopath, sizes=[(imgh, imgw)])
    except:
        print()
    try:
        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + choose_disk_combobox.get()[0] + '\DefaultIcon')
    except:
        errdisk = 1
    try:
        winreg.SetValue(winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + choose_disk_combobox.get()[0] + '\DefaultIcon',
        winreg.REG_SZ, xizopath)
        msg.showinfo('安装图标', '通过修改注册表安装图标成功')
    except:
        if errimg == 1 and errdisk == 1:
            msg.showerror('安装图标', '错误：你既没有选择磁盘也没有选择图片')
        if errimg == 1 and errdisk == 0:
            msg.showerror('安装图标', '错误：你没有选择图片')
        if errimg == 0 and errdisk == 1:
            msg.showerror('安装图标', '错误：你没有选择磁盘')

#autorun.inf方式安装图标
def infinstall():
    errimg = 0
    errdisk = 0
    try:
        os.remove(choose_disk_combobox.get()[0] + r":\autorun.inf")
    except:
        print()
    try:
        originimgpath = str(choose_img_bar.get(1.0, END).replace('/','\\'))[:-1]
        inputimg = Image.open(originimgpath)
    except:
        errimg = 1
    try:
        imgh, imgw = inputimg.size
        if imgh >= 48 and imgw >= 48:
            inputimg.save(choose_disk_combobox.get()[0]+r":\icon.ico", sizes=[(48, 48)])
        else:
            inputimg.save(choose_disk_combobox.get()[0]+r":\icon.ico", sizes=[(imgh, imgw)])
    except:
        print()
    try:
        inf = open(choose_disk_combobox.get()[0]+r":\autorun.inf", "w+")
    except:
        errdisk = 1
    try:
        inf.write("[autorun]\nicon=icon.ico")
        inf.close()
    except:
        print()
    try:
        win32api.SetFileAttributes(choose_disk_combobox.get()[0] + r":\icon.ico", win32con.FILE_ATTRIBUTE_HIDDEN)
    except:
        print()
    try:
        win32api.SetFileAttributes(choose_disk_combobox.get()[0] + r":\autorun.inf", win32con.FILE_ATTRIBUTE_HIDDEN)
    except:
        print()
    if errimg == 1 and errdisk == 1:
        msg.showerror('inf安装', '错误：你既没有选择磁盘也没有选择图片')
    if errimg == 1 and errdisk == 0:
        msg.showerror('inf安装', '错误：你没有选择图片')
    if errimg == 0 and errdisk == 1:
        msg.showerror('inf安装', '错误：你没有选择磁盘')
    if errimg == 0 and errdisk == 0:
        msg.showinfo('inf安装', '通过添加autorun.inf安装图标成功\n重启电脑或重新连接移动驱动器生效')

#恢复所有硬盘图标并清除缓存
def delete():
    for diski in disk_list:
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE,
            'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + diski[0] + '\DefaultIcon')
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE,
            'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons\\' + diski[0])
        except:
            print()
        try:
            os.remove(diski[0] + r":\icon.ico")
        except:
            print()
        try:
            os.remove(diski[0] + r":\autorun.inf")
        except:
            print()
    try:
        rmtree(r'C:\Windows\_XizoDiskIconChanger\icon')
    except:
        print()
    msg.showinfo('彻底还原清缓存', '已恢复所有硬盘图标并清除缓存')

#重启资源管理器
def restartexp():
    os.system("taskkill /f /im explorer.exe & start explorer.exe")

#修复网络驱动器
def fix():
    if os.path.exists(r'C:\Windows\_XizoDiskIconChanger\NetDeskFixed') == False:
        if msg.askyesno('修复网络驱动器','稍后会弹出一个“注册表编辑器”窗口，\n请确认安装注册表\n点击“是”开始'):
            regtmp = open("regtmp.reg", "wb+")
            regtmp.write(base64.b64decode(regfix))
            regtmp.close()
            win32api.ShellExecute(0, 'open', 'regtmp.reg', '', '', 1)
            if msg.askyesno('fix','成功添加后点击是\n然后手动重启电脑 即可完成修复'):
                os.remove("regtmp.reg")
                os.makedirs(r'C:\Windows\_XizoDiskIconChanger\NetDeskFixed')
    else:
        msg.showerror('修复网络驱动器', '你已经修复过了')

# 获取盘符
def getdisk():
    global disk_list
    kernel32 = ctypes.windll.kernel32
    volumeNameBuffer = ctypes.create_unicode_buffer(1024)
    fileSystemNameBuffer = ctypes.create_unicode_buffer(1024)
    serial_number = None
    max_component_length = None
    file_system_flags = None
    disk_list = []
    for diski in string.ascii_uppercase:
        disk = diski + ":\\"
        if os.path.isdir(disk):
            rc = kernel32.GetVolumeInformationW(
                ctypes.c_wchar_p(disk),
                volumeNameBuffer,
                ctypes.sizeof(volumeNameBuffer),
                serial_number,
                max_component_length,
                file_system_flags,
                fileSystemNameBuffer,
                ctypes.sizeof(fileSystemNameBuffer)
            )
            disk_list.append('%s: %s' % (diski, volumeNameBuffer.value))
            try:
                choose_disk_combobox['values'] = disk_list
            except:
                print()

if __name__ == "__main__":

    getdisk()

    #主窗口
    root = tk.Tk()
    root.title("硬盘图标修改器  V1.1  by  Xizo")

    #窗口位置大小
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    width = 750
    height = 450
    x = int((screenwidth - width) / 2)
    y = int((screenheight - height) / 2)
    root.geometry("%dx%d+%d+%d" % (width, height, x, y))
    root.resizable(0, 0)

    #图标
    icontmp = open("icontmp.ico", "wb+")
    icontmp.write(base64.b64decode(imgicon))
    icontmp.close()
    root.iconbitmap('icontmp.ico')
    os.remove('icontmp.ico')

    #背景
    bgtmp = open("bgtmp.png", "wb+")
    bgtmp.write(base64.b64decode(imgbg))
    bgtmp.close()
    bgphoto = ImageTk.PhotoImage(Image.open('bgtmp.png'))
    background = Label()
    background.pack()
    background['image'] = bgphoto
    os.remove('bgtmp.png')

    #Menubars
    menubar = tk.Menu(root)
    menubar.add_command(label="使用方法", command=usehelp)
    menubar.add_command(label="关于作者", command=about)
    menubar.add_command(label="退出程序", command=root.quit)
    root.config(menu=menubar)

    #选择磁盘下拉框
    choose_disk_combobox = ttk.Combobox(root, justify=tk.CENTER, height=10, width=11,
    state="readonly", font=("微软雅黑", 10), values=disk_list)

    #选择图标地址框
    choose_img_bar = tk.Text(root, width=17, height=2,
    state="disabled", font=("微软雅黑", 10))
    choose_img_bar.configure(state='normal')
    choose_img_bar.insert("insert", "      请选择图片文件→")
    choose_img_bar.configure(state='disabled')

    #Buttons
    choose_img_btn = tk.Button(root, text="打开\n图片", command=choose_img, font=("微软雅黑", 8))
    refresh_btn = tk.Button(root, text="刷新", command=getdisk, font=("微软雅黑", 8))
    rest_btn = tk.Button(root, text="  重置图标  ", command=rest, fg='black', bg='pink', font=("微软雅黑", 10))
    install_btn = tk.Button(root, text="  安装图标  ", command=install, fg='black', bg='light green', font=("微软雅黑", 10))
    inf_install_btn = tk.Button(root, text="inf安装", command=infinstall, fg='black', bg='yellow', font=("微软雅黑", 8))
    restartexplore_btn = tk.Button(root, text="重启资源管理器", command=restartexp, fg='black', font=("微软雅黑", 10))
    fix_btn = tk.Button(root, text="修复网络驱动器", command=fix, fg='black', font=("微软雅黑", 10))
    delete_btn = tk.Button(root, text="彻底还原清缓存", command=delete, font=("微软雅黑", 10))

    #Place elements
    choose_disk_combobox.place(x=121, y=201)
    refresh_btn.place(x=235, y=200)
    choose_img_bar.place(x=95, y=250)
    choose_img_btn.place(x=235, y=248)
    rest_btn.place(x=44, y=323)
    install_btn.place(x=190, y=323)
    inf_install_btn.place(x=134, y=323,height=33)
    restartexplore_btn.place(x=310, y=233)
    fix_btn.place(x=310, y=278)
    delete_btn.place(x=310, y=323)

    #预览小图标
    ShowImage = Label()
    ShowImage.place(x=44, y=245)

    #检测文件夹判断初次运行
    if os.path.exists(r'C:\Windows\_XizoDiskIconChanger') == False:
        os.makedirs(r'C:\Windows\_XizoDiskIconChanger')
        usehelp()

    root.mainloop()
