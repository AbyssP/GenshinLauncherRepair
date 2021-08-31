import os, winshell
import time
from win32com.client import Dispatch

PORT = 8080

s = None
thread = []
stop = False
gp = None


def run():
    def create_s(launch_after=False):
        print('start create' + f'{launch_after}')
        desktop = winshell.desktop()
        path = os.path.join(desktop, "Genshin Impact.lnk")
        target = gp + '\launcher.exe'
        wDir = gp
        icon = gp + '\launcher.exe'
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wDir
        shortcut.IconLocation = icon
        shortcut.save()
        if launch_after:
            print('create launch')
            launch_p()

    def create_launch():
        create_s(True)
        launch_p()

    def launch_p():
        global gp
        os.startfile(gp + '\launcher.exe')

    def open_dir():
        global gp
        os.startfile(gp)

    def check(app, label, path):
        time.sleep(2)
        for each_p in path:
            label['text'] = 'Checking for launcher path...\n' + each_p
            app.update()
            if os.path.exists(each_p) and os.path.isfile(each_p + '\launcher.exe'):
                label['text'] = 'Launcher found.\n' + each_p + '\\launcher.exe'
                app.update()
                return True
            time.sleep(.2)
        label['text'] = 'Launcher not found.\nPlease download and install a new launcher from\nhttps://genshin.mihoyo.com/en/download'
        app.update()
        return False

    path = []
    path.append('C:\Program Files\Genshin Impact')
    path.append('C:\Program Files (x86)\Genshin Impact')
    path.append('D:\Program Files\Genshin Impact')
    path.append('D:\Program Files (x86)\Genshin Impact')
    path.append('E:\Program Files\Genshin Impact')
    path.append('E:\Program Files (x86)\Genshin Impact')

    import tkinter as tk
    app = tk.Tk()
    app.geometry('350x290')
    app.title('Genshin Impact Launcher Repair')

    for each_p in path:
        if os.path.exists(each_p) and os.path.isfile(each_p + '\launcher.exe'):
            global gp
            gp = each_p
            label = tk.Label(app)
            label['text'] = 'Checking for launcher path...'
            label.place(relx=.5, y=10, anchor=tk.N)
            app.update()
            res =check(app, label, path)
            if not res:
                break

            add_shortcut_lunch_btn = tk.Button(app)
            add_shortcut_lunch_btn['text'] = 'Create a desktop shortcut\nand Launch'
            add_shortcut_lunch_btn.place(relx=.5, y=50, anchor=tk.N, width=300, heigh=50)
            add_shortcut_lunch_btn['command'] = create_launch

            add_shortcut_btn = tk.Button(app)
            add_shortcut_btn['text'] = 'Create a desktop shortcut'
            add_shortcut_btn.place(relx=.5, y=105, anchor=tk.N, width=300, heigh=50)
            add_shortcut_btn['command'] = create_s

            launch_btn = tk.Button(app)
            launch_btn['text'] = 'Launch Genshin launcher'
            launch_btn.place(relx=.5, y=160, anchor=tk.N, width=300, heigh=50)
            launch_btn['command'] = launch_p

            open_dir_btn = tk.Button(app)
            open_dir_btn['text'] = 'Open file directory'
            open_dir_btn.place(relx=.5, y=215, anchor=tk.N, width=300, heigh=50)
            open_dir_btn['command'] = open_dir
            app.mainloop()
            break
    app.mainloop()

if __name__ == '__main__':
    run()