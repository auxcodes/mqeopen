# GNU GENERAL PUBLIC LICENSE
#
# Copyright (C) 2015 MQEopen - Phill Banks - https://github.com/Phill-B/mqeopen
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
# You should have received a copy of the GNU General Public License along
# with this program; if not, write to the Free Software Foundation, Inc.,
# 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#
from Tkinter import *
import ttk
import tkFileDialog
import tkMessageBox
import win32api
import os
import pickle
import glob


class Application(Frame):
    def __init__(self, parent):
        # frame instance
        Frame.__init__(self, parent)
        self.parent = parent
        self.grid(column=0, row=0, sticky=(N, W, E))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        # menu instances
        self.menubar = Menu(self)
        self.filemenu = Menu(self.menubar, tearoff=0)
        self.helpmenu = Menu(self.menubar, tearoff=0)
        # path variables
        self.dirpath = StringVar()
        self.dirpath.trace("w", self.on_path_trace)
        self.longest = IntVar()
        self.configpath = StringVar()
        self.wrkpath = StringVar()
        # widget instances
        self.mqeLabelFrame = LabelFrame(self)
        self.dirLabel = ttk.Label(self.mqeLabelFrame)
        self.valCombo = []
        self.dirList = ttk.Combobox(self.mqeLabelFrame, value=self.valCombo, textvariable=self.dirpath)
        self.browse = Button(self.mqeLabelFrame)
        self.QUIT = Button(self.mqeLabelFrame)
        self.Open = Button(self.mqeLabelFrame)
        # clean directory of MQE crash files
        self.clean_dir()
        # check required ancillary files exist
        self.check_filesexist()
        self.read_save()
        # setup menu and widgets
        self.create_menu()
        self.create_widgets()

    def create_menu(self):
        # add menu title File
        self.menubar.add_cascade(label="Options", menu=self.filemenu)
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)
        self.parent.config(menu=self.menubar)
        # add submenu options
        self.filemenu.add_command(label="Display MQE Paths", command=self.display_mqewcfg)
        self.filemenu.add_command(label="Change MQE Install Path", command=self.open_cfgchange)
        self.filemenu.add_command(label="Change MQE Workspace Path", command=self.open_wrkchange)
        self.filemenu.add_command(label="Clear Workspace History", command=self.clear_combolist)
        # add help menu
        self.helpmenu.add_command(label="About", command=self.open_about)

    def create_widgets(self):
        # ttk styling
        style = ttk.Style()
        style.configure("BW.TLabel", background="white")
        # MQ Explorer Label Frame
        self.mqeLabelFrame["text"] = "MQ Explorer"
        self.mqeLabelFrame.grid(row=0, column=0, sticky=(N, E, W),
                                padx=5, pady=5, ipadx=5, ipady=5)
        self.mqeLabelFrame.columnconfigure(1, weight=1)
        self.mqeLabelFrame.rowconfigure(0, weight=1)

        # Directory selected Label and Entry
        self.dirLabel["text"] = "Select Workspace:"

        # self.dirLabel["style"] = "BW.TLabel"
        self.dirLabel.grid(row=0, column=0, padx=5, pady=5, sticky='W')

        # load existing paths from path config file
        self.read_mqewcfg()
        # load existing workspace list from save pickle file
        self.read_save()

        self.dirList.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky='WE')

        # Dir Dialog Button
        self.browse["text"] = "Browse"
        self.browse["command"] = self.open_dirbrowse
        self.browse.grid(row=0, column=4, padx=5, pady=5, sticky='W')

        # Quit button
        self.QUIT["text"] = "QUIT"
        self.QUIT["fg"] = "red"
        self.QUIT["command"] = self.parent.destroy
        self.QUIT.grid(row=2, column=2, padx=5, pady=5, sticky='W')

        # Open button
        self.Open["text"] = "Open"
        self.Open["fg"] = "dark green"
        self.Open["command"] = self.check_workspace
        self.Open.grid(row=2, column=3, padx=5, pady=5, sticky='W')

    def open_dirbrowse(self):
        # Directory dialog
        dirname = tkFileDialog.askdirectory(parent=self.parent, initialdir=self.dirpath.get(),
                                            title='Please select a directory')
        if len(dirname) > 0:
            # if a new path is selected update the path
            self.dirpath.set(dirname)
        else:
            # if cancel is selected pass back the original path
            self.dirpath.set(self.dirpath.get())

    def open_cfgchange(self):
        # Directory dialog
        dirname = tkFileDialog.askdirectory(parent=self.parent, initialdir=self.configpath.get(),
                                            title='Please select a directory')
        if len(dirname) > 0:
            # if a new path is selected update the path
            self.configpath.set(dirname)
            self.write_mqewcfg()
            self.read_mqewcfg()
        else:
            # if cancel is selected pass back the original path
            self.configpath.set(self.configpath.get())

    def open_wrkchange(self):
        # Directory dialog
        dirname = tkFileDialog.askdirectory(parent=self.parent, initialdir=self.wrkpath.get(),
                                            title='Please select a directory')
        if len(dirname) > 0:
            # if a new path is selected update the path
            self.wrkpath.set(dirname)
            self.write_mqewcfg()
            self.read_mqewcfg()
        else:
            # if cancel is selected pass back the original path
            self.wrkpath.set(self.wrkpath.get())

    @staticmethod
    def open_about():
        tkMessageBox.showinfo(title="About",
                              message='GNU GENERAL PUBLIC LICENSE\n\n'
                                      'MQEopen - https://github.com/Phill-B/mqeopen\n\n'
                                      'Copyright (C) 2015 - Phill Banks\n\n'
                                      'This program is free software; you can redistribute it and/or modify '
                                      'it under the terms of the GNU General Public License as published by '
                                      'the Free Software Foundation; either version 2 of the License, or '
                                      '(at your option) any later version.\n\n'
                                      'This program is distributed in the hope that it will be useful,'
                                      'but WITHOUT ANY WARRANTY; without even the implied warranty of '
                                      'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the '
                                      'GNU General Public License for more details.'
                                      'You should have received a copy of the GNU General Public License along '
                                      'with this program; if not, write to the Free Software Foundation, Inc., '
                                      '51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.')

    def check_workspace(self):
        # Check workspace path is a valid MQ Explorer workspace
        if os.path.isdir(self.dirpath.get() + "/.metadata") is False:
            if tkMessageBox.askyesno(title="Create Workspace",
                                     message='The workspace you have selected is not an ' +
                                             'existing MQ Explorer workspace\n\n' +
                                             'Attempted to open path:\n' + self.dirpath.get() + "\n\n" +
                                             'Would you like MQ Explorer to create a new workspace here?'):
                self.edit_config()
                return
            else:
                return
        # workspace looks valid edit config file
        self.edit_config()

    def edit_config(self):
        # Check config file path
        if os.path.isfile(self.configpath.get() + '/configuration/config.ini'):
            with open(self.configpath.get() + "/configuration/config.ini") as f:
                file_str = f.readlines()
        else:
            tkMessageBox.showerror(title='Missing config.ini',
                                   message='I was unable to open config.ini!\n\n' +
                                           'I attempted to open path:\n' + self.configpath.get() +
                                           '/configuration/config.ini\n\n' +
                                           'Please check your MQ Explorer install path.')
            return
        # default config text 4th line of config.ini
        # osgi.instance.area=@user.home/IBM/WebSphereMQ/workspace
        match = "osgi.instance.area"
        newtext = "osgi.instance.area=" + self.dirpath.get() + "\n"
        # find the unwanted path in the config file
        for t in file_str:
            if match in t:
                # replace the old path with the new
                file_str[file_str.index(t)] = newtext
        # write back to the file
        with open(self.configpath.get() + "/configuration/config.ini", "w") as f:
            f.writelines(file_str)
            f.close()

        # after all that attempt to open MQ Explorer
        self.open_mqe()

    def open_mqe(self):
        # open mq explorer with the newly selected workspace **requires win32api extension
        if os.path.isfile(self.configpath.get() + '/MQExplorer.exe'):
            win32api.WinExec(self.configpath.get() + '/MQExplorer.exe')
        else:
            tkMessageBox.showerror(title='MQ Explorer Error',
                                   message='I was unable to open MQ Explorer!\n\n' +
                                           'I attempted to open path:\n' + self.configpath.get() +
                                           '/MQExplorer.exe\n\n' + 'Please check your MQ Explorer install path.', )
            return
        # check to see if this workspace has been used before
        if self.dirList.get() in self.valCombo:
            return
        else:
            self.valCombo.append(self.dirList.get())
            pickle.dump(self.valCombo, open("save.p", "wb"))
            self.read_save()

    def read_mqewcfg(self):
        # reading values from path config file and setting values 
        if os.path.isfile("mqewcfg.ini"):
            with open("mqewcfg.ini") as f:
                file_str = f.readlines()
                # set paths from info in config file
            self.configpath.set(file_str[0].rstrip())
            self.dirpath.set(file_str[1].rstrip())
            self.wrkpath.set(file_str[1].rstrip())
            f.close()
        else:
            with open("mqewcfg.ini", "w+") as f:
                # appdata = os.getenv('APPDATA')
                file_str = ['C:/Program Files (x86)/IBM/WebSphere MQ Explorer\n',
                            '..\\MQ Workspaces']
                f.writelines(file_str)
                f.close()
            # Call myself after creating file to set values
            self.read_mqewcfg()

    def read_save(self):
        # read saved workspaces
        if os.path.isfile("save.p"):
            self.valCombo = pickle.load(open("save.p", "rb"))
            self.dirList["value"] = self.valCombo
            self.longest = max(len(s) for s in self.valCombo)
        else:
            with open("save.p", "w+") as f:
                f.close()
            self.valCombo = [self.dirpath.get(), ]
            pickle.dump(self.valCombo, open("save.p", "wb"))

    def display_mqewcfg(self):
        # display current path config settings
        tkMessageBox.showinfo(title='Workspace Tool Path Settings',
                              message='Current MQ Explorer directory paths are set to:\n\n' +
                                      'Workspace path: ' + self.wrkpath.get() + '\n\n' +
                                      'Install path: ' + self.configpath.get() + '\n\n')

    def write_mqewcfg(self):
        # write new values to the path config file
        if os.path.isfile("mqewcfg.ini"):
            with open("mqewcfg.ini", "w") as f:
                file_str = [self.configpath.get() + "\n", self.wrkpath.get()]
                f.writelines(file_str)
                f.close()
                return
        else:
            # files doesn't exist create it and call myself to write new values
            with open("mqewcfg.ini", "w+") as f:
                f.close()
            # call myself
            self.write_mqewcfg()

    def clear_combolist(self):
        # clear stored list of workspaces
        self.valCombo = [self.dirpath.get(), ]
        pickle.dump(self.valCombo, open("save.p", "wb"))
        self.read_save()
        self.dirList["width"] = self.longest

    def on_path_trace(self, *args):
        # resize workspace list for readability
        if len(self.dirpath.get()) > self.longest:
            self.dirList["width"] = len(self.dirpath.get())
        else:
            self.dirList["width"] = self.longest

    def check_filesexist(self):
        # check for path config file
        if os.path.isfile("mqewcfg.ini"):
            return
        else:
            with open("mqewcfg.ini", "w+") as f:
                # appdata = os.getenv('APPDATA')
                file_str = ['C:/Program Files (x86)/IBM/WebSphere MQ Explorer\n',
                            '..\\MQ Workspaces']
                f.writelines(file_str)
        # check pickle file exists
        if os.path.isfile("save.p"):
            self.valCombo = pickle.load(open("save.p", "rb"))
        else:
            with open("save.p", "w+") as f:
                f.close()
            self.valCombo = [self.dirpath.get(), ]
            pickle.dump(self.valCombo, open("save.p", "wb"))

    @staticmethod
    def clean_dir():
        heapdump = glob.glob("heapdump*")
        javacore = glob.glob("javacore*")
        snap = glob.glob("Snap*")

        if len(heapdump) > 0:
            for i in heapdump:
                os.remove(i)

        if len(javacore) > 0:
            for i in javacore:
                os.remove(i)

        if len(snap) > 0:
            for i in snap:
                os.remove(i)
