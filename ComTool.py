import tkinter
from tkinter import filedialog
import os
import re
import xlsxwriter

PATTERN = '[ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz~&.]'
COLOR_ERROR = '#b30000'
COLOR_GOOD = '#009933'

"""
GUI interface for this program
"""
class Gui:

    def __init__(self):
        self.frame = tkinter.Tk()

        self.lbInstruction = tkinter.Label(self.frame)
        self.lbOrgPath = tkinter.Label(self.frame)
        self.lbTarPath = tkinter.Label(self.frame)
        self.lbOutPath = tkinter.Label(self.frame)

        self.btSetOrgPath = tkinter.Button(self.frame)
        self.btSetTarPath = tkinter.Button(self.frame)
        self.btSetOutPath = tkinter.Button(self.frame)
        self.btOutputAction = tkinter.Button(self.frame)
        self.InitGui()

        self.core = None

    def InitGui(self):
        self.btSetOrgPath.config(text='Set original path', state=tkinter.DISABLED, command=self.SetOrgPath)
        self.btSetTarPath.config(text='Set target path', state=tkinter.DISABLED, command=self.SetTarPath)
        self.btSetOutPath.config(text='Set output path', state=tkinter.DISABLED, command=self.SetOutPath)
        self.btOutputAction.config(text='Output non-match table', state=tkinter.DISABLED, command=self.OutputAction)

        self.btSetOrgPath.pack()
        self.lbOrgPath.pack()

        self.btSetTarPath.pack()
        self.lbTarPath.pack()

        self.btSetOutPath.pack()
        self.lbOutPath.pack()

        self.btOutputAction.pack()

    def SetCore(self, core):
        assert (isinstance(core, Core))
        self.core = core
        self.btSetOrgPath.config(state=tkinter.NORMAL)
        self.btSetTarPath.config(state=tkinter.NORMAL)
        self.btSetOutPath.config(state=tkinter.NORMAL)

    def SetOrgPath(self):
        assert (isinstance(self.core, Core))
        path = filedialog.askdirectory(initialdir='./')
        self.lbOrgPath.config(text=path)
        if not self.core.SetOrgPath(path):
            print("Invalid input for original path, please try again")
            self.lbOrgPath.config(fg=COLOR_ERROR)
            return
        self.lbOrgPath.config(fg=COLOR_GOOD)
        if self.core.IsReadyToOutput():
            self.btOutputAction.config(state=tkinter.NORMAL)
        else:
            self.btOutputAction.config(state=tkinter.DISABLED)

    def SetTarPath(self):
        assert (isinstance(self.core, Core))
        path = filedialog.askdirectory(initialdir='./')
        self.lbTarPath.config(text=path)
        if not self.core.SetTarPath(path):
            print("Invalid input for original path, please try again")
            self.lbTarPath.config(fg=COLOR_ERROR)
            return
        self.lbTarPath.config(fg=COLOR_GOOD)
        if self.core.IsReadyToOutput():
            self.btOutputAction.config(state=tkinter.NORMAL)
        else:
            self.btOutputAction.config(state=tkinter.DISABLED)

    def SetOutPath(self):
        assert (isinstance(self.core, Core))
        path = filedialog.askdirectory(initialdir='./')
        self.lbOutPath.config(text=path)
        if not self.core.SetOutPath(path):
            print("Invalid input for original path, please try again")
            self.lbOutPath.config(fg=COLOR_ERROR)
            return
        self.lbOutPath.config(fg=COLOR_GOOD)
        if self.core.IsReadyToOutput():
            self.btOutputAction.config(state=tkinter.NORMAL)
        else:
            self.btOutputAction.config(state=tkinter.DISABLED)

    def OutputAction(self):
        print("Output")
        self.core.Output()

    def Run(self):
        self.frame.mainloop()


"""
Data structure to save file info.
"""
class FileInfo:

    def __init__(self, name, dirpath, key = 'N/A'):
        self.name = name
        self.dirpath = dirpath
        self.key = key

    def GetFileInfo(self):
        return self.name, self.dirpath

"""
Core class is the core of this program. It is responsible for read and compare the list
"""
class Core:
    def __init__(self):
        self.orgPath = None
        self.tarPath = None
        self.outPath = None
        """
        Both files hash table are used to save FileInfo. The key is pure num.
        """
        self.orgFiles = {}
        self.tarFiles = {}

    def SetOrgPath(self, path):
        """
        :param path: Root folder as a reference, should be a larger path.
        :return: bool
        """
        assert (isinstance(path, str))
        if not os.path.exists(path):
            return False
        self.orgPath = path
        return True

    def SetTarPath(self, path):
        """
        :param path: Root folder to be checked, should be a smaller path.
        :return: bool
        """
        assert (isinstance(path, str))
        if not os.path.exists(path):
            return False
        self.tarPath = path
        return True

    def SetOutPath(self, path):
        assert (isinstance(path, str))
        if not os.path.exists(path):
            return False
        self.outPath = path
        return True

    def ReadOrgFiles(self):
        self.orgFiles = {}
        for dirpath, dirname, filenames in os.walk(self.orgPath, topdown=True, onerror=None, followlinks=False):
            for filename in filenames:
                key = re.split(PATTERN, filename, 1)[0]
                if len(key) >= 1:
                    AddFileInfo(self.orgFiles, key, FileInfo(filename, dirpath, key))

    def ReadTarFiles(self):
        self.tarFiles = {}
        for dirpath, dirname, filenames in os.walk(self.tarPath, topdown=True, onerror=None, followlinks=False):
            for filename in filenames:
                key = re.split(PATTERN, filename, 1)[0]
                if len(key) >= 1:
                    AddFileInfo(self.tarFiles, key, FileInfo(filename, dirpath, key))

    def IsReadyToOutput(self):
        if (self.tarPath is None) or (self.orgPath is None) or (self.outPath is None):
            return False
        return os.path.exists(self.tarPath) and os.path.exists(self.orgPath) and (self.outPath is not None)

    def GetNonMatchList(self):
        self.ReadTarFiles()
        self.ReadOrgFiles()
        print('File keys in Original: ', len(self.orgFiles))
        print('File keys in Target: ', len(self.tarFiles))

        nonMatchFiles = []
        for key in self.orgFiles.keys():
            if key not in self.tarFiles:
                for fileInfo in self.orgFiles[key]:
                    nonMatchFiles.append(fileInfo)
        print('Non match numbers: ', len(nonMatchFiles))
        return nonMatchFiles

    def Output(self):
        ofiles = self.GetNonMatchList()
        workbook = xlsxwriter.Workbook(self.outPath + '/CompResult.xlsx')
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, 'File Name')
        worksheet.write(0, 1, 'File Path')
        worksheet.write(0, 2, 'Key')

        row = 1
        for info in ofiles:
            worksheet.write(row, 0, info.name)
            worksheet.write(row, 1, info.dirpath)
            worksheet.write(row, 2, info.key)
            row += 1

        workbook.close()


def AddFileInfo(dict, key, info):
    assert (isinstance(info, FileInfo))
    if key not in dict:
        dict[key] = [info]
    else:
        assert(isinstance(dict[key], list))
        dict[key].append(info)

if __name__ == '__main__':

    c = Core()
    """
    assert (c.SetOrgPath('./org'))
    assert (c.SetTarPath('./tar'))
    c.ReadOrgFiles()
    c.ReadTarFiles()

    print('OrgNum: ', len(c.orgFiles))
    print('TarNum: ', len(c.tarFiles))

    nonMatchFiles = []
    for key in c.orgFiles.keys():
        if key not in c.tarFiles:
            for fileInfo in c.orgFiles[key]:
                nonMatchFiles.append(fileInfo)
    workbook = xlsxwriter.Workbook('non-match.xlsx')
    worksheet = workbook.add_worksheet()

    # Write Title
    worksheet.write(0, 0, 'File Name')
    worksheet.write(0, 1, 'File Path')

    row = 1
    for fileInfo in nonMatchFiles:
        worksheet.write(row, 0, fileInfo.name)
        worksheet.write(row, 1, fileInfo.dirpath)
        row += 1
        
    workbook.close()
    """
    launcher = Gui()
    launcher.SetCore(c)
    launcher.Run()




