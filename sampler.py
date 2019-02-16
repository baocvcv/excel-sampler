import wx
import numpy as np
import os.path as path
import xlrd
import xlwt

class Sampler(wx.Frame):
    
    def __init__(self, *args, **kw):
        # ensure the parent's __init__ is called
        super(Sampler, self).__init__(*args, **kw)

        # Data
        self.sampleSize = 10
        self.classNo = 0
        self.nEntries = 0

        self.fileName = ''
        self.filePath = ''
        
        self.isFileOpen = False

        # create UI
        self.panel = wx.Panel(self)

        self.initUI()
        self.SetSize(400, 200)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

    def initUI(self):
        
        box1 = wx.BoxSizer(wx.HORIZONTAL)

        self.btnOpenFile = wx.Button(self.panel, label='选择原始文件')
        self.btnOpenFile.Bind(wx.EVT_BUTTON, self.OnOpenFile)
        box1.Add(self.btnOpenFile, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
        
        self.btnSave = wx.Button(self.panel, label='导出文件')
        self.btnSave.Bind(wx.EVT_BUTTON, self.OnSaveFile)
        box1.Add(self.btnSave, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)

        btnExit = wx.Button(self.panel, label='退出')
        btnExit.Bind(wx.EVT_BUTTON, self.OnExit)
        box1.Add(btnExit, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)

        self.panel.SetSizer(box1)


    def OnExit(self, event):
        self.Destroy()

    def OnOpenFile(self, event):
        # show the open file dialog
        with wx.FileDialog(self, "选择文件", defaultDir="D:\Documents\Programming\Projects\excel-sampler", 
                        style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind

            # Proceed loading the file chosen by the user
            self.fileName = fileDialog.GetPath()
            self.filePath, name = path.split(path.abspath(self.fileName))

            # Check if file type is valid
            suffix = name.split('.')[1]
            if suffix == 'xls' or suffix == 'xlsx':
                self.LoadData()
                self.btnOpenFile.SetBackgroundColour(wx.GREEN)

    def OnSaveFile(self, event):
        # sampling
        dpts = self.workSheet.col_values(7)
        indices = []
        start = 1
        while dpts[start] == '':
            start += 1
        while start < self.nEntries:
            end = start
            while dpts[start] == dpts[end]:
                end += 1
            step = min(int((end - start) / self.sampleSize), 1)
            for i in range(10):
                val = start + i * step
                if val >= end:
                    break
                indices.append(val)
            start = end
        noSamples = len(indices)

        # create workbook
        result = xlwt.Workbook()
        sheet1 = result.add_sheet('Sheet1')

        sheet1.write(0, 0, '序号')
        for index, cell in enumerate(self.workSheet.row_slice(0, 0, self.workSheet.ncols-1)):
            sheet1.row(0).write(index+1, cell.value)

        for i in range(0, noSamples):
            row = sheet1.row(i+1)
            source = self.workSheet.row_slice(indices[i], 0, self.workSheet.ncols-1)
            row.write(0, i+1)
            for index, cell in enumerate(source):
                if cell.ctype == 3:
                    row.write(index+1, cell.value, xlwt.easyxf(num_format_str='YY.M.D'))
                else:
                    row.write(index+1, cell.value)
        
        result.save(path.join(self.filePath, '抽样结果.xls'))
        self.btnOpenFile.SetBackgroundColour(wx.NullColour)
        wx.MessageBox('导出完成', '', wx.OK | wx.ICON_INFORMATION)


    def LoadData(self):
        self.workSheet = xlrd.open_workbook(self.fileName).sheet_by_index(0)
        self.isFileOpen = True

        self.nEntries = self.workSheet.nrows - 1
    
if __name__ == '__main__':
    app = wx.App()
    frm = Sampler(None, title='病例抽取软件')
    frm.Show()
    app.MainLoop()