import wx
from wx.lib.mixins.listctrl import CheckListCtrlMixin, ListCtrlAutoWidthMixin
import numpy as np
import os.path as path
import xlrd
import xlwt

class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin, ListCtrlAutoWidthMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, wx.ID_ANY, style=wx.LC_REPORT |
                wx.SUNKEN_BORDER)
        CheckListCtrlMixin.__init__(self)
        ListCtrlAutoWidthMixin.__init__(self)

class Sampler(wx.Frame):
    
    def __init__(self, *args, **kw):
        # ensure the parent's __init__ is called
        super(Sampler, self).__init__(*args, **kw)

        # Global vars
        self.fileName = ''
        self.filePath = ''
        self.isFileOpen = False

        # Sampling options
        self.sampleSize = 10
        self.sampleRate = 10
        self.sampleMode = 1 # 1: fixed size 2: fixed rate 3: max of 1&2
        self.nEntries = 0
        self.indexColumn = 0

        # Save options
        self.columnSelection = []

        # create UI
        self.panel = wx.Panel(self)
        self.initUI()
        self.SetSize(450, 600)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

    def initUI(self):
        
        # main box
        mainBox = wx.BoxSizer(wx.VERTICAL)

        # button box
        btnBox1 = wx.BoxSizer(wx.HORIZONTAL)

        self.btnOpenFile = wx.Button(self.panel, label='选择文件')
        self.btnOpenFile.Bind(wx.EVT_BUTTON, self.OnOpenFile)
        btnBox1.Add(self.btnOpenFile, proportion=1, flag=wx.ALL | wx.EXPAND, border=5)
        
        self.btnSave = wx.Button(self.panel, label='导出文件')
        self.btnSave.Bind(wx.EVT_BUTTON, self.OnSaveFile)
        btnBox1.Add(self.btnSave, proportion=1, flag=wx.ALL | wx.EXPAND, border=5)

        btnClose = wx.Button(self.panel, label='关闭文件')
        btnClose.Bind(wx.EVT_BUTTON, self.OnCloseFile)
        btnBox1.Add(btnClose, proportion=1, flag=wx.ALL | wx.EXPAND, border=5)

        btnExit = wx.Button(self.panel, label='退出')
        btnExit.Bind(wx.EVT_BUTTON, self.OnExit)
        btnBox1.Add(btnExit, proportion=1, flag=wx.ALL | wx.EXPAND, border=5)

        # sample options
        optBox1 = wx.BoxSizer(wx.HORIZONTAL)

        optBox2 = wx.BoxSizer(wx.VERTICAL)
        self.checkBox1 = wx.CheckBox(self.panel, label='固定数量:')
        self.checkBox1.SetValue(True)
        self.checkBox2 = wx.CheckBox(self.panel, label='固定比例(%):')
        self.spinCtrl1 = wx.SpinCtrl(self.panel, value='%d' % self.sampleSize)
        self.spinCtrl1.SetRange(0, 500)
        self.spinCtrl2 = wx.SpinCtrl(self.panel, value='%d' % self.sampleRate)
        self.spinCtrl2.SetRange(0, 100)
        optBox21 = wx.BoxSizer(wx.HORIZONTAL)
        optBox22 = wx.BoxSizer(wx.HORIZONTAL)
        optBox21.Add(self.checkBox1, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
        optBox21.Add(self.spinCtrl1, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
        optBox22.Add(self.checkBox2, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
        optBox22.Add(self.spinCtrl2, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
        optBox2.Add(optBox21, proportion=1)
        optBox2.Add(optBox22, proportion=1)
        optBox1.Add(optBox2, proportion=2)

        txt='''如果两个都选，则按
较大数量抽取'''
        st1 = wx.StaticText(self.panel, label=txt, style=wx.ALIGN_CENTRE)
        optBox1.Add(st1, proportion=1, flag=wx.EXPAND)

        # selection box
        selectBox = wx.BoxSizer(wx.VERTICAL)
        st2 = wx.StaticText(self.panel, label='请在下表选择需要导出的数据', style=wx.ALIGN_LEFT)
        self.listCtrl = CheckListCtrl(self.panel)
        self.listCtrl.InsertColumn(0, '可导出的列')
        self.listCtrl.InsertColumn(1, '列序号')
        self.listCtrl.setResizeColumn(0)
        selectBox.Add(st2, proportion=1, flag=wx.LEFT | wx.EXPAND, border=20)
        selectBox.Add(self.listCtrl, proportion=5, flag=wx.ALL | wx.EXPAND, border=10)

        mainBox.Add(btnBox1, proportion=1, flag=wx.ALL | wx.EXPAND, border=10)
        mainBox.Add(optBox1, proportion=2, flag=wx.EXPAND | wx.TOP, border=20)
        mainBox.Add(selectBox, proportion=4, flag=wx.EXPAND)
        self.panel.SetSizer(mainBox)


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
        if not self.isFileOpen:
            return

        indices = self.Sampling()
        self.GetColumnSelection()
        
        # create workbook
        result = xlwt.Workbook()
        sheet1 = result.add_sheet('Sheet1')

        # formats
        formatDate = xlwt.easyxf(num_format_str='yyyy.mm.dd')
        formatNum = xlwt.easyxf(num_format_str='0')

        sheet1.write(0, 0, '序号')
        targetIdx = 1
        for sourceIdx in self.columnSelection: 
            cell = self.workSheet.row(0)[sourceIdx]
            sheet1.row(0).write(targetIdx, cell.value)
            targetIdx += 1

        for i in range(0, len(indices)):
            targetRow = sheet1.row(i+1)
            sourceRow = self.workSheet.row(indices[i])
            targetRow.write(0, i+1)
            targetIdx = 1
            for sourceIdx in self.columnSelection:
                cell = sourceRow[sourceIdx]
                if cell.ctype == 0:
                    pass
                elif cell.ctype == 2:
                    targetRow.write(targetIdx, cell.value, formatNum) 
                elif cell.ctype == 3:
                    targetRow.write(targetIdx, cell.value, formatDate) 
                else :
                    targetRow.write(targetIdx, cell.value)
                targetIdx += 1
        
        result.save(path.join(self.filePath, '抽样结果.xls'))
        wx.MessageBox('导出完成', '', wx.OK | wx.ICON_INFORMATION)

    def OnCloseFile(self, event):
        self.isFileOpen = False
        self.btnOpenFile.SetBackgroundColour(wx.NullColour)
        self.listCtrl.DeleteAllItems()
        self.workSheet = None

    def LoadData(self):
        self.workSheet = xlrd.open_workbook(self.fileName).sheet_by_index(0)
        self.isFileOpen = True
        self.nEntries = self.workSheet.nrows - 1

        row = self.workSheet.row(0)
        for idx, cell in enumerate(row):
            if cell.value == '出院科别描述':
                # identify the department column
                self.indexColumn = idx
        
            # set the column selection box
            index = self.listCtrl.InsertItem(idx, cell.value)
            self.listCtrl.SetItem(index, 1, str(idx))
            self.listCtrl.CheckItem(idx)
    
    def Sampling(self):
         # sampling
        dpts = self.workSheet.col_values(self.indexColumn) 
        indices = []

        flag1 = self.checkBox1.IsChecked()
        if flag1:
            self.sampleSize = int(self.spinCtrl1.GetValue())
        flag2 = self.checkBox2.IsChecked()
        if flag2:
            self.sampleRate = int(self.spinCtrl2.GetValue())

        start = 1
        while dpts[start] == '':
            start += 1

        while start < self.nEntries:
            end = start
            while end < self.nEntries and dpts[start] == dpts[end]:
                end += 1

            if flag1 and flag2:
                size = end - start
                samplesize2 = int(size * self.sampleRate / 100)

                if samplesize2 < self.sampleSize:
                    step = max(int((end - start) / self.sampleSize), 1)
                    for i in range(self.sampleSize):
                        val = start + i * step
                        if val >= end:
                            break
                        indices.append(val)
                else:
                    step2 = int(100 / self.sampleSize)
                    for i in range(samplesize2):
                        val = start + i * step2
                        if val >= end:
                            break
                        indices.append(val)
            elif flag1:
                step = max(int((end - start) / self.sampleSize), 1)
                for i in range(self.sampleSize):
                    val = start + i * step
                    if val >= end:
                        break
                    indices.append(val)
            elif flag2:
                size = end - start
                samplesize2 = int(size * self.sampleRate / 100)
                step2 = int(100/self.sampleRate)
                for i in range(samplesize2):
                    val = start + i * step2
                    if val >= end:
                        break
                    indices.append(val)
            start = end
        return indices

    def GetColumnSelection(self):
        self.columnSelection.clear()
        colCount = self.listCtrl.GetItemCount()
        for idx in range(colCount):
            if self.listCtrl.IsChecked(idx):
                self.columnSelection.append(idx)
        


if __name__ == '__main__':
    app = wx.App()
    frm = Sampler(None, title='病例抽取')
    frm.Show()
    app.MainLoop()