import xlwings  # minor refactoring can remove this dependency
import pandas  # future refactoring to use dask is preferred
from pandasql import sqldf
import win32com.client
import os
import shutil
import re
from numpy import *

import win32com.client
import sys
PowerPoint = win32com.client.Dispatch("PowerPoint.Application")
Excel = win32com.client.Dispatch("Excel.Application")


class template:

    @staticmethod
    def relativePaths(input, output):
        dir = os.path.dirname(__file__).replace('/', '\\')
        relativeFile = os.path.join(dir, input)
        relativeCwd = os.path.join(os.getcwd(), input)
        relativeFileOutput = os.path.join(dir, output)
        relativeCwdOutput = os.path.join(os.getcwd(), output)
        return {
            'input': (relativeCwd if os.path.exists(relativeCwd) else relativeFile),
            'output': (relativeCwdOutput if os.path.exists(relativeCwd) else relativeFileOutput)
        }

    @classmethod
    def excel(self, input, output, replace):
        Excel = win32com.client.Dispatch('Excel.Application')
        # copy file
        relative = self.relativePaths(input, output)
        shutil.copy(relative['input'], relative['output'])
        # open file
        wb = Excel.Workbooks.Open(relative['output'])
        # replace each item
        for key in replace:
            if '!' not in key:
                wb.Names(key).RefersToRange.Value = replace[key]
            else:
                sheet, cell = key.split('!')
                wb.Sheets(sheet).Range(cell).Value = replace[key]
        # recalculate
        wb.RefreshAll()
        wb.Sheets.Select
        wb.Sheets(1).Calculate  # maybe need to be []
        # save file
        wb.Close(True)
        Excel = None
        return relative['output']

    @classmethod
    def powerpoint(self, input, output, replace, link=''):
        PowerPoint = win32com.client.Dispatch('PowerPoint.Application')
        # copy file
        relative = self.relativePaths(input, output)
        shutil.copy(relative['input'], relative['output'])
        # open file
        pres = PowerPoint.Presentations.Open(relative['input'])
        # replace each item
        for key in replace:
            value = str(replace[key])
            pattern = re.compile(key, re.IGNORECASE)
            for sld in pres.Slides:
                for shp in sld.Shapes:
                    if shp.HasTextFrame:
                        if shp.TextFrame.HasText:
                            originalText = shp.TextFrame.TextRange.Text
                            changedText = pattern.sub(value, originalText)
                            shp.TextFrame.TextRange.Text = changedText
                        if shp.HasTable:
                            for i in range(1, shp.Table.Rows.Count):
                                for j in range(1, shp.Table.Columns.Count):
                                    originalText = shp.Table.Rows.Item(i).Cells(
                                        j).Shape.TextFrame.TextRange.Text
                                    changedText = pattern.sub(
                                        value, originalText)
                                    shp.Table.Rows.Item(i).Cells(
                                        j).Shape.TextFrame.TextRange.Text = changedText

        # replace links
        msoChart = 3
        opAuto = 2
        opMan = 1
        if len(link) > 0:
            for sld in pres.Slides:
                for shp in sld.Shapes:
                    if shp.Type == msoChart:
                        shp.LinkFormat.SourceFullName = link
                        shp.LinkFormat.AutoUpdate = opAuto
                        shp.LinkFormat.Update
                        shp.LinkFormat.AutoUpdate = opMan

        # save file
        pres.Save()
        pres.Close()
        pres = None
        return relative['output']


def sql(q): return sqldf(q, globals())


def consolidate(folder, sheet_name, skiprows=0):
    import fnmatch
    import os
    import pandas
    import types

    dir = os.path.dirname(__file__).replace("/", "\\")
    folderRelFile = os.path.join(dir, folder)
    folderRelCwd = os.path.join(os.getcwd(), folder)
    folder = folderRelCwd if os.path.isdir(folderRelCwd) else folderRelFile

    folderPath = folder.rsplit("\\", 1)[0]
    filePattern = folder.rsplit("\\", 1)[1]
    output = pandas.DataFrame()

    for file in os.listdir(folderPath):
        if fnmatch.fnmatch(file, filePattern):
            filepath = os.path.join(folderPath, file)
            xldf = pandas.read_excel(
                filepath, sheet_name=sheet_name, skiprows=skiprows)
            xldf['source'] = file
            output = output.append(xldf)

    def save(self, filename):
        filename = os.path.join(dir, filename)
        self.to_excel(filename)

    output.save = types.MethodType(save, output)
    return output

#test = consolidate("consolidate\*.*", "Sheet1")


class tableConnection(object):

    def __init__(self, workbookPath=None):
        if workbookPath == None:
            self.wb = xlwings
        else:
            self.wb = xlwings.Book(workbookPath)

    def __getitem__(self, tableName):
        self.getTable(tableName)
        if self.sheet is None:
            return pandas.DataFrame()

        if self.table is None:
            df = self.sheet.range('A1').options(
                pandas.DataFrame,
                header=1,
                index=False,
                expand='table').value
        else:
            df = pandas.DataFrame(self.sheet.range(self.table.Range.Address).options(headers=True).value)
            df = df.rename(columns=df.iloc[0]).drop(df.index[0])

        # coverts decimal types to floats so that they work with numpy as expected
        return df.apply(pandas.to_numeric, errors='ignore').reset_index(drop=True)

    def __setitem__(self, tableName, data):
        try:
            self.getTable(tableName)
            self.createTable(tableName)
            data = pandas.DataFrame(data)
            proposedHeaders = data.columns.tolist()
            existingHeaders = [
                x for x in self.table.HeaderRowRange.Value[0]][0:len(proposedHeaders)]
            Excel.DisplayAlerts = False
            Excel.ScreenUpdating = False
            Excel.EnableEvents = False

            if existingHeaders == proposedHeaders:
                # check if table is already consistent with existing dataframe
                oldLastRow = self.table.Range.Cells(
                    1, 1).Row + self.table.Range.Rows.Count
                self.resize(data.shape[0]+1, self.table.HeaderRowRange.Count)
                newLastRow = self.table.Range.Cells(
                    1, 1).Row + self.table.Range.Rows.Count
                if(newLastRow < oldLastRow):
                    self.table.Range.Worksheet.Range(
                        str(newLastRow) + ":" + str(oldLastRow)).Delete()

                for i in range(0, len(proposedHeaders)):
                    self.table.DataBodyRange.Columns(i+1).ClearContents()
                    self.table.DataBodyRange.Columns(
                        i+1).value = [[val] for val in data[data.columns[i]]]

            else:
                # clear old data and build new table
                oldLastRow = self.table.Range.Cells(
                    1, 1).Row + self.table.Range.Rows.Count
                oldLastCol = self.table.Range.Cells(
                    1, 1).Column + self.table.Range.Columns.Count
                self.resize(data.shape[0]+1, data.shape[1])
                newLastRow = self.table.Range.Cells(
                    1, 1).Row + self.table.Range.Rows.Count
                newLastCol = self.table.Range.Cells(
                    1, 1).Column + self.table.Range.Columns.Count
                if(newLastRow < oldLastRow):
                    self.table.Range.Worksheet.Range(
                        str(newLastRow) + ":" + str(oldLastRow)).Delete()
                if(newLastCol < oldLastCol):
                    [self.table.Range.Worksheet.Columns(
                        i).Delete() for i in range(newLastCol, oldLastCol)]

                for i in range(0, len(proposedHeaders)):
                    self.table.HeaderRowRange.Columns(
                        i+1).Value = data.columns[i]
                    if self.table.DataBodyRange is None:
                        self.table.ListRows.Add()
                    self.table.DataBodyRange.Columns(i+1).ClearContents()
                    self.table.DataBodyRange.Columns(
                        i+1).value = [[val] for val in data[data.columns[i]]]

                if self.table.DataBodyRange is not None and self.table.Range.Cells.Rows.Count < 1000:
                    leftAligned = -4131  # VBA constant value for left-aligned
                    self.table.Range.Cells.HorizontalAlignment = leftAligned
        finally:
            if self.sheet is not None:
                Excel.DisplayAlerts = True
                Excel.ScreenUpdating = True
                Excel.EnableEvents = False

    def getTable(self, tablename):
         # create table and associated worksheet if it does not exist
        wb = self.wb
        self.tablename = tablename
        t = None
        s = None
        for sh in wb.sheets:
            for li in sh.api.ListObjects:
                if li.Name == tablename or sh.name == tablename:
                    t = li
                    s = sh
            if s is None and sh.name == tablename:
                s = sh
        if s is not None and t is None:
            for li in s.api.ListObjects:
                t = li  # set the table to a list item on the selected sheet, needs refactor
        self.table = t  # this is a com listobject
        self.sheet = s  # this is an xlwings object, probably unnecessarily now that we're not using expand() anymore

    def createTable(self, tablename):
        if self.sheet is None:
            activeSh = xlwings.Range("A1").api.Worksheet
            self.sheet = self.wb.sheets.add(tablename)
            activeSh.Activate()
        if self.table is None:
            topleft = self.sheet.range("A1:A2").api
            self.table = topleft.Worksheet.ListObjects.Add(1, topleft)
            self.table.ListRows.Add()
            self.table.Name = tablename
            self.resize()

    def resize(self, rows=0, cols=0):
        topleft = self.table.Range.Cells(1, 1)
        if rows == 0 or cols == 0:
            # could throw an error but would prefer to just to execute a default behaviour
            bottomright = topleft.Offset(2, 1)
        else:
            bottomright = topleft.Offset(rows, cols)
        newRange = self.sheet.api.Range(topleft, bottomright)
        self.table.Resize(newRange)


def refreshPivots(): return xlwings.Range("A1").api.Worksheet.Parent.RefreshAll


table = tableConnection()
cell = xlwings.Range
sheet = table
this = Excel.ActiveSheet.Name

