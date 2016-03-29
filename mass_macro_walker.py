import os, win32com.client
'''
Python Script: Mass Macro Runner
By: Matthew Gary Evans
Date: March 27, 2016

ABOUT: This script loops through all Excel files in the same folder, runs macros in each file, and saves each file with a new name. The user is prompted to enter the path to the
directory with the files in it, and a prefix to add to the file name to indicate that it has been processed by the script.

This project was made for use with .ssx files. .ssx files are proprietary files created by a company that manufactures CPR manikins that collect data about the CPR session for
the purpose of training. The data is downloadable from the manikin in the form of a .ssx file. The files can be opened by converting them to zip files, which can then
be extracted. The CPR session data are stored in xml files which can be opened with Excel.

This script and the macros were created to assist in a medical research project. The macros in this script parse the xml data and put them in a new worksheet in a format specified
by the research director of the project. The project contained 1,369 files that needed to be parsed using the following script.
'''

#prompt user for the path to the directory with the Excel files
print( 'Enter the path to the directory with the Excel files in it (e.g., C:\\path\\to\\your\\folder):' )
root = input()

#prompt the user to provide a prefix to be added to the files when saved after the script is done running
print( 'Enter the prefix you wish to give each file name once it has been processed (e.g., if you choose a prefix of myPrefix_ a file named myFile.xls will be saved as myPrefix_myFile.xls; if you don\'t enter anything, the file will be saved as \'COPY_OF_myFile.xls\'):')
done_prefix = input()

#assign default value to done_prefix if user didn't enter anything, otherwise use user's input
if done_prefix == '':
        done_prefix = 'COPY_OF_'
else:
        done_prefix = done_prefix

#dispatch Excel; add a workbook with a module with a macro to delete
#the macro once done running in each file; this is necessary to avoid an
#ambiguous name error because otherwise this script will recreate the other
#macros with the same name in the same workbook every time it loops
xl = win32com.client.Dispatch( "Excel.Application" )
wb = xl.Workbooks.Add()
xlmodule = wb.VBProject.VBComponents.Add(1)
code = '''
Sub DeleteModule()
    ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents("Module2")   
End Sub'''
xlmodule.CodeModule.AddFromString( code )

#loop through all Excel files in the directory and run the macros
#the macros could be anything
for foldername, subfolders, filenames in os.walk( root ):
        for filename in filenames:
                if not filename.startswith( done_prefix ) and not os.path.exists( root + "\\" + done_prefix + filename):
                        currentFile = root + "\\" + filename
                        print( 'Completed: ' + filename )
                        xlmodule2 = wb.VBProject.VBComponents.Add(1)
                        code = '''
Sub Open_XML_Table()
'
' this script opens an excel file as an xml table
'

'
    ChDir "{srcDir}"
    Workbooks.OpenXML Filename:="{name}", _
        LoadOption:=xlXmlLoadImportToList
End Sub
'''.format(srcDir = root, name = currentFile)
                        xlmodule2.CodeModule.AddFromString( code )
                        xl.Application.Run( "Open_XML_Table" )
                        code = '''
Sub Vent_Check()

' Vent_Check Macro
' @Author:  Matthew Gary Evans
' @Date: March 13, 2016
' @About: this script checks the CPR Events sheet for the value of ventCount; if that value is 0, the script
' calls the CprVentEvents script, otherwise it calls the CPREvents script.
' Those scripts parse data from the CPREvents sheets and pastes them in Sheet2.
    Application.ScreenUpdating = False
    Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3
    Range("A1").Select
    Cells.Find(What:="ventCount", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Select
    Dim ventCount As Integer
    ventCount = Val(ActiveCell.Value)
    If ventCount > 0 Then
        Call CprVentEvents
    Else
        Call CPREvents
    End If
        Application.ScreenUpdating = True
        
End Sub

Sub CprVentEvents()
'
' CprVentEvents Macro
'
' @Author: Matthew Gary Evans
' @Date: March 13, 2016
' @About: this script finds data in Sheet1 for ID, ManikinName, msecs, compDepth, compLeaningDepth,
' compReleaseDepth, compMeanRate, ventVolume, compInactivity, and ventDuringCompression, and pastes them in Sheet2

    'Part 1: turn off screen updating; make sure Sheet1 is the active sheet and clear filters
    Application.ScreenUpdating = False
    Sheets("Sheet1").Select
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3

'   Part 2: Prep the new sheet with column headers
    Sheets("Sheet2").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "startTime"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "ManikinName"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "msecs"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "compDepth"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "compLeaningDepth"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "compReleaseDepth"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "compMeanRate"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "ventVolume"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "compInactivity"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Epoch"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "comp_greater_20"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "comp_greater_40"
        
    'Part 3: clear filters and sort by msecs and type2 fields, ascending, in that order
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add _
        Key:=Range("Table1[msecs]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add _
        Key:=Range("Table1[type2]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Part 4: find the first cell in the msecs column that is not 0
    Range("B2").Select
    While Val(ActiveCell.Value) = 0
        ActiveCell.Offset(1, 0).Select
    Wend
    
    'Part 5: get the values for the comp and vent fields for the current msec value
    'variable for current value of msec
    Dim currentMsec As Long
        
    'get range of currently active cell
    Dim returnCell As String
    returnCell = ActiveCell.Address
    
    'variables for values to be pasted in Sheet2
    Dim compDepth As Double
    Dim compLeaningDepth As Double
    Dim compReleaseDepth As Double
    Dim compMeanRate As Double
    Dim ventVolume As Double
    Dim compInactivity As Double
    
    'currentRow is the row currently being written to in Sheet2; nextRow is the next unused row in Sheet2
    Dim currentRow As Integer
        
    Dim nextRow As Integer
    
    'loop through all active rows of Sheet1
    While ActiveCell.Row <= Sheets("Sheet1").UsedRange.Rows.Count
        returnCell = ActiveCell.Address
        currentMsec = Val(ActiveCell.Value)
        'go to sheet2 and paste the current value of currentMsec in the next unused row of column D
        Sheets("Sheet2").Select
        nextRow = Sheets("Sheet2").UsedRange.Rows.Count + 1
        Range("D" & nextRow).Select
        ActiveCell.Value = currentMsec
        currentRow = Sheets("Sheet2").UsedRange.Rows.Count
        
        'go back to Sheet1 and get the values for cols E-J in Sheet 2
        Sheets("Sheet1").Select
        Range("" & returnCell & "").Select
    
        'loop through all rows with the same msec value
        Do While Val(ActiveCell.Value) = currentMsec
            If ActiveCell.Offset(0, 1).Value = "compDepth" Then
                compDepth = Val(ActiveCell.Offset(0, 2).Value)
                Sheets("Sheet2").Select
                Range("E" & currentRow).Select
                ActiveCell.Value = compDepth
                Sheets("Sheet1").Select
                Range("" & returnCell & "").Select
                
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
                
            ElseIf ActiveCell.Offset(0, 1).Value = "compLeaningDepth" Then
                compLeaningDepth = Val(ActiveCell.Offset(0, 2).Value)
                Sheets("Sheet2").Select
                Range("F" & currentRow).Select
                ActiveCell.Value = compLeaningDepth
                Sheets("Sheet1").Select
                Range("" & returnCell & "").Select
                
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
                
            ElseIf ActiveCell.Offset(0, 1).Value = "compReleaseDepth" Then
                compReleaseDepth = Val(ActiveCell.Offset(0, 2).Value)
                Sheets("Sheet2").Select
                Range("G" & currentRow).Select
                ActiveCell.Value = compReleaseDepth
                Sheets("Sheet1").Select
                Range("" & returnCell & "").Select
                
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
                
            ElseIf ActiveCell.Offset(0, 1).Value = "compMeanRate" Then
                compMeanRate = Val(ActiveCell.Offset(0, 2).Value)
                Sheets("Sheet2").Select
                Range("H" & currentRow).Select
                ActiveCell.Value = compMeanRate
                Sheets("Sheet1").Select
                Range("" & returnCell & "").Select
                
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
                
            ElseIf ActiveCell.Offset(0, 1).Value = "ventVolume" Then
                ventVolume = Val(ActiveCell.Offset(0, 2).Value)
                Sheets("Sheet2").Select
                Range("I" & currentRow).Select
                ActiveCell.Value = ventVolume
                Sheets("Sheet1").Select
                Range("" & returnCell & "").Select
                
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
                
            ElseIf ActiveCell.Offset(0, 1).Value = "compInactivity" Then
                compInactivity = Val(ActiveCell.Offset(0, 2).Value)
                Sheets("Sheet2").Select
                Range("J" & currentRow).Select
                ActiveCell.Value = compDepth
                Sheets("Sheet1").Select
                Range("" & returnCell & "").Select
                
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
                
            Else
                If Val(ActiveCell.Offset(1, 0).Value) <> currentMsec Then
                    Exit Do
                Else
                    ActiveCell.Offset(1, 0).Select
                    returnCell = ActiveCell.Address
                End If
            End If
        Loop
        
        ActiveCell.Offset(1, 0).Select
        
    Wend
        
    'Part 6: write formula for epoch field
    'epoch 1 is the range of milliseconds from 0 to 30,000
    '30,000 ms < epoch 2 <= 60,000 ms
    '60,000 ms < epoch 3 <= 90,000 ms
    'epoch 4 > 90,000 ms
    Sheets("sheet2").Select
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-7]<=30000,1,IF(RC[-7]<=60000,2,IF(RC[-7]<=90000,3,4)))"
    ActiveCell.Copy
    Range("K2:K" & Sheets("Sheet2").UsedRange.Rows.Count).Select
    ActiveSheet.Paste
    
    'Part 7: add the formula for computing the comp_greater_20 values in the new worksheet
    'if compDepth >= 20, comp_greater_20 = 1, else comp_greater_20 = 0
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]>=20,1,0)"
    ActiveCell.Copy
    Range("L2:L" & Sheets("Sheet2").UsedRange.Rows.Count).Select
    ActiveSheet.Paste
    
    'Part 8: add the formula for computing the comp_greater_60 values
    'if compDepth >= 40, comp_greater_40 = 1, else 0
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-8]>=40,1,0)"
    ActiveCell.Copy
    Range("M2:M" & Sheets("Sheet2").UsedRange.Rows.Count).Select
    ActiveSheet.Paste
    
    'Part 9: get the student ID and startTime and paste in new worksheet
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "students"
    Cells.Find(What:="students", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A2").Select
    ActiveSheet.Paste

    'get startTime
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "startTime"
    Cells.Find(What:="startTime", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("B2").Select
    ActiveSheet.Paste

    'Part 10: get the manikin name and paste in new worksheet
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "manikinName"
    Cells.Find(What:="manikinName", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("C2").Select
    ActiveSheet.Paste
    
    'Part 11: calculate/write the ventRate (# of ventVolume entries in the data sheet / 120,000
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "ventVolume"
    Cells.Find(What:="ventVolume", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Dim ventVolCount As Integer
    ventVolCount = Application.CountIf(Selection, "ventVolume")
    Sheets("Sheet2").Select
    Range("N1").Select
    ActiveCell.Value = "ventRate"
    ActiveCell.Offset(1, 0).Value = ventVolCount / 120000
    Selection.NumberFormat = "General"
    
    'Part 12: undo the filters and sort criteria on Sheet1
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
    
    'Part 13: clean up the formatting in the new worksheet
    'paste the ID
    Sheets("Sheet2").Select
    Range("A2").Select
    ActiveCell.Copy
    Range("A2:A" & Sheets("Sheet2").UsedRange.Rows.Count).Select
    ActiveSheet.Paste
    
    'paste the startTime
    Range("B2").Select
    ActiveCell.Copy
    Range("B2:B" & Sheets("Sheet2").UsedRange.Rows.Count).Select
    ActiveSheet.Paste
    
    'paste the ManikinName
    Range("C2").Select
    ActiveCell.Copy
    Range("C2:C" & Sheets("Sheet2").UsedRange.Rows.Count).Select
    ActiveSheet.Paste
    
    'make sure the width of the columns is big enough for the content
    Columns("A:N").EntireColumn.AutoFit
    Range("A1").Select
            
    'Notify user that the script is done running
    'MsgBox ("Work complete.")
    Application.ScreenUpdating = True
End Sub

Sub CPREvents()
'
' CPREvents Macro
'
' @Author: Matthew Gary Evans
' @Date: March 12, 2016
' @About: this script finds data in Sheet1 for ID, ManikinName, msecs, compDepth, compLeaningDepth,
' compReleaseDepth, and compMeanRate and pastes it in Sheet2

    'Part 1: turn off screen updating
    Application.ScreenUpdating = False
    
    'Part 2: make sure Sheet1 is the active sheet and clear filters
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
    
'   Part 3: Prep the new sheet with column headers
    Sheets("Sheet2").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "startTime"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "ManikinName"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "msecs"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "compDepth"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "compLeaningDepth"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "compReleaseDepth"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "compMeanRate"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Epoch"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "comp_greater_50"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "comp_greater_60"
    
    'Part 4: go back to the sheet with data
    Sheets("Sheet1").Select
    
    'Part 5: get the student ID and startTime and paste in new worksheet
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "students"
    Cells.Find(What:="students", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    'get startTime
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "startTime"
    Cells.Find(What:="startTime", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("B2").Select
    ActiveSheet.Paste
    
    'Part 6: get the manikin name and paste in new worksheet
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "manikinName"
    Cells.Find(What:="manikinName", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("C2").Select
    ActiveSheet.Paste
    
    'Part 7: get the compDepth msecs and paste in new worksheet
    Sheets("Sheet1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "compDepth"
    Cells.Find(What:="compDepth", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("D2").Select
    ActiveSheet.Paste
    'the number of rows with data not including the header row
    Dim numDataRows As Integer
    numDataRows = Selection.Count
        
    'Part 8:get the compDepth values and paste in new worksheet
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "compDepth"
    Cells.Find(What:="compDepth", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("E2").Select
    ActiveSheet.Paste
        
    'Part 9:get the compLeaningDepth values and paste in new worksheet
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "compLeaningDepth"
    Cells.Find(What:="compLeaningDepth", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("F2").Select
    ActiveSheet.Paste
        
    'Part 10: get the compReleaseDepth values and paste in new worksheet
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "compReleaseDepth"
    Cells.Find(What:="compReleaseDepth", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("G2").Select
    ActiveSheet.Paste
        
    'Part 11: get the compMeanRate values and paste in new worksheet
    Sheets("Sheet1").Select
    Range("A1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "compMeanRate"
    Cells.Find(What:="CprCompEvent", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("H2").Select
    ActiveSheet.Paste
        
    'Part 12: add the formula for computing the epoch to the new worksheet
    'epoch 1 is the range of milliseconds from 0 to 30,000
    '30,000 ms < epoch 2 <= 60,000 ms
    '60,000 ms < epoch 3 <= 90,000 ms
    'epoch 4 > 90,000 ms
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]<=30000,1,IF(RC[-5]<=60000,2,IF(RC[-5]<=90000,3,4)))"
    ActiveCell.Copy
    Range("I2:I" & numDataRows + 1).Select
    ActiveSheet.Paste
    
    'Part 13: loop through used data cells in cols C-G and convert the text to numbers
    'the forumals in the next 2 parts will not work without this
    For Each Cell In Range("D2:H" & numDataRows + 1).Cells
        Cell.Value = Val(Cell)
    Next
    
    'Part 14:add the formula for computing the comp_greater_50 values in the new worksheet
    'if compDepth >= 50, comp_greater_50 = 1, else comp_greater_50 = 0
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]>=50,1,0)"
    ActiveCell.Copy
    Range("J2:J" & numDataRows + 1).Select
    ActiveSheet.Paste
    
    'Part 15: add the formula for computing the comp_greater_60 values
    'if compDepth >= 60, comp_greater_60 = 1, else 0
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-6]>=60,1,0)"
    ActiveCell.Copy
    Range("K2:K" & numDataRows + 1).Select
    ActiveSheet.Paste
    
    'Part 16: clean up the formatting in the new worksheet
    'paste the ID
    Range("A2").Select
    ActiveCell.Copy
    Range("A2:A" & numDataRows + 1).Select
    ActiveSheet.Paste
    
    'paste the startTime
    Range("B2").Select
    ActiveCell.Copy
    Range("B2:B" & numDataRows + 1).Select
    ActiveSheet.Paste
    
    'paste the ManikinName
    Range("C2").Select
    ActiveCell.Copy
    Range("C2:C" & numDataRows + 1).Select
    ActiveSheet.Paste
    
    'make sure the width of the columns is big enough for the content
    Columns("A:K").EntireColumn.AutoFit
    Range("A1").Select
    

        Application.ScreenUpdating = True
End Sub'''
                        xlmodule2.CodeModule.AddFromString( code )
                        macro = wb.Name + "!Vent_Check"
                        xl.Application.Run( macro )
                        #save the file, close the workbook, and then run the macro that deletes the macros above( Subs_1, 2 and 3 in this demo )
                        xl.ActiveWorkbook.SaveAs( root + "\\" + done_prefix + filename )
                        xl.ActiveWorkbook.Close()
                        xl.Application.Run( "DeleteModule" )
                        
print( "Work complete!" )
print( "Press enter to exit.")

#this is necessary to keep the CLI from automatically closing once the script is done
finish = input()
