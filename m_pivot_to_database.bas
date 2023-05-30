Attribute VB_Name = "PivotToDatabase"
Option Explicit

' In this project, I extract data from multiple files in pivot form and add those data to a database.
' In the process, I create a backup of all files and move the extracted files from a new files directory to a used files directory.

    Dim sDataPath As String, sRootPath As String, sNewFilesPath As String, sUsedFilesPath As String, sBackupPath As String
    Const constDataDir As String = "01_Data", constNewFilesDir As String = "02_NewFiles", constUsedFilesDir As String = "03_UsedFiles", constBackupDir As String = "04_Backup"
    Dim r As Long, c As Long, i As Long
    
    Dim fso As Object
    
    Dim sFile As String, sExtension As String 'MergeData
    Dim arrData() As Variant 'DePivot
    Dim lngRecords As Long 'DePivot
    Dim iColumns As Integer 'DePivot


Sub ExecuteProgram()

    Call ExcelActionsOff
    Call MultipleSubVariables
    Call BackUp
    Call HandleFiles
    ''' Call DePivot ' is called in MergeData procedure
    ''' Call DataIntoDatabase ' is called in MergeData procedure
    Call JoinData
    Call NewToUsedFiles
    Call ExcelActionsOn
    
End Sub


Sub MultipleSubVariables()
' initialize variables that are used in multiple procedures

    ' initialize the relevant directories
    sDataPath = ThisWorkbook.Path & "\"
    sRootPath = Left(sDataPath, VBA.Strings.Len(sDataPath) - VBA.Strings.Len(constDataDir) - 2) & "\"
    sNewFilesPath = sRootPath & constNewFilesDir & "\"
    sUsedFilesPath = sRootPath & constUsedFilesDir & "\"
    sBackupPath = sRootPath & constBackupDir & "\"
        
    ' create object that moves files and directories around
    Set fso = CreateObject("Scripting.FileSystemObject")

End Sub


Sub HandleFiles()
' new files are opened, handed over to the data-extraction procedures, and closed again
 
    ' find all files in the newFiles directory
    sExtension = "*.xlsx"
    sFile = Dir(sNewFilesPath & sExtension, vbNormal)
   
    Do While sFile <> ""

        ' open workbook
        Workbooks.Open (sNewFilesPath & "\" & sFile)
        
        ' code to bring the pivot data from all workbooks to the database
        Call DePivot
        Call DataIntoDatabase
    
        ' close workbook
        Workbooks(sFile).Close

        ' next workbook's name in directory
        sFile = Dir
    Loop

End Sub


Sub DePivot()
' the data in pivot format is extracted to an array in database format

    Dim iYear As Integer
    Dim lngRows As Long
    Dim bytSheets As Byte
    
    ' basics and initialization
    Const bytHeaderRow As Byte = 3
    Const bytAttributes As Byte = 4 'attributes are 1. company, 2. year, 3. month, 4. value
      
    ' determine the number of records to dim the array
    lngRows = Cells(Rows.Count, 1).End(xlUp).Row - bytHeaderRow
    iColumns = Cells(bytHeaderRow, Columns.Count).End(xlToLeft).Column - 1
    bytSheets = Worksheets.Count
    lngRecords = lngRows * iColumns * bytSheets
    
    ' redim the array
    ReDim arrData(1 To lngRecords, 1 To bytAttributes)
    
    ' determine the year
    iYear = CInt(Mid(sFile, InStr(1, sFile, "_") + 1, 4))
    
    ' assign the pivot table's values to the array
    For i = 1 To bytSheets
        Worksheets(i).Activate
        For r = 1 To lngRows
            For c = 1 To iColumns
                ' fill in the firm
                arrData(r + lngRows * (c - 1) + (lngRows * iColumns) * (i - 1), 1) = Cells(bytHeaderRow + r, 1)
                ' fill in the year
                arrData(r + lngRows * (c - 1) + (lngRows * iColumns) * (i - 1), 2) = iYear
                ' fill in the month
                arrData(r + lngRows * (c - 1) + (lngRows * iColumns) * (i - 1), 3) = Cells(bytHeaderRow, 1 + c)
                ' fill in the value
                arrData(r + lngRows * (c - 1) + (lngRows * iColumns) * (i - 1), 4) = Cells(bytHeaderRow + r, 1 + c)
            Next c
        Next r
    Next i

End Sub


Sub DataIntoDatabase()
' the array's values are handed over to the database

    Dim lngEntryRow As Long
    Dim rngFirstCell As Range
    Dim rngEntryRange As Range

    ThisWorkbook.Activate
    
    ' determine first data entry range
    lngEntryRow = Cells(Rows.Count, 2).End(xlUp).Row + 1
    Set rngFirstCell = Cells(lngEntryRow, 2)
    Set rngEntryRange = Range(rngFirstCell, Cells(rngFirstCell.Row + UBound(arrData, 1) - 1, 2 + UBound(arrData, 2) - 1))
    
    ' insert the data
    rngEntryRange = arrData
    
End Sub


Sub JoinData()
' the region data is joined to the database

    Dim baseSh As Worksheet, baseRowFirst As Long, baseRowLast As Long, baseRngSearch As Range, baseRngTarget As Range
    Dim lookupSh As Worksheet, lookupRowLast As Long, lookupRng As Range, lookupRngSearch As Range

    Const baseColSearch As Integer = 2
    Const baseColTarget As Integer = 1
    Const lookupFile As String = "LookupTable"
    Const lookupColSearch As Integer = 1
    Const lookupColTarget As Integer = 2
    Const lookupRowFirst As Long = 4
    
    ' open lookup workbook
    Workbooks.Open (sDataPath & lookupFile & ".xlsx")
 
    ' initialize the ranges of the database
    ThisWorkbook.Activate
    Set baseSh = ThisWorkbook.Worksheets(1)
    baseRowFirst = Cells(Rows.Count, baseColTarget).End(xlUp).Row + 1
    baseRowLast = Cells(Rows.Count, baseColSearch).End(xlUp).Row
    Set baseRngSearch = baseSh.Range(Cells(baseRowFirst, baseColSearch), Cells(baseRowLast, baseColSearch))
    Set baseRngTarget = baseSh.Range(Cells(baseRowFirst, baseColTarget), Cells(baseRowLast, baseColTarget))
    
    ' initialize the ranges of the lookup file
    Workbooks(lookupFile).Activate
    Set lookupSh = Workbooks(lookupFile).Worksheets(1)
    lookupRowLast = Cells(Rows.Count, 1).End(xlUp).Row
    Set lookupRng = lookupSh.Range(Cells(lookupRowFirst, lookupColSearch), Cells(lookupRowLast, lookupColTarget))
    Set lookupRngSearch = lookupSh.Range(Cells(lookupRowFirst, lookupColSearch), Cells(lookupRowLast, lookupColSearch))
    
    ' do the lookup using index-match
    For r = baseRowFirst To baseRowLast
        baseSh.Cells(r, baseColTarget).Value = WorksheetFunction.Index(lookupRng, _
                                                WorksheetFunction.Match(baseSh.Cells(r, baseColSearch), lookupRngSearch, 0), _
                                                lookupColTarget)
    Next r
    
    ' finally close the workbook
    Workbooks(lookupFile & ".xlsx").Close (False)
    
End Sub


Sub BackUp()
' creates a backup directory just in case something breakes

    Dim sToday As String
    Dim sBackUpDir As String
    Dim sDataRoot As String, sDataDir As String
    Dim sNewFilesRoot As String, sNewFilesDir As String
    Dim sUsedFilesRoot As String, sUsedFilesDir As String

    ' initialize backup directory
    sToday = Format(Now, "YYYY-MM-DD") 'find date
    sBackUpDir = sRootPath & constBackupDir & "\" & sToday 'directory name
    
    ' create backup directory if it does not exist yet
    If Dir(sBackUpDir, vbDirectory) = VBA.Constants.vbNullString Then
        MkDir sBackUpDir
    End If
    
    ' backup data directory
    sDataRoot = sRootPath & constDataDir
    sDataDir = sBackUpDir & "\" & constDataDir
    If Dir(sDataDir, vbDirectory) = VBA.Constants.vbNullString Then
        MkDir sDataDir
    End If
    fso.CopyFolder sDataRoot, sDataDir
    
    ' backup NewFiles directory
    sNewFilesRoot = sRootPath & constNewFilesDir
    sNewFilesDir = sBackUpDir & "\" & constNewFilesDir
    If Dir(sNewFilesDir, vbDirectory) = VBA.Constants.vbNullString Then
        MkDir sNewFilesDir
    End If
    fso.CopyFolder sNewFilesRoot, sNewFilesDir
    
    ' backup UsedFiles directory
    sUsedFilesRoot = sRootPath & constUsedFilesDir
    sUsedFilesDir = sBackUpDir & "\" & constUsedFilesDir
    If Dir(sUsedFilesDir, vbDirectory) = VBA.Constants.vbNullString Then
        MkDir sUsedFilesDir
    End If
    fso.CopyFolder sUsedFilesRoot, sUsedFilesDir

End Sub


Sub NewToUsedFiles()
' moves the files with the extracted data to the used files directory

    Dim objFile As Object
    
    For Each objFile In fso.getfolder(sNewFilesPath).Files
        objFile.Move sUsedFilesPath
    Next objFile

End Sub


Sub ExcelActionsOff()
' turn off excel "events" that slow down the code

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With

End Sub


Sub ExcelActionsOn()
' turn the excel "events" back on for the usual excel experience
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = True
        .Calculation = xlCalculationAutomatic
        .CutCopyMode = False
    End With

End Sub

