Attribute VB_Name = "ExcelIO"
'===========================================================
' SW Point Importer
'
' Module    : ExcelIO
' Revision  : 1 (Draft)
'
'===========================================================

Option Explicit

Private Const xlUp As Long = -4162

Public Function BrowseExcelFile() As Boolean

    On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim dlg As Object

    Set xlApp = CreateObject("Excel.Application")

    Set dlg = xlApp.FileDialog(3)    'msoFileDialogFilePicker

    With dlg

        .Title = "Select Coordinate Excel Workbook"

        .AllowMultiSelect = False

        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xlsx;*.xlsm;*.xls"

        If .Show <> -1 Then

            BrowseExcelFile = False

            GoTo CleanExit

        End If

        gExcelFile = .SelectedItems(1)

    End With

    BrowseExcelFile = True

CleanExit:

    On Error Resume Next

    xlApp.Quit

    Set dlg = Nothing
    Set xlApp = Nothing

    Exit Function

ErrorHandler:

    MsgBox _
        "Unable to open file picker." & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    BrowseExcelFile = False

    Resume CleanExit

End Function

Public Function ReadExcelCoordinates() As Boolean

    On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object

    Dim LastRow As Long
    Dim RowIndex As Long

    Dim pt As clsPoint

    Set xlApp = CreateObject("Excel.Application")

    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlBook = xlApp.Workbooks.Open(gExcelFile)

    Set xlSheet = xlBook.Worksheets(1)

    LastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(xlUp).Row

    If LastRow < 1 Then

        MsgBox "Worksheet contains no coordinate data.", vbCritical

        GoTo CleanFail

    End If

    Set gPoints = New Collection

    For RowIndex = 1 To LastRow

        If Trim$(CStr(xlSheet.Cells(RowIndex, 1).Value)) = "" Then Exit For

        If Trim$(CStr(xlSheet.Cells(RowIndex, 2).Value)) = "" Then

            MsgBox "Blank Y value found at row " & RowIndex, vbCritical

            GoTo CleanFail

        End If

        If Not IsNumeric(xlSheet.Cells(RowIndex, 1).Value) Then

            MsgBox "Invalid X value at row " & RowIndex, vbCritical

            GoTo CleanFail

        End If

        If Not IsNumeric(xlSheet.Cells(RowIndex, 2).Value) Then

            MsgBox "Invalid Y value at row " & RowIndex, vbCritical

            GoTo CleanFail

        End If

        Set pt = New clsPoint

        pt.ID = "P" & CStr(RowIndex)

        pt.XInch = CDbl(xlSheet.Cells(RowIndex, 1).Value)
        pt.YInch = CDbl(xlSheet.Cells(RowIndex, 2).Value)

        pt.X = pt.XInch * INCH_TO_METER
        pt.Y = pt.YInch * INCH_TO_METER

        gPoints.Add pt

    Next RowIndex

    xlBook.Close False
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    ReadExcelCoordinates = True

    Exit Function

CleanFail:

    On Error Resume Next

    xlBook.Close False
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    ReadExcelCoordinates = False

    Exit Function

ErrorHandler:

    MsgBox _
        "Excel Error" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    On Error Resume Next

    xlBook.Close False
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    ReadExcelCoordinates = False

End Function
