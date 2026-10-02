Attribute VB_Name = "Main"
'===========================================================
' SW Point Importer
'
' Module    : Main
' Revision  : 2
'
'===========================================================

Option Explicit

Public Sub Main()

    On Error GoTo ErrorHandler

    If Not InitializeGlobals() Then

        MsgBox _
        "Initialization failed." & vbCrLf & vbCrLf & _
        "Ensure:" & vbCrLf & _
        "• SolidWorks is running" & vbCrLf & _
        "• A Part document is active" & vbCrLf & _
        "• A 2D sketch is being edited", _
        vbCritical

        Exit Sub

    End If

    If Not BrowseExcelFile() Then GoTo CleanExit

If Not GetOriginCoordinates() Then GoTo CleanExit

If Not ReadExcelCoordinates() Then GoTo CleanExit

If Not BuildCoordinateGroups() Then GoTo CleanExit

If Not CreateSketchPoints() Then GoTo CleanExit

If Not CreateSketchRelations() Then GoTo CleanExit

    swModel.GraphicsRedraw2

    MsgBox _
        gPoints.Count & _
        " point(s) imported successfully.", _
        vbInformation

CleanExit:

    ClearGlobals

    Exit Sub

ErrorHandler:

    MsgBox _
        "Error " & Err.Number & vbCrLf & Err.Description, _
        vbCritical

    ClearGlobals

End Sub
