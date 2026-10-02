Attribute VB_Name = "SketchTools"
'===========================================================
' SW Point Importer
'
' Module    : SketchTools
' Revision  : 1 (Draft)
'
' Compatible with:
' SolidWorks 2024
'
'===========================================================

Option Explicit

Public Function CreateSketchPoints() As Boolean

    On Error GoTo ErrorHandler

Dim pt As clsPoint
Dim swPt As SldWorks.SketchPoint

Dim dblX As Double
Dim dblY As Double

    '-------------------------------------------------------
    ' Speed up sketch creation
    '-------------------------------------------------------

    swSketchMgr.AddToDB = True
    swSketchMgr.DisplayWhenAdded = False

    '-------------------------------------------------------
    ' Create every point
    '-------------------------------------------------------

    For Each pt In gPoints

        dblX = gOriginX + pt.X
dblY = gOriginY + pt.Y

Set swPt = swSketchMgr.CreatePoint(dblX, dblY, 0#)

        If swPt Is Nothing Then

            MsgBox _
                "Failed to create sketch point: " & pt.ID, _
                vbCritical

            GoTo CleanFail

        End If

        Set pt.SketchPoint = swPt

    Next pt

    '-------------------------------------------------------
    ' Restore normal sketch behavior
    '-------------------------------------------------------

    swSketchMgr.AddToDB = False
    swSketchMgr.DisplayWhenAdded = True

    swModel.EditRebuild3
    swModel.GraphicsRedraw2

    CreateSketchPoints = True

    Exit Function

CleanFail:

    swSketchMgr.AddToDB = False
    swSketchMgr.DisplayWhenAdded = True

    CreateSketchPoints = False

    Exit Function

ErrorHandler:

    swSketchMgr.AddToDB = False
    swSketchMgr.DisplayWhenAdded = True

    MsgBox _
        "Sketch Error" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    CreateSketchPoints = False

End Function
