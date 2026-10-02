Attribute VB_Name = "OriginManager"
'===========================================================
' SW Point Importer
'
' Module    : OriginManager
' Revision  : 2
'
'===========================================================

Option Explicit

Public Function GetSketchOriginPoint() As Boolean
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr

Set swApp = Application.SldWorks

If swApp Is Nothing Then

    MsgBox "SolidWorks is not available."

    Exit Function

End If

Set swModel = swApp.ActiveDoc

If swModel Is Nothing Then

    MsgBox "No active document."

    Exit Function

End If

Set swSelMgr = swModel.SelectionManager

If swSelMgr Is Nothing Then

    MsgBox "Unable to obtain Selection Manager."

    Exit Function

End If

    Dim swObj As Object

    Set swSelMgr = swModel.SelectionManager

    MsgBox "Selected objects: " & swSelMgr.GetSelectedObjectCount2(-1)

    Set swObj = swSelMgr.GetSelectedObject6(1, -1)

    If swObj Is Nothing Then

        MsgBox "Selected object is Nothing"

        Exit Function

    End If

    MsgBox TypeName(swObj)

    Set gOriginPoint = swObj
    Debug.Print "Origin X = "; gOriginPoint.X
Debug.Print "Origin Y = "; gOriginPoint.Y

MsgBox "Origin X = " & gOriginPoint.X & vbCrLf & _
       "Origin Y = " & gOriginPoint.Y

    If gOriginPoint Is Nothing Then

        MsgBox "Assignment failed"

        Exit Function

    End If

    MsgBox "Origin stored successfully"

    GetSketchOriginPoint = True

End Function
