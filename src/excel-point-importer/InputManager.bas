Attribute VB_Name = "InputManager"
'===========================================================
' SW Point Importer
'
' Module    : InputManager
' Revision  : 1
'
' Compatible with:
' SolidWorks 2024
'
'===========================================================

Option Explicit

Public Function GetOriginCoordinates() As Boolean

    On Error GoTo ErrorHandler

    Dim strInput As String
    Dim dblOriginX As Double
    Dim dblOriginY As Double

    '-------------------------------------------------------
    ' Origin X
    '-------------------------------------------------------

    strInput = Trim$(InputBox( _
                "Enter Origin X coordinate (inches)", _
                "Origin X"))

    If Len(strInput) = 0 Then Exit Function

    If Not IsNumeric(strInput) Then

        MsgBox _
            "Origin X must be a numeric value.", _
            vbCritical

        Exit Function

    End If

    dblOriginX = CDbl(strInput)

    '-------------------------------------------------------
    ' Origin Y
    '-------------------------------------------------------

    strInput = Trim$(InputBox( _
                "Enter Origin Y coordinate (inches)", _
                "Origin Y"))

    If Len(strInput) = 0 Then Exit Function

    If Not IsNumeric(strInput) Then

        MsgBox _
            "Origin Y must be a numeric value.", _
            vbCritical

        Exit Function

    End If

    dblOriginY = CDbl(strInput)

    '-------------------------------------------------------
    ' Store internally as meters
    '-------------------------------------------------------

    gOriginX = dblOriginX * INCH_TO_METER
    gOriginY = dblOriginY * INCH_TO_METER

    GetOriginCoordinates = True

    Exit Function

ErrorHandler:

    MsgBox _
        "Input Error" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    GetOriginCoordinates = False

End Function
