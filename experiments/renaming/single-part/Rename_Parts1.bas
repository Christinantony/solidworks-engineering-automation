Attribute VB_Name = "Rename_Parts1"
Option Explicit

Const USE_UPDATE_HANDLER As Boolean = True   ' set False to test without the event handler

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long
Dim partEvents As Class1

Sub main()

    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    If Part Is Nothing Then
        MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If
    If Part.GetType <> 1 Then          ' 1 = swDocPART
        MsgBox "The active document is not a part.", vbExclamation
        Exit Sub
    End If

    Dim fullPath As String, folder As String, fileName As String
    Dim baseName As String, ext As String, newName As String

    fullPath = Part.GetPathName
    If fullPath = "" Then
        MsgBox "This part has no file on disk.", vbExclamation
        Exit Sub
    End If

    folder = Left(fullPath, InStrRev(fullPath, "\"))
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    ext = Mid(fileName, InStrRev(fileName, "."))
    baseName = Left(fileName, InStrRev(fileName, ".") - 1)
    newName = baseName & "_01"

    ' Overwrite: remove an existing file with the target name
    On Error Resume Next
    If Dir(folder & newName & ext) <> "" Then
        SetAttr folder & newName & ext, vbNormal
        Kill folder & newName & ext
    End If
    On Error GoTo 0

    ' Hook the event BEFORE renaming so "Update where used" is forced on
    If USE_UPDATE_HANDLER Then
        Set partEvents = New Class1
        Set partEvents.swPart = Part
    End If

    ' Rebuild FIRST - it can clear the selection
    boolstatus = Part.EditRebuild3()

    ' Select LAST, immediately before RenameDocument
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByID2(fileName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    If Not boolstatus Then
        boolstatus = Part.Extension.SelectByID2(baseName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    End If
    If Not boolstatus Then
        boolstatus = Part.Extension.SelectByID2(Part.GetTitle, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    End If
    If Not boolstatus Then
        MsgBox "Could not select the part as a component." & vbCrLf & _
               "Tried: " & fileName & " / " & baseName & " / " & Part.GetTitle, vbCritical
        Exit Sub
    End If

    ' Rename
    longstatus = Part.Extension.RenameDocument(newName)
    If longstatus <> 0 Then
        MsgBox "Rename failed (swRenameDocumentError_e = " & longstatus & ").", vbCritical
        Exit Sub
    End If

    ' Save
    Dim swErrors As Long, swWarnings As Long
    Dim firstErrors As Long

    ' first save: NOT silent, so the Rename Documents prompt can appear
    boolstatus = Part.Save3(0, swErrors, swWarnings)
    firstErrors = swErrors

    ' retry: still not silent, also saving referenced documents
    If swErrors <> 0 Then
        swErrors = 0: swWarnings = 0
        boolstatus = Part.Save3(swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, swErrors, swWarnings)
    End If

    If swErrors = 0 Then
        MsgBox "Renamed and saved as " & newName & ext & _
               IIf(firstErrors <> 0, vbCrLf & "(first save returned " & firstErrors & "; retry with referenced docs succeeded)", ""), _
               vbInformation
    Else
        ' Report raw flags and what is actually on disk
        Dim bits As String, b As Long
        For b = 0 To 30
            If (swErrors And (2 ^ b)) <> 0 Then bits = bits & (2 ^ b) & " "
        Next b

        MsgBox "Renamed to " & newName & ", but save still failed." & vbCrLf & _
               "First save error: " & firstErrors & vbCrLf & _
               "Retry error: " & swErrors & "  (flags: " & bits & ")" & vbCrLf & _
               "New file on disk: " & (Dir(folder & newName & ext) <> "") & vbCrLf & _
               "Old file on disk: " & (Dir(folder & fileName) <> ""), vbExclamation
    End If

End Sub
