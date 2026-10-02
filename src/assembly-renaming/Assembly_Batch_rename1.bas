Attribute VB_Name = "Assembly_Batch_rename1"
Option Explicit

' ============================= SETTINGS =============================
Const SUFFIX_TO_FIND As String = "_01"
Const RENAME_DRAWINGS As Boolean = True       ' rename matching .SLDDRW (same folder, same base name as old part)
Const DELETE_OLD_DRAWINGS As Boolean = False   ' delete old drawing file after it is saved under the new name
' ====================================================================

Private Type PartItem
    oldBase As String
    folder As String
    ext As String
    newName As String
    drwPath As String
    comp As Object
    doc As Object
    drwDoc As Object
    renamed As Boolean
    saved As Boolean
End Type

Private items() As PartItem
Private itemCount As Long
Private seenParts As Object
Private seenAsms As Object
Private asmDocs As Collection
Private handlers As Collection
Private skipNotes As String

Sub main()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swAsm As SldWorks.AssemblyDoc
    Dim swConf As SldWorks.Configuration
    Dim i As Long, res As Long, errs As Long, warns As Long
    Dim v As Variant
    Dim ok As Boolean
    Dim report As String
    Dim topPath As String

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If
    If swModel.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
        MsgBox "Run this from the top-level assembly.", vbExclamation
        Exit Sub
    End If
    topPath = swModel.GetPathName
    If topPath = "" Then
        MsgBox "Save the assembly first.", vbExclamation
        Exit Sub
    End If

    Set swAsm = swModel
    swAsm.ResolveAllLightWeightComponents False

    ' ---------- 1. collect the _01 parts in tree order ----------
    Set seenParts = CreateObject("Scripting.Dictionary")
    seenParts.CompareMode = 1
    Set seenAsms = CreateObject("Scripting.Dictionary")
    seenAsms.CompareMode = 1
    Set asmDocs = New Collection
    itemCount = 0
    skipNotes = ""
    Erase items

    Set swConf = swModel.ConfigurationManager.ActiveConfiguration
    Traverse swConf.GetRootComponent3(True)

    If itemCount = 0 Then
        MsgBox "No parts ending in " & SUFFIX_TO_FIND & " were found." & vbCrLf & skipNotes, vbInformation
        Exit Sub
    End If

    ' ---------- 2. ask for the first part number ----------
    Dim firstInput As String, prefix As String
    Dim startNum As Long, width As Long

    firstInput = Trim$(InputBox("Found " & itemCount & " part(s) ending in " & SUFFIX_TO_FIND & "." & vbCrLf & vbCrLf & _
                 "Enter the FIRST new part number (e.g. HC-01010)." & vbCrLf & _
                 "The rest are numbered +1 in tree order.", "Rename " & SUFFIX_TO_FIND & " parts"))
    If firstInput = "" Then Exit Sub
    If Not SplitPartNumber(firstInput, prefix, startNum, width) Then
        MsgBox "The part number must end in digits, e.g. HC-01010.", vbExclamation
        Exit Sub
    End If

    ' ---------- 3. build names, find drawings, write preview ----------
    Dim txt As String, conflicts As Long, p As String
    txt = "PREVIEW - nothing has been changed yet" & vbCrLf & String$(70, "-") & vbCrLf

    For i = 1 To itemCount
        With items(i)
            .newName = prefix & Format$(startNum + i - 1, String$(width, "0"))
            .drwPath = ""
            If RENAME_DRAWINGS Then
                p = .folder & .oldBase & ".SLDDRW"
                If Dir(p) <> "" Then .drwPath = p
            End If

            txt = txt & Format$(i, "00") & "  " & .oldBase & .ext & "  ->  " & .newName & .ext
            If RENAME_DRAWINGS Then
                If .drwPath <> "" Then
                    txt = txt & "   | drawing -> " & .newName & ".SLDDRW"
                Else
                    txt = txt & "   | drawing: none found"
                End If
            End If
            If Dir(.folder & .newName & .ext) <> "" Then
                txt = txt & "   ** CONFLICT: " & .newName & .ext & " already exists"
                conflicts = conflicts + 1
            End If
            If .drwPath <> "" Then
                If Dir(.folder & .newName & ".SLDDRW") <> "" Then
                    txt = txt & "   ** CONFLICT: " & .newName & ".SLDDRW already exists"
                    conflicts = conflicts + 1
                End If
            End If
            txt = txt & vbCrLf
        End With
    Next i

    If Len(skipNotes) > 0 Then txt = txt & vbCrLf & "Skipped:" & vbCrLf & skipNotes

    Dim previewPath As String, repPath As String
    previewPath = Environ$("TEMP") & "\rename_preview.txt"
    repPath = Environ$("TEMP") & "\rename_report.txt"
    WriteText previewPath, txt
    Shell "notepad.exe """ & previewPath & """", vbNormalFocus

    If conflicts > 0 Then
        MsgBox conflicts & " conflict(s): target files already exist. Nothing was changed.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Proceed with the renames shown in the preview?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    ' ---------- 4. open the drawings so their references follow the rename ----------
    If RENAME_DRAWINGS Then
        For i = 1 To itemCount
            If items(i).drwPath <> "" Then
                errs = 0: warns = 0
                Set items(i).drwDoc = swApp.OpenDoc6(items(i).drwPath, swDocumentTypes_e.swDocDRAWING, _
                                      swOpenDocOptions_e.swOpenDocOptions_Silent, "", errs, warns)
                If items(i).drwDoc Is Nothing Then
                    report = report & "Could not open drawing: " & items(i).drwPath & " (error " & errs & ")" & vbCrLf
                End If
            End If
        Next i

        ' bring the assembly back to the front (0 = don't rebuild on activation)
        swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
        If LCase$(swApp.ActiveDoc.GetPathName) <> LCase$(topPath) Then
            MsgBox "Could not re-activate the assembly. Nothing has been renamed yet.", vbCritical
            Exit Sub
        End If
    End If

    ' ---------- 5. hook the "Update where used" events ----------
    Set handlers = New Collection
    AttachAsm swModel
    For Each v In asmDocs
        AttachAsm v
    Next v

    ' ---------- 6. rename the parts ----------
    For i = 1 To itemCount
        swModel.ClearSelection2 True
        ok = items(i).comp.Select4(False, Nothing, False)
        If Not ok Then
            report = report & "SELECT FAILED: " & items(i).oldBase & vbCrLf
        Else
            AttachPart items(i).doc
            res = swModel.Extension.RenameDocument(items(i).newName)
            If res = 0 Then
                items(i).renamed = True
            Else
                report = report & "RENAME FAILED: " & items(i).oldBase & " -> " & items(i).newName & _
                         "  (error " & res & " = " & RenameErrName(res) & ")" & vbCrLf
            End If
        End If
    Next i
    swModel.ClearSelection2 True

    ' ---------- 7. save: parts, then sub-assemblies (deepest first), then top ----------
    For i = 1 To itemCount
        If items(i).renamed Then
            res = SaveDoc(items(i).doc)
            If res = 0 Then
                items(i).saved = True
            Else
                report = report & "SAVE FAILED (part " & items(i).newName & "): error " & res & vbCrLf
            End If
        End If
    Next i

    For Each v In asmDocs
        res = SaveDoc(v)
        If res <> 0 Then report = report & "SAVE FAILED (sub-assembly " & v.GetTitle & "): error " & res & vbCrLf
    Next v

    res = SaveDoc(swModel)
    If res <> 0 Then report = report & "SAVE FAILED (top assembly): error " & res & vbCrLf

    ' ---------- 8. rename the drawings (Save As new name, close, delete old) ----------
    Dim nDrw As Long, newDrw As String
    If RENAME_DRAWINGS Then
        For i = 1 To itemCount
            If items(i).saved And Not items(i).drwDoc Is Nothing Then
                newDrw = items(i).folder & items(i).newName & ".SLDDRW"
                errs = 0: warns = 0
                items(i).drwDoc.Extension.SaveAs3 newDrw, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                    swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, Nothing, errs, warns

                If Dir(newDrw) <> "" Then
                    nDrw = nDrw + 1
                    If errs <> 0 Then report = report & "Drawing saved with error " & errs & ": " & newDrw & vbCrLf

                    On Error Resume Next
                    swApp.CloseDoc newDrw
                    If Not swApp.GetOpenDocumentByName(newDrw) Is Nothing Then swApp.CloseDoc items(i).drwDoc.GetTitle
                    If DELETE_OLD_DRAWINGS Then
                        SetAttr items(i).drwPath, vbNormal
                        Kill items(i).drwPath
                    End If
                    On Error GoTo 0

                    If DELETE_OLD_DRAWINGS And Dir(items(i).drwPath) <> "" Then
                        report = report & "Could not delete old drawing: " & items(i).drwPath & vbCrLf
                    End If
                Else
                    report = report & "DRAWING SAVE-AS FAILED: " & items(i).drwPath & " (error " & errs & ")" & vbCrLf
                End If
            ElseIf items(i).drwPath <> "" And Not items(i).saved Then
                report = report & "Drawing not renamed (part was not renamed/saved): " & items(i).drwPath & vbCrLf
            End If
        Next i
    End If

    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs

    ' ---------- 9. summary ----------
    Dim nRen As Long, nSaved As Long, summary As String
    For i = 1 To itemCount
        If items(i).renamed Then nRen = nRen + 1
        If items(i).saved Then nSaved = nSaved + 1
    Next i

    summary = "Parts renamed: " & nRen & " of " & itemCount & vbCrLf & "Parts saved: " & nSaved
    If RENAME_DRAWINGS Then summary = summary & vbCrLf & "Drawings renamed: " & nDrw

    If Len(report) > 0 Then
        WriteText repPath, report
        MsgBox summary & vbCrLf & vbCrLf & "Some items need attention - opening the report.", vbExclamation
        Shell "notepad.exe """ & repPath & """", vbNormalFocus
    Else
        MsgBox summary, vbInformation
    End If

End Sub

' ------------------------------------------------------------------
'  Tree walk: depth-first, in FeatureManager order
' ------------------------------------------------------------------
Private Sub Traverse(ByVal swComp As SldWorks.Component2)

    Dim v As Variant
    Dim i As Long
    Dim child As SldWorks.Component2
    Dim d As SldWorks.ModelDoc2
    Dim p As String, baseName As String

    v = swComp.GetChildren
    If IsEmpty(v) Then Exit Sub

    For i = 0 To UBound(v)
        Set child = v(i)
        If Not child.IsSuppressed Then
            Set d = child.GetModelDoc2
            If d Is Nothing Then
                skipNotes = skipNotes & "  Not loaded: " & child.Name2 & vbCrLf
            Else
                p = d.GetPathName
                If p <> "" Then
                    Select Case d.GetType
                        Case swDocumentTypes_e.swDocPART
                            baseName = FileBase(p)
                            If LCase$(Right$(baseName, Len(SUFFIX_TO_FIND))) = LCase$(SUFFIX_TO_FIND) Then
                                If child.IsVirtual Then
                                    skipNotes = skipNotes & "  Virtual component: " & child.Name2 & vbCrLf
                                ElseIf Not seenParts.Exists(LCase$(p)) Then
                                    seenParts.Add LCase$(p), 1
                                    AddItem p, baseName, child, d
                                End If
                            End If
                        Case swDocumentTypes_e.swDocASSEMBLY
                            Traverse child
                            If Not seenAsms.Exists(LCase$(p)) Then
                                seenAsms.Add LCase$(p), 1
                                asmDocs.Add d
                            End If
                    End Select
                End If
            End If
        End If
    Next i

End Sub

Private Sub AddItem(ByVal p As String, ByVal baseName As String, ByVal comp As Object, ByVal d As Object)
    itemCount = itemCount + 1
    ReDim Preserve items(1 To itemCount)
    With items(itemCount)
        .oldBase = baseName
        .folder = Left$(p, InStrRev(p, "\"))
        .ext = Mid$(p, InStrRev(p, "."))
        Set .comp = comp
        Set .doc = d
    End With
End Sub

' ------------------------------------------------------------------
'  Event hooks (they force "Update where used" on save)
' ------------------------------------------------------------------
Private Sub AttachAsm(ByVal doc As Object)
    Dim h As Class2
    Set h = New Class2
    Set h.swAsm = doc
    handlers.Add h
End Sub

Private Sub AttachPart(ByVal doc As Object)
    Dim h As Class1
    Set h = New Class1
    Set h.swPart = doc
    handlers.Add h
End Sub

' ------------------------------------------------------------------
'  Save: not silent (a silent save can't complete a pending rename),
'  then retry once with referenced documents
' ------------------------------------------------------------------
Private Function SaveDoc(ByVal doc As Object) As Long
    Dim e As Long, w As Long
    doc.Save3 0, e, w
    If e <> 0 Then
        e = 0: w = 0
        doc.Save3 swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, e, w
    End If
    SaveDoc = e
End Function

' ------------------------------------------------------------------
'  Helpers
' ------------------------------------------------------------------
Private Function SplitPartNumber(ByVal s As String, ByRef prefix As String, _
                                 ByRef num As Long, ByRef width As Long) As Boolean
    Dim i As Long
    i = Len(s)
    Do While i > 0
        If Mid$(s, i, 1) Like "[0-9]" Then i = i - 1 Else Exit Do
    Loop
    If i = Len(s) Then Exit Function              ' no trailing digits
    width = Len(s) - i
    If width > 9 Then Exit Function
    prefix = Left$(s, i)
    num = CLng(Mid$(s, i + 1))
    SplitPartNumber = True
End Function

Private Function FileBase(ByVal p As String) As String
    Dim f As String
    f = Mid$(p, InStrRev(p, "\") + 1)
    If InStrRev(f, ".") > 0 Then f = Left$(f, InStrRev(f, ".") - 1)
    FileBase = f
End Function

Private Sub WriteText(ByVal path As String, ByVal txt As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, txt
    Close #f
End Sub

Private Function RenameErrName(ByVal c As Long) As String
    If c >= 1 And c <= 19 Then
        RenameErrName = Choose(c, "UnspecifiedInternalError", "InvalidSelection", "InvalidForDrawings", _
            "NoModelLoaded", "ComponentNotResolved", "LightWeightComponent", "RoutingComponent", _
            "FileAlreadyExists", "InvalidCharactersInName", "InvalidVirtualComponent", "NameTooLong", _
            "DocumentNameInUse", "PendingNameAlreadyInUse", "ReadOnlyDocument", "DocumentNotSaved", _
            "VirtualComponent", "NotAllowedWithPDM", "ToolboxComponent", "PatternedComponent")
    Else
        RenameErrName = "unknown"
    End If
End Function
