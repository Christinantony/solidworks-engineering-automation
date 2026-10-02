Attribute VB_Name = "Batch_Rename1"
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swAssy As SldWorks.AssemblyDoc
Dim swComp As SldWorks.Component2
Dim swChild As SldWorks.Component2
Dim vComps As Variant

Dim errors As Long
Dim warnings As Long

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "No document open."
        Exit Sub
    End If
    
    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "Please open an Assembly."
        Exit Sub
    End If
    
    Set swAssy = swModel
    
    ' Ask for starting number
    Dim startNum As Long
    startNum = CLng(InputBox("Enter starting number (e.g. 3401 for HC-03401):", "Start Number", 3401))
    
    Dim prefix As String
    prefix = "HC-"
    
    Dim counter As Long
    counter = startNum
    
    ' Get all top-level components
    vComps = swAssy.GetComponents(False)
    
    Dim i As Long
    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        ProcessComponent swComp, prefix, counter
    Next i
    
    MsgBox "Renaming complete."
End Sub

Sub ProcessComponent(swComp As SldWorks.Component2, prefix As String, ByRef counter As Long)
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swComp.GetModelDoc2
    
    If swModel Is Nothing Then Exit Sub
    
    ' Only process PARTS (skip subassemblies & standard library parts)
    If swModel.GetType = swDocPART Then
        Dim oldPath As String
        oldPath = swModel.GetPathName
        
        If oldPath <> "" Then
            Dim fileName As String
            fileName = Dir(oldPath)
            
            ' Ask user whether to rename
            If MsgBox("Rename file: " & fileName & " ?", vbYesNo + vbQuestion) = vbYes Then
                Dim folderPath As String
                folderPath = Left(oldPath, InStrRev(oldPath, "\"))
                
                Dim newName As String
                newName = folderPath & prefix & Format(counter, "00000") & ".SLDPRT"
                
                Dim status As Boolean
                status = swModel.Extension.SaveAs(newName, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent + swSaveAsOptions_UpdateReferences + swSaveAsOptions_UpdateInactiveViews, _
                            Nothing, errors, warnings)
                
                If status = True Then
                    Debug.Print "Renamed: " & oldPath & " -> " & newName
                    counter = counter + 1
                Else
                    MsgBox "Error renaming: " & oldPath
                End If
            End If
        End If
    End If
End Sub
