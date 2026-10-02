Attribute VB_Name = "test_macro_for_rename1"
' ******************************************************************************
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
boolstatus = Part.Extension.SelectByID2("HC-00998.SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("HC-00998.SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.EditRebuild3()
Part.ClearSelection2 True

' Rename
longstatus = Part.Extension.RenameDocument("HC-00998_01")

' Save
Dim swErrors As Long
Dim swWarnings As Long
boolstatus = Part.Save3(1, swErrors, swWarnings)
End Sub
