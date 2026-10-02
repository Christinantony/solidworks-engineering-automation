Attribute VB_Name = "Create_Drawing1"
' ******************************************************************************
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Dim myModelView As Object
Set myModelView = Part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized

' New Document
Dim swSheetWidth As Double
swSheetWidth = 0.4318
Dim swSheetHeight As Double
swSheetHeight = 0.2794
Set Part = swApp.NewDocument("C:\CAD-Templates\Drawing.drwdot", 12, swSheetWidth, swSheetHeight)
Dim swDrawing As DrawingDoc
Set swDrawing = Part
Set swDrawing = Part
Dim swSheet As sheet
Set swSheet = swDrawing.GetCurrentSheet()
swSheet.SetProperties2 12, 12, 1, 1, False, swSheetWidth, swSheetHeight, True
swSheet.SetTemplateName "C:\CAD-Templates\SheetFormat.slddrt"
swSheet.ReloadTemplate True
swApp.ActivateDoc2 "Draw2 - Sheet1", False, longstatus
Set Part = swApp.ActiveDoc
Dim myView As Object
Set myView = Part.CreateDrawViewFromModelView3("C:\CAD-Examples\ExamplePart.SLDPRT", "*Front", 0.162902231578947, 0.150866196491228, 0)
boolstatus = Part.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.ActivateView("Drawing View1")
Part.ClearSelection2 True
End Sub
