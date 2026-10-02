Attribute VB_Name = "Name_the_circles_drawn1"
Dim swApp As Object
Dim Part As Object
Dim SelMgr As Object
Dim Sketch As Object
Dim swSketchSegment As Object
Dim nCounter As Integer

Sub main()

    ' Initialize SolidWorks App and Part
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    ' Get the selection manager
    Set SelMgr = Part.SelectionManager

    ' Check if a sketch is active
    If Part.SketchManager.ActiveSketch Is Nothing Then
        MsgBox "No active sketch found"
        Exit Sub
    End If
    
    ' Prompt for starting number if this is the first circle
    If nCounter = 0 Then
        nCounter = InputBox("Enter starting number:", "Starting Number")
    End If

    ' Get the selected circle
    Set swSketchSegment = SelMgr.GetSelectedObject5(1)

    ' Check if the selected object is a circle
    If swSketchSegment.GetType = swSketchSegment.swSketchARC Then
        ' Increment the circle number
        nCounter = nCounter + 1
        
        ' Place the text (circle number) on the circle
        Dim xPos As Double
        Dim yPos As Double
        Dim zPos As Double
        swSketchSegment.GetCenterPoint2 xPos, yPos, zPos
        
        Part.SketchManager.InsertSketchText xPos, yPos, zPos, "Circle " & nCounter, False, 0, 0

        ' Rebuild the sketch to update
        Part.SketchManager.InsertSketch True
    Else
        MsgBox "Please select a valid circle"
    End If

End Sub
