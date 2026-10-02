Attribute VB_Name = "Name_the_circles_selected1"
Dim swApp As Object
Dim Part As ModelDoc2
Dim swSelMgr As SelectionMgr
Dim swSketchSegment As SketchSegment
Dim swCurve As Curve
Dim userInput As String
Dim swSketchText As SketchText
Dim centerPoint As Variant
Dim offset As Double
Dim currentNumber As Long

Sub main()
    ' Get SolidWorks application and active part
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    ' Check if a part document is active
    If Part Is Nothing Then
        MsgBox "Please open a part document and enter sketch mode."
        Exit Sub
    End If

    ' Prompt user for starting number
    userInput = InputBox("Enter the starting number:", "Starting Number")
    
    ' Validate user input
    If Not IsNumeric(userInput) Or Val(userInput) < 1 Then
        MsgBox "Please enter a valid positive number."
        Exit Sub
    End If
    
    ' Initialize the current number
    currentNumber = Val(userInput)
    
    ' Start a loop to listen for selections
    Do
        ' Check if a circle or arc is selected
        Set swSelMgr = Part.SelectionManager
        If swSelMgr.GetSelectedObjectCount2(0) > 0 Then
            Set swSketchSegment = swSelMgr.GetSelectedObject6(1, 0)

            ' Check if the selected object is an arc or circle
            If Not swSketchSegment Is Nothing And swSketchSegment.GetType = 1 Then
                ' Get the curve associated with the sketch segment
                Set swCurve = swSketchSegment.GetCurve()
                
                ' Get the center point of the curve (arc or circle)
                If Not swCurve Is Nothing Then
                    centerPoint = swCurve.Evaluate(0.5) ' Get the midpoint of the curve
                    
                    ' Define the smaller vertical offset to place text closer to the curve
                    offset = 0.005 ' Adjust as necessary
                    
                    ' Insert the text just above the center point of the circle or arc
                    Set swSketchText = Part.InsertSketchText(centerPoint(0), centerPoint(1) + offset, centerPoint(2), CStr(currentNumber), 0, 0, 0, 100, 100)
                    
                    ' Increment the current number
                    currentNumber = currentNumber + 1
                    
                    ' Redraw the window to update
                    Part.WindowRedraw
                Else
                    MsgBox "Failed to get curve for the selected segment."
                End If
            Else
                MsgBox "Please select a valid circle or arc."
            End If
            
            ' Clear selection to continue listening for the next selection
            Part.ClearSelection2 True
        End If
        
        ' Optional: Add a small delay to prevent high CPU usage
        DoEvents
    Loop While True ' Continue the loop indefinitely; you can implement a termination condition if desired
End Sub
