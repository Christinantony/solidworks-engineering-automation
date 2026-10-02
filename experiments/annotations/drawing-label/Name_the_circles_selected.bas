Attribute VB_Name = "Name_the_circles_selected"
Dim swApp As Object
Dim Part As Object
Dim currentNumber As Long
Dim swSelMgr As Object
Dim boolstatus As Boolean
Dim swNote As Object
Dim selectedPoint As Variant

Sub main()

    ' Initialize SolidWorks application and active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    ' Check if the document is a drawing
    If Part Is Nothing Or Part.GetType <> swDocDRAWING Then
        MsgBox "Please open a drawing document."
        Exit Sub
    End If

    ' Ask for the starting number
    currentNumber = InputBox("Enter the starting number:", "Starting Number")

    ' Validate input
    If Not IsNumeric(currentNumber) Or currentNumber < 1 Then
        MsgBox "Please enter a valid starting number."
        Exit Sub
    End If

    ' Start the listening loop
    Do
        ' Get the selection manager
        Set swSelMgr = Part.SelectionManager

        ' Check if a point has been clicked (in the drawing view)
        If swSelMgr.GetSelectedObjectCount2(0) > 0 Then
            ' Get the selected object (expected to be a point or entity in the drawing view)
            Set selectedEntity = swSelMgr.GetSelectedObject6(1, 0)

            ' Only process if a valid entity is selected
            If Not selectedEntity Is Nothing Then
                ' Get the coordinates of the selected entity in the drawing view
                selectedPoint = selectedEntity.GetPosition

                ' Insert the note at the selected point (position in drawing view)
                Set swNote = Part.InsertNote(CStr(currentNumber))

                ' Ensure the note has no arrow and is placed at the correct location
                If Not swNote Is Nothing Then
                    swNote.LockPosition = False
                    swNote.Angle = 0
                    boolstatus = swNote.SetBalloon(0, 0) ' No balloon for the note
                    boolstatus = swNote.GetAnnotation().SetPosition(selectedPoint(0), selectedPoint(1), selectedPoint(2)) ' Place the note at the click point
                End If

                ' Increment the current number for the next note
                currentNumber = currentNumber + 1
            End If

            ' Clear the selection to listen for the next click
            Part.ClearSelection2 True
        End If

        ' Optional: Add a small delay to prevent high CPU usage
        DoEvents

    Loop

End Sub
