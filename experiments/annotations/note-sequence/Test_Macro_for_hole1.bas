Attribute VB_Name = "Test_Macro_for_hole1"
Dim swApp As Object
Dim Part As Object
Dim currentNumber As Long
Dim endNumber As Long
Dim noteX As Double
Dim noteY As Double
Dim i As Long
Dim myNote As Object
Dim myAnnotation As Object
Dim myTextFormat As Object

Sub main()
    ' Initialize SolidWorks application and get the active document
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    
    ' Ensure there is an active document
    If Part Is Nothing Then
        MsgBox "No document is open!"
        Exit Sub
    End If
    
    ' Ask for starting and ending number
    currentNumber = InputBox("Enter the starting number:", "Starting Number")
    endNumber = InputBox("Enter the ending number:", "Ending Number")
    
    ' Ensure ending number is greater than starting number
    If endNumber < currentNumber Then
        MsgBox "Ending number must be greater than starting number."
        Exit Sub
    End If
    
    ' Initialize position for the first note
    noteX = 0.1 ' Starting X position (adjust as needed)
    noteY = 0.1 ' Starting Y position (adjust as needed)
    
    ' Loop to create notes from starting number to ending number
    For i = currentNumber To endNumber
        ' Insert note with the current number
        Set myNote = Part.InsertNote(CStr(i))
        
        If Not myNote Is Nothing Then
            ' Get the annotation object to modify properties
            Set myAnnotation = myNote.GetAnnotation()
            
            ' Apply leader style with no arrows
            If Not myAnnotation Is Nothing Then
                myAnnotation.SetLeader3 0, 0, False, False, False, False ' Remove arrows (leader type 0)
                
                ' Set text format to bold
                Set myTextFormat = myAnnotation.GetTextFormat(0)
                If Not myTextFormat Is Nothing Then
                    myTextFormat.Bold = True
                    myAnnotation.SetTextFormat 0, True, myTextFormat
                    myAnnotation.SetTextFormat 1, True, myTextFormat
                End If
            End If
            
            ' Increment the position for the next note (moving it further to the right)
            noteX = noteX + 0.15 ' Horizontal spacing
            noteY = noteY ' Keep vertical position constant
        End If
    Next i
    
    ' Clear selection and redraw the part
    Part.ClearSelection2 True
    Part.WindowRedraw
End Sub
