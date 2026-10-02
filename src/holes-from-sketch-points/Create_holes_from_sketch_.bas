Attribute VB_Name = "Create_holes_from_sketch_"
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swDoc As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swFace As SldWorks.Face2
Dim vSelectedPoints As Collection
Dim pointData As Collection
Dim continueSelecting As Boolean
Dim angleDegrees As Double
Dim angleRadians As Double

' Main Subroutine
Sub main()
    
    ' Initialize variables
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc
    Set swSelMgr = swDoc.SelectionManager
    
    ' Check for SolidWorks environment
    If swApp Is Nothing Or swDoc Is Nothing Then
        MsgBox "SolidWorks is not open or no document is active."
        Exit Sub
    End If
    
    ' Initialize point collection
    Set vSelectedPoints = New Collection
    
    ' Wait for face selection
    Set swFace = WaitForFaceSelection()
    If swFace Is Nothing Then
        MsgBox "No face selected. Exiting macro."
        Exit Sub
    End If
    

    ' Prompt the user to input the drill angle in degrees
    angleDegrees = CDbl(InputBox("Enter the drill angle in degrees:", "Drill Angle"))

    ' Convert the angle from degrees to radians
    angleRadians = angleDegrees * (3.14159265358979 / 180)
    
    ' Wait for points selection and gather data
    continueSelecting = True
    Do While continueSelecting
        Dim point As SldWorks.SketchPoint
        Set point = WaitForPointSelection()
        
        ' Get user input for diameter and depth in inches
        Dim dia As Double, depth As Double
        dia = InputBox("Enter hole dia in inches:") ' Get diameter directly from user input
        depth = InputBox("Enter hole depth in inches:")
        
        ' Store data as a new Collection item for each point
        Set pointData = New Collection
        pointData.Add point
        pointData.Add dia
        pointData.Add depth * 0.0254  ' Convert inches to meters
        vSelectedPoints.Add pointData
        
        ' Deselect the current point
        swDoc.ClearSelection2 True
        
        ' Ask user to continue or start hole creation
        Dim response As Integer
        response = MsgBox("Select further points or start hole creation?", vbYesNo)
        If response = vbNo Then
            continueSelecting = False
        End If
    Loop
    
    ' Create holes based on collected data
    CreateHoles vSelectedPoints
    swDoc.ViewZoomtofit2
    swDoc.ClearSelection2 True
    
    MsgBox "Hole creation completed."
End Sub

' Waits for a face selection with a 1-second interval
Function WaitForFaceSelection() As SldWorks.Face2
    Do While swSelMgr.GetSelectedObjectCount2(-1) = 0 Or swSelMgr.GetSelectedObjectType3(1, -1) <> swSelectType_e.swSelFACES
        DoEvents
        Delay 1  ' 1-second delay
    Loop
    Set WaitForFaceSelection = swSelMgr.GetSelectedObject6(1, -1)
End Function

' Waits for a point selection with a 1-second interval
Function WaitForPointSelection() As SldWorks.SketchPoint
    Do
        DoEvents
        If swSelMgr.GetSelectedObjectCount2(-1) > 0 And swSelMgr.GetSelectedObjectType3(1, -1) = swSelectType_e.swSelEXTSKETCHPOINTS Then
            Set WaitForPointSelection = swSelMgr.GetSelectedObject6(1, -1)
            Exit Function
        End If
        Delay 1  ' 1-second delay
    Loop
End Function

Sub CreateHoles(points As Collection)
    Dim pointData As Collection, point As SldWorks.SketchPoint
    Dim dia As Double, depth As Double
    Dim swHoleFeature As SldWorks.Feature
    Dim drillCode As String
    
    For Each pointData In points
        Set point = pointData(1)
        dia = pointData(2)
        depth = pointData(3)
        
        ' Clear any previous selections to start fresh
        swDoc.ClearSelection2 True
        
        ' Get the closest drill size code
        drillCode = GetClosestDrillSizeCode(dia)
        
        ' Select the point
        If Not point Is Nothing Then
            point.Select True
            
            ' Hole creation with the '#' symbol before drillCode
            Set swHoleFeature = swDoc.FeatureManager.HoleWizard5(2, 0, 18, drillCode, 0, dia * 0.0254, depth, -1, 1, _
                angleRadians, 0, 0, 0, 0, 0, -1, -1, -1, -1, -1, "", False, True, True, True, True, False)
            
            If swHoleFeature Is Nothing Then
                MsgBox "Failed to create hole at one of the points."
                Exit Sub
            End If
        Else
            MsgBox "A selected point is missing or invalid."
            Exit Sub
        End If
        
        ' Deselect the current point after creating the hole
        swDoc.ClearSelection2 True
    Next pointData
End Sub

' Custom delay function to replace Application.Wait
Sub Delay(seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents  ' Allows the system to process other events
    Loop
End Sub

' Function to return the closest ANSI inch drill size code to the given diameter
Function GetClosestDrillSizeCode(dia As Double) As String
    ' Array of ANSI inch drill sizes and their corresponding codes
    Dim drillSizes() As Variant
    drillSizes = Array( _
        Array(0.0135, "#80"), Array(0.0145, "#79"), Array(0.015, "#78"), Array(0.016, "#77"), Array(0.018, "#76"), Array(0.02, "#75"), Array(0.021, "#74"), Array(0.0225, "#73"), Array(0.024, "#72"), Array(0.025, "#71"), _
        Array(0.026, "#70"), Array(0.028, "#69"), Array(0.0292, "#68"), Array(0.031, "#67"), Array(0.032, "#66"), Array(0.033, "#65"), Array(0.035, "#64"), Array(0.037, "#63"), Array(0.038, "#62"), Array(0.04, "#61"), _
        Array(0.041, "#60"), Array(0.042, "#59"), Array(0.043, "#58"), Array(0.0465, "#57"), Array(0.052, "#55"), Array(0.055, "#54"), Array(0.0595, "#53"), Array(0.0635, "#52"), Array(0.067, "#51"), Array(0.07, "#50"), _
        Array(0.073, "#49"), Array(0.076, "#48"), Array(0.0785, "#47"), Array(0.081, "#46"), Array(0.082, "#45"), Array(0.086, "#44"), Array(0.089, "#43"), Array(0.0935, "#42"), Array(0.096, "#41"), Array(0.098, "#40"), _
        Array(0.0995, "#39"), Array(0.1015, "#38"), Array(0.104, "#37"), Array(0.1065, "#36"), Array(0.106, "#35"), Array(0.111, "#34"), Array(0.113, "#33"), Array(0.116, "#32"), Array(0.12, "#31"), Array(0.125, "#30"), _
        Array(0.1285, "#29"), Array(0.136, "#28"), Array(0.1405, "#27"), Array(0.144, "#26"), Array(0.147, "#25"), Array(0.1495, "#24"), Array(0.154, "#23"), Array(0.157, "#22"), Array(0.159, "#21"), Array(0.166, "#19"), _
        Array(0.18, "#16"), Array(0.191, "#11"), Array(0.196, "#9"), Array(0.199, "#8"), Array(0.205, "#6"), Array(0.209, "#4"), Array(0.221, "#2"), Array(0.228, "#1"), Array(0.234, "A"), Array(0.238, "B"), _
        Array(0.242, "C"), Array(0.246, "D"), Array(0.25, "E"), Array(0.257, "F"), Array(0.261, "G"), Array(0.266, "H"), Array(0.272, "I"), Array(0.277, "J"), Array(0.281, "K"), Array(0.29, "L"), _
        Array(0.295, "M"), Array(0.302, "N"), Array(0.316, "O"), Array(0.323, "P"), Array(0.332, "Q"), Array(0.339, "R"), Array(0.348, "S"), Array(0.358, "T"), Array(0.368, "U"), Array(0.377, "V"), Array(0.386, "W"), Array(0.397, "X"), Array(0.404, "Y"), Array(0.413, "Z"), Array(0.03125, "1/32"), Array(0.046875, "3/64"), Array(0.0625, "1/16"), Array(0.078125, "5/64"), _
        Array(0.09375, "3/32"), Array(0.109375, "7/64"), Array(0.125, "1/8"), Array(0.140625, "9/64"), Array(0.15625, "5/32"), Array(0.171875, "11/64"), Array(0.1875, "3/16"), Array(0.203125, "13/64"), Array(0.21875, "7/32"), Array(0.234375, "15/64"), Array(0.25, "1/4"), Array(0.265625, "17/64"), Array(0.28125, "9/32"), Array(0.296875, "19/64"), Array(0.3125, "5/16"), Array(0.328125, "21/64"), _
        Array(0.34375, "11/32"), Array(0.359375, "23/64"), Array(0.375, "3/8"), Array(0.390625, "25/64"), Array(0.40625, "13/32"), Array(0.421875, "27/64"), Array(0.4375, "7/16"), Array(0.453125, "29/64"), _
        Array(0.46875, "15/32"), Array(0.484375, "31/64"), Array(0.5, "1/2"), Array(0.515625, "33/64"), Array(0.53125, "17/32"), Array(0.546875, "35/64"), Array(0.5625, "9/16"), Array(0.578125, "37/64"), _
        Array(0.59375, "19/32"), Array(0.609375, "39/64"), Array(0.625, "5/8"), Array(0.640625, "41/64"), Array(0.65625, "21/32"), Array(0.671875, "43/64"), Array(0.6875, "11/16"), Array(0.703125, "45/64"), _
        Array(0.71875, "23/32"), Array(0.734375, "47/64"), Array(0.75, "3/4"), Array(0.765625, "49/64"), Array(0.78125, "25/32"), Array(0.796875, "51/64"), Array(0.8125, "13/16"), Array(0.828125, "53/64"), _
        Array(0.84375, "27/32"), Array(0.859375, "55/64"), Array(0.875, "7/8"), Array(0.890625, "57/64"), Array(0.90625, "29/32"), Array(0.921875, "59/64"), Array(0.9375, "15/16"), Array(0.953125, "61/64"), _
        Array(0.96875, "31/32"), Array(0.984375, "63/64"), Array(1, "1"))

    Dim closestSize As Double
    Dim closestCode As String
    closestSize = drillSizes(0)(0)
    closestCode = drillSizes(0)(1)

    Dim i As Integer
    For i = 0 To UBound(drillSizes)
        Dim size As Double
        Dim code As String
        size = drillSizes(i)(0)
        code = drillSizes(i)(1)

        ' Check if this size is closer to the desired diameter
        If Abs(size - dia) < Abs(closestSize - dia) Then
            closestSize = size
            closestCode = code
        End If
    Next i

    GetClosestDrillSizeCode = closestCode
End Function
