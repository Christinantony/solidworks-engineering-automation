Attribute VB_Name = "Create_holes_cd1"
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
boolstatus = Part.Extension.SelectByRay(-1.06826136151525E-02, 0.060000000000116, 4.24451730842179E-03, 7.67274749200749E-02, -0.928404127046341, -0.363563847866795, 7.31303949306057E-04, 2, False, 0, 0)

' Hole Wizard
Dim swHoleFeature As Feature
Set swHoleFeature = Part.FeatureManager.HoleWizard5(5, -1, -1, "", 0, 0.0254, 0.0635, 1, 0.0508, 0.03048, 2.05948851735331, -1, -1, -1, -1, -1, -1, -1, -1, -1, "", False, True, True, True, True, False)
Dim swSketchFeature As Feature
Set swSketchFeature = swHoleFeature.GetFirstSubFeature
swSketchFeature.Select2 False, 0
Part.EditSketch
Dim swSelectionManager As SelectionMgr
Set swSelectionManager = Part.SelectionManager
Dim swSketch As Sketch
Set swSketch = swSketchFeature.GetSpecificFeature2()
Dim swSketchPointArray As Variant
swSketchPointArray = swSketch.GetSketchPoints2()
Dim swMaxPointNumber As Long
swMaxPointNumber = UBound(swSketchPointArray)
Dim swSketchPoint As Object
Dim swCurrentPointNumber As Long
For swCurrentPointNumber = 0 To swMaxPointNumber Step 1
   Set swSketchPoint = swSketchPointArray(swCurrentPointNumber)
   boolstatus = swSelectionManager.AddSelectionListObject(swSketchPoint, Nothing)
   Part.EditDelete
Next swCurrentPointNumber
Dim skPoint As Object
Set skPoint = Part.SketchManager.CreatePoint(-3.35207491185851E-02, 0, 0)
Set skPoint = Part.SketchManager.CreatePoint(-3.35207491185851E-02, 0, 0)
Part.SketchManager.InsertSketch True
End Sub
