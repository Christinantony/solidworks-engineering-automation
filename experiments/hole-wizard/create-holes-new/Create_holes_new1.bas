Attribute VB_Name = "Create_holes_new1"
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
boolstatus = Part.Extension.SelectByRay(-5.99084027038543E-02, 5.99999999998886E-02, -4.83809587852591E-03, 4.49472641230875E-02, -0.998446994872705, -3.29141592253758E-02, 5.35324898507481E-04, 2, False, 0, 0)

' Hole Wizard
Dim swHoleFeature As Feature
Set swHoleFeature = Part.FeatureManager.HoleWizard5(2, 0, 18, "#21", 0, 0.0043688, 0.03175, -1, 1, 2.05948851735331, 0, 0, 0, 0, 0, -1, -1, -1, -1, -1, "", False, True, True, True, True, False)
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
Set skPoint = Part.SketchManager.CreatePoint(-5.99084027038543E-02, 4.83809587852591E-03, 0)
Part.SketchManager.InsertSketch True
End Sub
