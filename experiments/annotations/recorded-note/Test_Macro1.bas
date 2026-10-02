Attribute VB_Name = "Test_Macro1"
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
boolstatus = Part.Extension.SelectByID2("Point32", "SKETCHPOINT", -0.5461, 0.1143, 0, False, 0, Nothing, 0)

' Zoom In/Out (MouseWheel)
Dim swModelView As Object
Set swModelView = Part.ActiveView
swModelView.Scale2 = 0.596556083846222
Dim swTranslation() As Double
ReDim swTranslation(0 To 2) As Double
swTranslation(0) = 0.243805339304692
swTranslation(1) = -5.25878722859675E-02
swTranslation(2) = -3.71236850977504E-02
Dim swTranslationVar As Variant
swTranslationVar = swTranslation
Dim swMathUtils As Object
Set swMathUtils = swApp.GetMathUtility()
Dim swTranslationVector As MathVector
Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
swModelView.Translation3 = swTranslationVector

' Zoom In/Out (MouseWheel)
Set swModelView = Part.ActiveView
swModelView.Scale2 = 1.08530082723831
ReDim swTranslation(0 To 2) As Double
swTranslation(0) = 0.51633854748522
swTranslation(1) = -0.113862820117964
swTranslation(2) = -6.75382704790398E-02
swTranslationVar = swTranslation
Set swMathUtils = swApp.GetMathUtility()
Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
swModelView.Translation3 = swTranslationVector

' Zoom In/Out (MouseWheel)
Set swModelView = Part.ActiveView
swModelView.Scale2 = 1.30759135811844
ReDim swTranslation(0 To 2) As Double
swTranslation(0) = 0.644469224696654
swTranslation(1) = -0.142516590503571
swTranslation(2) = -8.13714102157107E-02
swTranslationVar = swTranslation
Set swMathUtils = swApp.GetMathUtility()
Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
swModelView.Translation3 = swTranslationVector

' Zoom In/Out (MouseWheel)
Set swModelView = Part.ActiveView
swModelView.Scale2 = 1.5754112748415
ReDim swTranslation(0 To 2) As Double
swTranslation(0) = 0.798843534589947
swTranslation(1) = -0.17703920542599
swTranslation(2) = -9.80378436333863E-02
swTranslationVar = swTranslation
Set swMathUtils = swApp.GetMathUtility()
Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
swModelView.Translation3 = swTranslationVector

' Zoom In/Out (MouseWheel)
Set swModelView = Part.ActiveView
swModelView.Scale2 = 1.89808587330301
ReDim swTranslation(0 To 2) As Double
swTranslation(0) = 0.984836679039698
swTranslation(1) = -0.21863271738071
swTranslation(2) = -0.118117883895646
swTranslationVar = swTranslation
Set swMathUtils = swApp.GetMathUtility()
Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))
swModelView.Translation3 = swTranslationVector
boolstatus = Part.Extension.SelectByRay(-0.544605317739567, 0.124459999999999, -0.119757252677115, 0, -1, 0, 4.57135911148266E-04, 2, False, 0, 0)

Dim myNote As Object
Dim myAnnotation As Object
Dim myTextFormat As Object
Set myNote = Part.InsertNote("312")
If Not myNote Is Nothing Then
   myNote.LockPosition = False
   myNote.Angle = 0
   boolstatus = myNote.SetBalloon(0, 0)
   Set myAnnotation = myNote.GetAnnotation()
   If Not myAnnotation Is Nothing Then
      longstatus = myAnnotation.SetLeader3(swLeaderStyle_e.swUNDERLINED, 0, True, False, False, False)
      boolstatus = myAnnotation.SetPosition(-0.544605317739567, 0.12446, -0.119757252677115)
      boolstatus = myAnnotation.SetTextFormat(0, True, myTextFormat)
   End If
End If
Part.ClearSelection2 True
Part.WindowRedraw
End Sub
