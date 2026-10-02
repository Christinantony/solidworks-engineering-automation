Attribute VB_Name = "Create_holes1"
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
boolstatus = Part.Extension.SelectByRay(-5.95547179983011E-02, 6.00000000000023E-02, -1.16224882859694E-06, 0, -1, 0, 5.66812106804309E-04, 2, False, 0, 0)
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.14658723650489E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.79696643244879E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.79696643244879E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.14658723650489E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.79696643244879E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.14658723650489E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.79696643244879E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.79696643244879E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.44734562839269E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -2.09772482433659E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, 1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, 1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, 1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -4.54507045272928E-02, -4.67896778194226E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.74810402028049E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -1.39848321622439E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter -3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, -1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, -1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, -1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, -1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, -1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 2.09772482433659E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, -1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, -7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.74810402028049E-02, 1.55965592731409E-02
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 0.010488624121683, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 1.39848321622439E-02, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 6.99241608112197E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 7.79827963657043E-03
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0
Set myModelView = Part.ActiveView
myModelView.RotateAboutCenter 3.49620804056099E-03, 0

' Hole Wizard
Dim swHoleFeature As Feature
Set swHoleFeature = Part.FeatureManager.HoleWizard5(5, -1, -1, "", 0, 0.0254, 0.0635, 6, 0.0508, 2.05948851735331, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, "", False, True, True, True, True, False)
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
