VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' SW Point Importer
'
' Class Module : clsPoint
' Revision     : 1
'
'===========================================================

Option Explicit

Private mID As String

Private mXInch As Double
Private mYInch As Double

Private mX As Double
Private mY As Double

Private mSketchPoint As SldWorks.SketchPoint

Private mXGroup As Long
Private mYGroup As Long

Private mHasXOrdinate As Boolean
Private mHasYOrdinate As Boolean

Public Property Get ID() As String
    ID = mID
End Property

Public Property Let ID(ByVal Value As String)
    mID = Value
End Property

Public Property Get XInch() As Double
    XInch = mXInch
End Property

Public Property Let XInch(ByVal Value As Double)
    mXInch = Value
End Property

Public Property Get YInch() As Double
    YInch = mYInch
End Property

Public Property Let YInch(ByVal Value As Double)
    mYInch = Value
End Property

Public Property Get X() As Double
    X = mX
End Property

Public Property Let X(ByVal Value As Double)
    mX = Value
End Property

Public Property Get Y() As Double
    Y = mY
End Property

Public Property Let Y(ByVal Value As Double)
    mY = Value
End Property

Public Property Get SketchPoint() As SldWorks.SketchPoint
    Set SketchPoint = mSketchPoint
End Property

Public Property Set SketchPoint(ByVal Value As SldWorks.SketchPoint)
    Set mSketchPoint = Value
End Property

Public Property Get XGroup() As Long
    XGroup = mXGroup
End Property

Public Property Let XGroup(ByVal Value As Long)
    mXGroup = Value
End Property

Public Property Get YGroup() As Long
    YGroup = mYGroup
End Property

Public Property Let YGroup(ByVal Value As Long)
    mYGroup = Value
End Property

Public Property Get HasXOrdinate() As Boolean
    HasXOrdinate = mHasXOrdinate
End Property

Public Property Let HasXOrdinate(ByVal Value As Boolean)
    mHasXOrdinate = Value
End Property

Public Property Get HasYOrdinate() As Boolean
    HasYOrdinate = mHasYOrdinate
End Property

Public Property Let HasYOrdinate(ByVal Value As Boolean)
    mHasYOrdinate = Value
End Property
