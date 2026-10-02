Attribute VB_Name = "Globals"
'===========================================================
' SW Point Importer
'
' Module    : Globals
' Revision  : 1
'
' Compatible with:
' SolidWorks 2024
'
'===========================================================

Option Explicit

'-----------------------------------------------------------
' SolidWorks Objects
'-----------------------------------------------------------

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2
Public swPart As SldWorks.PartDoc
Public swSketchMgr As SldWorks.SketchManager
Public gOriginX As Double
Public gOriginY As Double

'-----------------------------------------------------------
' Global Data
'-----------------------------------------------------------

Public gPoints As Collection

Public gExcelFile As String

Public Const INCH_TO_METER As Double = 0.0254

Public Const DEBUG_MODE As Boolean = True

'-----------------------------------------------------------
' Initialize Globals
'-----------------------------------------------------------

Public Function InitializeGlobals() As Boolean

    On Error GoTo ErrorHandler

    Set swApp = Application.SldWorks

    If swApp Is Nothing Then Exit Function

    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then Exit Function

    If swModel.GetType <> swDocPART Then Exit Function

    Set swPart = swModel

    Set swSketchMgr = swModel.SketchManager

    If swSketchMgr Is Nothing Then Exit Function

    If swSketchMgr.ActiveSketch Is Nothing Then Exit Function

    Set gPoints = New Collection

    InitializeGlobals = True

    Exit Function

ErrorHandler:

    InitializeGlobals = False

End Function

'-----------------------------------------------------------
' Cleanup
'-----------------------------------------------------------

Public Sub ClearGlobals()

    Set gPoints = Nothing

    gExcelFile = ""

    gOriginX = 0#
    gOriginY = 0#

End Sub
