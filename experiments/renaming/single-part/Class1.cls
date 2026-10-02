VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents swPart As SldWorks.PartDoc
Attribute swPart.VB_VarHelpID = -1

Public Function swPart_RenamedDocumentNotify(ByRef swObj As Object) As Long
    Dim swRef As SldWorks.RenamedDocumentReferences
    Set swRef = swObj

    swRef.UpdateWhereUsedReferences = True    ' forces the checkbox on
    swRef.IncludeFileLocations = True
    swRef.Search
    swRef.CompletionAction = swRenamedDocumentFinalAction_e.swRenamedDocumentFinalAction_Ok
End Function
