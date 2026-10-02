Attribute VB_Name = "Relations"
Option Explicit

Private Const COORD_TOLERANCE As Double = 0.000000001

'===========================================================
' Build Coordinate Groups
'===========================================================

Public Function BuildCoordinateGroups() As Boolean

    On Error GoTo ErrorHandler

    Call BuildXGroups
    Call BuildYGroups

    BuildCoordinateGroups = True
    Exit Function

ErrorHandler:

    MsgBox _
        "Relation Group Error" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    BuildCoordinateGroups = False

End Function

'===========================================================
' X Groups
'===========================================================

Private Sub BuildXGroups()

    Dim Anchor As clsPoint
    Dim Candidate As clsPoint

    Dim lngGroup As Long

    lngGroup = 0

    For Each Anchor In gPoints

        If Anchor.XGroup = 0 Then

            lngGroup = lngGroup + 1
            Anchor.XGroup = lngGroup

            For Each Candidate In gPoints

                If Candidate.XGroup = 0 Then

                    If Abs(Candidate.X - Anchor.X) < COORD_TOLERANCE Then
                        Candidate.XGroup = lngGroup
                    End If

                End If

            Next Candidate

        End If

    Next Anchor

End Sub

'===========================================================
' Y Groups
'===========================================================

Private Sub BuildYGroups()

    Dim Anchor As clsPoint
    Dim Candidate As clsPoint

    Dim lngGroup As Long

    lngGroup = 0

    For Each Anchor In gPoints

        If Anchor.YGroup = 0 Then

            lngGroup = lngGroup + 1
            Anchor.YGroup = lngGroup

            For Each Candidate In gPoints

                If Candidate.YGroup = 0 Then

                    If Abs(Candidate.Y - Anchor.Y) < COORD_TOLERANCE Then
                        Candidate.YGroup = lngGroup
                    End If

                End If

            Next Candidate

        End If

    Next Anchor

End Sub

'===========================================================
' Create Sketch Relations
'===========================================================

Public Function CreateSketchRelations() As Boolean

    On Error GoTo ErrorHandler

    CreateVerticalRelations
    CreateHorizontalRelations

    CreateSketchRelations = True

    Exit Function

ErrorHandler:

    MsgBox _
        "Relation Creation Error" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    CreateSketchRelations = False

End Function

'===========================================================
' Vertical Relations
'===========================================================

Private Sub CreateVerticalRelations()

    Dim Anchor As clsPoint
    Dim Candidate As clsPoint

    Dim GroupIndex As Long
    Dim MaxGroup As Long

    MaxGroup = GetMaximumXGroup()

    For GroupIndex = 1 To MaxGroup

        Set Anchor = Nothing

        For Each Candidate In gPoints

            If Candidate.XGroup = GroupIndex Then
                Set Anchor = Candidate
                Exit For
            End If

        Next Candidate

        If Anchor Is Nothing Then GoTo NextGroup

        For Each Candidate In gPoints

            If Candidate.XGroup = GroupIndex Then

                If Not Candidate Is Anchor Then

                    swModel.ClearSelection2 True

                    Anchor.SketchPoint.Select4 False, Nothing
                    Candidate.SketchPoint.Select4 True, Nothing
Debug.Print swModel.SelectionManager.GetSelectedObjectCount2(-1)
                    swModel.SketchAddConstraints "sgVERTICALPOINTS2D"

                End If

            End If

        Next Candidate

NextGroup:

    Next GroupIndex

End Sub

'===========================================================
' Horizontal Relations
'===========================================================

Private Sub CreateHorizontalRelations()

    Dim Anchor As clsPoint
    Dim Candidate As clsPoint

    Dim GroupIndex As Long
    Dim MaxGroup As Long

    MaxGroup = GetMaximumYGroup()

    For GroupIndex = 1 To MaxGroup

        Set Anchor = Nothing

        For Each Candidate In gPoints

            If Candidate.YGroup = GroupIndex Then
                Set Anchor = Candidate
                Exit For
            End If

        Next Candidate

        If Anchor Is Nothing Then GoTo NextGroup

        For Each Candidate In gPoints

            If Candidate.YGroup = GroupIndex Then

                If Not Candidate Is Anchor Then

                    swModel.ClearSelection2 True

                    Anchor.SketchPoint.Select4 False, Nothing
                    Candidate.SketchPoint.Select4 True, Nothing
Debug.Print swModel.SelectionManager.GetSelectedObjectCount2(-1)
                    swModel.SketchAddConstraints "sgHORIZONTALPOINTS2D"

                End If

            End If

        Next Candidate

NextGroup:

    Next GroupIndex

End Sub

'===========================================================
' Helpers
'===========================================================

Private Function GetMaximumXGroup() As Long

    Dim pt As clsPoint

    For Each pt In gPoints

        If pt.XGroup > GetMaximumXGroup Then
            GetMaximumXGroup = pt.XGroup
        End If

    Next pt

End Function

Private Function GetMaximumYGroup() As Long

    Dim pt As clsPoint

    For Each pt In gPoints

        If pt.YGroup > GetMaximumYGroup Then
            GetMaximumYGroup = pt.YGroup
        End If

    Next pt

End Function
