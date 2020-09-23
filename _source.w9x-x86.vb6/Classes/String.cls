VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--STRUCTURES--
' Private:
    Private Type POINT_DESC
        Position As POINTF
        Locked As Boolean
    End Type
    Private Type POINTS_DESC
        Count As Long
        Item() As POINT_DESC
    End Type

'--VARIABLES--
' Private:
    Private i_Gravity As Single
    Private i_MaxLength As Single
    Private i_Points As POINTS_DESC
    ' For Collision Detection:
    Private i_SceneWidth As Long
    Private i_SceneHeight As Long
    ' For Statistics Only:
    Private i_CurLength As Single
Public Sub Continue(ByVal DeltaTime As Single)
    On Error Resume Next
    
    ' Calculate the maximum length of a connection line.
    Dim fPartLength As Single
    fPartLength = (i_MaxLength / i_Points.Count)
    
    Dim vecDistance(0 To 1) As POINTF
    Dim fDistance(0 To 1) As Single
    Dim fIntensity(0 To 1) As Single
    Dim computed_force As POINTF
    
    ' For statistics purposes.
    i_CurLength = 0#
    
    ' Go through each point ...
    Dim dwIndex As Long
    For dwIndex = 0 To (i_Points.Count - 1)
        ' ... Get the distance between this point and the previous ...
        If dwIndex = 0 Then
            vecDistance(0) = PointFSet(0#, 0#)
        Else
            vecDistance(0) = PointFSubtract(i_Points.Item(dwIndex - 1).Position, i_Points.Item(dwIndex).Position)
        End If
        fDistance(0) = PointFMagnitude(vecDistance(0))
        If fDistance(0) = 0 Then
            fIntensity(0) = 0#
        Else
            fIntensity(0) = (fDistance(0) - fPartLength) / fDistance(0)
        End If
        
        ' ... Get the distance between this point and the next ...
        If dwIndex < (i_Points.Count - 1) Then
            vecDistance(1) = PointFSubtract(i_Points.Item(dwIndex + 1).Position, i_Points.Item(dwIndex).Position)
        Else
            vecDistance(1) = PointFSet(0#, 0#)
        End If
        fDistance(1) = PointFMagnitude(vecDistance(1))
        If fDistance(1) = 0 Then
            fIntensity(1) = 0#
        Else
            fIntensity(1) = (fDistance(1) - fPartLength) / fDistance(1)
        End If
        
        computed_force = PointFMultiplyValue(vecDistance(0), fIntensity(0))
        computed_force = PointFAdd(computed_force, PointFMultiplyValue(vecDistance(1), fIntensity(1)))
        computed_force = PointFAdd(computed_force, PointFSet(0#, (-i_Gravity * DeltaTime)))
        
        ' Add the force unto the point.
        If i_Points.Item(dwIndex).Locked = False Then
            i_Points.Item(dwIndex).Position = PointFAdd(i_Points.Item(dwIndex).Position, computed_force)
            Call CheckPointCollisions(dwIndex)
        End If
        
        ' ... STATISTICS ONLY: Store the length into the curent length.
        i_CurLength = (i_CurLength + fDistance(0))
    Next dwIndex
End Sub
Public Function GetCurLength() As Single
    On Error Resume Next
    
    GetCurLength = i_CurLength
End Function
Public Sub Initialize(ByVal MaximumLength As Single, ByVal PointsCount As Long)
    On Error Resume Next
    
    ' Store the rope's maximum length (only effected in the loop).
    Call SetMaximumLength(MaximumLength)
    
    ' Set the rope's connection points count.
    Call SetPointsCount(PointsCount)
End Sub
Private Function IsPointValid(ByVal Index As Long) As Boolean
    On Error Resume Next
    
    If Index < 0 Or Index > (i_Points.Count - 1) Then
        IsPointValid = False
    Else
        IsPointValid = True
    End If
End Function
Public Sub SetGravity(ByVal newValue As Single)
    On Error Resume Next
    
    i_Gravity = newValue
End Sub
Public Sub SetMaximumLength(ByVal newValue As Single)
    On Error Resume Next
    
    i_MaxLength = newValue
End Sub
Public Function GetMaximumLength() As Single
    On Error Resume Next
    
    GetMaximumLength = i_MaxLength
End Function
Public Sub SetPointsCount(ByVal newValue As Long)
    On Error Resume Next
    
    If Not i_Points.Count = newValue Then
        i_Points.Count = newValue
        
        If i_Points.Count = 0 Then
            Erase i_Points.Item
        Else
            ReDim Preserve i_Points.Item(0 To (i_Points.Count - 1))
        End If
    End If
End Sub
Public Function GetPointsCount() As Long
    On Error Resume Next
    
    GetPointsCount = i_Points.Count
End Function
Public Sub SetPointPosition(ByVal PointIndex As Long, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    
    If IsPointValid(PointIndex) = False Then Exit Sub
    
    i_Points.Item(PointIndex).Position = PointFSet(x, y)
End Sub
Public Sub SetPointIsLocked(ByVal PointIndex As Long, ByVal IsLocked As Boolean)
    On Error Resume Next
    
    If IsPointValid(PointIndex) = False Then Exit Sub
    
    i_Points.Item(PointIndex).Locked = IsLocked
End Sub
Public Function GetPointIsLocked(ByVal PointIndex As Long) As Boolean
    On Error Resume Next
    
    If IsPointValid(PointIndex) = False Then Exit Function
    
    GetPointIsLocked = i_Points.Item(PointIndex).Locked
End Function
Public Sub GetPointPosition(ByVal PointIndex As Long, Optional ByRef outX As Single, Optional ByRef outY As Single)
    On Error Resume Next
    
    If IsPointValid(PointIndex) = False Then Exit Sub
    
    outX = i_Points.Item(PointIndex).Position.x
    outY = i_Points.Item(PointIndex).Position.y
End Sub
Public Sub SetSceneSize(ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    
    i_SceneWidth = Width
    i_SceneHeight = Height
End Sub
Private Sub CheckPointCollisions(ByVal Index As Long)
    On Error Resume Next
    
    If i_Points.Item(Index).Position.y < 1 Then
        i_Points.Item(Index).Position.y = 1
    ElseIf i_Points.Item(Index).Position.y > i_SceneHeight Then
        i_Points.Item(Index).Position.y = i_SceneHeight
    End If
    
    If i_Points.Item(Index).Position.x < 0 Then
        i_Points.Item(Index).Position.x = 0
    ElseIf i_Points.Item(Index).Position.x > i_SceneWidth Then
        i_Points.Item(Index).Position.x = i_SceneWidth
    End If
End Sub
