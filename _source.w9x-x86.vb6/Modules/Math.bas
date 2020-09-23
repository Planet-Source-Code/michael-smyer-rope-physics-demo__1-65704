Attribute VB_Name = "basMath"
Option Explicit

'--STRUCTURES--
' Public:
    Public Type POINTF  'My custom structure, float version of the original 'POINTAPI'
        x As Single
        y As Single
    End Type
Public Function MIN_float(ByVal a As Single, ByVal b As Single) As Single
    On Error Resume Next
    
    If a <= b Then
        MIN_float = a
    Else
        MIN_float = b
    End If
End Function
Public Function MAX_float(ByVal a As Single, ByVal b As Single) As Single
    On Error Resume Next
    
    If a >= b Then
        MAX_float = a
    Else
        MAX_float = b
    End If
End Function
Public Function PointFSet(ByVal x As Single, ByVal y As Single) As POINTF
    On Error Resume Next
    
    PointFSet.x = x
    PointFSet.y = y
End Function
Public Function PointFSubtract(a As POINTF, b As POINTF) As POINTF
    On Error Resume Next
    
    PointFSubtract.x = (a.x - b.x)
    PointFSubtract.y = (a.y - b.y)
End Function
Public Function PointFAdd(a As POINTF, b As POINTF) As POINTF
    On Error Resume Next
    
    PointFAdd.x = (a.x + b.x)
    PointFAdd.y = (a.y + b.y)
End Function
' Also known as the length of a vector, or the dot product of the vector.
Public Function PointFMagnitude(a As POINTF) As Single
    On Error Resume Next
    
    PointFMagnitude = CSng(Sqr((a.x * a.x) + (a.y * a.y)))
End Function
Public Function PointFMultiplyValue(a As POINTF, ByVal b As Single) As POINTF
    On Error Resume Next
    
    With PointFMultiplyValue
        .x = (a.x * b)
        .y = (a.y * b)
    End With
End Function
