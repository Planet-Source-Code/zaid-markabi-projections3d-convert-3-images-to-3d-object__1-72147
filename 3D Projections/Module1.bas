Attribute VB_Name = "Global"
Global StopExtracting As Boolean

Function IsDidX(Value As String) As Boolean
For i = 0 To frmMain.Points.ListCount - 1
If frmMain.Points.List(i) = Value Then GoTo 2
Next
IsDidX = False
Exit Function
2:
IsDidX = True
End Function

Function IsFaceAdded(Value As String) As Boolean
For i = 0 To frmMain.FacesTemp.ListCount - 1
If frmMain.FacesTemp.List(i) = Value Then GoTo 2
Next
IsFaceAdded = False
Exit Function
2:
IsFaceAdded = True
End Function

Function IsTextureAdded(Value As String) As Boolean
For i = 0 To View3D.TextureTemp.ListCount - 1
If View3D.TextureTemp.List(i) = Value Then GoTo 2
Next
IsTextureAdded = False
Exit Function
2:
IsTextureAdded = True
End Function

Function IsFaceAddedX(Value As String) As Boolean
For i = 0 To frmMain.ListX.ListCount - 1
If frmMain.ListX.List(i) = Value Then GoTo 2
Next
IsFaceAddedX = False
Exit Function
2:
IsFaceAddedX = True
End Function

Function IsFaceAddedy(Value As String) As Boolean
For i = 0 To frmMain.ListY.ListCount - 1
If frmMain.ListY.List(i) = Value Then GoTo 2
Next
IsFaceAddedy = False
Exit Function
2:
IsFaceAddedy = True
End Function

Function IsFaceAddedz(Value As String) As Boolean
For i = 0 To frmMain.ListZ.ListCount - 1
If frmMain.ListZ.List(i) = Value Then GoTo 2
Next
IsFaceAddedz = False
Exit Function
2:
IsFaceAddedz = True
End Function
