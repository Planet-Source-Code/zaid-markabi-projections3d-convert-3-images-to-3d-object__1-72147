VERSION 5.00
Begin VB.Form View3D 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " 3D View"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.ListBox TextureTemp 
      Height          =   2400
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   1680
   End
End
Attribute VB_Name = "View3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TV3d As New TVEngine
Dim World As New TVScene
Dim Object As TVMesh

Private Sub Command1_Click()
Timer1.Enabled = False
Object.Destroy
Set Object = Nothing
Set World = Nothing
Me.Hide
End Sub

Private Sub Form_Activate()
Set Object = Nothing
Set World = New TVScene

If frmMain.Check5.Value = 1 Then
World.SetRenderMode TV_LINE
Else
World.SetRenderMode TV_SOLID
End If

World.DestroyAllMeshes
Set Object = New TVMesh

Set Object = World.CreateMeshBuilder
Object.SetColor RGBA(1, 1, 1, 1), True

For i = 0 To frmMain.Faces.ListCount - 1

Dim X() As String
X() = Split(frmMain.Faces.List(i), " ")

If IsTextureAdded(X(3)) = False Then World.LoadTexture App.Path + "\Data\Colors\" + X(3) + ".bmp", , , X(3)

Dim PosX() As String
PosX() = Split(frmMain.Points.List(Int(X(0))), " ")
Dim PosY() As String
PosY() = Split(frmMain.Points.List(Int(X(1))), " ")
Dim PosZ() As String
PosZ() = Split(frmMain.Points.List(Int(X(2))), " ")

Object.AddTriangle GetTex(X(3)), PosX(0) - 50, PosX(1) - 50, PosX(2) - 50, PosY(0) - 50, PosY(1) - 50, PosY(2) - 50, PosZ(0) - 50, PosZ(1) - 50, PosZ(2) - 50

Next


Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Set TV3d = New TVEngine
Set World = New TVScene

Me.Hide
Picture1.Width = Me.Width
Picture1.Height = Me.Height + 220

TV3d.Init3DWindowedMode Picture1.hwnd
TV3d.SetAngleSystem TV_ANGLE_DEGREE

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
World.SetCamera -150 + ((-Me.Width \ 2) + X) \ 10, ((-Me.Height \ 2) + Y) \ 10, 0, 0, 0, 0
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
TV3d.Clear
World.RenderAllMeshes
Object.RotateY 3
Object.RotateX 0.6
Object.RotateZ 0.3
TV3d.RenderToScreen
End Sub
