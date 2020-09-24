VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6360
      Width           =   975
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   6480
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "* New Project"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6420
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stop"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2D Object"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   6120
      Width           =   4575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remove Unimportant Faces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   5760
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remove Unimportant Points"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   5400
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   5280
      ScaleHeight     =   135
      ScaleWidth      =   3135
      TabIndex        =   24
      Top             =   6900
      Width           =   3135
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   135
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hide Lines in 3D Render"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   5040
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   1320
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox Proj 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   1980
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3D View"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   4695
   End
   Begin VB.ListBox ListZ 
      Height          =   645
      Left            =   4680
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListY 
      Height          =   645
      Left            =   3600
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListX 
      Height          =   645
      Left            =   3600
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox FacesTemp 
      Height          =   2400
      Left            =   6480
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox Faces 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2760
      Left            =   6000
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox ColorCap 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1800
      ScaleHeight     =   1545
      ScaleWidth      =   1545
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Extract points and faces"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   4095
   End
   Begin VB.ListBox Points 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2760
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox PictureZ 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.PictureBox PictureY 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   1800
      Picture         =   "Form1.frx":7572
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox PictureX 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":EAE4
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   33
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " 3D Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   28
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Projects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Status 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   6840
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Faces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   18
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Points"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   17
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Right Side"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Front"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " 3D Projections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1_Click
End Sub

Private Sub Command1_Click()
On Error Resume Next

StopExtracting = False

Dim PosX, PosY, PosZ, PosW As Integer
Dim BackPosX, BackPosY, BackPosZ, BackPosW, BackPosU As Integer
Dim PosXCol, PosYCol, PosZCol As Long
Dim OldPosXCol, OldPosYCol, OldPosZCol As Long

Dim FaceColor, FaceColorB, FaceColorC As Long
Dim Face As String

Points.Clear
Faces.Clear
FacesTemp.Clear

Faces.Visible = False
Points.Visible = False

Command2.Enabled = False

Status.Caption = " Extracting Points ..."
Picture1.ScaleWidth = 99 * 99
Picture2.Width = 0
Picture2.Visible = True
' Extract Points
For PosY = 0 To 99
DoEvents
If StopExtracting = True Then GoTo 32

OldPosXCol = PictureX.Point(0, 0)
OldPosYCol = PictureY.Point(0, 0)
OldPosZCol = PictureZ.Point(0, 0)

For PosX = 1 To 99
Picture2.Width = Picture2.Width + 1
DoEvents
PosXCol = PictureX.Point(PosX, PosY)
BackPosW = 0

If PosXCol = PictureX.Point(0, 0) Then BackPosX = -1 Else BackPosX = 0

If Not PictureX.Point(PosX + BackPosX - 1, PosY - 1) = PictureX.Point(PosX + BackPosX, PosY - 1) And Check4.Value = 0 Then
PosXCol = PictureX.Point(PosX, PosY - 1)
OldPosXCol = PictureX.Point(0, 0)
BackPosU = -1
End If

If Not PictureX.Point(PosX + 1, PosY + 1) = PictureX.Point(PosX, PosY + 1) And Check4.Value = 0 Then
PosXCol = PictureX.Point(PosX + 1, PosY + 1)
OldPosXCol = PictureX.Point(0, 0)
BackPosU = 1
End If
If Not PictureX.Point(PosX + 1, PosY - 1) = PictureX.Point(PosX, PosY - 1) And Check4.Value = 0 Then
PosXCol = PictureX.Point(PosX + 1, PosY - 1)
OldPosXCol = PictureX.Point(0, 0)
BackPosU = 1
End If

If Not PictureX.Point(PosX - 1, PosY - 1) = PictureX.Point(PosX, PosY - 1) And Check4.Value = 0 Then
PosXCol = PictureX.Point(PosX - 1, PosY - 1)
OldPosXCol = PictureX.Point(0, 0)
BackPosU = -1
End If

If Not PictureX.Point(PosX - 1, PosY) = PictureX.Point(PosX, PosY) And Not PictureX.Point(PosX, PosY) = PictureX.Point(PosX, PosY - 1) And Check4.Value = 0 Then
PosXCol = PictureX.Point(PosX, PosY)
OldPosXCol = PictureX.Point(0, 0)
BackPosU = -1
End If

If PictureX.Point(PosX + BackPosX, PosY) = PictureX.Point(PosX + BackPosX, PosY - 1) Then OldPosXCol = PosXCol

If Not OldPosXCol = PosXCol Then
PosXCol = PictureX.Point(PosX + BackPosX, PosY)

For PosZ = 1 To 99
DoEvents
PosZCol = PictureZ.Point(PosX + BackPosX, PosZ)
If PosZCol = PictureZ.Point(0, 0) Then BackPosZ = -1 Else BackPosZ = 0

If Not PosZCol = OldPosZCol Then
PosZCol = PictureZ.Point(PosX + BackPosX, PosZ + BackPosZ)

For PosW = 1 To 99
DoEvents
PosYCol = PictureY.Point(PosZ + BackPosZ, PosW)
If PosYCol = PictureY.Point(0, 0) Then BackPosY = -1 Else BackPosY = 0

If Not PosYCol = OldPosYCol Then
PosYCol = PictureY.Point(PosZ + BackPosZ, PosW + BackPosY)

FaceColor = PosXCol
FaceColorB = PictureX.Point(0, 0)
FaceColorC = PictureX.Point(0, 0)
If FaceColor = PictureX.Point(0, 0) Then

If Not PosYCol = PictureY.Point(0, 0) And PosZCol = PictureZ.Point(0, 0) Then
FaceColor = PosYCol
End If
If Not PosZCol = PictureZ.Point(0, 0) And PosYCol = PictureY.Point(0, 0) Then
FaceColor = PosZCol
End If
If Not PosYCol = PictureY.Point(0, 0) And Not PosZCol = PictureZ.Point(0, 0) Then
FaceColor = PosYCol
FaceColorB = PosZCol
End If
If Not PictureX.Point(PosX - 1, PosY - 1) = PictureX.Point(0, 0) Then
FaceColorC = PictureX.Point(PosX - 1, PosY - 1)
Else
FaceColorC = PictureX.Point(PosX, PosY - 1)
End If

Else
If Not PosYCol = PictureY.Point(0, 0) Then
FaceColorB = PosYCol
End If
If Not PosZCol = PictureZ.Point(0, 0) Then
FaceColorC = PosZCol
End If
End If

BackPosW = 0
If PictureX.Point(PosX, PosY) = PictureY.Point(0, 0) And Not PictureX.Point(PosX + 1, PosY) = PictureX.Point(0, 0) Then BackPosW = 1
If PictureX.Point(PosX - 2, PosY + BackPosU) = PictureY.Point(0, 0) And Not PictureX.Point(PosX - 1, PosY + BackPosU) = PictureX.Point(0, 0) Then BackPosW = -1

Face = Format(PosX + BackPosW, "000") + " " + Format(PosY, "000") + " " + Format(PosZ, "000") + " " + Format(FaceColor)
If IsDidX(Face) = False Then
Points.AddItem Face
End If

If Not FaceColorB = PictureX.Point(0, 0) Then
Face = Format(PosX + BackPosW, "000") + " " + Format(PosY, "000") + " " + Format(PosZ, "000") + " " + Format(FaceColorB)
If IsDidX(Face) = False Then
ColorCap.BackColor = FaceColorB
Points.AddItem Face
End If
End If
If Not FaceColorC = PictureX.Point(0, 0) Then
If PictureX.Point(PosX, PosY) = PictureX.Point(0, 0) Then BackPosX = -1 Else BackPosX = 0
Face = Format(PosX + BackPosW, "000") + " " + Format(PosY, "000") + " " + Format(PosZ, "000") + " " + Format(FaceColorC)
If IsDidX(Face) = False Then
ColorCap.BackColor = FaceColorC
Points.AddItem Face
End If
End If

End If


OldPosYCol = PictureY.Point(PosZ + BackPosZ, PosW + BackPosY)
Next

End If

OldPosZCol = PictureZ.Point(PosX + BackPosX, PosZ + BackPosZ)
Next
End If

OldPosXCol = PictureX.Point(PosX + BackPosX, PosY)

DoEvents
Next

Next

If Check2.Value = 1 Then
' Delete Unimportant Points
Dim Ipoint As Integer
For Ipoint = 0 To Points.ListCount - 1
DoEvents
If StopExtracting = True Then GoTo 32
Dim XPoint() As String
XPoint() = Split(Points.List(Ipoint), " ")
If CLng(XPoint(3)) = PictureX.Point(0, 0) Then Points.RemoveItem Ipoint
Next
End If

Status.Caption = " Extracting Faces ..."
Picture1.ScaleWidth = 99
Picture2.Width = 0
DoEvents
' Extract Faces
Dim i, II, III As Integer
Dim p() As String
Dim PP() As String
Dim PPP() As String

For i = 0 To Points.ListCount - 1
Picture2.Width = Picture2.Width + 1
For II = i + 1 To Points.ListCount - 1
For III = II + 1 To Points.ListCount - 1

DoEvents
If StopExtracting = True Then GoTo 32

p() = Split(Points.List(i), " ")
PP() = Split(Points.List(II), " ")
PPP() = Split(Points.List(III), " ")

If p(3) = PP(3) And PP(3) = PPP(3) And ((p(0) = PP(0) And PP(0) = PPP(0)) Or (p(1) = PP(1) And PP(1) = PPP(1)) Or (p(2) = PP(2) And PP(2) = PPP(2))) Then

FaceColor = p(3)

Face = Format(i, "000") + " " + Format(II, "000") + " " + Format(III, "000") + " " + Format(FaceColor)
If IsFaceAdded(Face) = False Then
Faces.AddItem Face
FacesTemp.AddItem Face
Face = Format(i, "000") + " " + Format(III, "000") + " " + Format(II, "000") + " " + Format(FaceColor)
FacesTemp.AddItem Face
Face = Format(II, "000") + " " + Format(i, "000") + " " + Format(III, "000") + " " + Format(FaceColor)
FacesTemp.AddItem Face
Face = Format(II, "000") + " " + Format(III, "000") + " " + Format(i, "000") + " " + Format(FaceColor)
FacesTemp.AddItem Face
Face = Format(III, "000") + " " + Format(i, "000") + " " + Format(II, "000") + " " + Format(FaceColor)
FacesTemp.AddItem Face
Face = Format(III, "000") + " " + Format(II, "000") + " " + Format(i, "000") + " " + Format(FaceColor)
FacesTemp.AddItem Face
End If
End If

Next
Next
Next

If Check3.Value = 1 Then
' Delete Unimportant Faces
Picture1.ScaleWidth = Faces.ListCount
Picture2.Width = 0

For i = 0 To Faces.ListCount - 1

Picture2.Width = Picture2.Width + 1
DoEvents
If StopExtracting = True Then GoTo 32

Dim XX() As String
XX() = Split(Faces.List(i), " ")

PictureX.Cls
PictureY.Cls
PictureZ.Cls

ListX.Clear
ListY.Clear
ListZ.Clear

If i < Faces.ListCount Then
Points.ListIndex = XX(0)
Points_Click
Points.ListIndex = XX(1)
Points_Click
Points.ListIndex = XX(2)
Points_Click
End If

PictureX.Cls
PictureY.Cls
PictureZ.Cls

Dim Col1, Col2, Col3 As Long
Dim XXX() As String

If ListX.ListCount = 3 Then
XXX() = Split(ListX.List(0), " ")
Col1 = PictureX.Point(XXX(0), XXX(1))
If Col1 = PictureX.Point(0, 0) Then Col1 = PictureX.Point(Int(XXX(0)) - 1, XXX(1))
If Col1 = PictureX.Point(0, 0) Then Col1 = PictureX.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col1 = PictureX.Point(0, 0) Then Col1 = PictureX.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
XXX() = Split(ListX.List(1), " ")
Col2 = PictureX.Point(XXX(0), XXX(1))
If Col2 = PictureX.Point(0, 0) Then Col2 = PictureX.Point(Int(XXX(0)) - 1, XXX(1))
If Col2 = PictureX.Point(0, 0) Then Col2 = PictureX.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col2 = PictureX.Point(0, 0) Then Col2 = PictureX.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
XXX() = Split(ListX.List(2), " ")
Col3 = PictureX.Point(XXX(0), XXX(1))
If Col3 = PictureX.Point(0, 0) Then Col3 = PictureX.Point(Int(XXX(0)) - 1, XXX(1))
If Col3 = PictureX.Point(0, 0) Then Col3 = PictureX.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col3 = PictureX.Point(0, 0) Then Col3 = PictureX.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
'ListX.BackColor = Col1
'ListY.BackColor = Col2
'ListZ.BackColor = Col3
If Col1 = Col2 And Col2 = Col3 And Not Col3 = CLng(XX(3)) And i > -1 Then Faces.RemoveItem i: i = i - 1
End If

If ListZ.ListCount = 3 Then
XXX() = Split(ListZ.List(0), " ")
Col1 = PictureY.Point(XXX(0), XXX(1))
If Col1 = PictureY.Point(0, 0) Then Col1 = PictureY.Point(Int(XXX(0)) - 1, XXX(1))
If Col1 = PictureY.Point(0, 0) Then Col1 = PictureY.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col1 = PictureY.Point(0, 0) Then Col1 = PictureY.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
XXX() = Split(ListZ.List(1), " ")
Col2 = PictureY.Point(XXX(0), XXX(1))
If Col2 = PictureY.Point(0, 0) Then Col2 = PictureY.Point(Int(XXX(0)) - 1, XXX(1))
If Col2 = PictureY.Point(0, 0) Then Col2 = PictureY.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col2 = PictureY.Point(0, 0) Then Col2 = PictureY.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
XXX() = Split(ListZ.List(2), " ")
Col3 = PictureY.Point(XXX(0), XXX(1))
If Col3 = PictureY.Point(0, 0) Then Col3 = PictureY.Point(Int(XXX(0)) - 1, XXX(1))
If Col3 = PictureY.Point(0, 0) Then Col3 = PictureY.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col3 = PictureY.Point(0, 0) Then Col3 = PictureY.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
'ListX.BackColor = Col1
'ListY.BackColor = Col2
'ListZ.BackColor = Col3
If Col1 = Col2 And Col2 = Col3 And Not Col3 = CLng(XX(3)) And i > -1 Then Faces.RemoveItem i: i = i - 1
End If

If ListY.ListCount = 3 Then
XXX() = Split(ListY.List(0), " ")
Col1 = PictureZ.Point(XXX(0), XXX(1))
If Col1 = PictureZ.Point(0, 0) Then Col1 = PictureZ.Point(Int(XXX(0)) - 1, XXX(1))
If Col1 = PictureZ.Point(0, 0) Then Col1 = PictureZ.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col1 = PictureZ.Point(0, 0) Then Col1 = PictureZ.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
XXX() = Split(ListY.List(1), " ")
Col2 = PictureZ.Point(XXX(0), XXX(1))
If Col2 = PictureZ.Point(0, 0) Then Col2 = PictureZ.Point(Int(XXX(0)) - 1, XXX(1))
If Col2 = PictureZ.Point(0, 0) Then Col2 = PictureZ.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col2 = PictureZ.Point(0, 0) Then Col2 = PictureZ.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
XXX() = Split(ListY.List(2), " ")
Col3 = PictureZ.Point(XXX(0), XXX(1))
If Col3 = PictureZ.Point(0, 0) Then Col3 = PictureZ.Point(Int(XXX(0)) - 1, XXX(1))
If Col3 = PictureZ.Point(0, 0) Then Col3 = PictureZ.Point(Int(XXX(0)), Int(XXX(1)) - 1)
If Col3 = PictureZ.Point(0, 0) Then Col3 = PictureZ.Point(Int(XXX(0)) - 1, Int(XXX(1)) - 1)
'ListX.BackColor = Col1
'ListY.BackColor = Col2
'ListZ.BackColor = Col3
If Col1 = Col2 And Col2 = Col3 And Not Col3 = CLng(XX(3)) And i > -1 Then Faces.RemoveItem i: i = i - 1
End If

Next

End If

32:
Faces.Visible = True
Points.Visible = True

Picture2.Visible = False

Command2.Enabled = True

Status.Caption = " Extract Colors ..."
' Save Colors as Images
For i = 0 To Faces.ListCount - 1
XX() = Split(Faces.List(i), " ")
ColorCap.BackColor = CLng(XX(3))

DoEvents
SavePicture CaptureScreen((Me.Left \ 15.4) + ColorCap.Left + (Check1.Value * 10), (Me.Top \ 15.4) + ColorCap.Top + (Check1.Value * 10), 32, 32), App.Path + "\Data\Colors\" + XX(3) + ".bmp"
DoEvents

Next

Status.Caption = " Ready"
Label1(6).Caption = " Points (" + Format(Points.ListCount) + ")"
Label1(7).Caption = "  Faces (" + Format(Faces.ListCount) + ")"
End Sub

Private Sub Command2_Click()
View3D.Show
End Sub

Private Sub Command3_Click()
StopExtracting = True
End Sub

Private Sub Command4_Click()
frmAbout.Show 1
End Sub

Private Sub Command5_Click()
frmNewProject.Show 1
Proj.ListIndex = Proj.ListCount - 1
End Sub

Private Sub Faces_Click()
Dim XX() As String
XX() = Split(Faces.List(Faces.ListIndex), " ")

PictureX.Cls
PictureY.Cls
PictureZ.Cls

ListX.Clear
ListY.Clear
ListZ.Clear

Points.ListIndex = XX(0)
Points_Click
Points.ListIndex = XX(1)
Points_Click
Points.ListIndex = XX(2)
Points_Click
End Sub

Private Sub Form_Load()
Dir1.Path = App.Path + "\Data\Projects\"

Dim i As Integer
For i = 0 To Dir1.ListCount - 1
Proj.AddItem Right(Dir1.List(i), Len(Dir1.List(i)) - Len(Dir1.Path) - 1)
Next
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Points_Click()
Dim XX() As String
XX() = Split(Points.List(Points.ListIndex), " ")

PictureX.PSet (XX(0), XX(1))
PictureZ.PSet (XX(0), XX(2))
PictureY.PSet (XX(2), XX(1))

If IsFaceAddedX(XX(0) + " " + XX(1)) = False Then ListX.AddItem XX(0) + " " + XX(1)
If IsFaceAddedy(XX(0) + " " + XX(2)) = False Then ListY.AddItem XX(0) + " " + XX(2)
If IsFaceAddedz(XX(2) + " " + XX(1)) = False Then ListZ.AddItem XX(2) + " " + XX(1)

ColorCap.BackColor = CLng(XX(3))
DoEvents
End Sub

Private Sub Points_KeyUp(KeyCode As Integer, Shift As Integer)
PictureX.Cls
PictureY.Cls
PictureZ.Cls
Points_Click
ListX.Clear
ListY.Clear
ListZ.Clear
End Sub

Private Sub Points_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureX.Cls
PictureY.Cls
PictureZ.Cls
Points_Click
ListX.Clear
ListY.Clear
ListZ.Clear
End Sub

Private Sub Proj_Click()
On Error Resume Next
PictureX.Picture = LoadPicture(App.Path + "\Data\Projects\" + Proj.List(Proj.ListIndex) + "\Front.bmp")
PictureY.Picture = LoadPicture(App.Path + "\Data\Projects\" + Proj.List(Proj.ListIndex) + "\Right.bmp")
PictureZ.Picture = LoadPicture(App.Path + "\Data\Projects\" + Proj.List(Proj.ListIndex) + "\Top.bmp")
End Sub
