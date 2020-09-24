VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewProject 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "* Save"
      Height          =   315
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog OPN 
      Left            =   3720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "* Load"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "* Load"
      Height          =   315
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "* Load"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
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
      Picture         =   "frmNewProject.frx":0000
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   600
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
      Picture         =   "frmNewProject.frx":7572
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox PictureZ 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      Picture         =   "frmNewProject.frx":EAE4
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "4- Simplest projection should be in Front image ."
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   15
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3- Background can be any color (Black,Red,Green,...)"
      Height          =   495
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2- Images Size = ( 100 * 100 ) Pixels ."
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1- Import 3 Images ( Front,Side,Top )"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
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
      Left            =   4440
      TabIndex        =   10
      Top             =   0
      Width           =   255
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
      TabIndex        =   6
      Top             =   360
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
      TabIndex        =   5
      Top             =   360
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
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " New Project"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo 1
OPN.ShowOpen
PictureY.Picture = LoadPicture(OPN.FileName)
1:
End Sub

Private Sub Command2_Click()
On Error GoTo 1
OPN.ShowOpen
PictureZ.Picture = LoadPicture(OPN.FileName)
1:
End Sub

Private Sub Command3_Click()
SavePicture PictureX.Picture, App.Path + "\Data\Projects\Wizard New Project\Front.bmp"
SavePicture PictureY.Picture, App.Path + "\Data\Projects\Wizard New Project\Right.bmp"
SavePicture PictureZ.Picture, App.Path + "\Data\Projects\Wizard New Project\Top.bmp"
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error GoTo 1
OPN.ShowOpen
PictureX.Picture = LoadPicture(OPN.FileName)
1:
End Sub

Private Sub Form_Load()
On Error Resume Next
PictureX.Picture = LoadPicture(App.Path + "\Data\Projects\Wizard New Project\Front.bmp")
PictureY.Picture = LoadPicture(App.Path + "\Data\Projects\Wizard New Project\Right.bmp")
PictureZ.Picture = LoadPicture(App.Path + "\Data\Projects\Wizard New Project\Top.bmp")
End Sub

Private Sub Label2_Click()
Unload Me
End Sub
