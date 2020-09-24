VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email : ZaidMarkabi@yahoo.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Http://www.yazanmarkabi.webs.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   2640
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For more VB applications, visit my website :"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   3030
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Written by  Zaid Markabi"
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded in VB 6.0"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Projections "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   2670
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convert 3 Images to 3D Object"
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   2190
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   360
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   4695
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
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " 3D Projections - About"
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
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
Unload Me
End Sub
