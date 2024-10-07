VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   LinkTopic       =   "Form8"
   ScaleHeight     =   8295
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Show Report 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Show Report 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Report 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   4920
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Report 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Report 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Report 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Report 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   480
      Picture         =   "data report.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub Command2_Click()
DataReport2.Show
End Sub

Private Sub Command3_Click()
DataReport3.Show
End Sub

Private Sub Command4_Click()
DataReport4.Show
End Sub

Private Sub Command5_Click()
DataReport5.Show
End Sub

Private Sub Command6_Click()
DataReport6.Show
End Sub

Private Sub Command7_Click()
DataReport7.Show
End Sub
