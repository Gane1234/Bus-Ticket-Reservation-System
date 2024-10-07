VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form24 
   Caption         =   "Form24"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16065
   LinkTopic       =   "Form24"
   ScaleHeight     =   10350
   ScaleWidth      =   16065
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   85
      Left            =   840
      Top             =   7680
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   8640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Image Image4 
      Height          =   1455
      Index           =   1
      Left            =   4560
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   5295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING FILES...."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   8160
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LIKE BREATHING"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10200
      TabIndex        =   4
      Top             =   8400
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "it's something you do."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "is not something you're good at."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   7440
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TRAVELLING"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10560
      TabIndex        =   1
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE TICKET BOOKING SYSTEM"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   9615
      Left            =   240
      Picture         =   "Splash.frx":177C9
      Stretch         =   -1  'True
      Top             =   240
      Width           =   14295
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label1.Caption = "Loading Successful...."
Label2.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
Form1.Show
End If
End Sub

