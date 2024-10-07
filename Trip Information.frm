VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   7815
      Begin VB.CommandButton cmdnew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         TabIndex        =   8
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         TabIndex        =   7
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   6
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Txttno 
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txtrno 
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Txtbno 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Txtstime 
         Height          =   405
         Left            =   3120
         TabIndex        =   2
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Txtetime 
         Height          =   405
         Left            =   3120
         TabIndex        =   1
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Number of Trip "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Route  Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bus Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "TRIP INFORMATION"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   14
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
MsgBox ("Do You want to Exit")
Me.Hide

End Sub

Private Sub Cmdnew_Click()
MsgBox ("Do You want to Clear")
Txttno.Text = " "
Txtrno.Text = " "
Txtbno.Text = " "
Txtstime.Text = " "
Txtetime.Text = " "
End Sub

Private Sub cmdsave_Click()

If Txttno.Text = "" Then
MsgBox "Please Enter the Trip number.", vbInformation
Txttno.SetFocus
Exit Sub
End If

If Txtrno.Text = "" Then
MsgBox "Please Enter the Route Number .", vbInformation
Txtrno.SetFocus
Exit Sub
End If

If Txtbno.Text = "" Then
MsgBox "Please Enter the bus Number .", vbInformation
Txtbno.SetFocus
Exit Sub
End If

If Txtstime.Text = "" Then
MsgBox "Please Enter the start time .", vbInformation
Txtstime.SetFocus
Exit Sub
End If


If Txtetime.Text = "" Then
MsgBox "Please Enter the End Time .", vbInformation
Txtetime.SetFocus
Exit Sub
End If

con.Execute ("insert into trip values(" + Txttno.Text + "," + Txtstime.Text + ", " + Txtetime.Text + ", " + Txtrno.Text + "," + Txtbno.Text + ")")

MsgBox ("successfully saved")
End Sub

Private Sub Form_Load()
connectdb
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

