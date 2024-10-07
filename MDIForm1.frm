VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9945
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16215
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10695
      Left            =   0
      ScaleHeight     =   10635
      ScaleWidth      =   16155
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "BUS ROUTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPLAINTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   15
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ENQUIRY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "ANY QUERIES CONTACT TO OUR WEBSTE OR NUMBER 080-2561 8776, ganeshediga51@gmail.com"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   9000
         Width           =   16215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "LOGOUT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   12120
         TabIndex        =   12
         Top             =   7080
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "BOOK TICKET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   8160
         TabIndex        =   11
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   4440
         TabIndex        =   10
         Top             =   7080
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   840
         TabIndex        =   9
         Top             =   7080
         Width           =   2895
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Ph.no : - 8105983033              8431676544"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   13680
         TabIndex        =   4
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Image Image6 
         Height          =   1695
         Left            =   14160
         Picture         =   "MDIForm1.frx":0000
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"MDIForm1.frx":1A75
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   15375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Under the Maintenance of KSRTC,NWKSRTC,APSRTC,TSRTC ALL NON-GOVERNMENT VEHICALS CAN BE BOOKED TO COMPLETE YOUR JOURNEY"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   3600
         Width           =   13695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "    Traveling IT GIVES YOU in thousand strange places THEN LEAVES YOU a stranger in your OWN LAND"
         BeginProperty Font 
            Name            =   "@PMingLiU-ExtB"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   14055
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "E-MAIL ID : - mrpenterrpises@gmail.com          "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   4920
         Width           =   7095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "                                                  MRP ENTERPRISES"
         BeginProperty Font 
            Name            =   "@Microsoft YaHei UI"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   13335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "       #79/1 1STCROSS,GEETHA CLINIC ROAD, DEVASANDRA, KRPURAM, BANGALORE-560036"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   14535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "FAX ID : - deepakdeepu0118@gmail.com"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   10020
         Index           =   0
         Left            =   0
         Picture         =   "MDIForm1.frx":1B5C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   16200
      End
   End
   Begin VB.Menu Mnmenu 
      Caption         =   "Menu"
      Begin VB.Menu mlogin 
         Caption         =   "Login"
         Index           =   0
      End
      Begin VB.Menu mcomplaint 
         Caption         =   "Complaint"
         Index           =   0
      End
      Begin VB.Menu mlogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mroutemap 
         Caption         =   "routemap"
      End
      Begin VB.Menu mregister 
         Caption         =   "Register"
      End
      Begin VB.Menu MEnd 
         Caption         =   "End"
      End
   End
   Begin VB.Menu mnbooking 
      Caption         =   "Booking"
      Begin VB.Menu mticket 
         Caption         =   "Ticket"
      End
      Begin VB.Menu mcancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mconfirm 
         Caption         =   "Confirm"
      End
   End
   Begin VB.Menu mncomplaint 
      Caption         =   "Complaint"
      Begin VB.Menu mwithdraw 
         Caption         =   "Withdraw"
      End
   End
   Begin VB.Menu mnlogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Label10_Click()
Form5.Show
Unload Me
End Sub

Private Sub Label12_Click()
Form7.Show
Unload Me
End Sub

Private Sub Label13_Click()
Form3.Show
Unload Me
End Sub

Private Sub Label15_Click()
Form9.Show
Unload Me
End Sub

Private Sub Label5_Click()
Form2.Show
Unload Me
End Sub

Private Sub Label7_Click()
Form10.Show
Unload Me
End Sub

Private Sub mcomplaint_Click(Index As Integer)
Form3.Show
Unload Me
End Sub

Private Sub MEnd_Click()
Form5.Show
Unload Me
End Sub

Private Sub mlogin_Click(Index As Integer)
Form1.Show
Unload Me
End Sub

Private Sub mlogout_Click()
Form1.Show
Unload Me
End Sub

Private Sub mnbooking_Click()
Form10.Show
Unload Me
End Sub

Private Sub mncomplaint_Click()
Form3.Show
Unload Me
End Sub

Private Sub mnlogout_Click()
Form1.Show
Unload Me
End Sub

Private Sub mregister_Click()
Form2.Show
Unload Me
End Sub

Private Sub mroutemap_Click()
Form9.Show
Unload Me
End Sub

Private Sub mticket_Click()
Form10.Show
Unload Me
End Sub
