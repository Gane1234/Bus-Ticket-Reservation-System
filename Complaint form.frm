VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13155
   LinkTopic       =   "Form3"
   ScaleHeight     =   9360
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataField       =   "Registercomplaint"
      DataSource      =   "Adodc1"
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
      Left            =   5640
      TabIndex        =   9
      Top             =   2640
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   960
      Top             =   8160
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   1
      EOFAction       =   1
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "enquiry"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   8
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   7
      Top             =   6360
      Width           =   2895
   End
   Begin VB.TextBox Text1 
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
      Left            =   5640
      TabIndex        =   6
      Top             =   4920
      Width           =   3975
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Routenumber"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Complaint form.frx":0000
      Left            =   5640
      List            =   "Complaint form.frx":0002
      TabIndex        =   5
      Text            =   "---SELECT---"
      Top             =   3840
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "SLno"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Text            =   "---SELECT---"
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Any Other Reasons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Define route"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Complaint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Register Complaint Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   120
      Picture         =   "Complaint form.frx":0004
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cns As ADODB.Connection
Dim rst As ADODB.Recordset

Private Sub Combo1_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
rst.Open "select Registercomplaint from enquiry where SLno =" & Combo1.List(Combo1.ListIndex), cns
Text2.Text = rst.Fields("registercomplaint")
End Sub

Private Sub Combo3_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
'rst.Open "select Routenumber from enquiry" & Combo3.List(Combo3.ListIndex), cns
End Sub

Private Sub Command1_Click()
MsgBox "Complaint Registered Successfully. It will be resolved within 7 working days"
MDIForm1.Show
Unload Me
End Sub

Private Sub Command2_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Form_Load()
Set cns = New ADODB.Connection
Set rst = New ADODB.Recordset
cns.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
rst.Open "select * from enquiry", cns
Do While Not rst.EOF
Combo1.AddItem rst.Fields("SLno")
Combo3.AddItem rst.Fields("Routenumber")
rst.MoveNext
Loop
End Sub
