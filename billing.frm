VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   LinkTopic       =   "Form4"
   ScaleHeight     =   10785
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   8760
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      BOFAction       =   0
      EOFAction       =   0
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
      RecordSource    =   "amount1"
      Caption         =   "Adodc3"
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
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   7440
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   5400
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2040
      Top             =   9000
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "billing"
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
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   12
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OrderId"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   7440
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "Foodaccomidation"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6360
      TabIndex        =   9
      Top             =   4440
      Width           =   3255
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Noofseats"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6360
      TabIndex        =   8
      Top             =   3600
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Busname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6360
      TabIndex        =   7
      Top             =   2880
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Routenumber"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6360
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "amount"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   6180
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Of Food"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare Stages"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Your Tickets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   2040
      TabIndex        =   10
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable Amount"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   6240
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Food Accomidation"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Seats"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bus Name"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RouteNumber"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   10215
      Left            =   480
      Picture         =   "billing.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cns As ADODB.Connection
Dim rst As ADODB.Recordset
Dim cns1 As ADODB.Connection
Dim rst1 As ADODB.Recordset
Dim num As Integer
Dim s1 As String
Dim confirm As Integer

Private Sub Combo1_Click()
Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
rst1.Open "select Farestages from enquiry where Routenumber =" & Combo1.List(Combo1.ListIndex), cns1
Text2.Text = rst1.Fields("Farestages")
End Sub

Private Sub Combo2_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
End Sub

Private Sub Combo3_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
End Sub

Private Sub Combo4_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
If Combo4.Text = "Biryani" Then
Text3.Text = 100
ElseIf Combo4.Text = "Rice" Then
Text3.Text = 100
ElseIf Combo4.Text = "LemonRice" Then
Text3.Text = 50
ElseIf Combo4.Text = "VegPalav" Then
Text3.Text = 70
ElseIf Combo4.Text = "Chapathi6Pcs" Then
Text3.Text = 80
ElseIf Combo4.Text = "Meals" Then
Text3.Text = 150
ElseIf Combo4.Text = "MiniMeals" Then
Text3.Text = 80
ElseIf Combo4.Text = "Kushka" Then
Text3.Text = 70
End If
End Sub

Private Sub Command1_Click()
s1 = Text1.Text = ""
s2 = Text2.Text = ""
s3 = Text3.Text = ""
s4 = Combo1.Text = ""
s5 = Combo2.Text = ""
s6 = Combo3.Text = ""
s7 = Combo4.Text = ""
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
MsgBox "please enter all the details"
Text1.SetFocus
Text2.SetFocus
Text3.SetFocus
Combo3.SetFocus
Combo1.SetFocus
Combo2.SetFocus
Combo4.SetFocus
Exit Sub
End If
If Combo1.Text = "18897" Then
MsgBox "Order Id:- AAY1123221"
ElseIf Combo1.Text = "16754" Then
MsgBox "Order Id:- AAY1123222"
ElseIf Combo1.Text = "13324" Then
MsgBox "Order Id:- AAY1123223"
ElseIf Combo1.Text = "44567" Then
MsgBox "Order Id:- AAY1123224"
ElseIf Combo1.Text = "84384" Then
MsgBox "Order Id:- AAY1123225"
ElseIf Combo1.Text = "66654" Then
MsgBox "Order Id:- AAY1123226"
ElseIf Combo1.Text = "76556" Then
MsgBox "Order Id:- AAY1123227"
ElseIf Combo1.Text = "18897" Then
MsgBox "Order Id:- AAY1123228"
ElseIf Combo1.Text = "66758" Then
MsgBox "Order Id:- AAY1123229"
ElseIf Combo1.Text = "77778" Then
MsgBox "Order Id:- AAY1123230"
ElseIf Combo1.Text = "55636" Then
MsgBox "Order Id:- AAY1123231"
ElseIf Combo1.Text = "7768" Then
MsgBox "Order Id:- AAY1123232"
ElseIf Combo1.Text = "88756" Then
MsgBox "Order Id:- AAY1123233"
ElseIf Combo1.Text = "77785" Then
MsgBox "Order Id:- AAY1123234"
ElseIf Combo1.Text = "11234" Then
MsgBox "Order Id:- AAY1123235"
ElseIf Combo1.Text = "11564" Then
MsgBox "Order Id:- AAY1123236"
ElseIf Combo1.Text = "76676" Then
MsgBox "Order Id:- AAY1123237"
End If
s1 = Text1.Text = ""
If Text1.Text = "" Then
Exit Sub
End If
Adodc3.Recordset.AddNew
Adodc3.Recordset("amount") = Text1.Text
Adodc3.Recordset.Update
Text1.Text = ""
Form11.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form10.Show
Unload Me
End Sub

Private Sub Command3_Click()
Sum = (Val(Text2.Text) * Val(Combo3.Text)) + (Val(Combo3.Text) * Val(Text3.Text))
Text1.Text = Sum
End Sub

Private Sub Form_Load()
Set cns = New ADODB.Connection
Set rst = New ADODB.Recordset
cns.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
rst.Open "select * from billing", cns
Do While Not rst.EOF
Combo2.AddItem rst.Fields("Busname")
Combo3.AddItem rst.Fields("Noofseats")
Combo4.AddItem rst.Fields("Foodaccomidation")
rst.MoveNext
Loop
Set cns1 = New ADODB.Connection
Set rst1 = New ADODB.Recordset
cns1.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
rst1.Open "select * from enquiry", cns1
Do While Not rst1.EOF
Combo1.AddItem rst1.Fields("Routenumber")
rst1.MoveNext
Loop
Adodc3.Recordset.AddNew
End Sub

