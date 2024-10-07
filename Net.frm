VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13995
   LinkTopic       =   "Form14"
   ScaleHeight     =   10620
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   2040
      TabIndex        =   20
      Top             =   6480
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Verify OTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   18
      Top             =   7800
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   450
      Left            =   4320
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   794
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
      RecordSource    =   "payment"
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
      Height          =   375
      Left            =   1560
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      RecordSource    =   "amount1"
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
      Caption         =   "Fetch Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   17
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text5 
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
      Left            =   2040
      TabIndex        =   16
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox Text1 
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
      Left            =   7080
      TabIndex        =   14
      Top             =   2520
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "CYBH"
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
      Left            =   7080
      TabIndex        =   13
      Top             =   1920
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "CYBK"
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
      Left            =   7080
      TabIndex        =   12
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   8880
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pay Now"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   8880
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send OTP"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
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
      Left            =   8760
      TabIndex        =   4
      Top             =   8880
      Width           =   3015
   End
   Begin VB.TextBox Text4 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
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
      Left            =   2040
      TabIndex        =   19
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "IFSC CODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CHOOSE YOUR BRANCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CHOOSE YOUR BANK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Banking"
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
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "OTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT TO PAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   10215
      Left            =   1200
      Picture         =   "Net.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   11295
   End
End
Attribute VB_Name = "Form14"
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
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
End Sub

Private Sub Combo2_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
If Combo2.Text = "Marthalli" Then
Text1.Text = "LKJKA786213"
ElseIf Combo2.Text = "Krpuram" Then
Text1.Text = "HASL21238"
ElseIf Combo2.Text = "TCPALYA" Then
Text1.Text = "HJASH89712"
ElseIf Combo2.Text = "Seggahalli" Then
Text1.Text = "BABRB8123"
ElseIf Combo2.Text = "Kadugudi" Then
Text1.Text = "CORYU3212"
ElseIf Combo2.Text = "Chintamani" Then
Text1.Text = "SBIIN3211"
ElseIf Combo2.Text = "Bhatalapalli" Then
Text1.Text = "KARN789123"
ElseIf Combo2.Text = "Whitefeild" Then
Text1.Text = "BARBO897132"
ElseIf Combo2.Text = "Varthur" Then
Text1.Text = "SBBIN878321"
ElseIf Combo2.Text = "Gunjur" Then
Text1.Text = "KOTA89Y312"
ElseIf Combo2.Text = "Kolar" Then
Text1.Text = "987073KJASD"
ElseIf Combo2.Text = "Bangerpet" Then
Text1.Text = "OJPA07321"
ElseIf Combo2.Text = "Mulbagal" Then
Text1.Text = "GANE2384E"
ElseIf Combo2.Text = "Mysore" Then
Text1.Text = "KKIIE87211"
ElseIf Combo2.Text = "Mangalore" Then
Text1.Text = "HHYYAI8132"
ElseIf Combo2.Text = "Shivvamoga" Then
Text1.Text = "GGWYY3211"
ElseIf Combo2.Text = "Dharwad" Then
Text1.Text = "KRPU88721"
End If
End Sub

Private Sub Command1_Click()
Combo1.Clear
Combo2.Clear
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command2_Click()
Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
If Text5 = "" Then
MsgBox "please enter all the details"
Text5.SetFocus
Exit Sub
End If
rst1.Open "Select amount from Amount1 where amount", cns1
Text2.Text = rst1.Fields("amount")
End Sub

Private Sub Command3_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command4_Click()
s1 = Text1.Text = ""
s2 = Text2.Text = ""
s3 = Text3.Text = ""
s4 = Text5.Text = ""
s5 = Combo1.Text = ""
s6 = Combo2.Text = ""
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Text3.Text = "" Then
MsgBox "please enter all the details"
Text1.SetFocus
Text2.SetFocus
Text5.SetFocus
Combo1.SetFocus
Combo2.SetFocus
Text3.SetFocus
Exit Sub
End If
If Len(Text3.Text) = 10 Then
MsgBox "OTP sent Sucessfully"
Else
MsgBox "Enter Valid Mobile Number"
End If
End Sub

Private Sub Command5_Click()
Dim rst1 As ADODB.Recordset
Set rst1 = New ADODB.Recordset
s1 = Text1.Text = ""
s2 = Text2.Text = ""
s3 = Text4.Text = ""
s4 = Text5.Text = ""
s5 = Combo1.Text = ""
s6 = Combo2.Text = ""
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "please enter all the details"
Text1.SetFocus
Text2.SetFocus
Text4.SetFocus
Text5.SetFocus
Combo1.SetFocus
Combo2.SetFocus
Exit Sub
End If
MsgBox "Thanks for Using our Website your ticket has been confirmed"
Form5.Show
Unload Me
rst1.Open "delete from amount1 where amount", cns
End Sub

Private Sub Command6_Click()
s1 = Text4.Text = ""
If Text4.Text = "" Then
MsgBox "Please enter otp"
Text4.SetFocus
Exit Sub
End If
MsgBox "OTP Verified"
End Sub

Private Sub Form_Load()
Set cns = New ADODB.Connection
Set rst = New ADODB.Recordset
cns.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
rst.Open "select * from payment", cns
Do While Not rst.EOF
Combo1.AddItem rst.Fields("CYBK")
Combo2.AddItem rst.Fields("CYBH")
rst.MoveNext
Loop
Set cns1 = New ADODB.Connection
Set rst1 = New ADODB.Recordset
cns1.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
rst1.Open "select * from amount1", cns1
Do While Not rst1.EOF
rst1.MoveNext
Loop
End Sub

