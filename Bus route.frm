VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   LinkTopic       =   "Form9"
   ScaleHeight     =   9795
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Back To Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   5760
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   7920
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "busroute"
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
   Begin VB.TextBox Text1 
      DataField       =   "State"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "SLNO"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Text            =   "---SELECT---"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   5655
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "VIEW YOUR ROUTE MAP AND REACH YOUR DESTINY"
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
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   10335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your State Code"
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
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   -240
      Picture         =   "Bus route.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18015
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cns As ADODB.Connection
Dim rst As ADODB.Recordset

Private Sub Combo1_Click()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
rst.Open "select State from busroute where SLNO =" & Combo1.List(Combo1.ListIndex), cns
Text1.Text = rst.Fields("State")
If Combo1.Text = "1221" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\1.jpg")
ElseIf Combo1.Text = "1889" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\2.jpg")
ElseIf Combo1.Text = "1223" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\3.jpg")
ElseIf Combo1.Text = "2312" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\4.jpg")
ElseIf Combo1.Text = "1241" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\5.gif")
ElseIf Combo1.Text = "1432" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\6.jpg")
ElseIf Combo1.Text = "1234" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\7.jpg")
ElseIf Combo1.Text = "3564" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\8.jpg")
ElseIf Combo1.Text = "2342" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\9.jpg")
ElseIf Combo1.Text = "2435" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\10.jpg")
ElseIf Combo1.Text = "1243" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\11.jpg")
ElseIf Combo1.Text = "1255" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\12.jpg")
ElseIf Combo1.Text = "1235" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\13.jpg")
ElseIf Combo1.Text = "6333" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\14.jpg")
ElseIf Combo1.Text = "4322" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\15.jpg")
ElseIf Combo1.Text = "1532" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\16.jpg")
ElseIf Combo1.Text = "6754" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\17.jpg")
ElseIf Combo1.Text = "8954" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\18.jpg")
ElseIf Combo1.Text = "9834" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\19.jpg")
ElseIf Combo1.Text = "9082" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\20.jpg")
ElseIf Combo1.Text = "8922" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\21.jpg")
ElseIf Combo1.Text = "9023" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\22.jpg")
ElseIf Combo1.Text = "0923" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\23.jpg")
ElseIf Combo1.Text = "8972" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\24.jpg")
ElseIf Combo1.Text = "2923" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\25.jpg")
ElseIf Combo1.Text = "2992" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\26.jpg")
ElseIf Combo1.Text = "9290" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\27.jpg")
ElseIf Combo1.Text = "8875" Then
Image2.Picture = LoadPicture("C:\Users\gane0\OneDrive\Desktop\maps\28.jpg")
End If
End Sub


Private Sub Command1_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Form_Load()
Set cns = New ADODB.Connection
Set rst = New ADODB.Recordset
cns.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=database"
rst.Open "select * from busroute", cns
Do While Not rst.EOF
Combo1.AddItem rst.Fields("SLNO")
rst.MoveNext
Loop
End Sub

