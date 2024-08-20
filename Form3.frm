VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18150
   LinkTopic       =   "Form3"
   ScaleHeight     =   9465
   ScaleWidth      =   18150
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12480
      TabIndex        =   7
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12480
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   15120
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Height          =   2685
      Left            =   0
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   2625
      ScaleWidth      =   2370
      TabIndex        =   0
      Top             =   0
      Width           =   2430
   End
   Begin VB.Label Label5 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Login ID"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   2520
      Top             =   0
      Width           =   20175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   20175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "RTO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   20175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Management"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   20175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String
Private Sub Command1_Click()
If (Text1.Text = res.Fields(0) And Text2.Text = res.Fields(1)) Then
 MsgBox ("Login Successfull")
'form1 load'
Load Form1
Form1.Height = Screen.Height
Form1.Width = Screen.Width + 100
Form1.Visible = True
Unload Me
ElseIf (res.Fields(0) <> Text1.Text) Then
  MsgBox ("Login ID is Incorrect")
ElseIf (res.Fields(0) <> Text1.Text) Then
  MsgBox ("Password is Incorrect")
End If
End Sub

Private Sub Form_Load()
'form display  setting'
Form3.Height = Screen.Height
Form3.Width = Screen.Width + 100
'adodc connectivity code'
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "select * from Login"
res.Open str, con, adOpenDynamic, adLockOptimistic
End Sub
