VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   39
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OKAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   8
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      TabIndex        =   7
      Top             =   7680
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   10680
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
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
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8880
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   38
      Top             =   240
      Width           =   11175
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   37
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   36
      Top             =   960
      Width           =   11175
   End
   Begin VB.Label Label1 
      Caption         =   "Aadhaar number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   480
      TabIndex        =   35
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   34
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Applicant Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   33
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Email ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   32
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   31
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Blood Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   30
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   10800
      TabIndex        =   29
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   28
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   27
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   26
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Mobile number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   8160
      TabIndex        =   25
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   10800
      TabIndex        =   24
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   10800
      TabIndex        =   23
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   22
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   8160
      TabIndex        =   21
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   8160
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Village/Town"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   19
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Identification Mark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   18
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   17
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   16
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   15
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   9
      Left            =   10800
      TabIndex        =   14
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   10800
      TabIndex        =   13
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   12
      Top             =   5280
      Width           =   11175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   11
      Top             =   6000
      Width           =   11175
   End
   Begin VB.Label Label1 
      Caption         =   "LL_RID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   10
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   9
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Applicant Name"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "LL_RID"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7440
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String
Dim a As String
Dim str1, space As String
Private Sub Command1_Click()
While (res.EOF <> True)
  If (res.Fields(23) = Text2.Text) Then
  str1 = "Applicant's"
  space = "  "
Label2(0).Visible = False
Label1(0).Visible = False
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Command1.Visible = False
  Label2(1).Visible = True
  Label2(2).Visible = True
  Label2(3).Visible = True
  Label2(4).Visible = True
  Label2(5).Visible = True
  Label2(6).Visible = True
  Label2(7).Visible = True
  Label2(8).Visible = True
  Label2(9).Visible = True
  Label2(10).Visible = True
  Label2(11).Visible = True
  Label2(12).Visible = True
  Label2(13).Visible = True
  Label4.Visible = True
  Label5.Visible = True
Label1(1).Visible = True
Label1(2).Visible = True
Label1(3).Visible = True
Label1(4).Visible = True
Label1(5).Visible = True
Label1(6).Visible = True
Label1(7).Visible = True
Label1(8).Visible = True
Label1(9).Visible = True
Label1(10).Visible = True
Label1(11).Visible = True
Label1(12).Visible = True
Label1(13).Visible = True
Label1(14).Visible = True
Command2.Visible = True
Command3.Visible = True
  'label caption setting'
Label5.Caption = str1 + space + res.Fields(5)

'appicant's detail'
Label2(14).Caption = res.Fields(2) + space + res.Fields(3) + space + res.Fields(4)
Label4.Caption = res.Fields(6) + space + res.Fields(7) + space + res.Fields(8)
Label2(1).Caption = res.Fields(10)
Label2(2).Caption = res.Fields(13)
Label2(3).Caption = res.Fields(11)
Label2(4).Caption = res.Fields(15)
Label2(5).Caption = res.Fields(10)
Label2(6).Caption = res.Fields(14)
Label2(7).Caption = res.Fields(0)
Label2(8).Caption = res.Fields(19)
Label2(9).Caption = res.Fields(18)
Label2(10).Caption = res.Fields(20)
Label2(11).Caption = res.Fields(16)
Label2(12).Caption = res.Fields(21) + space + res.Fields(22) + space + res.Fields(24)
Label2(13).Caption = res.Fields(23)
  End If
 res.MoveNext
 Wend
 Form8.Command4.Visible = True
End Sub

Private Sub Command4_Click()
While (res.EOF <> True)
If (res.Fields(23) = Text2.Text) Then
If (MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Delete?") = vbYes) Then
If (res.EOF <> True) Then
res.Delete
MsgBox ("Deleted Successfully")
End If
End If
  End If
 res.MoveNext
 Wend
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "SELECT * FROM LL_registration"
res.Open str, con, adOpenDynamic, adLockPessimistic


Label2(1).Visible = False
Label2(2).Visible = False
Label2(3).Visible = False
Label2(4).Visible = False
Label2(5).Visible = False
Label2(6).Visible = False
Label2(7).Visible = False
Label2(8).Visible = False
Label2(9).Visible = False
Label2(10).Visible = False
Label2(11).Visible = False
Label2(12).Visible = False
Label2(13).Visible = False
Label4.Visible = False
Label5.Visible = False
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False
Label1(4).Visible = False
Label1(5).Visible = False
Label1(6).Visible = False
Label1(7).Visible = False
Label1(8).Visible = False
Label1(9).Visible = False
Label1(10).Visible = False
Label1(11).Visible = False
Label1(12).Visible = False
Label1(13).Visible = False
Label1(14).Visible = False
Command2.Visible = False
Command3.Visible = False
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Load Form1
Form1.Height = Screen.Height
Form1.Width = Screen.Width + 100
Form1.Visible = True
Unload Me
End Sub
