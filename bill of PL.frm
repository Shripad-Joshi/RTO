VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H80000005&
   Caption         =   "Form6"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17685
   LinkTopic       =   "Form6"
   ScaleHeight     =   11025
   ScaleWidth      =   17685
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   8040
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command1 
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
      TabIndex        =   1
      Top             =   8760
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
      Left            =   8880
      TabIndex        =   0
      Top             =   8760
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   18600
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Shape Shape3 
      Height          =   1095
      Left            =   8760
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   3120
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   2880
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Bill Receipt"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   2880
      TabIndex        =   32
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   2760
      TabIndex        =   31
      Top             =   1320
      Width           =   11175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   29
      Top             =   2040
      Width           =   11175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   27
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   7800
      TabIndex        =   23
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   10440
      TabIndex        =   22
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   21
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   20
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   7800
      TabIndex        =   19
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   7800
      TabIndex        =   18
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   10440
      TabIndex        =   17
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   10440
      TabIndex        =   16
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   7800
      TabIndex        =   14
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   7800
      TabIndex        =   13
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   9
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   8
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   9
      Left            =   10440
      TabIndex        =   7
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   10440
      TabIndex        =   6
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   5
      Top             =   6360
      Width           =   11175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   4
      Top             =   7080
      Width           =   11175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "PL_RID"
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
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Left            =   2760
      TabIndex        =   2
      Top             =   7800
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
con.Close
Load Form1
Form1.Height = Screen.Height
Form1.Width = Screen.Width + 100
Form1.Visible = True
Unload Me
End Sub

Private Sub Form_Load()
Dim str1, space As String
str1 = "Applicant's"
space = "  "
'adodc connectivity code'
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "select * from PL_registration"
res.Open str, con, adOpenDynamic, adLockOptimistic

'label caption setting'
res.MoveFirst
Label3(0).Caption = str1 + space + res.Fields(5)

'appicant's detail'
Label2(0).Caption = res.Fields(2) + space + res.Fields(3) + space + res.Fields(4)
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
End Sub


