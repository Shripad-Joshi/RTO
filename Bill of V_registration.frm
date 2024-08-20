VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   BackColor       =   &H80000005&
   Caption         =   "Form11Form11"
   ClientHeight    =   10770
   ClientLeft      =   3075
   ClientTop       =   660
   ClientWidth     =   16500
   LinkTopic       =   "Form11"
   ScaleHeight     =   10770
   ScaleWidth      =   16500
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   12240
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   2040
      TabIndex        =   31
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
      Left            =   7680
      TabIndex        =   30
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   1200
      Top             =   120
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
      Left            =   1200
      TabIndex        =   32
      Top             =   120
      Width           =   9255
   End
   Begin VB.Shape Shape3 
      Height          =   1095
      Left            =   7560
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   1920
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Index           =   16
      Left            =   3240
      TabIndex        =   29
      Top             =   6960
      Width           =   7815
   End
   Begin VB.Label Label3 
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
      Left            =   9120
      TabIndex        =   28
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   27
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   9120
      TabIndex        =   26
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   25
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Index           =   9
      Left            =   9120
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   23
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   9120
      TabIndex        =   22
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   21
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   9120
      TabIndex        =   20
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   19
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   9120
      TabIndex        =   18
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   17
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   16
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   7815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Vehicler color"
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
      Left            =   6600
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Engine No"
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
      Left            =   6600
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Chassis No"
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
      TabIndex        =   12
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Vehicle  Manfacturer"
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
      TabIndex        =   11
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Vehicle type"
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
      Left            =   480
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner Name "
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
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner Email ID"
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
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner's Adhaar Card No."
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
      TabIndex        =   7
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Dealer name (To be printed on smartcard)"
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
      Left            =   6600
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Dealer Registration ID"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner Gender"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner's Phone no"
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
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label District 
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
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   16
      Left            =   6600
      TabIndex        =   2
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label State 
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
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   18
      Left            =   480
      TabIndex        =   1
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Address "
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
      TabIndex        =   0
      Top             =   6960
      Width           =   1335
   End
End
Attribute VB_Name = "Form11"
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
Dim space As String
space = "  "
'adodc connectivity code'
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "select * from Vehicle_registration"
res.Open str, con, adOpenDynamic, adLockOptimistic

'owner's detail'
res.MoveFirst
Label3(0).Caption = res.Fields(0) + space + res.Fields(1) + space + res.Fields(2)
Label3(1).Caption = res.Fields(12)
Label3(2).Caption = res.Fields(4)
Label3(3).Caption = res.Fields(5)
Label3(4).Caption = res.Fields(3)
Label3(5).Caption = res.Fields(7)
Label3(6).Caption = res.Fields(6)
Label3(7).Caption = res.Fields(9)
Label3(8).Caption = res.Fields(8)
Label3(9).Caption = res.Fields(11)
Label3(10).Caption = res.Fields(9)
Label3(11).Caption = res.Fields(13)
Label3(12).Caption = res.Fields(14)
Label3(13).Caption = res.Fields(15)
Label3(16).Caption = res.Fields(18) + space + res.Fields(19) + space + res.Fields(20)
End Sub

