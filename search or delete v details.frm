VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   BackColor       =   &H80000005&
   Caption         =   "Form12"
   ClientHeight    =   3015
   ClientLeft      =   -75
   ClientTop       =   90
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
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
      Left            =   5040
      TabIndex        =   39
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14040
      Top             =   6240
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7920
      TabIndex        =   38
      Text            =   "Combo1"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6960
      TabIndex        =   37
      Top             =   4920
      Width           =   1335
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
      Left            =   7800
      TabIndex        =   1
      Top             =   8040
      Width           =   1575
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
      Left            =   2160
      TabIndex        =   0
      Top             =   8040
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   12360
      Top             =   4920
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
   Begin VB.Shape Shape4 
      Height          =   855
      Left            =   6840
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   1095
      Left            =   7680
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   4920
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   2040
      Top             =   7920
      Width           =   1815
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
      Index           =   13
      Left            =   4920
      TabIndex        =   36
      Top             =   4080
      Width           =   2415
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
      Left            =   600
      TabIndex        =   35
      Top             =   6960
      Width           =   1335
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
      Left            =   600
      TabIndex        =   34
      Top             =   5520
      Width           =   735
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
      Left            =   6720
      TabIndex        =   33
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   11
      Left            =   6720
      TabIndex        =   32
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Index           =   1
      Left            =   600
      TabIndex        =   31
      Top             =   6240
      Width           =   1455
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
      Left            =   600
      TabIndex        =   30
      Top             =   1920
      Width           =   1815
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
      Left            =   6720
      TabIndex        =   29
      Top             =   1920
      Width           =   1575
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
      Left            =   6720
      TabIndex        =   28
      Top             =   4800
      Width           =   2175
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
      Left            =   6720
      TabIndex        =   27
      Top             =   4080
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
      Left            =   600
      TabIndex        =   26
      Top             =   1200
      Width           =   2415
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
      Left            =   600
      TabIndex        =   25
      Top             =   2640
      Width           =   1455
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
      Left            =   600
      TabIndex        =   24
      Top             =   480
      Width           =   1335
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
      Left            =   600
      TabIndex        =   23
      Top             =   3360
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
      Left            =   600
      TabIndex        =   22
      Top             =   4800
      Width           =   2175
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
      Left            =   600
      TabIndex        =   21
      Top             =   4080
      Width           =   1335
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
      Left            =   6720
      TabIndex        =   20
      Top             =   3360
      Width           =   1335
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
      Left            =   6720
      TabIndex        =   19
      Top             =   2640
      Width           =   1455
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
      Left            =   3360
      TabIndex        =   18
      Top             =   480
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
      Index           =   1
      Left            =   3360
      TabIndex        =   17
      Top             =   1200
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
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
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
      Left            =   9240
      TabIndex        =   15
      Top             =   1920
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
      Left            =   3360
      TabIndex        =   14
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
      Index           =   5
      Left            =   9240
      TabIndex        =   13
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
      Index           =   6
      Left            =   3360
      TabIndex        =   12
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
      Index           =   7
      Left            =   9240
      TabIndex        =   11
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
      Index           =   8
      Left            =   3360
      TabIndex        =   10
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
      Index           =   9
      Left            =   9240
      TabIndex        =   9
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
      Index           =   10
      Left            =   3360
      TabIndex        =   8
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
      Index           =   11
      Left            =   9240
      TabIndex        =   7
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
      Index           =   12
      Left            =   3360
      TabIndex        =   6
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
      Index           =   13
      Left            =   9240
      TabIndex        =   5
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
      Index           =   14
      Left            =   3360
      TabIndex        =   4
      Top             =   6240
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
      Index           =   15
      Left            =   9240
      TabIndex        =   3
      Top             =   6240
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
      Index           =   16
      Left            =   3360
      TabIndex        =   2
      Top             =   6960
      Width           =   7815
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String
Dim a As String
Dim space As String
Dim cnt As Integer
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

Private Sub Command3_Click()
Shape1.Visible = True
Shape3.Visible = True
Shape4.Visible = False
Label1(0).Visible = True
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
Label6(1).Visible = True
Label6(3).Visible = True
Label2(11).Visible = True
State(18).Visible = True
District(16).Visible = True
Command1.Visible = True
Command2.Visible = True
Label1(13).Visible = False
Command3.Visible = False
Combo1.Visible = False
res.MoveFirst
While (res.EOF <> True)
  If (res.Fields(12) = Combo1.Text) Then
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
Label3(10).Caption = res.Fields(10)
Label3(11).Caption = res.Fields(13)
Label3(12).Caption = res.Fields(14)
Label3(13).Caption = res.Fields(15)
Label3(14).Caption = res.Fields(16)
Label3(15).Caption = res.Fields(17)
Label3(16).Caption = res.Fields(18) + space + res.Fields(19) + space + res.Fields(20)
  End If
 res.MoveNext
 Wend
End Sub
Private Sub Command4_Click()
res.MoveFirst
While (res.EOF <> True)
If (res.Fields(12) = Combo1.Text) Then
 If (MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Delete?") = vbYes) Then
  If (res.EOF <> True) Then
   res.Delete adAffectCurrent
   MsgBox ("Deleted record successfully")
   res.Update
   Exit Sub
  End If
 End If
Else
 res.MoveNext
End If
Wend
End Sub

Private Sub Form_Load()
space = " "
cnt = 0
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "SELECT * FROM vehicle_registration"
res.Open str, con, adOpenDynamic, adLockPessimistic
Label1(0).Visible = False
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
Label6(1).Visible = False
Label6(3).Visible = False
Label2(11).Visible = False
State(18).Visible = False
District(16).Visible = False
Command1.Visible = False
Command2.Visible = False
Command4.Visible = False
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
'combo1 additem
res.MoveFirst
While (res.EOF <> True)
Combo1.AddItem (res.Fields(12))
res.MoveNext
Wend
End Sub

Private Sub Timer1_Timer()
If (cnt = 7) Then
Command4.Visible = True
Shape2.Visible = True
Else
cnt = cnt + 1
End If
End Sub

