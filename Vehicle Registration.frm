VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   12165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22515
   LinkTopic       =   "Form10"
   ScaleHeight     =   12165
   ScaleWidth      =   22515
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   14760
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
   Begin VB.CommandButton Save 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   41
      Top             =   10560
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   405
      Left            =   3960
      TabIndex        =   40
      Top             =   9360
      Width           =   3945
   End
   Begin VB.TextBox Text17 
      Height          =   405
      Left            =   3960
      TabIndex        =   39
      Top             =   8640
      Width           =   3945
   End
   Begin VB.TextBox Text16 
      Height          =   405
      Left            =   3960
      TabIndex        =   38
      Top             =   7920
      Width           =   3945
   End
   Begin VB.TextBox Text15 
      Height          =   405
      Left            =   10440
      TabIndex        =   37
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Height          =   405
      Left            =   3960
      TabIndex        =   36
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Height          =   405
      Left            =   10440
      TabIndex        =   35
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   10440
      TabIndex        =   34
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   3960
      TabIndex        =   33
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   405
      Left            =   10440
      TabIndex        =   32
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Height          =   405
      Left            =   3960
      TabIndex        =   31
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   405
      Left            =   10440
      TabIndex        =   30
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3960
      TabIndex        =   29
      Text            =   "Select State"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   28
      Text            =   "Select Vehicle Type"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000005&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   20
      Top             =   2880
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000005&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   19
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   3960
      TabIndex        =   16
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   10440
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   3960
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   3960
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11880
      TabIndex        =   9
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1440
      Width           =   3645
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   3960
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   5160
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "  Vehicle Registration"
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
      Left            =   5160
      TabIndex        =   42
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Address Line 3"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   27
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Address Line 2"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   26
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Address Line 1"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   25
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label State 
      BackColor       =   &H80000005&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   18
      Left            =   1440
      TabIndex        =   24
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label District 
      BackColor       =   &H80000005&
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   16
      Left            =   7920
      TabIndex        =   23
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Pincode"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   11
      Left            =   7920
      TabIndex        =   22
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Village/Town"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   21
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner's Phone no"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1440
      TabIndex        =   18
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner Gender"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7920
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Dealer Registration ID"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   7920
      TabIndex        =   15
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Dealer name (To be printed on smartcard)"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner's Adhaar Card No."
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   13
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner Email ID"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Owner Name "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Vehicle type"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Vehicle  Manfacturer"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Chassis No"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Engine No"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Vehicler color"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String

'Function for validation

'Function for character
Function validate(KeyAscii As Integer) As Boolean
 If (KeyAscii >= 0 And KeyAscii < 65) Or (KeyAscii > 90 And KeyAscii < 96) Or (KeyAscii > 122) Then
 If (KeyAscii = 8 Or KeyAscii = 32) Then
    Exit Function
 End If
 KeyAscii = 0
 MsgBox ("Enter Correct Information.")
 End If
End Function

'Function for integer
Function validate_i(KeyAscii As Integer) As Boolean
 If (KeyAscii < 48 Or KeyAscii > 57) Then
 If (KeyAscii = 8) Then
    Exit Function
 End If
 KeyAscii = 0
 MsgBox ("Enter Correct Information.")
 End If
End Function

Private Sub Form_Load()
'state name for combo2'
Combo2.AddItem ("Assam")
Combo2.AddItem ("Andhra Pradesh")
Combo2.AddItem ("Odisha")
Combo2.AddItem ("Punjab")
Combo2.AddItem ("Delhi")
Combo2.AddItem ("Gujarat")
Combo2.AddItem ("Karnataka")
Combo2.AddItem ("Haryana")
Combo2.AddItem ("Rajasthan")
Combo2.AddItem ("Himachal Pradesh")
Combo2.AddItem ("Uttarakhand")
Combo2.AddItem ("Jharkhand")
Combo2.AddItem ("Chhattisgarh")
Combo2.AddItem ("Kerala")
Combo2.AddItem ("Tamil Nadu")
Combo2.AddItem ("Madhya Pradesh")
Combo2.AddItem ("West Bengal")
Combo2.AddItem ("Bihar")
Combo2.AddItem ("Maharashtra")
Combo2.AddItem ("Uttar Pradesh")
Combo2.AddItem ("Chandigard")
Combo2.AddItem ("Telengana")
Combo2.AddItem ("Jammu and Kashamir")
Combo2.AddItem ("Tripura")
Combo2.AddItem ("Meghalaya")
Combo2.AddItem ("Goa")
Combo2.AddItem ("Arunachal Pradesh")
Combo2.AddItem ("Manipur")
Combo2.AddItem ("Mizoram")
Combo2.AddItem ("Sikkim")
Combo2.AddItem ("Puducherry")
Combo2.AddItem ("Nagaland")
Combo2.AddItem ("Andaman and Nicobar")
Combo2.AddItem ("Dadra and Nagar Haveli")
Combo2.AddItem ("Daman and Diu")
Combo2.AddItem ("Lakshadweep")

'Vehicle type for combo1
Combo1.AddItem ("MCWOG")
Combo1.AddItem ("MCWG")
Combo1.AddItem ("LMV")
Combo1.AddItem ("3W-NT")
Combo1.AddItem ("LMV-TR")
Combo1.AddItem ("BUS")
Combo1.AddItem ("RDRLR")
Combo1.AddItem ("OTHER")

End Sub


Private Sub Save_Click()
If (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text16.Text = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1.Text = "Select Vehicle Type" Or Combo2.Text = "Select State") Then
   MsgBox ("Fill all information")
   Exit Sub
End If
'adodc connectivity code'
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "select * from Vehicle_registration"
res.Open str, con, adOpenDynamic, adLockOptimistic
'Save to database'
res.AddNew
res.Fields(0) = Text1.Text
res.Fields(1) = Text2.Text
res.Fields(2) = Text3.Text
res.Fields(3) = Text7.Text
res.Fields(4) = Text5.Text
If (Option1.Value = True) Then
  res.Fields(5) = Option1.Caption
ElseIf (Option2.Value = True) Then
  res.Fields(5) = Option2.Caption
End If
res.Fields(6) = Combo1.Text
res.Fields(7) = Text6.Text
res.Fields(8) = Text9.Text
res.Fields(9) = Text8.Text
res.Fields(10) = Text11.Text
res.Fields(11) = Text10.Text
res.Fields(12) = Text4.Text
res.Fields(13) = Text12.Text
res.Fields(14) = Combo2.Text
res.Fields(15) = Text13.Text
res.Fields(16) = Text14.Text
res.Fields(17) = Text15.Text
res.Fields(18) = Text16.Text
res.Fields(19) = Text17.Text
res.Fields(20) = Text18.Text
res.Update
MsgBox ("Successfull")
Load Form11
Form11.Height = Screen.Height + 100
Form11.Width = Screen.Height + 500
Form11.Visible = True
Unload Me
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub

