VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H80000005&
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   LinkTopic       =   "Form5"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   9960
      TabIndex        =   51
      Top             =   5520
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13320
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   495
      Left            =   1560
      TabIndex        =   25
      Top             =   9720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   9000
      Width           =   4935
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2280
      TabIndex        =   22
      Text            =   "Combo4"
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   8520
      Width           =   4935
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   8040
      Width           =   4935
   End
   Begin VB.TextBox Text18 
      Height          =   405
      Left            =   8400
      TabIndex        =   18
      Top             =   7560
      Width           =   2070
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   9960
      TabIndex        =   15
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   3000
      TabIndex        =   14
      Top             =   4680
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
      Left            =   3960
      TabIndex        =   13
      Top             =   4200
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
      Left            =   3000
      TabIndex        =   12
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   11160
      TabIndex        =   10
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   12360
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11160
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "Combo3"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   4080
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Permanant License Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   52
      Top             =   240
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1440
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000005&
      Caption         =   "LL_RID"
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
      Left            =   8880
      TabIndex        =   50
      Top             =   5520
      Width           =   855
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
      Left            =   720
      TabIndex        =   49
      Top             =   9000
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
      Left            =   720
      TabIndex        =   48
      Top             =   8520
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
      Left            =   720
      TabIndex        =   47
      Top             =   8040
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
      Left            =   1320
      TabIndex        =   46
      Top             =   7080
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
      Left            =   7440
      TabIndex        =   45
      Top             =   7080
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
      Left            =   7440
      TabIndex        =   44
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   " (To be printed on smartcard)"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   43
      Top             =   6600
      Width           =   2295
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
      Left            =   840
      TabIndex        =   42
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Present Address"
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
      Left            =   840
      TabIndex        =   41
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Select state and RTO office  from where PL  is being applied"
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
      Left            =   0
      TabIndex        =   40
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "Pin Code"
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
      Left            =   11280
      TabIndex        =   39
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "To know your RTO office Enter pincode of applicant's permanant  address line"
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
      Left            =   7200
      TabIndex        =   38
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   " State"
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
      Index           =   0
      Left            =   600
      TabIndex        =   37
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000005&
      Caption         =   "Blood Group"
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
      Left            =   8400
      TabIndex        =   36
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000005&
      Caption         =   "Email Id"
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
      Left            =   8880
      TabIndex        =   35
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "Date of Birth"
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
      Left            =   8400
      TabIndex        =   34
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000005&
      Caption         =   "Mobile number"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   33
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000005&
      Caption         =   "Identification Mark "
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Caption         =   "Place of Birth"
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
      Left            =   1560
      TabIndex        =   31
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Name of Applicant"
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
      Index           =   7
      Left            =   960
      TabIndex        =   30
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Aadhaar Number"
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
      Index           =   3
      Left            =   1080
      TabIndex        =   29
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   " Relation"
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
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   " (To be printed on smartcard)"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   27
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Gender"
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
      Index           =   4
      Left            =   1920
      TabIndex        =   26
      Top             =   4200
      Width           =   735
   End
End
Attribute VB_Name = "Form5"
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
'state name for combo1'
Combo1.AddItem ("Assam")
Combo1.AddItem ("Andhra Pradesh")
Combo1.AddItem ("Odisha")
Combo1.AddItem ("Punjab")
Combo1.AddItem ("Delhi")
Combo1.AddItem ("Gujarat")
Combo1.AddItem ("Karnataka")
Combo1.AddItem ("Haryana")
Combo1.AddItem ("Rajasthan")
Combo1.AddItem ("Himachal Pradesh")
Combo1.AddItem ("Uttarakhand")
Combo1.AddItem ("Jharkhand")
Combo1.AddItem ("Chhattisgarh")
Combo1.AddItem ("Kerala")
Combo1.AddItem ("Tamil Nadu")
Combo1.AddItem ("Madhya Pradesh")
Combo1.AddItem ("West Bengal")
Combo1.AddItem ("Bihar")
Combo1.AddItem ("Maharashtra")
Combo1.AddItem ("Uttar Pradesh")
Combo1.AddItem ("Chandigard")
Combo1.AddItem ("Telengana")
Combo1.AddItem ("Jammu and Kashamir")
Combo1.AddItem ("Tripura")
Combo1.AddItem ("Meghalaya")
Combo1.AddItem ("Goa")
Combo1.AddItem ("Arunachal Pradesh")
Combo1.AddItem ("Manipur")
Combo1.AddItem ("Mizoram")
Combo1.AddItem ("Sikkim")
Combo1.AddItem ("Puducherry")
Combo1.AddItem ("Nagaland")
Combo1.AddItem ("Andaman and Nicobar")
Combo1.AddItem ("Dadra and Nagar Haveli")
Combo1.AddItem ("Daman and Diu")
Combo1.AddItem ("Lakshadweep")

'state name for combo4'
Combo4.AddItem ("Assam")
Combo4.AddItem ("Andhra Pradesh")
Combo4.AddItem ("Odisha")
Combo4.AddItem ("Punjab")
Combo4.AddItem ("Delhi")
Combo4.AddItem ("Gujarat")
Combo4.AddItem ("Karnataka")
Combo4.AddItem ("Haryana")
Combo4.AddItem ("Rajasthan")
Combo4.AddItem ("Himachal Pradesh")
Combo4.AddItem ("Uttarakhand")
Combo4.AddItem ("Jharkhand")
Combo4.AddItem ("Chhattisgarh")
Combo4.AddItem ("Kerala")
Combo4.AddItem ("Tamil Nadu")
Combo4.AddItem ("Madhya Pradesh")
Combo4.AddItem ("West Bengal")
Combo4.AddItem ("Bihar")
Combo4.AddItem ("Maharashtra")
Combo4.AddItem ("Uttar Pradesh")
Combo4.AddItem ("Chandigard")
Combo4.AddItem ("Telengana")
Combo4.AddItem ("Jammu and Kashamir")
Combo4.AddItem ("Tripura")
Combo4.AddItem ("Meghalaya")
Combo4.AddItem ("Goa")
Combo4.AddItem ("Arunachal Pradesh")
Combo4.AddItem ("Manipur")
Combo4.AddItem ("Mizoram")
Combo4.AddItem ("Sikkim")
Combo4.AddItem ("Puducherry")
Combo4.AddItem ("Nagaland")
Combo4.AddItem ("Andaman and Nicobar")
Combo4.AddItem ("Dadra and Nagar Haveli")
Combo4.AddItem ("Daman and Diu")
Combo4.AddItem ("Lakshadweep")

'relation for combo2'
Combo3.AddItem ("Father")
Combo3.AddItem ("Mother")
Combo3.AddItem ("Husband")
Combo3.AddItem ("Guardian")

'adodc connectivity code'
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\BCA\My Project\RTO Database.mdb;Persist Security Info=False"
str = "select * from PL_registration"
res.Open str, con, adOpenDynamic, adLockOptimistic
End Sub


Private Sub Save_Click()
If (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text16.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Or Text21.Text = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1.Text = "select State" Or Combo4.Text = "select State") Then
   MsgBox ("Fill all information")
   Exit Sub
End If
res.AddNew
res.Fields(0) = Combo1.Text
res.Fields(1) = Text4.Text
res.Fields(2) = Text1.Text
res.Fields(3) = Text2.Text
res.Fields(4) = Text3.Text
res.Fields(5) = Combo3.Text
res.Fields(6) = Text5.Text
res.Fields(7) = Text6.Text
res.Fields(8) = Text7.Text
res.Fields(9) = Text8.Text
If (Option1.Value = True) Then
  res.Fields(10) = Option1.Caption
ElseIf (Option2.Value = True) Then
  res.Fields(10) = Option1.Caption
End If
res.Fields(11) = Text10.Text
res.Fields(12) = Text11.Text
res.Fields(13) = Text12.Text
res.Fields(14) = Text13.Text
res.Fields(15) = Text9.Text
res.Fields(16) = Text14.Text
res.Fields(17) = Combo4.Text
res.Fields(18) = Text15.Text
res.Fields(19) = Text17.Text
res.Fields(20) = Text18.Text
res.Fields(21) = Text19.Text
res.Fields(22) = Text20.Text
res.Fields(24) = Text21.Text
res.Fields(23) = Text16.Text
res.Update
MsgBox ("Successfull")
con.Close
Load Form6
Form6.Height = Screen.Height + 100
Form6.Width = Screen.Height + 100
Form6.Visible = True
Unload Me
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text10_Click()
MsgBox ("Date should be in DD/MM/YYYY Format")
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
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
Private Sub Text15_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text17_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate(KeyAscii)
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub
Private Sub Text18_KeyPress(KeyAscii As Integer)
Dim b As Boolean
b = validate_i(KeyAscii)
End Sub


