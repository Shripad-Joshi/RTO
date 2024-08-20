VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Height          =   2685
      Left            =   0
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   2625
      ScaleWidth      =   2370
      TabIndex        =   4
      Top             =   0
      Width           =   2430
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Tax"
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   3
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   14160
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   2520
      Top             =   0
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
      Index           =   1
      Left            =   2520
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   840
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
      TabIndex        =   5
      Top             =   1800
      Width           =   20175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select the state "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   2
      Top             =   3120
      Width           =   4935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If (Combo1.Text = Combo2.Text) Then
  MsgBox ("Same State!!")
  End
End If
Dim str1, str2 As Integer
str1 = Len(Combo1.Text)
str2 = Len(Combo2.Text)
Select Case (str1 Mod 2)
  Case 0
   Select Case (str2 Mod 2)
      Case 0
      MsgBox ("RS.200")
      Case 1
      MsgBox ("RS.250")
    End Select
  Case 1
   Select Case (str2 Mod 2)
      Case 0
      MsgBox ("RS.300")
      Case 1
      MsgBox ("RS.350")
  End Select
 End Select
End Sub

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
End Sub
