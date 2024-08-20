VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   Picture         =   "RTO Management system(Shripad Aad Yash).frx":0000
   ScaleHeight     =   3.16007e5
   ScaleMode       =   0  'User
   ScaleWidth      =   5.43492e8
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Height          =   2685
      Left            =   0
      Picture         =   "RTO Management system(Shripad Aad Yash).frx":4321
      ScaleHeight     =   2625
      ScaleWidth      =   2370
      TabIndex        =   0
      Top             =   0
      Width           =   2430
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   11
      Left            =   3120
      TabIndex        =   31
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H8000000E&
      Caption         =   "Report Generation"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3360
      MousePointer    =   11  'Hourglass
      TabIndex        =   30
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   10
      Left            =   13200
      TabIndex        =   29
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000E&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13440
      MousePointer    =   11  'Hourglass
      TabIndex        =   28
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   9
      Left            =   13080
      TabIndex        =   27
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Delete Vehicle Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13320
      MousePointer    =   11  'Hourglass
      TabIndex        =   26
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   8
      Left            =   8160
      TabIndex        =   25
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000E&
      Caption         =   "Know Vehicle Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8400
      MousePointer    =   11  'Hourglass
      TabIndex        =   24
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   7
      Left            =   3240
      TabIndex        =   23
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000E&
      Caption         =   "Vehicle Registration"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      MousePointer    =   11  'Hourglass
      TabIndex        =   22
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   6
      Left            =   13080
      TabIndex        =   21
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   5
      Left            =   8160
      TabIndex        =   20
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   4
      Left            =   13080
      TabIndex        =   19
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000E&
      Caption         =   "Delete Permanant Licence Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13320
      MousePointer    =   11  'Hourglass
      TabIndex        =   18
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000E&
      Caption         =   "Know Permanant Licence Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8400
      MousePointer    =   11  'Hourglass
      TabIndex        =   17
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Caption         =   "Delete Learing Licence Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13440
      MousePointer    =   11  'Hourglass
      TabIndex        =   16
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   3
      Left            =   8160
      TabIndex        =   15
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   2
      Left            =   8280
      TabIndex        =   14
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Know Learning Licence Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8400
      MousePointer    =   11  'Hourglass
      TabIndex        =   13
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Check Post Tax"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8520
      MousePointer    =   11  'Hourglass
      TabIndex        =   12
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Permanant Licence  Registration"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      MousePointer    =   11  'Hourglass
      TabIndex        =   10
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   2520
      Top             =   0
      Width           =   20175
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "Learning Licence Registration"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   3480
      MousePointer    =   11  'Hourglass
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Permanant Licence Registration"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Check Post Tax "
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Learning Licence Registration"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vehicle Registration"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
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
      TabIndex        =   2
      Top             =   840
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
      TabIndex        =   1
      Top             =   0
      Width           =   20175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Form1.Height = Screen.Height
Form1.Width = Screen.Width + 100
End Sub

Private Sub Label10_Click(Index As Integer)
Load Form2
Form2.Height = Screen.Height
Form2.Width = Screen.Width + 100
Form2.Visible = True
Unload Me
End Sub

Private Sub Label12_Click()
Load Form7
Form7.Height = Screen.Height
Form7.Width = Screen.Width + 100
Form7.Visible = True
Unload Me
End Sub

Private Sub Label13_Click()
Load Form8
Form8.Height = Screen.Height
Form8.Width = Screen.Width + 100
Form8.Visible = True
Unload Me
End Sub

Private Sub Label14_Click()
Load Form8
Form8.Height = Screen.Height
Form8.Width = Screen.Width + 100
Form8.Visible = True
Form8.Timer1.Enabled = True
Unload Me
End Sub

Private Sub Label15_Click()
Load Form9
Form9.Height = Screen.Height
Form9.Width = Screen.Width + 100
Form9.Visible = True
Unload Me
End Sub

Private Sub Label16_Click()
Load Form9
Form9.Height = Screen.Height
Form9.Width = Screen.Width + 100
Form9.Visible = True
Form9.Timer1.Enabled = True
Unload Me
End Sub

Private Sub Label17_Click()
Load Form10
Form10.Height = Screen.Height
Form10.Width = Screen.Width + 100
Form10.Visible = True
Unload Me
End Sub

Private Sub Label18_Click()
Load Form10
Form12.Height = Screen.Height + 100
Form12.Width = Screen.Width + 100
Form12.Visible = True
Unload Me
End Sub

Private Sub Label19_Click()
Unload Me
End Sub

Private Sub Label20_Click()
Load Form13
Form13.Height = Screen.Height
Form13.Width = Screen.Width + 100
Form13.Visible = True
Unload Me
End Sub

Private Sub Label7_Click()
Load Form5
Form5.Height = Screen.Height
Form5.Width = Screen.Width + 100
Form5.Visible = True
Unload Me
End Sub

Private Sub Label8_Click()
Load Form9
Form12.Height = Screen.Height
Form12.Width = Screen.Width + 100
Form12.Visible = True
Form12.Timer1.Enabled = True
Unload Me
End Sub
