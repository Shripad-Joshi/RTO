VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H80000005&
   Caption         =   "Form13"
   ClientHeight    =   6315
   ClientLeft      =   -75
   ClientTop       =   270
   ClientWidth     =   17520
   LinkTopic       =   "Form13"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   Begin VB.CommandButton Command5 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
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
      Left            =   11400
      TabIndex        =   4
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vehicle Registration Report"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   10800
      TabIndex        =   2
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Permanant License Registration Report"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      TabIndex        =   1
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Learning License Registration Report"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1560
      TabIndex        =   0
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      Height          =   1095
      Left            =   2280
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      Height          =   1095
      Left            =   11280
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      Height          =   1455
      Left            =   6000
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   1440
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   10680
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Report "
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport1.Show
DataReport1.Height = Screen.Height + 100
DataReport1.Width = Screen.Width + 100
End Sub

Private Sub Command2_Click()
DataReport2.Show
DataReport2.Height = Screen.Height + 100
DataReport2.Width = Screen.Width + 100
End Sub

Private Sub Command3_Click()
DataReport3.Show
DataReport3.Height = Screen.Height + 100
DataReport3.Width = Screen.Width + 100
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Load Form1
Form1.Height = Screen.Height
Form1.Width = Screen.Width + 100
Form1.Visible = True
Unload Me
End Sub

