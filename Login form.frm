VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   10680
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3615
      Left            =   0
      Picture         =   "Login form.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sign  In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim str As String

Private Sub Command1_Click()
If (res.Fields(0) = Text1.Text And res.Fields(1) = Text2.Text) Then
 MsgBox ("Login Successfull")
Else
  MsgBox ("Login Unsuccessfull")
End If
End Sub

Private Sub Form_Load()
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\shripad\SEM 5\PROJECT\My Project\RTO Database.mdb;Persist Security Info=False")
str = "select * from login"
res.Open str, con, adOpenDynamic, adLockOptimistic
End Sub
