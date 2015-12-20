VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   2655
      Left            =   16080
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   3555
      TabIndex        =   10
      Top             =   6240
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   720
      Picture         =   "Form2.frx":47FD
      ScaleHeight     =   3435
      ScaleWidth      =   3315
      TabIndex        =   9
      Top             =   5040
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   16920
      Picture         =   "Form2.frx":6638
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   1440
      Picture         =   "Form2.frx":A2D0
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   7
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Harry Potter Trivia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   10440
      TabIndex        =   5
      Top             =   5880
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   5880
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   10320
      TabIndex        =   3
      Top             =   4560
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   4560
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "QUIT "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      TabIndex        =   1
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Select your category. The quiz will start once the category is selected."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5880
      TabIndex        =   6
      Top             =   2040
      Width           =   9855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cat_name As String
Dim MSG1 As Integer
Dim str As String
Option Explicit
Dim RS As New ADODB.Recordset
Dim CMD As New ADODB.Command
Dim MSG As Double
Dim cnstr As String
Dim CON As New ADODB.Connection


Private Sub Command1_Click()
'MSG = MsgBox("ARE YOU SURE TO UPDATE..?", vbYesNoCancel)
MSG1 = MsgBox("Are you sure you want to quit?", vbYesNoCancel)
If (MSG1 = vbYes) Then
Form1.Enabled = True
Form1.Show
Unload Me
End If
End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then
cat_name = "history"
End If
If Index = 1 Then
cat_name = "sports"
End If
If Index = 2 Then
cat_name = "english"
End If
If Index = 3 Then
cat_name = "hp"
End If

Form5.Show
Form2.Enabled = False
Form2.Hide
'Unload Me '''''''''''''''
End Sub

Private Sub Form_Load()
cnstr = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
CON.ConnectionString = cnstr
'CON.Close
CON.Open
CMD.ActiveConnection = CON


End Sub
