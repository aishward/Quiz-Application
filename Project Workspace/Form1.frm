VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture5 
      Height          =   3615
      Left            =   16320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   3435
      TabIndex        =   13
      Top             =   240
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   360
      Picture         =   "Form1.frx":22AA
      ScaleHeight     =   3435
      ScaleWidth      =   3435
      TabIndex        =   12
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Quit Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16560
      TabIndex        =   11
      Top             =   8760
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New User"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   11400
      TabIndex        =   10
      Top             =   5040
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RULES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   3
      Top             =   8040
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "VIEW SCORES"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   2
      Top             =   6480
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "EXISTING USER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4200
      TabIndex        =   1
      Top             =   2400
      Width           =   5295
      Begin VB.TextBox Text2 
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
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text1 
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
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00808080&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   6
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "CREATED BY : Aishwarya Desai    Adesh Gupta    Nikhil Mehta    Apurva Ravetkar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   10200
      Width           =   10455
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   16320
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line Line1 
      X1              =   10440
      X2              =   10440
      Y1              =   2160
      Y2              =   9720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "**QUIZ**"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public key As String

Dim RS As New ADODB.Recordset
Dim CMD As New ADODB.Command
Dim MSG As Double
Dim cnstr As String
Dim CON As New ADODB.Connection
'Dim connect As ADODB.Connection

Private Sub Command1_Click()
Form6.Show
Form1.Enabled = True
End Sub

Private Sub Command2_Click()
Form3.Enabled = True                'rules form
Form3.Show

End Sub

Private Sub Command3_Click()
If (Text1.Text = "admin" And Text2.Text = "pvgcoet") Then
Form7.Show
Else

'login form for welcome and inserting values
If (Text1.Text = "" Or Text2.Text = "") Then
MsgBox "Enter all fields", vbCritical, "ERROR"
Form1.Show
Else
'RS.Close
RS.Open "select * from DUMMY where NAME='" & Text1.Text & "' ", CON, adOpenDynamic, adLockOptimistic
If (RS.EOF) Then
MsgBox "User does not exist.Please Register!!"
Text1.Text = ""
Text2.Text = ""
Else
If (RS(1) = Text2.Text) Then
'RS.Close
'Text3.Text = RS(0)
    MsgBox ("Login successful")
    key = Text1.Text
    'CMD.CommandText = "Update DUMMY set flag = 1 where name='" & Text1.Text & "' "
    'CMD.Execute
    Text1.Text = ""
    
    Text2.Text = ""
    Form2.Enabled = True
    Form2.Show
    Form1.Hide
Else
    MsgBox ("Password Incorrect!")
Text1.Text = ""
   Text2.Text = ""
End If
End If
RS.Close
End If
End If

'd:
'Form1.Show

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Form4.Enabled = True
Form4.Show
End Sub

Private Sub Form_Load()
cnstr = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
CON.ConnectionString = cnstr
CON.Open
CMD.ActiveConnection = CON
'RS.Open "SELECT * FROM DUMMY", CON, adOpenDynamic, adLockOptimistic
'Call disp
Text1.Text = ""
   Text2.Text = ""
End Sub
