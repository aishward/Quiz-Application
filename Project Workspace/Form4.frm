VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   12
      Top             =   7920
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   15480
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   11
      Top             =   4560
      Width           =   4695
   End
   Begin VB.PictureBox Picture7 
      Height          =   3855
      Left            =   720
      Picture         =   "Form4.frx":2325
      ScaleHeight     =   3795
      ScaleWidth      =   3075
      TabIndex        =   10
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      DataField       =   "CONNUM"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   9
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "PASSWORD"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   6480
      TabIndex        =   0
      Top             =   2040
      Width           =   8535
      Begin VB.CommandButton Command4 
         Caption         =   "REGISTER NOW"
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
         Left            =   2760
         TabIndex        =   2
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         DataField       =   "NAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
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
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
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
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "CONFIRM PASSWORD"
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
         Left            =   840
         TabIndex        =   4
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0FF&
         Caption         =   "CONTACT NUMBER"
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
         Left            =   960
         TabIndex        =   3
         Top             =   3480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As New ADODB.Recordset
Dim CMD As New ADODB.Command
Dim MSG As Double
Dim cnstr As String
Dim CON As New ADODB.Connection

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Command1_Click()
Form1.Show
Form4.Hide
End Sub


Private Sub Command4_Click()
If (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "") Then
MsgBox "Enter all fields", vbCritical, "ERROR"
Form4.Show
GoTo d
Else
CMD.CommandText = "select * from dummy"
CMD.CommandText = "commit"
CMD.Execute


RS.Open "select * from dummy", CON, adOpenDynamic, adLockOptimistic
top:
If Text3.Text = RS(0) Then
MsgBox "userexist"
GoTo quit
End If

If Text3.Text <> RS(0) Then
RS.MoveNext
If RS.EOF = True Then
GoTo down
Else
If Text3.Text = RS(0) Then
MsgBox ("user exist")
GoTo quit
End If
End If
GoTo top
End If


down:
If (Text1.Text = Text2.Text) Then
    CMD.CommandText = "insert into dummy values ('" & Text3.Text & "','" & Text2.Text & "'," & Text4.Text & ",0)"
    CMD.Execute
    CMD.CommandText = "commit"
    CMD.Execute
    MsgBox "Registration Successful"
    Form1.Show
    Form4.Hide
    Unload Me
Else
MsgBox ("password does not match")

End If


 
End If
quit:
RS.Close



d: End Sub


Private Sub Form_Load()
Form4.Width = 100000
Form4.Height = 10000
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
cnstr = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
CON.ConnectionString = cnstr
CON.Open
CMD.ActiveConnection = CON
'RS.Close
End Sub

