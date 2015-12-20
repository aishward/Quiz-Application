VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form7"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   8400
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Complete Database of Participants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   14
      Top             =   600
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1920
      TabIndex        =   12
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3480
      TabIndex        =   11
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   10
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Welcome Admin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   8640
      TabIndex        =   0
      Top             =   1800
      Width           =   8895
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF80FF&
         Caption         =   "insert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2880
         TabIndex        =   5
         Top             =   4800
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF80FF&
         Caption         =   "Harry Potter"
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
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   3360
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF80FF&
         Caption         =   "English"
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
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   2760
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF80FF&
         Caption         =   "Sports"
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
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   2160
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF80FF&
         Caption         =   "History"
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
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF80FF&
         Caption         =   "Selct the category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   6975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Enter your question here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   5415
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Option Explicit
Public key As String
Dim RS As New ADODB.Recordset
Dim CMD As New ADODB.Command
Dim MSG As Double
Dim cnstr As String
Dim CON As New ADODB.Connection



Private Sub Command1_Click(Index As Integer)
If Option1(0) Then
CMD.CommandText = "select * from history"
CMD.Execute
RS.Open "select * from history", CON, adOpenDynamic, adLockOptimistic
If RS.BOF Then
CMD.CommandText = "INSERT INTO history VALUES(1,'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
Else
n = RS(0)
RS.MoveLast
CMD.CommandText = " INSERT INTO history VALUES(" & n + 1 & ",'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
CMD.CommandText = "commit"
CMD.Execute
RS.Close
End If
End If

If Option1(1) Then
CMD.CommandText = "select * from sports"
CMD.Execute
RS.Open "select * from sports", CON, adOpenDynamic, adLockOptimistic
If RS.BOF Then
CMD.CommandText = "INSERT INTO sports VALUES(1,'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
Else
n = RS(0)
RS.MoveLast
CMD.CommandText = " INSERT INTO sports VALUES(" & n + 1 & ",'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
CMD.CommandText = "commit"
CMD.Execute
RS.Close
End If
End If

If Option1(2) Then
CMD.CommandText = "select * from english"
CMD.Execute

RS.Open "select * from english where q_id=" & n & "", CON, adOpenDynamic, adLockOptimistic
If RS.BOF Then
CMD.CommandText = "INSERT INTO english VALUES(1,'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
RS.Close
Else
n = RS(0)
n = n + 1
RS.MoveLast
CMD.CommandText = " INSERT INTO english VALUES(" & n & ",'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
CMD.CommandText = "commit"
CMD.Execute
RS.Close
End If
End If

If Option1(3) Then
CMD.CommandText = "select * from hp"
CMD.Execute
RS.Open "select * from hp", CON, adOpenDynamic, adLockOptimistic
If RS.BOF Then
CMD.CommandText = "INSERT INTO hp VALUES(1,'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
Else
n = RS(0)
RS.MoveLast
CMD.CommandText = " INSERT INTO hp VALUES(" & n + 1 & ",'" & Text1.Text & "','" & Text2(0).Text & "','" & Text2(1).Text & "','" & Text2(2).Text & "','" & Text2(3).Text & "','" & Text2(4).Text & "')"
CMD.Execute
CMD.CommandText = "commit"
CMD.Execute
RS.Close
End If

End If
MsgBox ("Record Inserted")
Text1.Text = ""
Text2(0).Text = ""
Text2(1).Text = ""
Text2(2).Text = ""
Text2(3).Text = ""
Text2(4).Text = ""
End Sub

Private Sub Command2_Click()
Form8.Show
Form1.Enabled = True
End Sub

Private Sub Command3_Click()
Form1.Show
Form7.Enabled = False
End Sub

Private Sub Form_Load()
cnstr = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
CON.ConnectionString = cnstr
CON.Open
CMD.ActiveConnection = CON
n = 1
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0) Then
n = 1
Else: n = 0
End If
'MsgBox Val(n)
End Sub
