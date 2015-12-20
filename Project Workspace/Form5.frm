VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   600
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   16680
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   3195
      TabIndex        =   10
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1100
      Left            =   1680
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "End Test"
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
      Left            =   10800
      TabIndex        =   4
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit Answer"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "---Options---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   2040
      TabIndex        =   2
      Top             =   4560
      Width           =   13575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   5055
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
         Caption         =   "Option4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7440
         TabIndex        =   7
         Top             =   2640
         Width           =   5055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Option3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1200
         TabIndex        =   6
         Top             =   2520
         Width           =   4935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Option2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   7440
         TabIndex        =   8
         Top             =   720
         Width           =   4695
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Score:"
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
      Left            =   12360
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label q 
      BackColor       =   &H00C0C000&
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   11775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "TIMER"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q_count As Integer
Dim random As Integer
Public score As String
Dim n As Integer
Dim MSG1 As Integer
Option Explicit
Dim RS As New ADODB.Recordset
Dim CMD As New ADODB.Command
Dim MSG As Double
Dim cnstr As String
Dim CON As New ADODB.Connection
Dim min As String
Dim sec As String
Dim colon As String
Dim cat As String
Dim key1 As String
Dim recordname As String
Dim recordpassword As String
Dim conno As Double



Private Sub Command1_Click()
If cat = "history" Then
n = n + 1
If n > 20 Then
Call qui
End If
If (Option1.Value) Then
  If Option1.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option2.Value) Then
  If Option2.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option3.Value) Then
  If Option3.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option4.Value) Then
  If Option4.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
  
  
RS.Close
Option1.Value = False
Option2.Value = False
Option3.Value = False

Option4.Value = False
' Randomize (n)
'n = CInt(Int((25 * Rnd()) + 1))




RS.Open "select * from history where q_id=" & n & "", CON, adOpenDynamic, adLockOptimistic
q.Caption = RS(1)
Option1.Caption = RS(2)
Option2.Caption = RS(3)
Option3.Caption = RS(4)
Option4.Caption = RS(5)

End If


If cat = "sports" Then
n = n + 1
If n > 20 Then
Call qui
End If
If (Option1.Value) Then
  If Option1.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option2.Value) Then
  If Option2.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option3.Value) Then
  If Option3.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option4.Value) Then
  If Option4.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
  
  
RS.Close
Option1.Value = False
Option2.Value = False
Option3.Value = False

Option4.Value = False
' Randomize (n)
'n = CInt(Int((25 * Rnd()) + 1))



RS.Open "select * from sports where q_id=" & n & "", CON, adOpenDynamic, adLockOptimistic
q.Caption = RS(1)
Option1.Caption = RS(2)
Option2.Caption = RS(3)
Option3.Caption = RS(4)
Option4.Caption = RS(5)
'Text1.Text = n
End If


'english
If cat = "english" Then
n = n + 1
If n > 20 Then
Call qui
End If
If (Option1.Value) Then
  If Option1.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option2.Value) Then
  If Option2.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option3.Value) Then
  If Option3.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option4.Value) Then
  If Option4.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
  
  
RS.Close
Option1.Value = False
Option2.Value = False
Option3.Value = False

Option4.Value = False
' Randomize (n)
'n = CInt(Int((25 * Rnd()) + 1))
RS.Open "select * from english where q_id=" & n & "", CON, adOpenDynamic, adLockOptimistic
q.Caption = RS(1)
Option1.Caption = RS(2)
Option2.Caption = RS(3)
Option3.Caption = RS(4)
Option4.Caption = RS(5)
End If

If cat = "hp" Then
n = n + 1
If n > 20 Then
Call qui
End If
If (Option1.Value) Then
  If Option1.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option2.Value) Then
  If Option2.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option3.Value) Then
  If Option3.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
If (Option4.Value) Then
  If Option4.Caption = RS(6) Then
  score = score + 2
  Label2.Caption = "Score: " + score
  Else
  score = score - 1
  Label2.Caption = "Score: " + score
  End If
  End If
  
  
RS.Close
Option1.Value = False
Option2.Value = False
Option3.Value = False

Option4.Value = False
' Randomize (n)
'n = CInt(Int((25 * Rnd()) + 1))

RS.Open "select * from hp where q_id=" & n & "", CON, adOpenDynamic, adLockOptimistic
q.Caption = RS(1)
Option1.Caption = RS(2)
Option2.Caption = RS(3)
Option3.Caption = RS(4)
Option4.Caption = RS(5)
End If

End Sub
Private Sub qui()
MsgBox ("quiz completed. your score is:" + score)
RS.Close
RS.Open "select * from dummy where name='" & key1 & "' ", CON, adOpenDynamic, adLockOptimistic
recordname = RS(0)
recordpassword = RS(1)
conno = RS(2)
Text1.Text = RS(0)
Text2.Text = RS(1)
CMD.CommandText = "delete from dummy where name='" & key1 & "'"
CMD.Execute

CMD.CommandText = "commit"
CMD.Execute

CMD.CommandText = "insert into dummy values ('" & Text1.Text & "','" & Text2.Text & "','" & conno & "' , " & score & " )"
CMD.Execute
CMD.CommandText = "commit"
CMD.Execute

Form5.Hide
Unload Me
Form1.Show
CON.Close

End Sub

Private Sub Command2_Click()
MSG1 = MsgBox("Are you sure you want to end the test??", vbYesNo)
If (MSG1 = vbYes) Then
Dim score1 As Integer
score1 = Val(score)

key1 = Form1.key
MsgBox ("Your score is:" + score)

RS.Close
RS.Open "select * from dummy where name='" & key1 & "' ", CON, adOpenDynamic, adLockOptimistic
recordname = RS(0)
recordpassword = RS(1)
conno = RS(2)
Text1.Text = RS(0)
Text2.Text = RS(1)
CMD.CommandText = "delete from dummy where name='" & key1 & "'"
CMD.Execute

CMD.CommandText = "commit"
CMD.Execute

CMD.CommandText = "insert into dummy values ('" & Text1.Text & "','" & Text2.Text & "','" & conno & "' , " & score & " )"
CMD.Execute
CMD.CommandText = "commit"
CMD.Execute
RS.Close
Form1.Enabled = True
Form1.Show
Form5.Enabled = False
CON.Close
Unload Me
End If

End Sub

Private Sub Form_Load()


key1 = Form1.key
min = 4
sec = 59
Timer1.Enabled = True
q_count = 0
'score = 0
colon = ":"


Label1.Caption = min + colon + sec

cnstr = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
CON.ConnectionString = cnstr
CON.Open
CMD.ActiveConnection = CON
RS.Open "select * from dummy where name='" & key1 & "'", CON, adOpenDynamic, adLockOptimistic
score = RS(3)
Label2.Caption = "Score: " + score
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
If Form2.cat_name = "history" Then
CMD.CommandText = "select * from history"
CMD.Execute

cat = "history"
End If

If Form2.cat_name = "sports" Then
CMD.CommandText = "select * from sports"
CMD.Execute
'RS.Close
cat = "sports"
End If

If Form2.cat_name = "english" Then
'MsgBox "history"
CMD.CommandText = "select * from english"
CMD.Execute
'RS.Close
cat = "english"
End If

If Form2.cat_name = "hp" Then
'MsgBox "history"
CMD.CommandText = "select * from hp"
CMD.Execute
cat = "hp"
End If

n = 1

RS.Close


RS.Open "select * from  " & cat & "  where q_id = " & n & " ", CON, adOpenDynamic, adLockOptimistic
'Text1.Text = cat
'Text2.Text = RS(1)

q.Caption = RS(1)
Option1.Caption = RS(2)
Option2.Caption = RS(3)
Option3.Caption = RS(4)
Option4.Caption = RS(5)

End Sub

Private Sub Timer1_Timer()
sec = sec - 1
If sec < 0 Then
sec = 59
min = min - 1
Label1.Caption = min + colon + sec
End If
If sec >= 0 Then
Label1.Caption = min + colon + sec
End If
End Sub
