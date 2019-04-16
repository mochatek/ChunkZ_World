VERSION 5.00
Begin VB.Form mainform 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHUNKZ WORLD"
   ClientHeight    =   6315
   ClientLeft      =   6840
   ClientTop       =   2355
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "mainform.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   7410
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   600
      Top             =   240
   End
   Begin VB.CommandButton logout 
      BackColor       =   &H00C0C0FF&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox onlinelist 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2175
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton online 
      BackColor       =   &H00FFFFC0&
      Caption         =   " ONLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox chatbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2640
      Width           =   4815
   End
   Begin VB.TextBox sendtext 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6240
      Picture         =   "mainform.frx":1ACC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Image refresh 
      Height          =   615
      Left            =   6360
      Picture         =   "mainform.frx":4804
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   615
   End
   Begin VB.Image send 
      Height          =   615
      Left            =   5280
      Picture         =   "mainform.frx":64E9
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CHUNKZ WORLD"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hdl As Integer
Dim file As String
Dim tm As String
Dim msg As String
Dim frnd As String
Dim ln As Integer
Private Sub Form_Load()
file = "" & App.Path & "\chat.txt"
'file = "E:\VB PROJECTS\chat.txt"
read
ln = lines()
Timer1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
'closing
writedata "update mocha set online=0 where Name ='" & loginform.username & "'"
con.Close
Timer1.Enabled = False
End Sub
Private Sub logout_Click()
'logout
writedata "update mocha set online=0 where Name ='" & loginform.username & "'"
con.Close
a = MsgBox("LOG-OUT SUCCESSFULL", 64, "CHUNKZ WORLD")
sendtext.Text = ""
Timer1.Enabled = False
mainform.Hide
loginform.Show
End Sub

Private Sub online_Click()
'friends online
onlinelist.Text = ""
readdata "select Name from mocha where online=1"
While Not rs.EOF
Dim name As String
name = rs.Fields("Name")
onlinelist.Text = onlinelist.Text + name + vbNewLine
rs.MoveNext
Wend
End Sub

Private Sub refresh_Click()
read
End Sub

Private Sub send_Click()
'send message
tm = "[" + Format$(Time, "hh:mm:ss AM/PM") + "]"
msg = loginform.username + " : " + sendtext.Text + " " + tm
hdl = FreeFile()
Open file For Append As #hdl
Print #hdl, msg
Close #hdl
sendtext.Text = ""
End Sub
Public Function lines() As Integer
'count number of lines in text file
Dim n As Integer
n = 0
hdl = FreeFile()
Open file For Input As #hdl
While Not EOF(hdl)
a = Input$(1, #hdl)
n = n + 1
Wend
Close #hdl
lines = n
End Function
Public Sub read()
'display chat
chatbox.Text = ""
hdl = FreeFile()
Open file For Input As #hdl
While Not EOF(hdl)
chatbox.Text = chatbox.Text + Input$(1, #hdl)
Wend
Close #hdl
chatbox.refresh
End Sub

Private Sub Timer1_Timer()
Dim nl As Integer
nl = lines()
If (ln <> nl) Then
read
ln = nl
End If
End Sub
