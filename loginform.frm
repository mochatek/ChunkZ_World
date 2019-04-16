VERSION 5.00
Begin VB.Form loginform 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHUNKZ WORLD"
   ClientHeight    =   6270
   ClientLeft      =   6840
   ClientTop       =   2355
   ClientWidth     =   7350
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "loginform.frx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   7350
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   3135
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
      Begin VB.CommandButton create 
         BackColor       =   &H00FFFFC0&
         Caption         =   "HERE"
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
         Left            =   2280
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton login 
         BackColor       =   &H0000FFFF&
         Caption         =   "LOG IN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox pswd 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox user 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Don't have one? Create"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CHUNKZ WORLD"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Integer
Public username As String
Private Sub Command1_Click()
MsgBox hdl
End Sub

Private Sub login_Click()
readdata "select Name,Password from mocha where Name='" & user.Text & "' and Password='" & pswd.Text & "'"
If rs.EOF = True Then
 MsgBox "USERNAME-PASSWORD MISMATCH", 16, "CHUNKZ WORLD"
user.Text = ""
pswd.Text = ""
Else
writedata "update mocha set online=1 where Name='" & user.Text & "' and Password='" & pswd.Text & "'"
username = user.Text
 MsgBox "LOG-IN SUCCESSFULL", 64, "CHUNKZ WORLD"
user.Text = ""
pswd.Text = ""
loginform.Hide
mainform.Show
End If
End Sub

Private Sub create_Click()
user.Text = ""
pswd.Text = ""
loginform.Hide
signupform.Show
End Sub

