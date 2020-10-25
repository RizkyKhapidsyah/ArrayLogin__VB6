VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "Login Form Using An Array"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPassWord 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox TxtUserName 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblWhatThisDoes 
      Alignment       =   2  'Center
      Caption         =   "What This Form Does"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   7
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblInstruct 
      Caption         =   $"FrmLogin.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblUserName 
      Caption         =   "UserName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblPassWord 
      Caption         =   "PassWord:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UserNames(2, 3) As String
Dim PassWords(2, 3) As String

Private Sub CmdLogin_Click()
On Error GoTo Exit_Sub:
'Now where gonna scan the UserName and PassWord Arrays to see
'if the users selected a correct UserName and PassWord
'And if they did they get a Login Message, otherwise they get the error message
'For your purposes, you can change the message boxes to show form blah
If TxtUserName.Text = UserNames(1, 1) And TxtPassWord.Text = PassWords(1, 1) Then
            MsgBox ("Your Logged In")
            GoTo Exit_Sub
ElseIf TxtUserName.Text = UserNames(2, 1) And TxtPassWord.Text = PassWords(2, 1) Then
            MsgBox ("Your Logged In")
            GoTo Exit_Sub
ElseIf TxtUserName.Text = UserNames(1, 2) And TxtPassWord.Text = PassWords(1, 2) Then
            MsgBox ("Your Logged In")
            GoTo Exit_Sub
ElseIf TxtUserName.Text = UserNames(2, 2) And TxtPassWord.Text = PassWords(2, 2) Then
            MsgBox ("Your Logged In")
            GoTo Exit_Sub
ElseIf TxtUserName.Text = UserNames(1, 3) And TxtPassWord.Text = PassWords(1, 3) Then
            MsgBox ("Your Logged In")
            GoTo Exit_Sub
ElseIf TxtUserName.Text = UserNames(2, 3) And TxtPassWord.Text = PassWords(2, 3) Then
            MsgBox ("Your Logged In")
            GoTo Exit_Sub
End If
'If your password and Username combination was wrong then you get this message
    MsgBox ("Invalid Login Name Or PassWord")

Exit_Sub:
    Exit Sub
End Sub

Private Sub CmdReset_Click()
'Clear the Textboxes
    TxtUserName = ""
    TxtPassWord.Text = ""
End Sub

Private Sub Form_Load()
'This Makes the first Table out of Arrays
'This is the UserName Table and corresponds with
'Whats in the UserName TextBox
UserNames(1, 1) = "Name1"
UserNames(2, 1) = "Name2"
UserNames(1, 2) = "Name3"
UserNames(2, 2) = "Name4"
UserNames(1, 3) = "Name5"
UserNames(2, 3) = "Name6"

'This makes the second table, the name of this table
'Is called the PassWord table, and it corresponds with the
'PassWord Textbox
PassWords(1, 1) = "PassWord1"
PassWords(2, 1) = "PassWord2"
PassWords(1, 2) = "PassWord3"
PassWords(2, 2) = "PassWord4"
PassWords(1, 3) = "PassWord5"
PassWords(2, 3) = "PassWord6"
End Sub
