VERSION 5.00
Begin VB.Form frmLoginTTC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login to TTCoffice.com"
   ClientHeight    =   3720
   ClientLeft      =   2832
   ClientTop       =   3480
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2197.899
   ScaleMode       =   0  'User
   ScaleWidth      =   6028.032
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   2325
   End
   Begin VB.CommandButton cmdLOGIN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&LOG IN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Picture         =   "frmLoginTTC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      Picture         =   "frmLoginTTC.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   900
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1800
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please use your TTCoffice.com User name and Password."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   912
      Left            =   120
      Picture         =   "frmLoginTTC.frx":08CC
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblUserID 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1830
      Width           =   1080
   End
End
Attribute VB_Name = "frmLoginTTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bolTTCLoginClickedLogin As Boolean


Private Sub cmdCancel_Click()

    'set the global var to false
    'to denote a failed login
    bolTTCLoginClickedLogin = False
    Me.Hide

End Sub

Private Sub cmdLogin_Click()

    bolTTCLoginClickedLogin = True
    Me.Hide

End Sub

Private Sub Form_Activate()

    txtUserID.SetFocus

End Sub

Private Sub Form_Load()

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    'Initialize login
    bolTTCLoginClickedLogin = False
    txtPassword = ""
    
End Sub

Public Sub subTTCLoginTryAgain()


End Sub

Private Sub txtPassword_GotFocus()

    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtUserID.Text)

End Sub

Private Sub txtUserID_GotFocus()

    txtUserID.SelStart = 0
    txtUserID.SelLength = Len(txtUserID.Text)

End Sub
