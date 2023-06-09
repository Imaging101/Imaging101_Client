VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessageForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Message"
   ClientHeight    =   5640
   ClientLeft      =   4185
   ClientTop       =   630
   ClientWidth     =   8595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8595
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   4965
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1191
      ButtonWidth     =   609
      ButtonHeight    =   1032
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton OKButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopyToClipboard 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy To Clipboard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3915
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   900
      Width           =   8415
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6015
      TabIndex        =   5
      Top             =   615
      Width           =   2445
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   6135
      Picture         =   "frmMessageForm.frx":0000
      Top             =   15
      Width           =   2400
   End
   Begin VB.Label lblWindowTitle 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Imaging101 Message"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCopyToClipboard_Click()

    Clipboard.Clear
    Clipboard.SetText txtMessage.Text, vbCFText

End Sub

Private Sub Form_Load()

    Me.Caption = App.Title
    
    lblVersion.Caption = " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    lblWindowTitle = Me.Caption
    
End Sub

Private Sub OKButton_Click()

    Unload Me

End Sub

Public Sub subDisplayMessage(strMessage)
    
'    frmMessageForm.Height = frmMessageForm.Height + 100
'    txtMessage.Height = txtMessage.Height + 100
'    txtMessage.Index = 0

    txtMessage = funcRemovePasswordFromText(strMessage)
    'Scroll Text Box to Bottom
    txtMessage.SelStart = Len(txtMessage.Text)

    cmdCopyToClipboard.SetFocus
    
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    OKButton.SetFocus
End Sub

Private Sub txtMessage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OKButton.SetFocus

End Sub

Private Sub txtMessage_Validate(Cancel As Boolean)
    Cancel = True
End Sub
