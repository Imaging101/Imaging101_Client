VERSION 5.00
Begin VB.Form frmImaging101BatchNotes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Batch Notes Expanded"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      Picture         =   "frmImaging101BatchNotes.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   0
      Width           =   1572
   End
   Begin VB.CommandButton cmdCopyToClipboard 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy To Clipboard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Picture         =   "frmImaging101BatchNotes.frx":0693
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox txtBatchNotes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6885
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   6240
      TabIndex        =   4
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch Notes Expanded"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmImaging101BatchNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCopyToClipboard_Click()

    Clipboard.Clear
    Clipboard.SetText txtBatchNotes.Text, vbCFText

End Sub


Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdCopyToClipboard.SetFocus
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    

End Sub

Private Sub txtBatchNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCopyToClipboard.SetFocus

End Sub

Private Sub txtBatchNotes_Validate(Cancel As Boolean)
    Cancel = True
End Sub

