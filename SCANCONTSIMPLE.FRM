VERSION 5.00
Begin VB.Form frmContinueBatchScanSimple 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Continue Scan?"
   ClientHeight    =   1770
   ClientLeft      =   5490
   ClientTop       =   4470
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4560
      Picture         =   "SCANCONTSIMPLE.frx":0000
      ScaleHeight     =   405
      ScaleWidth      =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   1440
   End
   Begin VB.CommandButton butStop 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&CANCEL SCAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Picture         =   "SCANCONTSIMPLE.frx":0693
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2292
   End
   Begin VB.CommandButton butGo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SCAN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Picture         =   "SCANCONTSIMPLE.frx":0C1D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2292
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "frmContinueBatchScanSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bolScanContinue As Boolean

Private Sub butGo_Click()
    
    bolScanContinue = True
    Me.Hide
    
End Sub

Private Sub butStop_Click()

    bolScanContinue = False
    Me.Hide
    
End Sub

Public Function subScanContinue() As Boolean

    Me.Show vbModal
    subScanContinue = bolScanContinue
    Unload Me
    
End Function

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub
