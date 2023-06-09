VERSION 5.00
Begin VB.Form frmScanBatchCancel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scan Batch"
   ClientHeight    =   1500
   ClientLeft      =   2430
   ClientTop       =   1710
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdScanCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel Scanning.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   672
      Width           =   1788
   End
   Begin VB.Label lblScanStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   200
      TabIndex        =   0
      Top             =   192
      Width           =   2500
   End
End
Attribute VB_Name = "frmScanBatchCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private EzImage As PixEzImage
Private Sub cmdScanCancel_Click()
    lblScanStatus = EW2W("Canceling...")
    EzImage.ScanCancel = 1
    cmdScanCancel.enabled = False
End Sub

Public Sub StartForm(image As PixEzImage)
    lblScanStatus = EW2W("Scanning...")
    cmdScanCancel.enabled = True
    Set EzImage = image
    Me.Show
End Sub

Private Sub Form_Load()
   PixLocalizeForm Me
   ' center form
   '
   Me.Move (Screen.width - Me.width) / 2, _
        (Screen.Height - Me.Height) / 2
End Sub















