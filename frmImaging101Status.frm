VERSION 5.00
Begin VB.Form frmImaging101Status 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imaging101 Status"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   5910
   Icon            =   "frmImaging101Status.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5910
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      Picture         =   "frmImaging101Status.frx":0442
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   6
      Top             =   120
      Width           =   1572
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
      Left            =   4320
      TabIndex        =   7
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "NOTE: Closing this Window will CLOSE DOWN Imaging101"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Label lblDateTime 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Message"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   23.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imaging101 Server Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblServerStatus 
      Caption         =   "lblServerStatus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   1
      Top             =   3600
      Width           =   2292
   End
   Begin VB.Label lblMessage 
      Caption         =   "lblMessage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5415
   End
End
Attribute VB_Name = "frmImaging101Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Kill Imaging101 completelly unloading ALL forms, etc.
    End

End Sub
