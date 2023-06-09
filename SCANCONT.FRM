VERSION 5.00
Begin VB.Form frmContinueBatchScan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prepare Scanner"
   ClientHeight    =   4590
   ClientLeft      =   5490
   ClientTop       =   4470
   ClientWidth     =   7095
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
   ScaleHeight     =   4590
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4560
      Picture         =   "SCANCONT.frx":0000
      ScaleHeight     =   405
      ScaleWidth      =   1440
      TabIndex        =   12
      Top             =   0
      Width           =   1440
   End
   Begin VB.CheckBox chkUseFlatbed 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Use Flatbed"
      Height          =   228
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   2100
   End
   Begin VB.CommandButton butStop 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "S&top Scanning"
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
      Height          =   1095
      Left            =   3840
      Picture         =   "SCANCONT.frx":0693
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton butGo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Continue Scanning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Picture         =   "SCANCONT.frx":0C1D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Page Count"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   3012
      Begin VB.TextBox NoOfPages 
         Height          =   288
         Left            =   1920
         TabIndex        =   8
         Text            =   "1"
         Top             =   480
         Width           =   852
      End
      Begin VB.OptionButton radioNumberofPages 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Number of pages:"
         Height          =   252
         Left            =   100
         TabIndex        =   7
         Top             =   480
         Width           =   1692
      End
      Begin VB.OptionButton radioAllSheets 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&All sheets in stack"
         Height          =   252
         Left            =   100
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1692
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3012
      Begin VB.OptionButton radioFront 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Back of sheet %n"
         Height          =   252
         Index           =   1
         Left            =   100
         TabIndex        =   4
         Top             =   720
         Width           =   1692
      End
      Begin VB.OptionButton radioFront 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Front of sheet %n"
         Height          =   252
         Index           =   0
         Left            =   100
         TabIndex        =   3
         Top             =   360
         Width           =   1692
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5532
      Begin VB.Label ScanName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scanner Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5364
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAGES SCANNED "
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label txtPagesScanned 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   1320
      Width           =   1410
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   4560
      TabIndex        =   13
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "frmContinueBatchScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ctrActivePixEzImage As PixEzImage
Private lReturn As Long
Private go As Boolean

Public Sub StartForm(ctrPixEzImage As PixEzImage)

  Dim NumPages As Long
  Dim StopScan As Long
  Dim UseFeeder As Long
  Dim ScanBacks As Long
  On Error Resume Next
    Set ctrActivePixEzImage = ctrPixEzImage
    Dim iPage As Integer
    If (ctrActivePixEzImage.ScanMaxPages > 0) Then
        NoOfPages = ctrActivePixEzImage.ScanMaxPages
        radioNumberofPages.enabled = True
        radioNumberofPages.Value = True
        radioAllSheets.enabled = True
    Else
        ctrActivePixEzImage.ScanMaxPages = -1
        NoOfPages = ""
        radioAllSheets.enabled = True
        radioNumberofPages.enabled = False
        radioAllSheets.Value = True
    End If
    
    txtPagesScanned = ctrActivePixEzImage.ScanMaxPages
    ScanName.Caption = ctrActivePixEzImage.ScanName
    
''''    Load frmContinueBatchScan
    
    go = False
    radioFront(0).Value = True
    radioFront(0).enabled = True
    iPage = ctrActivePixEzImage.ScanCurrentPage
    radioFront(0).Caption = EW2W("&Front of sheet ") & (iPage)
    radioFront(1).Caption = EW2W("&Back of sheet ") & (iPage - 1)
    radioFront(1).enabled = True
    
    frmContinueBatchScan.Show vbModal
    
    StopScan = Not go
    ScanBacks = radioFront(1).Value
    
   
End Sub
Public Function ShowPrepareScannerDialog(ctrPixEzImage As PixEzImage, NumPages As Long, StopScan As Long, UseFeeder As Long, ScanBacks As Long) As Long
    
    Static lStaticUseFlatBed As Long
    On Error Resume Next
    Set ctrActivePixEzImage = ctrPixEzImage
    Dim iPage As Integer
    Dim bCanInsert As Boolean
    Const IF_FEATURE_INSERTPAGE = &H40
    bCanInsert = (ctrActivePixEzImage.PackagingFeatures And IF_FEATURE_INSERTPAGE) = IF_FEATURE_INSERTPAGE

    If (ctrActivePixEzImage.ScanMaxPages > 0) Then
        NoOfPages = ctrActivePixEzImage.ScanMaxPages
        radioNumberofPages.enabled = True
        radioNumberofPages.Value = True
        radioAllSheets.enabled = True
    Else
        ctrActivePixEzImage.ScanMaxPages = -1
        NoOfPages = ""
        radioAllSheets.enabled = True
        radioNumberofPages.enabled = False
        radioAllSheets.Value = True
    End If
    ScanName.Caption = ctrActivePixEzImage.ScanName
    
''''    Load frmContinueBatchScan
    
    go = False
    radioFront(0).Value = True
    radioFront(0).enabled = True
    iPage = ctrActivePixEzImage.ScanCurrentPage
    radioFront(0).Caption = EW2W("&Front of sheet ") & (iPage)
    radioFront(1).Caption = EW2W("&Back of sheet ") & (iPage - 1)
    If bCanInsert And NumPages > 0 And Not ctrActivePixEzImage.ScanInsertMode = 2 Then
        radioFront(1).enabled = True
    Else
        radioFront(1).enabled = False
    End If
    chkUseFlatbed.Value = lStaticUseFlatBed
     
    frmContinueBatchScan.Show vbModal

    StopScan = Not go
    ScanBacks = radioFront(1).Value
    If chkUseFlatbed.Value = 1 Then
        UseFeeder = 0
    End If
    lStaticUseFlatBed = chkUseFlatbed.Value
    ShowPrepareScannerDialog = lReturn
    Unload Me
    
End Function

Private Sub butGo_Click()
    go = True
    If (radioAllSheets.Value) Then
        ctrActivePixEzImage.ScanMaxPages = -1
    Else
        If ctrActivePixEzImage.ScanDuplex = 1 And (Val(NoOfPages) Mod 2) <> 0 Then
            MsgBox EW2W("Number of pages should be even for a Duplex Scan")
            go = False
            Exit Sub
        End If
        ctrActivePixEzImage.ScanMaxPages = Val(NoOfPages)
    End If
    frmContinueBatchScan.Hide
End Sub

Private Sub butStop_Click()

    go = False
    frmContinueBatchScan.Hide
    
End Sub

Private Sub Form_Activate()

    DoEvents
    Me.SetFocus
    
    '2017-08-24 - Jacob - Changed FOCUS from Button GO (Continue Scanning) to "STOP"
    'butGo.SetFocus
    butStop.SetFocus

End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
   PixLocalizeForm Me
   ' center form
   '
   Me.Move (Screen.width - Me.width) / 2, _
        (Screen.Height - Me.Height) / 2
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set ctrActivePixEzImage = Nothing
End Sub

Private Sub NoOfPages_Change()
        radioNumberofPages.enabled = Val(NoOfPages) > 0
        radioNumberofPages.Value = Val(NoOfPages) > 0
        radioAllSheets.Value = Not (Val(NoOfPages) > 0)
End Sub















