VERSION 5.00
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.6#0"; "PIXEZOCX.OCX"
Begin VB.Form frmScannerSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scanner Settings"
   ClientHeight    =   3465
   ClientLeft      =   1245
   ClientTop       =   1890
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Region Selection"
      Height          =   732
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   6852
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting12 
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   2295
         _Version        =   65542
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         ControlType     =   21
         Data            =   0
         Caption         =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Re&gion:"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.CommandButton cmdArea 
      Caption         =   "&Area..."
      Height          =   400
      Left            =   7224
      TabIndex        =   34
      Top             =   1644
      Width           =   1212
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   400
      Left            =   7224
      TabIndex        =   33
      Top             =   1020
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   400
      Left            =   7224
      TabIndex        =   32
      Top             =   552
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   7224
      TabIndex        =   31
      Top             =   120
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Contrast"
      Height          =   972
      Left            =   3456
      TabIndex        =   23
      Top             =   1524
      Width           =   3564
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting9 
         Height          =   204
         Left            =   3072
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   576
         Width           =   396
         _Version        =   65542
         _ExtentX        =   688
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   14
         Data            =   0
         Caption         =   "ContrastValue"
      End
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting8 
         Height          =   204
         Left            =   480
         TabIndex        =   27
         Top             =   576
         Width           =   2220
         _Version        =   65542
         _ExtentX        =   3916
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   13
         Data            =   0
         Caption         =   ""
      End
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting7 
         Height          =   204
         Left            =   1836
         TabIndex        =   25
         Top             =   288
         Width           =   1200
         _Version        =   65542
         _ExtentX        =   2117
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   12
         Data            =   0
         Caption         =   "Automatic"
      End
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting6 
         Height          =   210
         Left            =   195
         TabIndex        =   24
         Top             =   285
         Width           =   1155
         _Version        =   65542
         _ExtentX        =   2037
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   11
         Data            =   0
         Caption         =   "Manual"
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   204
         Left            =   2760
         TabIndex        =   28
         Top             =   576
         Width           =   108
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   204
         Left            =   2784
         TabIndex        =   29
         Top             =   576
         Width           =   228
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   204
         Left            =   240
         TabIndex        =   26
         Top             =   576
         Width           =   204
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Brightness"
      Height          =   1356
      Left            =   3468
      TabIndex        =   13
      Top             =   0
      Width           =   3564
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting5 
         Height          =   270
         Left            =   2205
         TabIndex        =   22
         Top             =   960
         Width           =   885
         _Version        =   65542
         _ExtentX        =   1561
         _ExtentY        =   476
         _StockProps     =   77
         ControlType     =   10
         Data            =   0
         Caption         =   "&Lighten"
      End
      Begin PixezocxLib.PixEzScanSetting cmdNormal 
         Height          =   285
         Left            =   1245
         TabIndex        =   21
         Top             =   960
         Width           =   885
         _Version        =   65542
         _ExtentX        =   1561
         _ExtentY        =   503
         _StockProps     =   77
         ControlType     =   9
         Data            =   0
         Caption         =   "&Normal"
      End
      Begin PixezocxLib.PixEzScanSetting cmdDarken 
         Height          =   270
         Left            =   192
         TabIndex        =   20
         Top             =   960
         Width           =   972
         _Version        =   65542
         _ExtentX        =   1720
         _ExtentY        =   476
         _StockProps     =   77
         ControlType     =   8
         Data            =   0
         Caption         =   "Dar&ken"
      End
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting4 
         Height          =   210
         Left            =   2970
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   570
         Width           =   405
         _Version        =   65542
         _ExtentX        =   714
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   7
         Data            =   0
         Caption         =   "BrightnessValue"
      End
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting3 
         Height          =   204
         Left            =   480
         TabIndex        =   17
         Top             =   576
         Width           =   2220
         _Version        =   65542
         _ExtentX        =   3916
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   6
         Data            =   0
         Caption         =   ""
      End
      Begin PixezocxLib.PixEzScanSetting PixEzScanSetting2 
         Height          =   204
         Left            =   1836
         TabIndex        =   15
         Top             =   288
         Width           =   1260
         _Version        =   65542
         _ExtentX        =   2223
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   5
         Data            =   0
         Caption         =   "Automatic"
      End
      Begin PixezocxLib.PixEzScanSetting optManual 
         Height          =   204
         Left            =   192
         TabIndex        =   14
         Top             =   288
         Width           =   1068
         _Version        =   65542
         _ExtentX        =   1879
         _ExtentY        =   370
         _StockProps     =   77
         ControlType     =   4
         Data            =   0
         Caption         =   "Manual"
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   204
         Left            =   2760
         TabIndex        =   18
         Top             =   576
         Width           =   204
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   204
         Left            =   240
         TabIndex        =   16
         Top             =   576
         Width           =   204
      End
   End
   Begin PixezocxLib.PixEzScanSetting chkDuplex 
      Height          =   300
      Left            =   7230
      TabIndex        =   36
      Top             =   2910
      Width           =   1200
      _Version        =   65542
      _ExtentX        =   2117
      _ExtentY        =   529
      _StockProps     =   77
      ControlType     =   19
      Data            =   0
      Caption         =   "Duple&x"
   End
   Begin PixezocxLib.PixEzScanSetting cmdMore 
      Height          =   405
      Left            =   7230
      TabIndex        =   35
      Top             =   2145
      Width           =   1200
      _Version        =   65542
      _ExtentX        =   2117
      _ExtentY        =   714
      _StockProps     =   77
      ControlType     =   20
      Data            =   0
      Caption         =   "M&ore..."
   End
   Begin PixezocxLib.PixEzScanSetting cboPageSize 
      Height          =   315
      Left            =   195
      TabIndex        =   7
      Top             =   1530
      Width           =   2985
      _Version        =   65542
      _ExtentX        =   5265
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      ControlType     =   3
      Data            =   0
      Caption         =   ""
   End
   Begin PixezocxLib.PixEzScanSetting PixEzScanSetting1 
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   870
      Width           =   1890
      _Version        =   65542
      _ExtentX        =   3334
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      ControlType     =   15
      Data            =   0
      Caption         =   ""
   End
   Begin PixezocxLib.PixEzScanSetting cboDither 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   2220
      _Version        =   65542
      _ExtentX        =   3916
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      ControlType     =   2
      Data            =   0
      Caption         =   ""
   End
   Begin PixezocxLib.PixEzScanSetting cboMode 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   30
      Width           =   2220
      _Version        =   65542
      _ExtentX        =   3916
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      ControlType     =   1
      Data            =   0
      Caption         =   ""
   End
   Begin VB.Label lblScanDriverVersion 
      Caption         =   "SPEW Version 1.4"
      Height          =   300
      Left            =   192
      TabIndex        =   9
      Top             =   2208
      Width           =   2988
   End
   Begin VB.Label lblPixDfltVersion 
      Caption         =   "PIXDFLT Version 1.143"
      Height          =   204
      Left            =   204
      TabIndex        =   8
      Top             =   1920
      Width           =   2988
   End
   Begin VB.Label Label4 
      Caption         =   "&Page Size:"
      Height          =   204
      Left            =   192
      TabIndex        =   6
      Top             =   1248
      Width           =   876
   End
   Begin VB.Label Label3 
      Caption         =   "Dots pe&r inch:"
      Height          =   276
      Left            =   192
      TabIndex        =   4
      Top             =   864
      Width           =   1068
   End
   Begin VB.Label Label2 
      Caption         =   "D&ither:"
      Height          =   300
      Left            =   192
      TabIndex        =   2
      Top             =   480
      Width           =   588
   End
   Begin VB.Label Label1 
      Caption         =   "&Mode:"
      Height          =   264
      Left            =   180
      TabIndex        =   0
      Top             =   48
      Width           =   732
   End
End
Attribute VB_Name = "frmScannerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ctrActivePixEzImage As PixEzImage
Private lReturn As Long
Option Explicit


Private Sub cmdArea_Click()
    frmScanningArea.StartForm ctrActivePixEzImage
End Sub

Private Sub cmdCancel_Click()
    ctrActivePixEzImage.ScanStateRestore
    Me.Hide
    lReturn = 0
End Sub

Private Sub cmdDefault_Click()
    ctrActivePixEzImage.ScanStateRestoreDefault
End Sub

Private Sub cmdOK_Click()
    ctrActivePixEzImage.ScanStateFlush
    Me.Hide
    lReturn = 1
End Sub

Private Sub Form_Load()
   PixLocalizeForm Me
   ' center form
   '
   Me.Move (Screen.width - Me.width) / 2, _
        (Screen.Height - Me.Height) / 2
    lblPixDfltVersion = ctrActivePixEzImage.ScanPixdfltVersion
   lblScanDriverVersion = ctrActivePixEzImage.ScanDriverVersion
   
End Sub
Public Function StartForm(ctrPixEzImage As PixEzImage) As Long
    Set ctrActivePixEzImage = ctrPixEzImage
    If ctrActivePixEzImage.ScanLoaded = 0 Then
       On Error Resume Next
       ctrActivePixEzImage.ScanLoaded = 1
       If Err Then
         MsgBox (EW2W("Error Loading Scanner Driver: ") & Err.Description)
         StartForm = 0
         Exit Function
       End If
    End If
    Load Me
    ctrActivePixEzImage.ScanStateSave
    cmdOK.Default = 1
    Me.Show 1
    StartForm = lReturn
    Unload Me
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
            Cancel = True       'any non-zero value stops the unload
        ' if user closed thru control box
        cmdCancel_Click               '  treatit like it is cancel.
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ctrActivePixEzImage = Nothing
End Sub















