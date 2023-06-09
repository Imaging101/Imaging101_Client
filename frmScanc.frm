VERSION 5.00
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "ImgScan.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "imgedit.ocx"
Begin VB.Form frmScan 
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   9810
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Scanning"
      TabPicture(0)   =   "frmScanc.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ImgEdit1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdImport"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdScan"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdScanPreferences"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSelectScanner"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ImgScan1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Configuration"
      TabPicture(1)   =   "frmScanc.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtFileSpec"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtFilePath"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Check1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CheckBox Check1 
         Caption         =   "Search ALL Subdirectories under this one?"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   2280
         Width           =   3495
      End
      Begin ScanLibCtl.ImgScan ImgScan1 
         Left            =   -74760
         Top             =   7560
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
         DestImageControl=   "ImgEdit1"
      End
      Begin VB.TextBox txtFilePath 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Text            =   "C:\Program Files\eCapture\Batches\001\00000002"
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtFileSpec 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "000*.*"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdSelectScanner 
         Caption         =   "Select Scanner"
         Height          =   495
         Left            =   -74760
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdScanPreferences 
         Caption         =   "Scan Preferences"
         Height          =   495
         Left            =   -74760
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74760
         TabIndex        =   3
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   495
         Left            =   -74760
         TabIndex        =   2
         Top             =   3600
         Width           =   1335
      End
      Begin ImgeditLibCtl.ImgEdit ImgEdit1 
         Height          =   7455
         Left            =   -72720
         TabIndex        =   1
         Top             =   600
         Width           =   6135
         _Version        =   131074
         _ExtentX        =   10821
         _ExtentY        =   13150
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         UndoBufferSize  =   74115584
         OcrZoneVisibility=   -3548
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3548
         sReserved2      =   -3548
         lReserved1      =   1241696
         lReserved2      =   1241696
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "File "
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Base Directory for Import"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdImport_Click()
    frmImport.Show
End Sub

Private Sub cmdScan_Click()

    ImgScan1.ScanTo = DisplayAndUseFileTemplate '3
    'Set the image property to a "Template" name.
    '  Use the same destination as configured in frmIndex
    ImgScan1.Image = frmScan.txtFilePath & "\000"
    'MultiPage property must be set to true in order to create
    'files with more than one page.
    ImgScan1.MultiPage = False
    'Create x page image files.
    ImgScan1.PageCount = 0
    'Do not show the scanner's TWAIN UI.
    ImgScan1.ShowSetupBeforeScan = False
    'Scan without using dialog box.
    ImgScan1.DestImageControl = "ImgEdit1"
    'ScrollBars already default to true, but this is
    'modifiable if desired
    ImgEdit1.ScrollBars = True
    
    'Repaint any changes immediately (ex. zoom or resolution
    'changes, etc.)
    ImgEdit1.AutoRefresh = False
    
    'Allow user to scroll image using keyboard shortcuts
    ImgEdit1.ScrollShortcutsEnabled = False
    
    Result = vbOK
    While Result = vbOK
        ImgScan1.StartScan
        'Scale black and white images to grayscale.  Keep color image
        'displayed as color images.
        ImgEdit1.DisplayScaleAlgorithm = wiScaleOptimize
        
        'Display entire image in control
        ImgEdit1.FitTo BEST_FIT

        DoEvents
        Result = MsgBox("Scan Next Page?", vbOKCancel, "Scan Paused")
    Wend
    
End Sub

Private Sub cmdScanPreferences_Click()
   ImgScan1.ShowScanPreferences

End Sub

Private Sub cmdSelectScanner_Click()
    ImgScan1.ShowSelectScanner

End Sub

Private Sub Form_Load()
    ' Get saved settings from the registry
    On Error Resume Next
    frmScan.Top = VBGetPrivateProfileString(RegAppname, "frmScan.Top", RegFileName)
    frmScan.Left = VBGetPrivateProfileString(RegAppname, "frmScan.Left", RegFileName)
    frmScan.Width = VBGetPrivateProfileString(RegAppname, "frmScan.Width", RegFileName)
    frmScan.Height = VBGetPrivateProfileString(RegAppname, "frmScan.Height", RegFileName)
    On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Form settings to the registry
    Result = WritePrivateProfileString(RegAppname, "frmScan.Top", frmScan.Top, RegFileName)
    Result = WritePrivateProfileString(RegAppname, "frmScan.Left", frmScan.Left, RegFileName)
    Result = WritePrivateProfileString(RegAppname, "frmScan.Width", frmScan.Width, RegFileName)
    Result = WritePrivateProfileString(RegAppname, "frmScan.Height", frmScan.Height, RegFileName)
    
    Me.Hide
    frmDocumentList.Show

End Sub
