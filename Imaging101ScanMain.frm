VERSION 5.00
Object = "{C1AD690C-829F-4862-9CA2-61B9A6A815E4}#1.0#0"; "TwnPro3.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Imaging101ScanMain 
   BackColor       =   &H80000013&
   Caption         =   "Imaging101 Batch Scan"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   12255
   Icon            =   "Imaging101ScanMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   12255
   Begin VB.CommandButton cmdScannerAdvancedCapabilitiesLoad 
      Caption         =   "LOAD Advanced Scanner Capabilities"
      Height          =   495
      Left            =   9840
      TabIndex        =   127
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtBatchDirectory 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   107
      Top             =   1800
      Width           =   8175
   End
   Begin VB.CommandButton cmdBatchDirectoryFind 
      Caption         =   "..."
      Height          =   255
      Left            =   10800
      TabIndex        =   30
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtBatchRootDirectory 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   1440
      Width           =   8175
   End
   Begin VB.TextBox cmbBatchScanSettingsDesc 
      Height          =   495
      Left            =   2520
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   480
      Width           =   5295
   End
   Begin VB.ComboBox cmbBatchScanSettingsName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2520
      TabIndex        =   27
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   26
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtApplicationName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   25
      Top             =   1080
      Width           =   4935
   End
   Begin VB.CommandButton cmdScanBegin 
      BackColor       =   &H0080FF80&
      Caption         =   "&Begin Scanning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   1875
   End
   Begin VB.TextBox txtTwainAcquireStatus 
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   7800
      Width           =   4215
   End
   Begin VB.CommandButton cmdScanStop 
      BackColor       =   &H008080FF&
      Caption         =   "&End Scanning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
      Width           =   1875
   End
   Begin VB.Frame frameImageLayout 
      Caption         =   "Image Picking Rectangle"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   11895
      Begin VB.ComboBox cmbTwainDocumentSize 
         Height          =   315
         ItemData        =   "Imaging101ScanMain.frx":0442
         Left            =   6240
         List            =   "Imaging101ScanMain.frx":0467
         TabIndex        =   10
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtTwainImageBottom 
         Height          =   315
         Left            =   5040
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTwainImageRight 
         Height          =   315
         Left            =   3840
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTwainImageTop 
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTwainImageLeft 
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkUseFlatBed 
         Alignment       =   1  'Right Justify
         Caption         =   "Use FlatBed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   9360
         MaskColor       =   &H00008000&
         TabIndex        =   14
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Paper Size"
         Height          =   195
         Left            =   6240
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "ImageBottom"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "ImageRight"
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "ImageTop"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "ImageLeft"
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Batch Settings"
      TabPicture(0)   =   "Imaging101ScanMain.frx":0564
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtBatchBoxNumber"
      Tab(0).Control(1)=   "cmbBatchGroup"
      Tab(0).Control(2)=   "cmbBatchOwner"
      Tab(0).Control(3)=   "cmbBatchQueue"
      Tab(0).Control(4)=   "txtBatchRECID"
      Tab(0).Control(5)=   "txtBatchDesc"
      Tab(0).Control(6)=   "txtBatchNotes"
      Tab(0).Control(7)=   "txtBatchSuffix"
      Tab(0).Control(8)=   "txtBatchPrefix"
      Tab(0).Control(9)=   "txtBatchScanUser"
      Tab(0).Control(10)=   "cmbBatchPriority"
      Tab(0).Control(11)=   "cmbBatchStatus"
      Tab(0).Control(12)=   "txtBatchName"
      Tab(0).Control(13)=   "lblBatchBoxNumber"
      Tab(0).Control(14)=   "Label35"
      Tab(0).Control(15)=   "Label34"
      Tab(0).Control(16)=   "lblBatchQueue"
      Tab(0).Control(17)=   "Label30"
      Tab(0).Control(18)=   "Label29"
      Tab(0).Control(19)=   "Label24"
      Tab(0).Control(20)=   "Label23"
      Tab(0).Control(21)=   "Label22"
      Tab(0).Control(22)=   "Label21"
      Tab(0).Control(23)=   "Label20"
      Tab(0).Control(24)=   "lblBatchName"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Scanner Settings"
      TabPicture(1)   =   "Imaging101ScanMain.frx":0580
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdScannerSettingsSave"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameImageQuality"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdScannerSettingsCancel"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frameMisc"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frameBatchDefaults"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdScannerSettingsApply"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FrameBatchOptions"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbTwainPageSize"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Image Caption Settings"
      TabPicture(2)   =   "Imaging101ScanMain.frx":059C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrmCaption"
      Tab(2).Control(1)=   "cmdScannerCaptionSettingsSave"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Advanced Settings"
      TabPicture(3)   =   "Imaging101ScanMain.frx":05B8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmCaps"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtBatchBoxNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -66480
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ComboBox cmbTwainPageSize 
         Height          =   315
         ItemData        =   "Imaging101ScanMain.frx":05D4
         Left            =   4680
         List            =   "Imaging101ScanMain.frx":05D6
         TabIndex        =   126
         Top             =   4080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ComboBox cmbBatchGroup 
         Height          =   315
         ItemData        =   "Imaging101ScanMain.frx":05D8
         Left            =   -66480
         List            =   "Imaging101ScanMain.frx":05E2
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   3360
         Width           =   3135
      End
      Begin VB.ComboBox cmbBatchOwner 
         Height          =   315
         ItemData        =   "Imaging101ScanMain.frx":05F6
         Left            =   -66480
         List            =   "Imaging101ScanMain.frx":05F8
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   3135
      End
      Begin VB.ComboBox cmbBatchQueue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Imaging101ScanMain.frx":05FA
         Left            =   -66480
         List            =   "Imaging101ScanMain.frx":05FC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Frame FrameBatchOptions 
         Caption         =   "Batch Options"
         Height          =   3375
         Left            =   7320
         TabIndex        =   109
         Top             =   480
         Width           =   4215
         Begin VB.CheckBox chkScanPreviewOnly 
            Alignment       =   1  'Right Justify
            Caption         =   "Preview Image Only (NO Save)"
            Height          =   195
            Left            =   120
            TabIndex        =   131
            Top             =   840
            Width           =   2715
         End
         Begin VB.CheckBox chkAutoDetectPaperOut 
            Alignment       =   1  'Right Justify
            Caption         =   "Automatically Detect Paper Out"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   600
            Value           =   1  'Checked
            Width           =   2715
         End
         Begin VB.TextBox txtMinimumImageSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   122
            Text            =   "1000"
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbImageRotation 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":05FE
            Left            =   840
            List            =   "Imaging101ScanMain.frx":060E
            TabIndex        =   116
            Text            =   "0"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtScanImageSkipCount 
            Height          =   285
            Left            =   1680
            TabIndex        =   113
            Text            =   "10"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CheckBox chkScanDisplayImages 
            Alignment       =   1  'Right Justify
            Caption         =   "Display Images While Scanning"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   1440
            Width           =   2715
         End
         Begin VB.CheckBox chkScanResetScanner 
            Alignment       =   1  'Right Justify
            Caption         =   "Reset Scanner before each scan"
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2715
         End
         Begin VB.CheckBox chkScanShowUI 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Manuf. User Interface"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   360
            Width           =   2715
         End
         Begin VB.Label Label37 
            Caption         =   "Bytes"
            Height          =   255
            Left            =   2880
            TabIndex        =   124
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label36 
            Caption         =   "Delete Images Smaller Than"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label33 
            Caption         =   "Rotation"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Images"
            Height          =   255
            Left            =   2160
            TabIndex        =   115
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Display every"
            Height          =   255
            Left            =   600
            TabIndex        =   114
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdScannerSettingsApply 
         Caption         =   "&Apply Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9720
         TabIndex        =   93
         Top             =   3960
         Width           =   855
      End
      Begin VB.Frame frameBatchDefaults 
         Caption         =   "Batch Defaults"
         Height          =   2295
         Left            =   4680
         TabIndex        =   85
         Top             =   1680
         Width           =   2535
         Begin VB.CheckBox chkBatchBoxNumberRequired 
            Alignment       =   1  'Right Justify
            Caption         =   "BOX # Required"
            Height          =   195
            Left            =   240
            TabIndex        =   130
            Top             =   1920
            Width           =   1875
         End
         Begin VB.CheckBox chkBatchAutoUseDateTime 
            Alignment       =   1  'Right Justify
            Caption         =   "Use Date/Time"
            Height          =   195
            Left            =   240
            TabIndex        =   90
            Top             =   1560
            Width           =   1515
         End
         Begin VB.CheckBox chkBatchAutoUseBatchID 
            Alignment       =   1  'Right Justify
            Caption         =   "Use BatchID  "
            Height          =   195
            Left            =   240
            TabIndex        =   89
            Top             =   1320
            Width           =   1515
         End
         Begin VB.CheckBox chkBatchAutoName 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto-Assign BatchID"
            Height          =   195
            Left            =   240
            TabIndex        =   88
            Top             =   1080
            Width           =   1875
         End
         Begin VB.TextBox txtBatchSettingsPrefix 
            Height          =   285
            Left            =   120
            TabIndex        =   87
            ToolTipText     =   "Characters at the Beginning of the Batch name"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtBatchSettingsSuffix 
            Height          =   285
            Left            =   1320
            TabIndex        =   86
            ToolTipText     =   "Characters at the Beginning of the Batch name"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "Batch Prefix"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Batch Suffix"
            Height          =   255
            Left            =   1320
            TabIndex        =   91
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbApplicationList 
         Height          =   315
         Left            =   1920
         TabIndex        =   84
         Top             =   -2160
         Width           =   4935
      End
      Begin VB.CommandButton cmdScannerCaptionSettingsSave 
         Caption         =   "&Save Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -66120
         TabIndex        =   83
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtBatchRECID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70680
         TabIndex        =   82
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBatchDesc 
         Height          =   285
         Left            =   -73560
         TabIndex        =   3
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox txtBatchNotes 
         Height          =   765
         Left            =   -73560
         TabIndex        =   4
         Top             =   1920
         Width           =   5295
      End
      Begin VB.TextBox txtBatchSuffix 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -69000
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtBatchPrefix 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -74520
         TabIndex        =   0
         Top             =   1080
         Width           =   975
      End
      Begin VB.Frame frameMisc 
         Caption         =   "Scan Defaults"
         Height          =   1335
         Left            =   240
         TabIndex        =   75
         Top             =   360
         Width           =   6975
         Begin VB.ComboBox cmbTwainTransferMode 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":0623
            Left            =   1800
            List            =   "Imaging101ScanMain.frx":062D
            TabIndex        =   78
            Top             =   960
            Width           =   4215
         End
         Begin VB.ComboBox cmbTwainScanMethod 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":064A
            Left            =   1800
            List            =   "Imaging101ScanMain.frx":0654
            TabIndex        =   77
            Top             =   600
            Width           =   4215
         End
         Begin VB.ComboBox cmbTwainSourceName 
            Height          =   315
            Left            =   1800
            TabIndex        =   76
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label31 
            Caption         =   "TwainTrasferMode"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "TwainScanMethod"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "TwainSourceName"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdScannerSettingsCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   74
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox txtBatchScanUser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73920
         TabIndex        =   73
         Top             =   4080
         Width           =   2295
      End
      Begin VB.ComboBox cmbBatchPriority 
         Height          =   315
         ItemData        =   "Imaging101ScanMain.frx":0674
         Left            =   -66480
         List            =   "Imaging101ScanMain.frx":0676
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox cmbBatchStatus 
         Height          =   315
         ItemData        =   "Imaging101ScanMain.frx":0678
         Left            =   -66480
         List            =   "Imaging101ScanMain.frx":067A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtBatchName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73560
         TabIndex        =   1
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Frame frameImageQuality 
         Caption         =   "Image Quality"
         Height          =   2895
         Left            =   240
         TabIndex        =   60
         Top             =   1680
         Width           =   4335
         Begin VB.CommandButton cmdScannerAdvancedCapabilitiesSet 
            Caption         =   "Scanner Settings"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   1080
            TabIndex        =   128
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtTwainIntensity 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   64
            Top             =   2160
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtTwainContrast 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   63
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cmbTwainResolution 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":067C
            Left            =   1440
            List            =   "Imaging101ScanMain.frx":069E
            TabIndex        =   62
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cmbTwainColor 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":06D0
            Left            =   1440
            List            =   "Imaging101ScanMain.frx":06E0
            TabIndex        =   61
            Top             =   480
            Width           =   2655
         End
         Begin MSComctlLib.Slider sldTwainContrast 
            Height          =   495
            Left            =   480
            TabIndex        =   65
            Top             =   1440
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   20
            SmallChange     =   10
            Min             =   -2000
            Max             =   2000
            TickFrequency   =   100
         End
         Begin MSComctlLib.Slider sldTwainIntensity 
            Height          =   495
            Left            =   480
            TabIndex        =   66
            Top             =   2280
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   20
            SmallChange     =   10
            Min             =   -2000
            Max             =   2000
            TickFrequency   =   100
         End
         Begin VB.Label lblBrightnessDefault 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1440
            TabIndex        =   72
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblContrastDefault 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1440
            TabIndex        =   71
            Top             =   1200
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "TwainIntensity"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   2040
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "TwainContrast"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Image Resolution"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   68
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Save Image As"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdScannerSettingsSave 
         Caption         =   "Sa&ve Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10680
         TabIndex        =   59
         Top             =   3960
         Width           =   855
      End
      Begin VB.Frame FrmCaption 
         Caption         =   "Caption to Display on Image"
         Height          =   3615
         Left            =   -73320
         TabIndex        =   42
         Top             =   480
         Width           =   6375
         Begin VB.CheckBox chkCaptionClip 
            Alignment       =   1  'Right Justify
            Caption         =   "Clip Caption"
            Height          =   195
            Left            =   360
            TabIndex        =   51
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbCaptionVerticalAlign 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":073B
            Left            =   2400
            List            =   "Imaging101ScanMain.frx":0748
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   2565
            Width           =   1635
         End
         Begin VB.ComboBox cmbCaptionHorizontalAlign 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":0773
            Left            =   2400
            List            =   "Imaging101ScanMain.frx":0780
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1725
            Width           =   1635
         End
         Begin VB.TextBox txtCaptionWidth 
            Height          =   285
            Left            =   1380
            TabIndex        =   48
            Text            =   "0"
            Top             =   2160
            Width           =   675
         End
         Begin VB.TextBox txtCaptionHeight 
            Height          =   285
            Left            =   1380
            TabIndex        =   47
            Text            =   "0"
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkCaptionShadowText 
            Alignment       =   1  'Right Justify
            Caption         =   "Shadow Text"
            Height          =   195
            Left            =   2520
            TabIndex        =   46
            Top             =   960
            Width           =   1260
         End
         Begin VB.TextBox txtCaption 
            Height          =   285
            Left            =   1380
            TabIndex        =   45
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox txtCaptionLeft 
            Height          =   285
            Left            =   1380
            TabIndex        =   44
            Text            =   "0"
            Top             =   1320
            Width           =   675
         End
         Begin VB.TextBox txtCaptionTop 
            Height          =   285
            Left            =   1380
            TabIndex        =   43
            Text            =   "0"
            Top             =   1740
            Width           =   675
         End
         Begin VB.Label Label12 
            Caption         =   "Vertical Alignment"
            Height          =   255
            Left            =   2400
            TabIndex        =   58
            Top             =   2340
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Horizontal Alignment"
            Height          =   195
            Left            =   2400
            TabIndex        =   57
            Top             =   1500
            Width           =   1515
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionHeight:"
            Height          =   255
            Left            =   300
            TabIndex        =   56
            Top             =   2625
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionWidth:"
            Height          =   255
            Left            =   300
            TabIndex        =   55
            Top             =   2205
            Width           =   1035
         End
         Begin VB.Label Label9 
            Caption         =   "Caption:"
            Height          =   255
            Left            =   720
            TabIndex        =   54
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionLeft:"
            Height          =   255
            Left            =   420
            TabIndex        =   53
            Top             =   1365
            Width           =   915
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionTop:"
            Height          =   255
            Left            =   420
            TabIndex        =   52
            Top             =   1785
            Width           =   915
         End
      End
      Begin VB.Frame FrmCaps 
         Caption         =   "Capabilities"
         Height          =   3855
         Left            =   -72120
         TabIndex        =   32
         Top             =   480
         Width           =   6735
         Begin VB.TextBox txtCapsListIndex 
            Height          =   315
            Left            =   4680
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox CmbCaps 
            Height          =   315
            ItemData        =   "Imaging101ScanMain.frx":07AB
            Left            =   480
            List            =   "Imaging101ScanMain.frx":0806
            TabIndex        =   36
            Top             =   360
            Width           =   3495
         End
         Begin VB.ListBox List1 
            Height          =   2010
            Left            =   480
            TabIndex        =   35
            Top             =   840
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.TextBox EdtCurrent 
            Height          =   315
            Left            =   1740
            TabIndex        =   34
            Top             =   3120
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdateCapability 
            Caption         =   "Update"
            Height          =   315
            Left            =   3120
            TabIndex        =   33
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label LblDefault 
            Caption         =   "LblDefault"
            Height          =   195
            Left            =   480
            TabIndex        =   41
            Top             =   840
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label LblMin 
            Caption         =   "LblMin"
            Height          =   195
            Left            =   480
            TabIndex        =   40
            Top             =   1200
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label LblMax 
            Caption         =   "LblMax"
            Height          =   195
            Left            =   480
            TabIndex        =   39
            Top             =   1560
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Current Value:"
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   3180
            Width           =   1155
         End
      End
      Begin VB.Label lblBatchBoxNumber 
         BackColor       =   &H80000016&
         Caption         =   "Box #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -67200
         TabIndex        =   129
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000016&
         Caption         =   "Batch Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -67680
         TabIndex        =   121
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000016&
         Caption         =   "Route To User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68040
         TabIndex        =   119
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblBatchQueue 
         BackColor       =   &H80000016&
         Caption         =   "Route To Queue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -68040
         TabIndex        =   118
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "Prefix"
         Height          =   255
         Left            =   -74520
         TabIndex        =   101
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Suffix"
         Height          =   255
         Left            =   -69000
         TabIndex        =   100
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Scan User"
         Height          =   255
         Left            =   -74760
         TabIndex        =   99
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Batch Priority"
         Height          =   255
         Left            =   -67560
         TabIndex        =   98
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Batch Status"
         Height          =   255
         Left            =   -67560
         TabIndex        =   97
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Batch Notes"
         Height          =   255
         Left            =   -74880
         TabIndex        =   96
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Batch Description"
         Height          =   255
         Left            =   -74880
         TabIndex        =   95
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblBatchName 
         Caption         =   "Batch Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -73560
         TabIndex        =   94
         Top             =   840
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Settings Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   108
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   103
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   106
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Root Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   105
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   104
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   102
      Top             =   8160
      Width           =   11895
   End
   Begin TWNPRO3LibCtl.TwainPRO TwainPRO 
      Left            =   0
      Top             =   0
      ErrStr          =   "1MFH00S0GEP-GB3067SXEP"
      ErrCode         =   1575969092
      ErrInfo         =   445070108
      _cx             =   847
      _cy             =   847
      Caption         =   ""
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ProductName     =   ""
      ProductFamily   =   ""
      Manufacturer    =   ""
      VersionInfo     =   ""
      MaxImages       =   -1
      ShowUI          =   -1  'True
      SaveJPGLumFactor=   32
      SaveJPGChromFactor=   36
      SaveJPGSubSampling=   2
      SaveJPGProgressive=   0   'False
      PICPassword     =   ""
      FTPUserName     =   ""
      FTPPassword     =   ""
      ProxyServer     =   ""
      SaveTIFCompression=   4
      SaveMultiPage   =   0   'False
      CaptionLeft     =   0
      CaptionTop      =   0
      CaptionWidth    =   0
      CaptionHeight   =   0
      ShadowText      =   -1  'True
      ClipCaption     =   0   'False
      HAlign          =   0
      VAlign          =   0
      EnableExtCaps   =   -1  'True
      CloseOnCancel   =   -1  'True
      TransferMode    =   0
   End
End
Attribute VB_Name = "Imaging101ScanMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim m_ImageCount As Integer
    Dim m_PageCount As Integer
    Dim m_ImageSkipCount As Integer
    Dim intLoop As Integer
    
    
    Dim bolCancelPendingXfers As Boolean
    Dim bolBatchCreated As Boolean
    
    Dim strFullBatchDirectory
    Dim sErrMessage As String

    ' Set up for Imaging101 DB connection
    Dim connImaging101 As ADODB.Connection
    Dim cmdImaging101 As ADODB.Command
    Dim rsImaging101 As ADODB.Recordset
    
    ' Set up for Imaging101Batch DB connection
    Dim connImaging101Batch As ADODB.Connection
    Dim cmdImaging101Batch As ADODB.Command
    Dim rsImaging101Batch As ADODB.Recordset
    Dim rsImaging101BatchPage As ADODB.Recordset
    
    ' For the load / save Scanner Settings
    Dim fso As FileSystemObject
    Dim fsoScannerSettings
    Dim strScannerSettingsFileName As String



Private Sub chkBatchAutoName_Click()

    On Error GoTo error_handler
    
    If ItoB(chkBatchAutoName) = True Then
        chkBatchAutoUseBatchID.Enabled = True
        chkBatchAutoUseDateTime.Enabled = True
    Else
        chkBatchAutoUseBatchID.Enabled = False
        chkBatchAutoUseDateTime.Enabled = False
    End If
    
Exit Sub

error_handler:

'    MsgBox "chkBatchAutoName_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
    
End Sub

Private Sub chkUseFlatBed_Click()
    
    On Error GoTo error_handler
    
    ' Session MUST be OPEN to set the Feeder
    TwainPRO.Capability = TWCAP_FEEDERENABLED
    If TwainPRO.CapSupported Then
        If ItoB(chkUseFlatbed) = True Then
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = 0 ' Sets document acquisition source to FLATBED if available
            TwainPRO.SetCapOut
        Else
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = 1 ' Sets document acquisition source to ADF if available
            TwainPRO.SetCapOut
        End If
    End If
    

Exit Sub

error_handler:

'    MsgBox "chkUseFlatBed_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub

Private Sub cmbBatchQueue_Click()

    txtBatchName.SetFocus

End Sub

Private Sub cmbBatchScanSettingsName_click()

    On Error GoTo error_handler
    
    subScannerSettingsGetSettings
    

    
Exit Sub

error_handler:

'    MsgBox "cmbBatchScanSettingsName_click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub

' Cababilities selection
Private Sub CmbCaps_Click()
    
    
    On Error GoTo error_handler
    
    Dim Index As Long
    
    ' Only process when session is open and FrmCaps can only be visible
    ' after OpenSession is called
    ''If FrmCaps.Visible = False Then
    ''   Exit Sub
    ''End If
    
'JR-7/19/2004 Disabled Close/Open Session
'    TwainPRO.CloseSession
'    TwainPRO.OpenSession
    
    ' Reset all the controls in FrmCaps
    LblDefault.Visible = False
    LblMin.Visible = False
    LblMax.Visible = False
    List1.Visible = False
    List1.Clear
    EdtCurrent.Text = ""
    cmdUpdateCapability.Enabled = False
    TwainPRO.Capability = CmbCaps.ListIndex
    txtCapsListIndex = CmbCaps.ListIndex
        
    ' Check for support and stop if none
    If (TwainPRO.CapSupported = False) Then
       LblDefault.Caption = "Not supported"
       LblDefault.Visible = True
       Exit Sub
    End If
    
    ' What data type is the Cap?  We have to check the data type
    ' to know which properties are valid
    Select Case TwainPRO.CapType
       ' Type ONEVALUE only returns a single value
       Case TWCAP_ONEVALUE
          EdtCurrent.Text = TwainPRO.CapValue
          EdtCurrent.Visible = True
          
       ' Type ENUM returns a list of legal values as well as current and
       ' default values.  A list of constants is returned and the CapDesc
       ' property can be used to find out what the constants mean
       Case TWCAP_ENUM
          EdtCurrent.Text = TwainPRO.CapValue
          EdtCurrent.Visible = True
          LblDefault.Caption = "Default = " & TwainPRO.CapDefault
          LblDefault.Visible = True
          List1.AddItem "Legal Values:"
          For Index = 0 To TwainPRO.CapNumItems - 1
             List1.AddItem TwainPRO.CapItem(Index) & " - " & TwainPRO.CapDesc(TwainPRO.CapItem(Index))
          Next
          List1.Visible = True
          
       ' Type ARRAY returns a list of values, but no current or default values
       ' This is a less common type that many sources don't use
       Case TWCAP_ARRAY
          List1.AddItem "Legal Values:"
          For Index = 0 To TwainPRO.CapNumItems - 1
             List1.AddItem TwainPRO.CapItem(Index)
          Next
       
       ' Returns a range of values as well as current and default values
       Case TWCAP_RANGE
          EdtCurrent.Text = TwainPRO.CapValue
          EdtCurrent.Visible = True
          LblDefault.Caption = "Default = " & TwainPRO.CapDefault
          LblDefault.Visible = True
          LblMin.Caption = "MinValue = " & TwainPRO.CapMin
          LblMin.Visible = True
          LblMax.Caption = "MaxValue = " & TwainPRO.CapMax
          LblMax.Visible = True
         
'    TwainPRO.CloseSession
    
    End Select
    cmdUpdateCapability.Enabled = True
    
Exit Sub

error_handler:

'    MsgBox "CmbCaps_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
  End Sub

' Cababilities selection
Private Sub CmbCaps_Save()
    
    
    On Error GoTo error_handler
    
    Dim Index As Long
    
    ' Only process when session is open and FrmCaps can only be visible
    ' after OpenSession is called
    ''If FrmCaps.Visible = False Then
    ''   Exit Sub
    ''End If
    
'JR-7/19/2004 Disabled Close/Open Session
'    TwainPRO.CloseSession
'    TwainPRO.OpenSession
    
    ' Reset all the controls in FrmCaps
    LblDefault.Visible = False
    LblMin.Visible = False
    LblMax.Visible = False
    List1.Visible = False
    List1.Clear
    EdtCurrent.Text = ""
    cmdUpdateCapability.Enabled = False
    TwainPRO.Capability = CmbCaps.ListIndex
    txtCapsListIndex = CmbCaps.ListIndex
        
    ' Check for support and stop if none
    If (TwainPRO.CapSupported = False) Then
       LblDefault.Caption = "Not supported"
       LblDefault.Visible = True
       Exit Sub
    End If
    
    ' What data type is the Cap?  We have to check the data type
    ' to know which properties are valid
    Select Case TwainPRO.CapType
       ' Type ONEVALUE only returns a single value
       Case TWCAP_ONEVALUE
          EdtCurrent.Text = TwainPRO.CapValue
          EdtCurrent.Visible = True
          
       ' Type ENUM returns a list of legal values as well as current and
       ' default values.  A list of constants is returned and the CapDesc
       ' property can be used to find out what the constants mean
       Case TWCAP_ENUM
          EdtCurrent.Text = TwainPRO.CapValue
          EdtCurrent.Visible = True
          LblDefault.Caption = "Default = " & TwainPRO.CapDefault
          LblDefault.Visible = True
          List1.AddItem "Legal Values:"
          For Index = 0 To TwainPRO.CapNumItems - 1
             List1.AddItem TwainPRO.CapItem(Index) & " - " & TwainPRO.CapDesc(TwainPRO.CapItem(Index))
          Next
          List1.Visible = True
          
       ' Type ARRAY returns a list of values, but no current or default values
       ' This is a less common type that many sources don't use
       Case TWCAP_ARRAY
          List1.AddItem "Legal Values:"
          For Index = 0 To TwainPRO.CapNumItems - 1
             List1.AddItem TwainPRO.CapItem(Index)
          Next
       
       ' Returns a range of values as well as current and default values
       Case TWCAP_RANGE
          EdtCurrent.Text = TwainPRO.CapValue
          EdtCurrent.Visible = True
          LblDefault.Caption = "Default = " & TwainPRO.CapDefault
          LblDefault.Visible = True
          LblMin.Caption = "MinValue = " & TwainPRO.CapMin
          LblMin.Visible = True
          LblMax.Caption = "MaxValue = " & TwainPRO.CapMax
          LblMax.Visible = True
         
'    TwainPRO.CloseSession
    
    End Select
    cmdUpdateCapability.Enabled = True
    
Exit Sub

error_handler:

'    MsgBox "CmbCaps_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
  End Sub
  

Private Sub subGetSettings_TwainColor()
    
    On Error GoTo error_handler
    
    Dim intListIndex As Integer
    Dim intHoldListIndex1 As Integer
    Dim strTwainColor As String
    
'    lblStatus.Caption = "LOADING SCANNER TWAIN SETTINGS - TwainColor"
    
    'Set the PixelType (Color Mode)
    TwainPRO.Capability = TWCAP_PIXELTYPE
    
    ' Check for support and stop if none
    If (TwainPRO.CapSupported = False) Then
       MsgBox "PIXELTYPE (Color) Not supported"
       Exit Sub
    End If
    
    ' What data type is the Cap?  We have to check the data type
    ' to know which properties are valid
    Select Case TwainPRO.CapType
       ' Type ENUM returns a list of legal values as well as current and
       ' default values.  A list of constants is returned and the CapDesc
       ' property can be used to find out what the constants mean
       Case TWCAP_ENUM
          'Hold work variables
          strTwainColor = cmbTwainColor.Text
          intHoldListIndex1 = 0
          
          ' Clear the list
          cmbTwainColor.Clear
          For intListIndex = 0 To TwainPRO.CapNumItems - 1
                cmbTwainColor.AddItem TwainPRO.CapDesc(TwainPRO.CapItem(intListIndex))
                If strTwainColor = TwainPRO.CapDesc(TwainPRO.CapItem(intListIndex)) Then
                   intHoldListIndex1 = intListIndex
                End If
          Next
          'Select the First item
          cmbTwainColor.ListIndex = intHoldListIndex1
    End Select
    
Exit Sub

error_handler:

'    MsgBox "subGetSettings_TwainColor ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub


Private Sub subGetSettings_SupportedPageSizes()
    
    On Error GoTo error_handler
    
    Dim intListIndex As Integer
    Dim intHoldListIndex1 As Integer
    Dim strTwainPageSize As String
    
'    lblStatus.Caption = "LOADING SCANNER TWAIN SETTINGS - TwainColor"
    
    'Set the PixelType (Color Mode)
    TwainPRO.Capability = TWCAP_SUPPORTEDSIZES
    
    ' Check for support and stop if none
    If (TwainPRO.CapSupported = False) Then
       lblStatus = "SUPPORTEDSIZES Not supported"
       Exit Sub
    End If
    
    ' What data type is the Cap?  We have to check the data type
    ' to know which properties are valid
    Select Case TwainPRO.CapType
       ' Type ENUM returns a list of legal values as well as current and
       ' default values.  A list of constants is returned and the CapDesc
       ' property can be used to find out what the constants mean
       Case TWCAP_ENUM
          'Hold work variables
          strTwainPageSize = cmbTwainPageSize.Text
          intHoldListIndex1 = 0
          
          ' Clear the list
          cmbTwainPageSize.Clear
          For intListIndex = 0 To TwainPRO.CapNumItems - 1
                cmbTwainPageSize.AddItem TwainPRO.CapDesc(TwainPRO.CapItem(intListIndex))
                If cmbTwainPageSize = TwainPRO.CapDesc(TwainPRO.CapItem(intListIndex)) Then
                   intHoldListIndex1 = intListIndex
                End If
          Next
          'Select the First item
          cmbTwainPageSize.ListIndex = intHoldListIndex1
    End Select
    
Exit Sub

error_handler:

'    MsgBox "subGetSettings_TwainColor ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub


Private Sub subGetSettings_TwainResolution()
    
    On Error GoTo error_handler
    
    Dim intListIndex As Integer
    Dim intHoldListIndex2 As Integer
    Dim strTwainResolution As String
    
'    lblStatus.Caption = "LOADING SCANNER TWAIN SETTINGS - TwainResolution"

    'Set the PixelType (Color Mode)
    TwainPRO.Capability = TWCAP_XRESOLUTION
    
    ' Check for support and stop if none
    If (TwainPRO.CapSupported = False) Then
       MsgBox "XRESOLUTION (Resolution) Not supported"
       Exit Sub
    End If
    
    ' What data type is the Cap?  We have to check the data type
    ' to know which properties are valid
    Select Case TwainPRO.CapType
       ' Type ENUM returns a list of legal values as well as current and
       ' default values.  A list of constants is returned and the CapDesc
       ' property can be used to find out what the constants mean
       Case TWCAP_ENUM
          'Hold work variables
          strTwainResolution = cmbTwainResolution.Text
          intHoldListIndex2 = 0
       
          ' Clear the list
          cmbTwainResolution.Clear
          For intListIndex = 0 To TwainPRO.CapNumItems - 1
             cmbTwainResolution.AddItem TwainPRO.CapItem(intListIndex)
                If strTwainResolution = TwainPRO.CapItem(intListIndex) Then
                   intHoldListIndex2 = intListIndex
                End If
          Next
          'Select the First item
          cmbTwainResolution.ListIndex = intHoldListIndex2
          
          
    End Select
    
    
Exit Sub

error_handler:

'    MsgBox "subGetSettings_TwainResolution ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub


Private Sub cmbTwainDocumentSize_DropDown()

    cmbTwainDocumentSize.Clear
    
    cmbTwainDocumentSize.AddItem "Letter   - 8.5 x 11 in"
    cmbTwainDocumentSize.AddItem "TrafTick - 4.25 x 8.5 in"
    cmbTwainDocumentSize.AddItem "Legal    - 8.5 x 14 in"
    cmbTwainDocumentSize.AddItem "Photo    - 5 x 3.5 in"
    cmbTwainDocumentSize.AddItem "Photo    - 3.5 x 5 in"
    cmbTwainDocumentSize.AddItem "Photo    - 6 x 4 in"
    cmbTwainDocumentSize.AddItem "Photo    - 4 x 6 in"
    cmbTwainDocumentSize.AddItem "A4       - 8.27 x 11.69 in"
    cmbTwainDocumentSize.AddItem "A5       - 5.82 x 8.27 in"
    cmbTwainDocumentSize.AddItem "* Custom *"

End Sub

Private Sub cmbTwainResolution_Click()
''''    subSetScannerResolution cmbTwainResolution.Text

End Sub



Private Sub cmbTwainSourceName_Click()
    
    On Error GoTo error_handler

'    lblStatus.Caption = ""
'
    TwainPRO.CloseSession
'
    TwainPRO.OpenDSM
    TwainPRO.SetDataSource (cmbTwainSourceName.ListIndex)
'''    MsgBox "TwainSource " & TwainPRO.DataSourceList(cmbTwainSourceName.ListIndex)
'
    
    TwainPRO.CloseSession
'    TwainPRO.OpenSession
'    subGetSettings_TwainColor
'    subGetSettings_TwainResolution
'    subGetSettings_SupportedPageSizes
'    subScannerSettingsApply
    
Exit Sub

error_handler:

'    MsgBox "cmbTwainSourceName_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next

End Sub


Private Sub cmdBatchDirectoryFind_Click()

    On Error GoTo error_handler
    
    txtBatchRootDirectory = funcGetDirectoryLocation("C:\Workarea\Jacob")
    
Exit Sub

error_handler:

'    MsgBox "cmdBatchDirectoryFind_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub

Private Sub cmdScanBegin_Click()

    On Error GoTo error_handler
    
    '*** CHECK FOR REQUIRED FIELDS - BEGIN
    
    If Trim(txtBatchName.Text = "") Then
        MsgBox "Please enter a BATCH NAME.", vbInformation, "Batch Name Required"
        SSTab1.Tab = 0
        txtBatchName.SetFocus
        Exit Sub
    End If
    
    If chkBatchBoxNumberRequired = vbChecked Then
        If Trim(txtBatchBoxNumber = "") Then
            MsgBox "Please Enter a BOX #!", vbInformation, "Batch BOX # Required"
            SSTab1.Tab = 0
            txtBatchBoxNumber.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(cmbBatchQueue.Text = "") Then
        MsgBox "Please select a QUEUE to scan into!", vbInformation, "Batch Queue Required"
        SSTab1.Tab = 0
        cmbBatchQueue.SetFocus
        Exit Sub
    End If
    
    '*** CHECK FOR REQUIRED FIELDS - END
    
    subToggleScanButtons
    
    lblStatus.Caption = ""
    
    
    'Initialize Image and Page Counts
    m_ImageCount = 0
    m_PageCount = 0
    m_ImageSkipCount = 0

    ' Reset cancel scan flag
    bolCancelPendingXfers = False

'''''    ' ***  CREATE BATCH HEADER RECORD
'''''    subCreateBatchRecord
    bolBatchCreated = False

''    cmdScannerSettingsApply_Click
    
    ' Keep Scanning Until User Stops the Scanner.
    '  The scanner will "pause" until user places more paper in ADF
    While Not bolCancelPendingXfers
        
        ' Close a session
        ' This is necessary only after calling OpenDSM or OpenSession methods.
        ' StartSession automatically calls CloseSession before returning.
'JR-07-19-2004 Disabled CloseSession
'        TwainPRO.CloseSession
        
        'Show Manufacturer User Interface (UI) if selected
        lblStatus.Caption = "Set Show Manuf. User Interface"
        TwainPRO.ShowUI = ItoB(chkScanShowUI.Value)
''        TwainPRO.ShowUI = False

'JR-07-19-2004 Disabled OpenSession
'        TwainPRO.OpenSession
        
        If ItoB(chkUseFlatbed) Then
            'Scanning with Flatbed
            result = MsgBox("Place paper in Flatbed and Click [OK]... Click [Cancel] when done.", vbOKCancel)
            If result = vbCancel Then
                txtTwainAcquireStatus = "Batch Ended by User."
'JR-07-19-2004 Disabled CloseSession
'                TwainPRO.CloseSession
                subToggleScanButtons
                GoTo EXIT_SUB
            End If
        Else
            
            '*** Prompt for Retry after each Session if chkAutoDetectPaperOut is Unchecked.
            '     This is for scanners like the CANON 3080C Driver that works erratically
            If chkAutoDetectPaperOut = vbUnchecked Then
                    result = MsgBox("Place paper in Feeder and Click Retry or Cancel.", vbRetryCancel)
                    If result = vbCancel Then bolCancelPendingXfers = True
                    If bolCancelPendingXfers Then
                        txtTwainAcquireStatus = "Batch Ended by User."
                        subToggleScanButtons
                        GoTo EXIT_SUB
                    End If
                    
            Else

                'Scanning with ADF
                ' Check to make sure the scanner supports the FEEDERLOADED (i.e.-Paper Detection) capability
                ' NOTE:  We MUST first OpenSession for this capability to be detected!
                TwainPRO.OpenSession
                TwainPRO.Capability = TWCAP_FEEDERLOADED
                If TwainPRO.CapSupported Then
                    While Not TwainPRO.CapValue = 1
                        txtTwainAcquireStatus = "Waiting for Paper in Feeder " & Format(Now, "hh:mm:ss")
                        DoEvents
                        If bolCancelPendingXfers Then
                            txtTwainAcquireStatus = "Batch Ended by User."
    'JR-07-19-2004 Disabled CloseSession
    '                        TwainPRO.CloseSession
            ''                subToggleScanButtons
                            GoTo EXIT_SUB
                        End If
                        TwainPRO.Capability = TWCAP_FEEDERLOADED
                    Wend
                End If
            End If
        End If
        
        
        txtTwainAcquireStatus = ""
        
        ' ***  BEGIN SCANNING
        ' Start the session (move to state 5 in the Twain spec)
        ' This will bring up the UI if ShowUI is enabled or start
        ' a scan if ShowUI is disabled.
    
'        lblStatus.Visible = False
'        subScannerSettingsApply
'        subToggleScanButtons
'        lblStatus.Visible = True
        
        '*** Added chkScanResetScanner for problem scanners.
        If chkScanResetScanner = vbChecked Then
            cmdScannerAdvancedCapabilitiesLoad_Click
        End If
        
        subSetPickingRectangle
        
        bolCancelPendingXfers = False
        
        TwainPRO.StartSession
        
        ' See if the Manufacturer UI was enabled and the user Clicked Exit without scanning.
        If m_ImageCount < 1 Then
            cmdScanStop_Click
            Exit Sub
        End If
        
    Wend
        
        
EXIT_SUB:
        
    'RESET FIELDS FOR NEXT SCAN
    If chkBatchAutoName.Value = vbChecked Then
        txtBatchName.Text = "Auto"
    Else
        txtBatchName.Text = ""
    End If
    txtBatchDesc.Text = ""
    txtBatchNotes.Text = ""
    txtBatchName.SetFocus
    
Exit Sub


error_handler:

    GetError
    MsgBox "cmdScanBegin_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    Resume Next

End Sub

Private Sub cmdStopScanning_Click()

    bolCancelPendingXfers = True

End Sub

Private Sub subScannerSettingsGetSettings()

    On Error GoTo error_handler

''     Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbNormal

     Dim txtOvewriteScannerSettings As Integer
     

     '*****************************************************************
     ' Establish the Imaging101 Batch List Connection
     txtActionBeforeError = "Establish the Imaging101 Batch List Connection"
     
     Set cmdImaging101Batch.ActiveConnection = connImaging101Batch
     
     ' Open BatchScannerSettings table.
     Dim rsBatchScannerSettings As ADODB.Recordset
     Set rsBatchScannerSettings = New ADODB.Recordset
     rsBatchScannerSettings.CursorType = adOpenDynamic
     rsBatchScannerSettings.LockType = adLockReadOnly
     rsBatchScannerSettings.Source = "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'"
     rsBatchScannerSettings.Open "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'", connImaging101Batch
     
     If rsBatchScannerSettings.EOF Then
         txtOvewriteScannerSettings = MsgBox("Sorry!  I couldn't find Settings for [ " & cmbBatchScanSettingsName & "]... It might have been deleted.", vbOK)
        Exit Sub
     End If
        
            
    '*** Using the [ & "" ] at the end of each line to Prevent VB Error 94 = Invalid use of Null
    '***      for Chkboxes or numbers must use [ & 0 ]
    cmbBatchScanSettingsName = rsBatchScannerSettings("BatchScannerSettingsName") & ""
    cmbBatchScanSettingsDesc = rsBatchScannerSettings("BatchScannerSettingsDesc") & ""
    txtBatchRootDirectory = rsBatchScannerSettings("BatchRootDirectory") & ""
    
    cmbTwainSourceName = rsBatchScannerSettings("ScanSourceName") & ""
    funcFindItemInListObject cmbTwainSourceName, cmbTwainSourceName.Text
    
    cmbTwainScanMethod = rsBatchScannerSettings("ScanScanMethod") & ""
    funcFindItemInListObject cmbTwainScanMethod, cmbTwainScanMethod.Text
    
    cmbTwainTransferMode = rsBatchScannerSettings("ScanTransferMode") & ""
    funcFindItemInListObject cmbTwainTransferMode, cmbTwainTransferMode.Text


    cmbTwainColor = rsBatchScannerSettings("ScanTwainColor") & ""
    funcFindItemInListObject cmbTwainColor, cmbTwainColor.Text
    
    cmbTwainResolution = rsBatchScannerSettings("ScanResolution") & ""
    funcFindItemInListObject cmbTwainResolution, cmbTwainResolution.Text
    
    
    sldTwainContrast = rsBatchScannerSettings("ScanContrast") & ""
    sldTwainContrast_Change
    sldTwainIntensity = rsBatchScannerSettings("ScanIntensity") & ""
    sldTwainIntensity_Change
    
    cmbTwainDocumentSize = rsBatchScannerSettings("ScanDocumentSize") & ""
    txtTwainImageTop = rsBatchScannerSettings("ScanImageTop") & ""
    txtTwainImageLeft = rsBatchScannerSettings("ScanImageLeft") & ""
    txtTwainImageRight = rsBatchScannerSettings("ScanImageRight") & ""
    txtTwainImageBottom = rsBatchScannerSettings("ScanImageBottom") & ""
    
    
    chkScanShowUI = rsBatchScannerSettings("ScanShowUI") & ""
    chkAutoDetectPaperOut = rsBatchScannerSettings("ScanAutoDetectPaperOut") & ""
    txtMinimumImageSize = rsBatchScannerSettings("ScanMinimumImageSize")
    chkScanPreviewOnly = rsBatchScannerSettings("ScanPreviewOnly") & ""
    
    
    txtBatchSettingsPrefix = rsBatchScannerSettings("ScanBatchPrefix") & ""
    txtBatchSettingsSuffix = rsBatchScannerSettings("ScanBatchSuffix") & ""
    chkBatchAutoName = rsBatchScannerSettings("ScanBatchAutoName") & ""
    chkBatchAutoUseBatchID = rsBatchScannerSettings("ScanBatchAutoUseBatchID") & ""
    chkBatchAutoUseDateTime = rsBatchScannerSettings("ScanBatchAutoUseDateTime") & ""
    chkBatchBoxNumberRequired = rsBatchScannerSettings("ScanBatchBoxNumberRequired") & ""
    
    chkScanDisplayImages = rsBatchScannerSettings("ScanDisplayImages") & ""
    txtScanImageSkipCount = rsBatchScannerSettings("ScanImageSkipCount") & ""
    chkUseFlatbed = rsBatchScannerSettings("ScanUseFlatBed") & ""
    
    
    txtCaption = rsBatchScannerSettings("ScanCaption") & ""
    chkCaptionClip = rsBatchScannerSettings("ScanCaptionClip") & "0"
    chkCaptionShadowText = rsBatchScannerSettings("ScanCaptionShadowText") & ""
    txtCaptionLeft = rsBatchScannerSettings("ScanCaptionLeft") & ""
    txtCaptionTop = rsBatchScannerSettings("ScanCaptionTop") & ""
    txtCaptionWidth = rsBatchScannerSettings("ScanCaptionWidth") & ""
    txtCaptionHeight = rsBatchScannerSettings("ScanCaptionHeight") & ""
    cmbCaptionHorizontalAlign.ListIndex = rsBatchScannerSettings("ScanCaptionHorizontalAlign") & ""
    cmbCaptionVerticalAlign.ListIndex = rsBatchScannerSettings("ScanCaptionVerticalAlign") & ""
    
    
    rsBatchScannerSettings.Close
    
    Set rsBatchScannerSettings = Nothing
    
    ' Set Batch Scanning Variables
    If chkBatchAutoName = vbChecked Then
        txtBatchName = "Auto"
    Else
        txtBatchName = ""
    End If
    
    ' See if the Batch Box # is required... if so color the Label Red
    If chkBatchBoxNumberRequired = vbChecked Then
        lblBatchBoxNumber.ForeColor = vbRed
    Else
        lblBatchBoxNumber.ForeColor = vbBlack
    End If
    
    
    txtBatchPrefix = txtBatchSettingsPrefix
    txtBatchSuffix = txtBatchSettingsSuffix
    
    
    '******************************
'    lblStatus.Caption = "Set DataSource"
'    '*** SELECT the Twain Data Source
'
'    Dim txtHoldTwainSourceName
'    Dim txtHoldListIndex As Integer
'    txtHoldTwainSourceName = cmbTwainSourceName.Text
'    txtHoldListIndex = 0
'
'    ' Walk down the list to get the right ListIndex to send the TwainPro.SetDatasource
'    For txtHoldListIndex = 0 To cmbTwainSourceName.ListCount - 1
'        If txtHoldTwainSourceName = cmbTwainSourceName.List(txtHoldListIndex) Then
'            Exit For
'        End If
'    Next
    
'***JR 7/19/2004 Moved to subScannerSettingsApply()
'    TwainPRO.CloseSession
'    TwainPRO.OpenDSM
'    TwainPRO.SetDataSource (txtHoldListIndex)
'    TwainPRO.OpenSession
'    cmbTwainSourceName.ListIndex = txtHoldListIndex

'******************************
    
'''    ' Fill the TwainColor combo box
'''    lblStatus.Caption = "LOADING SCANNER TWAIN SETTINGS - TwainColor"
'''    subGetSettings_TwainColor
'''
'''    ' Fill the TwainColor combo box
'''    lblStatus.Caption = "LOADING SCANNER TWAIN SETTINGS - TwainResolution"
'''    subGetSettings_TwainResolution
    
'    'Set Contrast & Brigtness Min/Max for sliders
'    TwainPRO.Capability = TWCAP_CONTRAST
'    If TwainPRO.CapSupported Then
''          EdtCurrent.Text = TwainPRO.CapValue
''          EdtCurrent.Visible = True
'          lblContrastDefault.Caption = "Default = " & TwainPRO.CapDefault
'          sldTwainContrast.Min = TwainPRO.CapMin
'          sldTwainContrast.Max = TwainPRO.CapMax
'          sldTwainContrast.SmallChange = 10
'          sldTwainContrast.LargeChange = 20
'    End If
'
'    TwainPRO.Capability = TWCAP_BRIGHTNESS
'    If TwainPRO.CapSupported Then
''          EdtCurrent.Text = TwainPRO.CapValue
''          EdtCurrent.Visible = True
'          lblBrightnessDefault.Caption = "Default = " & TwainPRO.CapDefault
'          sldTwainIntensity.Min = TwainPRO.CapMin
'          sldTwainIntensity.Max = TwainPRO.CapMax
'          sldTwainIntensity.SmallChange = 10
'          sldTwainIntensity.LargeChange = 20
'    End If
     
    
Exit Sub

error_handler:

'    MsgBox "subScannerSettingsGetSettings ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
    
End Sub
Private Sub subScannerSettingsLoadList()
    
    On Error GoTo error_handler

''     Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbNormal

     ' Establish the Imaging101 Batch List Connection
     txtActionBeforeError = "Establish the Imaging101 Batch List Connection"
     Dim txtOvewriteScannerSettings As Integer
     
     Set cmdImaging101Batch.ActiveConnection = connImaging101Batch
     
     ' Open BatchScannerSettings table.
     Dim rsBatchScannerSettings As ADODB.Recordset
     
     Set rsBatchScannerSettings = New ADODB.Recordset
     rsBatchScannerSettings.CursorType = adOpenDynamic
     rsBatchScannerSettings.LockType = adLockReadOnly
     rsBatchScannerSettings.Open "SELECT BatchScannerSettingsName FROM I101BatchScannerSettings WHERE ApplicationRECID=" & txtApplicationRECID, connImaging101Batch
     
    
     If rsBatchScannerSettings.EOF Then
         txtOvewriteScannerSettings = MsgBox("Sorry!  I couldn't find Settings for [ " & cmbBatchScanSettingsName & "] in Application [ " & _
                                             txtApplicationName & " ]... " & vbCrLf & "It might have been deleted." & vbCrLf & _
                                             "Please create a new one!", vbOK)
        Exit Sub
    Else
        'If NOT EOF then get the First record
        rsBatchScannerSettings.MoveFirst
     End If
        
    ' Reset the List
    cmbBatchScanSettingsName.Clear
    
    While Not rsBatchScannerSettings.EOF()
        cmbBatchScanSettingsName.AddItem rsBatchScannerSettings("BatchScannerSettingsName")
        rsBatchScannerSettings.MoveNext
    Wend
    
    rsBatchScannerSettings.Close
    Set rsBatchScannerSettings = Nothing
    
    
Exit Sub

error_handler:

'    MsgBox "subScannerSettingsLoadList ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
    
End Sub

Private Sub subScannerSettingsApply()

    '******************************************************
    '***  JUST LOAD THE SCANNER SETTINGS AND GET OUT
    '***
    
    '*** Added chkScanResetScanner for problem scanners.
    If chkScanResetScanner = vbChecked Then
        cmdScannerAdvancedCapabilitiesLoad_Click
    End If
    
Exit Sub
    
    '*****************************************************
    
    
    
    On Error GoTo error_handler
    
    cmdScanBegin.Visible = False
    
    Dim l As Single, t As Single, r As Single, B As Single
    Dim ScannerPhysicalWidth As Single
    Dim txtHoldResolution As String
    Dim txtHoldListIndex As Integer
    
'    ' JR - 7/19/2004 Moved from ScannerSettingsGetSettings()
'    txtHoldListIndex = cmbTwainSourceName.ListIndex
'    TwainPRO.CloseSession
'    TwainPRO.OpenDSM
'    TwainPRO.SetDataSource (txtHoldListIndex)
'    TwainPRO.CloseSession
'
    TwainPRO.CloseSession
    TwainPRO.OpenSession
    
'    cmbTwainSourceName.ListIndex = txtHoldListIndex
    
    Screen.MousePointer = vbHourglass
    
    lblStatus.Caption = "APPLYING SCANNER SETTINGS..."
    
    
    lblStatus.Caption = "Set Scan Method"
   
    Select Case cmbTwainScanMethod
        Case "Single Sided"
            TwainPRO.Capability = TWCAP_DUPLEXENABLED
            If TwainPRO.CapSupported Then
                        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
                        TwainPRO.CapValueOut = False
                        lblStatus.Caption = "Set DUPLEX = (" & cmbTwainScanMethod.ListIndex & ") "
                        TwainPRO.SetCapOut
                Else
                    lblStatus.Caption = "Set DUPLEX = NOT SUPPORTED"
            End If
        Case "Double Sided"
            TwainPRO.Capability = TWCAP_DUPLEXENABLED
            If TwainPRO.CapSupported Then
                        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
                        TwainPRO.CapValueOut = True
                        lblStatus.Caption = "Set DUPLEX = (" & cmbTwainScanMethod.ListIndex & ") "
                        TwainPRO.SetCapOut
                Else
                    lblStatus.Caption = "Set DUPLEX = NOT SUPPORTED"
            End If
    End Select
    
        
    
    TwainPRO.Capability = TWCAP_PIXELTYPE
    If TwainPRO.CapSupported Then
                TwainPRO.CapTypeOut = TWCAP_ONEVALUE
                TwainPRO.CapValueOut = CStr(cmbTwainColor.ListIndex)
                lblStatus.Caption = "Set Pixel Type = (" & cmbTwainColor.ListIndex & ") " & cmbTwainColor
                TwainPRO.SetCapOut
        Else
            lblStatus.Caption = "Set Pixel Type = NOT SUPPORTED"
    End If
    
'    TwainPRO.Capability = TWCAP_PIXELTYPE
'    If TwainPRO.CapSupported Then
'        Select Case cmbTwainColor
'            Case "Black & White (Bi-Tonal)"
'                TwainPRO.CapTypeOut = TWCAP_ONEVALUE
'                TwainPRO.CapValueOut = TWPT_BW
'                TwainPRO.SetCapOut
'
'            Case "HalfTone"
'                TwainPRO.CapTypeOut = TWCAP_ONEVALUE
'                TwainPRO.CapValueOut = TWPT_GRAY
'                TwainPRO.SetCapOut
'
'            Case "8-Bit Gray"
'                TwainPRO.CapTypeOut = TWCAP_ONEVALUE
'                TwainPRO.CapValueOut = TWPT_GRAY
'                TwainPRO.SetCapOut
'
'            Case "8-Bit Color"
'                TwainPRO.CapTypeOut = TWCAP_ONEVALUE
'                TwainPRO.CapValueOut = TWPT_PALETTE
'                TwainPRO.SetCapOut
'
'            Case "24-Bit Color"
'                TwainPRO.CapTypeOut = TWCAP_ONEVALUE
'                TwainPRO.CapValueOut = TWPT_RGB
'                TwainPRO.SetCapOut
'
'        End Select
'    End If

        'set the TWCAP_UNITS capability equal to TWUN_PIXELS (5) before setting X and Y resolutions
        TwainPRO.Capability = TWCAP_UNITS
        If TwainPRO.CapSupported Then
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = TWUN_PIXELS
            lblStatus.Caption = "Set TWCAP_UNITS = TWUN_INCHES"
            TwainPRO.SetCapOut
        Else
            lblStatus.Caption = "Set TWCAP_UNITS = NOT SUPPORTED"
        End If
 
    
        'Set the X & Y resolutions the same
        TwainPRO.Capability = TWCAP_XRESOLUTION
        If TwainPRO.CapSupported Then
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            
            If cmbTwainResolution = "" Then
                cmbTwainResolution = 200
                txtHoldResolution = cmbTwainResolution
            Else
                txtHoldResolution = cmbTwainResolution
            End If
            
            TwainPRO.CapValueOut = cmbTwainResolution
            lblStatus.Caption = "Set XRESOLUTION = " & TwainPRO.CapValueOut
            TwainPRO.SetCapOut
        Else
            lblStatus.Caption = "Set XRESOLUTION = NOT SUPPORTED"
        End If
        
    
        TwainPRO.Capability = TWCAP_YRESOLUTION
        If TwainPRO.CapSupported Then
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            
            If cmbTwainResolution = "" Then
                cmbTwainResolution = 200
                txtHoldResolution = cmbTwainResolution
            Else
                txtHoldResolution = cmbTwainResolution
            End If
            
            TwainPRO.CapValueOut = cmbTwainResolution
            lblStatus.Caption = "Set YRESOLUTION = " & TwainPRO.CapValueOut
            TwainPRO.SetCapOut
        Else
            lblStatus.Caption = "Set YRESOLUTION = NOT SUPPORTED"
        End If

        
        'Set Contrast
        TwainPRO.Capability = TWCAP_CONTRAST
        If TwainPRO.CapSupported Then
        
            sldTwainContrast.Enabled = True
            txtTwainContrast.Enabled = False
            
            lblStatus.Caption = "GET Contrast Defaults"
            lblContrastDefault.Caption = "Default = " & TwainPRO.CapDefault
            sldTwainContrast.Min = TwainPRO.CapMin
            sldTwainContrast.Max = TwainPRO.CapMax
            sldTwainContrast.TickFrequency = CInt(sldTwainContrast.Max) * 0.1
            sldTwainContrast.LargeChange = CInt(sldTwainContrast.Max) * 0.1
            sldTwainContrast.SmallChange = CInt(sldTwainContrast.Max) * 0.01

            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = sldTwainContrast
            lblStatus.Caption = "Set Contrast = " & TwainPRO.CapValueOut
            TwainPRO.SetCapOut
        Else
            lblStatus.Caption = "Set Contrast = NOT SUPPORTED"
            sldTwainContrast.Enabled = False
            txtTwainContrast.Enabled = False
        End If

        
        'Set Brigtness
        TwainPRO.Capability = TWCAP_BRIGHTNESS
        If TwainPRO.CapSupported Then
        
            sldTwainIntensity.Enabled = True
            txtTwainIntensity.Enabled = False
            
            lblStatus.Caption = "GET Brightness Defaults"
            lblBrightnessDefault.Caption = "Default = " & TwainPRO.CapDefault
            sldTwainIntensity.Min = TwainPRO.CapMin
            sldTwainIntensity.Max = TwainPRO.CapMax
            sldTwainIntensity.TickFrequency = CInt(sldTwainIntensity.Max) * 0.1
            sldTwainIntensity.LargeChange = CInt(sldTwainIntensity.Max) * 0.1
            sldTwainIntensity.SmallChange = CInt(sldTwainIntensity.Max) * 0.01

            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = sldTwainIntensity
            lblStatus.Caption = "Set Brightness = " & TwainPRO.CapValueOut
            TwainPRO.SetCapOut
        Else
            lblStatus.Caption = "Set Brightness = NOT SUPPORTED"
            sldTwainIntensity.Enabled = False
            txtTwainIntensity.Enabled = False
        End If
        
        

    Select Case cmbTwainTransferMode
        Case "Native"
            lblStatus.Caption = "Set Transfer Mode = TWSX_NATIVE"
            TwainPRO.TransferMode = TWSX_NATIVE
        Case "Memory Buffered"
            lblStatus.Caption = "Set Transfer Mode = TWSX_MEMORY"
            TwainPRO.TransferMode = TWSX_MEMORY
    End Select
    
        
    TwainPRO.Capability = TWCAP_FEEDERENABLED
    If TwainPRO.CapSupported Then
        If chkUseFlatbed = vbChecked Then
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = 0 ' Sets document acquisition source to Flatbed if available
            lblStatus.Caption = "Set Feeder Enabled = (0) FLATBED"
    
            TwainPRO.SetCapOut
        Else
            TwainPRO.CapTypeOut = TWCAP_ONEVALUE
            TwainPRO.CapValueOut = 1 ' Sets document acquisition source to feeder if available
            lblStatus.Caption = "Set Feeder Enabled = (1) ADF"
            TwainPRO.SetCapOut
        End If
    Else
        lblStatus.Caption = "Set Feeder Enabled = NOT SUPPORTED"
    End If
    

        
    ' Suppress the Scanning Progress window if capability is available
    'CAP_INDICATORS capability to be FALSE.
    TwainPRO.Capability = TWCAP_INDICATORS
    If TwainPRO.CapSupported Then
        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
        TwainPRO.CapValueOut = False
        lblStatus.Caption = "Set Indicators (Scanning Progress Window)= FALSE"
        TwainPRO.SetCapOut
    Else
        lblStatus.Caption = "Set Indicators (Scanning Progress Window)= NOT SUPPORTED"
    End If
    
        
    TwainPRO.Capability = TWCAP_USECAPADVANCED
    TwainPRO.CapAdvanced = ICAP_AUTOMATICDESKEW
    If (TwainPRO.CapSupported) Then
        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
        TwainPRO.CapValueOut = 1
        lblStatus.Caption = "Set AutoDeskew = 1"
        TwainPRO.SetCapOut
    Else
        lblStatus.Caption = "Set AutoDeskew = NOT SUPPORTED"
    End If
    
    
        
    TwainPRO.Capability = TWCAP_UNITS
    If TwainPRO.CapSupported Then
        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
        TwainPRO.CapValueOut = TWUN_INCHES 'set the TWCAP_UNITS capability equal to TWUN_PIXELS (5)
        lblStatus.Caption = "Set TWCAP_UNITS = TWUN_INCHES"
        TwainPRO.SetCapOut
    Else
        lblStatus.Caption = "Set TWCAP_UNITS = NOT SUPPORTED"
    End If
    
    
    ' SET PAGE SIZE
    TwainPRO.Capability = TWCAP_SUPPORTEDSIZES
    If TwainPRO.CapSupported Then
        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
        
        TwainPRO.CapValueOut = cmbTwainPageSize.ListIndex
        lblStatus.Caption = "Set SUPPORTEDSIZES = " & TwainPRO.CapValueOut
        TwainPRO.SetCapOut
    Else
        lblStatus.Caption = "Set SUPPORTEDSIZES = NOT SUPPORTED"
    End If
       
        
        
        
    'SET THE PICKING RECTANGLE AREA FOR THE PAGE
    If Not ItoB(chkUseFlatbed) Then
        ScannerPhysicalWidth = 0
        ' Using ADF - must Center scan area
        TwainPRO.Capability = TWCAP_PHYSICALWIDTH
        If TwainPRO.CapSupported Then
            ScannerPhysicalWidth = TwainPRO.CapValue
        End If
        
        If ScannerPhysicalWidth <= 0 Then
            ScannerPhysicalWidth = 8.5
        End If
        
        If txtTwainImageLeft.Text = "" Then
            txtTwainImageLeft.Text = 0
        End If
        
        If txtTwainImageRight.Text = "" Then
            txtTwainImageRight.Text = ScannerPhysicalWidth
        End If
        
        If txtTwainImageTop.Text = "" Then
            txtTwainImageTop.Text = 0
        End If
        
        If txtTwainImageBottom.Text = "" Then
            txtTwainImageBottom.Text = 11
        End If
        
        ' Calculate left & right based on center of Physical Width as read from scanner
'        l = (ScannerPhysicalWidth - (txtTwainImageRight.Text - txtTwainImageLeft.Text)) / 2
        l = txtTwainImageLeft.Text
        t = txtTwainImageTop.Text
        r = l + txtTwainImageRight.Text
        B = txtTwainImageBottom.Text
    
    Else
        ' Using Flatbed - begin from defined left
        l = txtTwainImageLeft.Text
        t = txtTwainImageTop.Text
        r = txtTwainImageRight.Text
        B = txtTwainImageBottom.Text
    
    End If
    lblStatus.Caption = "Set ImageLayout = Left:" & l & " Top:" & t & " Right:" & r & " Bottom:" & B
    
    On Error Resume Next
    TwainPRO.SetImageLayout l, t, r, B
    If Err.Number <> 0 Then
            MsgBox "The Scanner does NOT support the current Page Size settings.", vbInformation, "Page Size Error"
            TwainPRO.SetImageLayout 0, 0, 8.5, 11
    End If
    On Error GoTo error_handler
    
    ' Set Text Caption to Display on Image
    
    lblStatus.Caption = "Set Caption = " & txtCaption.Text
    TwainPRO.Caption = txtCaption.Text
    
    lblStatus.Caption = "Set ClipCaption = " & chkCaptionClip.Value
    TwainPRO.ClipCaption = ItoB(chkCaptionClip.Value)
    
    lblStatus.Caption = "Set ShadowText = " & chkCaptionShadowText.Value
    TwainPRO.ShadowText = ItoB(chkCaptionShadowText.Value)
    
    lblStatus.Caption = "Set CaptionLeft = " & txtCaptionLeft.Text
    TwainPRO.CaptionLeft = TtoL(txtCaptionLeft.Text)
    
    lblStatus.Caption = "Set CaptionTop = " & txtCaptionTop.Text
    TwainPRO.CaptionTop = TtoL(txtCaptionTop.Text)
    
    lblStatus.Caption = "Set CaptionWidth = " & txtCaptionWidth.Text
    TwainPRO.CaptionWidth = TtoL(txtCaptionWidth.Text)
    
    lblStatus.Caption = "Set CaptionHeight = " & txtCaptionHeight.Text
    TwainPRO.CaptionHeight = TtoL(txtCaptionHeight.Text)
    
    If cmbCaptionHorizontalAlign.ListIndex >= 0 Then
        lblStatus.Caption = "Set CaptionHorizontalAlign = " & cmbCaptionHorizontalAlign.ListIndex
        TwainPRO.HAlign = cmbCaptionHorizontalAlign.ListIndex
    End If
    
    If cmbCaptionVerticalAlign.ListIndex >= 0 Then
        lblStatus.Caption = "Set CaptionVerticalAlign = " & cmbCaptionVerticalAlign.ListIndex
        TwainPRO.VAlign = cmbCaptionVerticalAlign.ListIndex
    End If
    
'    lblStatus.Caption = "Set Units of Measure"
'
'    TwainPRO.Capability = TWCAP_UNITS
'    If TwainPRO.CapSupported Then
'        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
'        TwainPRO.CapValueOut = TWUN_PIXELS 'set the TWCAP_UNITS capability equal to TWUN_PIXELS (5)
'        TwainPRO.SetCapOut
'    End If
    
        
    TwainPRO.Capability = TWCAP_ROTATION
    If TwainPRO.CapSupported Then
        TwainPRO.CapTypeOut = TWCAP_ONEVALUE
        If cmbImageRotation = "" Then cmbImageRotation = "0"
        TwainPRO.CapValueOut = CInt(cmbImageRotation)
        lblStatus.Caption = "Set ROTATION = " & cmbImageRotation
        TwainPRO.SetCapOut
    Else
        lblStatus.Caption = "Set ROTATION = NOT SUPPORTED"
    End If

    lblStatus.Caption = "Scanner Settings Applied."
    
    Screen.MousePointer = vbNormal
    
    cmdScanBegin.Visible = True

Exit Sub

error_handler:

'    MsgBox "subScannerSettingsApply ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub


Private Sub cmdScannerCaptionSettingsSave_Click()
    
    cmdScannerSettingsSave_Click

End Sub



Private Sub cmdScannerSettingsApply_Click()

    'Apply Settings to Scanner
    subScannerSettingsApply

End Sub

Private Sub cmdScannerSettingsSave_Click()
    
    On Error GoTo error_handler
    
    ''     Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbNormal

     ' Establish the Imaging101 Batch List Connection
     txtActionBeforeError = "Establish the Imaging101 Batch List Connection"
     Dim txtOvewriteScannerSettings As Integer
     
     ' Open BatchScannerSettings table.
     Dim rsBatchScannerSettings As ADODB.Recordset
     Set rsBatchScannerSettings = New ADODB.Recordset
     rsBatchScannerSettings.CursorType = adOpenDynamic
     rsBatchScannerSettings.LockType = adLockOptimistic
     rsBatchScannerSettings.Source = "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'"
     rsBatchScannerSettings.Open "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'", connImaging101Batch
     
     'User Transaction Tracking to prevent partial imports!
     connImaging101Batch.BeginTrans

     
     If Not rsBatchScannerSettings.EOF Then
         txtOvewriteScannerSettings = MsgBox("Do you wish to Overwrite Batch Scanner Settings for [ " & cmbBatchScanSettingsName & "]?", vbYesNo)
         If txtOvewriteScannerSettings = vbNo Then
             Exit Sub
         End If
     Else
        rsBatchScannerSettings.AddNew
        rsBatchScannerSettings("BatchScannerSettingsRECID") = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchScannerSettingsRECID")
     End If
        
            
    rsBatchScannerSettings("ApplicationRECID") = txtApplicationRECID
    
    rsBatchScannerSettings("BatchScannerSettingsName") = cmbBatchScanSettingsName
    rsBatchScannerSettings("BatchScannerSettingsDesc") = cmbBatchScanSettingsDesc
    rsBatchScannerSettings("BatchRootDirectory") = txtBatchRootDirectory
    rsBatchScannerSettings("ScanSourceName") = cmbTwainSourceName
    rsBatchScannerSettings("ScanScanMethod") = cmbTwainScanMethod
    rsBatchScannerSettings("ScanTransferMode") = cmbTwainTransferMode

    rsBatchScannerSettings("ScanTwainColor") = cmbTwainColor
    rsBatchScannerSettings("ScanResolution") = cmbTwainResolution
    rsBatchScannerSettings("ScanContrast") = sldTwainContrast
    rsBatchScannerSettings("ScanIntensity") = sldTwainIntensity
    
    rsBatchScannerSettings("ScanDocumentSize") = cmbTwainDocumentSize
    rsBatchScannerSettings("ScanImageTop") = txtTwainImageTop
    rsBatchScannerSettings("ScanImageLeft") = txtTwainImageLeft
    rsBatchScannerSettings("ScanImageRight") = txtTwainImageRight
    rsBatchScannerSettings("ScanImageBottom") = txtTwainImageBottom
    
    rsBatchScannerSettings("ScanShowUI") = chkScanShowUI
    rsBatchScannerSettings("ScanAutoDetectPaperOut") = chkAutoDetectPaperOut
    rsBatchScannerSettings("ScanMinimumImageSize") = txtMinimumImageSize

    rsBatchScannerSettings("ScanPreviewOnly") = chkScanPreviewOnly
    
    rsBatchScannerSettings("ScanBatchPrefix") = txtBatchSettingsPrefix
    rsBatchScannerSettings("ScanBatchSuffix") = txtBatchSettingsSuffix
    rsBatchScannerSettings("ScanBatchAutoName") = chkBatchAutoName
    rsBatchScannerSettings("ScanBatchAutoUseBatchID") = chkBatchAutoUseBatchID
    rsBatchScannerSettings("ScanBatchAutoUseDateTime") = chkBatchAutoUseDateTime
    rsBatchScannerSettings("ScanBatchBoxNumberRequired") = chkBatchBoxNumberRequired
    
    rsBatchScannerSettings("ScanDisplayImages") = chkScanDisplayImages
    rsBatchScannerSettings("ScanImageSkipCount") = txtScanImageSkipCount
    rsBatchScannerSettings("ScanUseFlatBed") = chkUseFlatbed
    
    rsBatchScannerSettings("ScanCaption") = txtCaption
    rsBatchScannerSettings("ScanCaptionClip") = chkCaptionClip & ""
    rsBatchScannerSettings("ScanCaptionShadowText") = chkCaptionShadowText & ""
    rsBatchScannerSettings("ScanCaptionLeft") = txtCaptionLeft
    rsBatchScannerSettings("ScanCaptionTop") = txtCaptionTop
    rsBatchScannerSettings("ScanCaptionWidth") = txtCaptionWidth
    rsBatchScannerSettings("ScanCaptionHeight") = txtCaptionHeight
    rsBatchScannerSettings("ScanCaptionHorizontalAlign") = cmbCaptionHorizontalAlign.ListIndex
    rsBatchScannerSettings("ScanCaptionVerticalAlign") = cmbCaptionVerticalAlign.ListIndex
    
    rsBatchScannerSettings.Update
    
    connImaging101Batch.CommitTrans
    
    rsBatchScannerSettings.Close
    Set rsBatchScannerSettings = Nothing
    
Exit Sub

error_handler:

'    MsgBox "cmdScannerSettingsSave_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
End Sub


Private Sub cmdScanStop_Click()
    
    'Set Flag to stop scanning
    bolCancelPendingXfers = True
    subToggleScanButtons
    
End Sub



Private Sub cmdUpdateImageLayout_Click()

    On Error GoTo error_handler

    ' Update the Image Layout
    ' This can only be done after calling OpenSession and before calling
    ' StartSession (ie. in state 4 if you've read the Twain spec)
    
    Dim l As Single, t As Single, r As Single, B As Single
    Dim ScannerPhysicalWidth As Single
    
    
    lblStatus.Caption = ""
    
    ' Let's make sure the user entered valid data
    If False = IsNumeric(txtTwainImageLeft.Text) Or False = IsNumeric(txtTwainImageTop.Text) Or False = IsNumeric(txtTwainImageRight.Text) Or False = IsNumeric(txtTwainImageBottom.Text) Then
       MsgBox "You must enter a real number for each field."
    End If
    
    
    Exit Sub

Exit Sub

error_handler:

'    MsgBox "cmdUpdateImageLayout_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next

End Sub
Private Sub cmdUpdateCapability_Click()

    On Error GoTo error_handler
    
    ' Update the current Capability
    ' This can only be done after calling OpenSession and before calling
    ' StartSession (ie. in state 4 if you've read the Twain spec)
    
    lblStatus.Caption = ""
    
    ' Let's make sure the user entered valid data
    If False = IsNumeric(EdtCurrent.Text) Then
       MsgBox "You must enter a real number for the new current value."
    End If
    
    ' In this sample we only set the current value.  Other data types can
    ' be used to provide the Source with a new list of legal values to
    ' present the user when the UI is enabled, however this requires strict
    ' Twain compliance by the Source which is not found in many Twain drivers
'''    TwainPRO.CloseSession
    TwainPRO.CapTypeOut = TWCAP_ONEVALUE
    TwainPRO.CapValueOut = EdtCurrent.Text
    TwainPRO.SetCapOut
    
    ' Force an update of the Caps frame with the new values
    CmbCaps.ListIndex = CmbCaps.ListIndex
    Exit Sub

Exit Sub

error_handler:

'    MsgBox "cmdUpdateCapability_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
    
 End Sub
' Update the status bar with the current error code and description
Sub GetError()

    Dim RetPos As Long
    Dim TwnProErr As Long
    
    ' Error from an ActiveX object or a user-defined error.
    If (Err.Number > vbObjectError And Err.Number < vbObjectError + 65536) Then
    
        ' This definitively lets us know that that TwainPRO generated the error
        If InStr(1, Err.Source, "TwainPRO", vbTextCompare) <> 0 Then
            
            TwnProErr = Err.Number - vbObjectError
            
            ' ErrorCode TWERR_TWAIN indicates the Twain Source returned an error
            ' and we should check the ErrorCode and ErrorString properties
            ' for extended information
            If TwnProErr = TWERR_TWAIN Then
                lblStatus.Caption = "Data Source Error " & TwainPRO.ErrorCode & ": " & TwainPRO.ErrorString
            Else
                ' Remove carriage return from Description
                RetPos = InStr(1, Err.Description, Chr$(10))
                If RetPos > 0 Then
                    lblStatus.Caption = TwnProErr & ": " & Left$(Err.Description, RetPos - 1) & "  " & Mid$(Err.Description, RetPos + 1)
                Else
                    lblStatus.Caption = TwnProErr & ": " & Err.Description
                End If
            End If
            
        End If
        
    End If

    sErrMessage = lblStatus.Caption
    
End Sub
' Helper functions for setting Checkbox values to boolean properties
Function ItoB(Value As Integer) As Boolean
    If (Value <> 0) Then
       ItoB = True
    Else
       ItoB = False
    End If
End Function

Function BtoI(Value As Boolean) As Integer
    ' Convert Boolean to Integer
    If (Value = True) Then
       BtoI = 1
    Else
       BtoI = 0
    End If
End Function
Function TtoL(Value As String) As Long
    ' Convert Edit string value to Long
    If True = IsNumeric(Value) Then
       TtoL = CDbl(Value)
    Else
       TtoL = 0
    End If
End Function



Private Sub cmbTwainDocumentSize_Click()

    On Error GoTo error_handler
    
    If cmbTwainDocumentSize = "* Custom *" Then
        ' Enable boxes for entry
        txtTwainImageTop.Enabled = True
        txtTwainImageLeft.Enabled = True
        txtTwainImageBottom.Enabled = True
        txtTwainImageRight.Enabled = True
    Else
        ' Disable boxes, will set parameters automatically
        txtTwainImageTop.Enabled = False
        txtTwainImageLeft.Enabled = False
        txtTwainImageBottom.Enabled = False
        txtTwainImageRight.Enabled = False
    End If
        
            
    Select Case cmbTwainDocumentSize
        Case "Letter   - 8.5 x 11 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 8.5
            txtTwainImageBottom.Text = 11
        Case "TrafTick - 4.25 x 8.5 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 4.5
            txtTwainImageBottom.Text = 8.5
        Case "Legal    - 8.5 x 14 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 8.5
            txtTwainImageBottom.Text = 14
        Case "Photo    - 5 x 3.5 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 5
            txtTwainImageBottom.Text = 3.5
        Case "Photo    - 3.5 x 5 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 3.5
            txtTwainImageBottom.Text = 5
        Case "Photo    - 6 x 4 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 6
            txtTwainImageBottom.Text = 4
        Case "Photo    - 4 x 6 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 4
            txtTwainImageBottom.Text = 6
        Case "A4       - 8.27 x 11.69 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 8.27
            txtTwainImageBottom.Text = 11.69
        Case "A5       - 5.82 x 8.27 in"
            txtTwainImageLeft.Text = 0
            txtTwainImageTop.Text = 0
            txtTwainImageRight.Text = 5.82
            txtTwainImageBottom.Text = 8.27
    End Select
    
'Disabled ScannerSettingsApply to Prevent running it multiple times
'    subScannerSettingsApply
    
Exit Sub

error_handler:

'    MsgBox "cmbTwainDocumentSize_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    Resume Next
     
End Sub


Private Sub cmdScannerAdvancedCapabilitiesLoad_Click()

    On Error GoTo error_handler
    
    cmdScanBegin.Visible = False
    lblStatus = "Loading Scanner Advanced Capabilites for [" & cmbBatchScanSettingsName & "]"
    
    TwainPRO.CloseSession
    TwainPRO.OpenSession
    
    Dim newcapnum
    Dim newval
    Dim newtype
    Dim i
    Dim arraycount

        ' We assume we are loading the SAME SCANNER!
      Set fso = CreateObject("Scripting.FileSystemObject")
      
      strScannerSettingsFileName = App.path & "\ScanSet_" & cmbBatchScanSettingsName & ".cap"
      
      If fso.FileExists(strScannerSettingsFileName) Then
      
         'MsgBox "This demo is designed for 'Load Settings' from the same scanner that 'Save Settings' was issued from", vbInformation
         Set fsoScannerSettings = fso.OpenTextFile(strScannerSettingsFileName, 1, 0)
         
         
         Do While fsoScannerSettings.AtEndOfStream <> True
           newcapnum = fsoScannerSettings.ReadLine
           newval = fsoScannerSettings.ReadLine
           newtype = fsoScannerSettings.ReadLine
           
           TwainPRO.Capability = TWCAP_USECAPADVANCED
           TwainPRO.CapAdvanced = newcapnum
           
           
           If (TwainPRO.CapSupported) And (newtype <> "NotSupported") And (newval <> "NotSupported") Then
             TwainPRO.CapTypeOut = newtype
             
             If (newtype = 2) Or (newtype = 1) Or (newtype = 0) Then
               TwainPRO.CapTypeOut = 0
               TwainPRO.CapValueOut = newval
               TwainPRO.SetCapOut
             End If
             
             If newtype = 3 Then 'ARRAY
               arraycount = fsoScannerSettings.ReadLine
               For i = 1 To arraycount
                 TwainPRO.CapItemOut(i - 1) = fsoScannerSettings.ReadLine
               Next i
                 TwainPRO.SetCapOut
             End If
           End If
         Loop
         fsoScannerSettings.Close
        
        Else
          MsgBox "This scanner has not been configured yet..." & _
                    vbCrLf & "Please click the [Scanner Settings] button to configure!", vbInformation
          
        End If ' file does not exist
    
    
'        cmdScanBegin.Visible = True
        lblStatus = "Scanner Advanced Capabilites Load Complete for [" & cmbBatchScanSettingsName & "]"
        
        'Clear errors so they don't come back to bite us later.
        Err.Clear
        
Exit Sub

error_handler:
    
    GetError ' we can skip for now
    bolCancelPendingXfers = True
    
End Sub

Private Sub cmdScannerAdvancedCapabilitiesSave_Click()

    Dim INPUTLINE As String
    Dim OUTPUTLINE As String
    
    Dim HexStr As String
    Dim HexVal As Long
    
    Dim TWCONSTANT() As String
    Dim TWSupportedCaps() As Long
    Dim TWCapValue As String
    Dim TWCapType As Integer
    Dim i As Integer
    
    '********************************************************************
    '*** Create Output File for Settings
    '***
   
    lblStatus = "Saving Scanner Advanced Capabilites for [" & cmbBatchScanSettingsName & "]"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoScannerSettings = fso.CreateTextFile(App.path & "\ScanSet_" & cmbBatchScanSettingsName & ".cap", True)
    
    ' A demo of how to save all of the capabilities for
    ' a SPECIFIC scanner
    
    ' The order of saving and restoring is CRITICAL
    ' This order was created from http://twain.org/docs/CapOrderForWeb.PDF
    
    'Individual capabilities should be explored at http://twain.org/docs/Spec1_9_197.pdf
    
    'It is STRONGLY suggested only to save / restore only the capabilties you need.
    
    ' Independent of Read/Write Order
      
      ' This set can ONLY be read NOT restored
      'Save_A_Twain_Cap (CAP_ENABLEDSUIONLY)
      'Save_A_Twain_Cap (CAP_CUSTOMDSDATA)
      'Save_A_Twain_Cap (CAP_UICONTROLLABLE)
      'Save_A_Twain_Cap (CAP_SERIALNUMBER)
      'Save_A_Twain_Cap (CAP_BATTERYMINUTES)
      'Save_A_Twain_Cap (CAP_BATTERYPERCENTAGE)
      'Save_A_Twain_Cap (CAP_POWERSUPPLY)
      'Save_A_Twain_Cap (CAP_CAMERAPREVIEWUI)
      
      Save_A_Twain_Cap (CAP_INDICATORS)
      Save_A_Twain_Cap (ICAP_LAMPSTATE)
      Save_A_Twain_Cap (ICAP_BITORDER)
      Save_A_Twain_Cap (CAP_DEVICETIMEDATE)
      Save_A_Twain_Cap (CAP_DEVICEEVENT)
      
    ' SEMI Independent of Read/Write Order
      
      Save_A_Twain_Cap (CAP_ALARMS)
      Save_A_Twain_Cap (CAP_ALARMVOLUME)
      
      Save_A_Twain_Cap (CAP_AUTOMATICCAPTURE)
      Save_A_Twain_Cap (CAP_TIMEBEFOREFIRSTCAPTURE)
      Save_A_Twain_Cap (CAP_TIMEBETWEENCAPTURES)
      
      Save_A_Twain_Cap (ACAP_XFERMECH)
      Save_A_Twain_Cap (ACAP_AUDIOFILEFORMAT)
      
      
      'Order Dependent
     
      'Save_A_Twain_Cap (CAP_SUPPORTEDCAPS) 'Read only
      Save_A_Twain_Cap (CAP_LANGUAGE)
      'Save_A_Twain_Cap (CAP_DEVICEONLINE) 'Read only
      
      Save_A_Twain_Cap (ICAP_XFERMECH)
      Save_A_Twain_Cap (ICAP_TILES)
      'Save_A_Twain_Cap (ICAP_IMAGEFILEFORMAT) automatically handled
      Save_A_Twain_Cap (ICAP_COMPRESSION)
      
      Save_A_Twain_Cap (CAP_FEEDERENABLED)
      'Save_A_Twain_Cap (CAP_DUPLEX) Read Only
      Save_A_Twain_Cap (CAP_DUPLEXENABLED)
      Save_A_Twain_Cap (CAP_FEEDERORDER)
      Save_A_Twain_Cap (CAP_FEEDERALIGNMENT)
      Save_A_Twain_Cap (CAP_AUTOFEED)
      Save_A_Twain_Cap (CAP_CLEARPAGE)
      Save_A_Twain_Cap (CAP_FEEDPAGE)
      Save_A_Twain_Cap (CAP_REWINDPAGE)
      'Save_A_Twain_Cap (CAP_PAPERDETECTABLE) Read only
      'Save_A_Twain_Cap (CAP_FEEDERLOADED ) Read only
      
      Save_A_Twain_Cap (CAP_PRINTER)
      Save_A_Twain_Cap (CAP_PRINTERENABLED)
      Save_A_Twain_Cap (CAP_PRINTERMODE)
      Save_A_Twain_Cap (CAP_PRINTERSTRING)
      Save_A_Twain_Cap (CAP_PRINTERINDEX)
      Save_A_Twain_Cap (CAP_PRINTERSUFFIX)
      
      'Save_A_Twain_Cap (CAP_EXTENDEDCAPS) don't use, TwainPro handles
      
      Save_A_Twain_Cap (ICAP_UNITS)
      
      Save_A_Twain_Cap (ICAP_IMAGEDATASET)
      
      Save_A_Twain_Cap (ICAP_PIXELTYPE)
      Save_A_Twain_Cap (ICAP_BITDEPTH)
      Save_A_Twain_Cap (ICAP_XRESOLUTION)
      Save_A_Twain_Cap (ICAP_YRESOLUTION)
      Save_A_Twain_Cap (ICAP_PIXELFLAVOR)
      Save_A_Twain_Cap (ICAP_PLANARCHUNKY)
      Save_A_Twain_Cap (ICAP_BITDEPTHREDUCTION)
      Save_A_Twain_Cap (ICAP_CUSTHALFTONE)
      Save_A_Twain_Cap (ICAP_HALFTONES)
      Save_A_Twain_Cap (ICAP_THRESHOLD)
      Save_A_Twain_Cap (ICAP_COMPRESSION)
      Save_A_Twain_Cap (ICAP_BITORDERCODES)
      Save_A_Twain_Cap (ICAP_CCITTKFACTOR)
      Save_A_Twain_Cap (ICAP_PIXELFLAVOR)
      Save_A_Twain_Cap (ICAP_TIMEFILL)
      Save_A_Twain_Cap (ICAP_JPEGPIXELTYPE)
      
      Save_A_Twain_Cap (ICAP_XSCALING)
      Save_A_Twain_Cap (ICAP_YSCALING)
      Save_A_Twain_Cap (ICAP_ZOOMFACTOR)
      
      Save_A_Twain_Cap (ICAP_AUTOBRIGHT)
      Save_A_Twain_Cap (ICAP_BRIGHTNESS)
      
      Save_A_Twain_Cap (ICAP_CONTRAST)
      Save_A_Twain_Cap (ICAP_GAMMA)
      Save_A_Twain_Cap (ICAP_HIGHLIGHT)
      Save_A_Twain_Cap (ICAP_SHADOW)
      Save_A_Twain_Cap (ICAP_EXPOSURETIME)
      Save_A_Twain_Cap (ICAP_FILTER)
      Save_A_Twain_Cap (ICAP_IMAGEFILTER)
      Save_A_Twain_Cap (ICAP_NOISEFILTER)
      
      Save_A_Twain_Cap (ICAP_UNDEFINEDIMAGESIZE)
      Save_A_Twain_Cap (ICAP_AUTOMATICBORDERDETECTION)
      Save_A_Twain_Cap (ICAP_AUTOMATICDESKEW)
      Save_A_Twain_Cap (ICAP_AUTOMATICROTATE)
      Save_A_Twain_Cap (ICAP_OVERSCAN)
      
      Save_A_Twain_Cap (ICAP_SUPPORTEDSIZES)
      
      Save_A_Twain_Cap (ICAP_MAXFRAMES)
      Save_A_Twain_Cap (ICAP_FRAMES)
      
      Save_A_Twain_Cap (ICAP_ORIENTATION)
      Save_A_Twain_Cap (ICAP_FLIPROTATION)
      Save_A_Twain_Cap (ICAP_ROTATION)
      
      Save_A_Twain_Cap (CAP_AUTHOR)
      Save_A_Twain_Cap (CAP_CAPTION)
      Save_A_Twain_Cap (ICAP_LIGHTSOURCE)
      Save_A_Twain_Cap (ICAP_LIGHTPATH)
      Save_A_Twain_Cap (ICAP_FLASHUSED2)
      
      Save_A_Twain_Cap (CAP_XFERCOUNT)
      Save_A_Twain_Cap (CAP_AUTOSCAN)
      Save_A_Twain_Cap (CAP_MAXBATCHBUFFERS)
      Save_A_Twain_Cap (CAP_CLEARBUFFERS)
      
      Save_A_Twain_Cap (ICAP_EXTIMAGEINFO)
      Save_A_Twain_Cap (ICAP_PATCHCODEDETECTIONENABLED)
      Save_A_Twain_Cap (ICAP_PATCHCODESEARCHMODE)
      Save_A_Twain_Cap (ICAP_PATCHCODEMAXRETRIES)
      Save_A_Twain_Cap (ICAP_PATCHCODETIMEOUT)
      Save_A_Twain_Cap (ICAP_PATCHCODEMAXSEARCHPRIORITIES)
      Save_A_Twain_Cap (ICAP_PATCHCODESEARCHPRIORITIES)
      Save_A_Twain_Cap (ICAP_BARCODEDETECTIONENABLED)
      Save_A_Twain_Cap (ICAP_BARCODESEARCHMODE)
      Save_A_Twain_Cap (ICAP_BARCODEMAXRETRIES)
      Save_A_Twain_Cap (ICAP_BARCODETIMEOUT)
      Save_A_Twain_Cap (ICAP_BARCODEMAXSEARCHPRIORITIES)
      Save_A_Twain_Cap (ICAP_BARCODESEARCHPRIORITIES)
      
      Save_A_Twain_Cap (CAP_ENDORSER)
      Save_A_Twain_Cap (CAP_JOBCONTROL)
      
    fsoScannerSettings.Close
    
    TwainPRO.CloseSession
    cmdScanBegin.Visible = True
    
    lblStatus = "Scanner Advanced Capabilites Save Complete for [" & cmbBatchScanSettingsName & "]"
    
End Sub



Private Sub cmdScannerAdvancedCapabilitiesSet_Click()

    On Error GoTo error_handler
    
    '********************************************************************
    '*** Set Module-Level Flag to CLOSESESSION INSTEAD OF SCANNING
    '*** after clicking the SCAN button.
    
    cmdScannerAdvancedCapabilitiesSet.Enabled = False
    cmdScanBegin.Enabled = False
    
    'Flag to Prevent Scanning in the TwainPRO.PreScan sub
    bolCancelPendingXfers = True
    
    '*** Added chkScanResetScanner for problem scanners.
    If chkScanResetScanner = vbChecked Then
        cmdScannerAdvancedCapabilitiesLoad_Click
    End If
    
    Me.Visible = False
    
    TwainPRO.ShowUI = True
    TwainPRO.StartSession
    TwainPRO.CloseSession
    
    Me.Visible = True
    cmdScannerAdvancedCapabilitiesSet.Enabled = True
    cmdScanBegin.Enabled = True
    Me.SetFocus
    bolCancelPendingXfers = False
    
error_handler:
    GetError
    bolCancelPendingXfers = True
    Me.Visible = True
    
End Sub


Private Sub Form_Load()

    On Error GoTo error_handler

    'Make objects Invisible while we load the form
    SSTab1.Visible = False
    frameImageLayout.Visible = False
    cmdScanBegin.Visible = False
    cmdScanStop.Visible = False
    
    lblStatus.Caption = " "
    lblStatus.Caption = Now() & " Imaging101 Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    
    '*** SET UP SECURITY
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    
    If gsecRightsScannerSettings = vbChecked Then
        SSTab1.TabVisible(1) = True
    End If
    
    If gsecRightsAdminSystem = vbChecked Then
        SSTab1.TabVisible(1) = True
        SSTab1.TabVisible(2) = True
        SSTab1.TabVisible(3) = True
    End If
 

    Me.Show
    
    lblStatus.Caption = "OPENING DATABASES..."
    DoEvents
    
    Screen.MousePointer = vbHourglass
    
    ' Establish the Imaging101 DB Connections
    txtActionBeforeError = "Prepare Imaging101 DB Connections"
    Set connImaging101 = New ADODB.Connection
    Set cmdImaging101 = New ADODB.Command
    Set rsImaging101 = New ADODB.Recordset
    
'''''''    On Error Resume Next
'''''''    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
'''''''    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    connImaging101.ConnectionString = RegImaging101ConnectionString
    connImaging101.ConnectionTimeout = 60
    connImaging101.Mode = adModeReadWrite
    connImaging101.Open
    Set cmdImaging101.ActiveConnection = connImaging101
    
    ' Establish the Imaging101Batch DB Connections
    txtActionBeforeError = "Prepare Imaging101Batch DB Connections"
    Set connImaging101Batch = New ADODB.Connection
    Set cmdImaging101Batch = New ADODB.Command
    Set rsImaging101Batch = New ADODB.Recordset
    
'''''''    On Error Resume Next
'''''''    RegImaging101BatchListConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionType", RegFileName)
'''''''    RegImaging101BatchListConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionString." & RegImaging101BatchListConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
    connImaging101Batch.ConnectionTimeout = 60
    connImaging101Batch.Mode = adModeReadWrite
    connImaging101Batch.Open
    Set cmdImaging101Batch.ActiveConnection = connImaging101Batch
    
    
    
    '*************************************************************
    '*** LOAD UserID's   - BEGIN
    
    lblStatus.Caption = "Loading Route-To USERID's ..."
    
    txtActionBeforeError = "Connect to Imaging101 DB"
    
    Dim rs As ADODB.Recordset
    Dim ssql As String

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = connImaging101
    
    rs.Source = "Select * from I101Security"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    rs.Open
    
    txtActionBeforeError = "Populate UserID List"
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbBatchOwner.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchOwner.AddItem rs.Fields("UserName")
        rs.MoveNext
        DoEvents
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    
    '*** LOAD UserID's   - END
    '*************************************************************
    
    '***************************************
    '*** LOAD BATCH QUEUES LIST DROP-DOWN
        
    Dim Con As ADODB.Connection
    
    Set Con = New ADODB.Connection
    Con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = Con
    
    rs.Source = "Select * from I101BatchQueues WHERE BatchQueueAllowScanInto = '1' ORDER BY BatchQueue"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    Con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    
    'Add a Blank value to allow clearing the BatchOwner
'    cmbBatchQueue.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchQueue.AddItem rs.Fields!BatchQueue
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing

    '*** Set the Top/First item as the default
'    If cmbBatchQueue.ListCount > 0 Then
'        cmbBatchQueue.ListIndex = cmbBatchQueue.TopIndex
'    End If
    
    '***************************************
    '*** LOAD BATCH STATUS LIST DROP-DOWN
        
    Set Con = New ADODB.Connection
    Con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = Con
    
    rs.Source = "Select * from I101BatchStatus ORDER BY BatchStatus"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    Con.Errors.Clear
    
    rs.Open
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Add a Blank value to allow clearing the BatchOwner
'        cmbBatchStatus.AddItem ""
        For intIndex = 0 To rs.RecordCount - 1
            cmbBatchStatus.AddItem rs.Fields!BatchStatus
            rs.MoveNext
        Next
    End If
    
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing

    '****************************
    
    '***************************************
    '*** LOAD BATCH PRIORITY LIST DROP-DOWN
        
    Set Con = New ADODB.Connection
    Con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = Con
    
    rs.Source = "Select * from I101BatchPriority ORDER BY BatchPriority"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    
    Con.Errors.Clear
    
    rs.Open
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Add a Blank value to allow clearing the BatchOwner
'        cmbBatchPriority.AddItem ""
        For intIndex = 0 To rs.RecordCount - 1
            cmbBatchPriority.AddItem rs.Fields!BatchPriority
            rs.MoveNext
        Next
    End If
    
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing

    '****************************

    '*** Set the Default Values for the DropDown Lists
    funcFindItemInListObject Me.cmbBatchGroup, "REGULAR"
    funcFindItemInListObject Me.cmbBatchPriority, "3- LOW"
    funcFindItemInListObject Me.cmbBatchStatus, "Unassigned"


    TwainPRO.CloseOnCancel = True
    TwainPRO.Debug = True
    
''******************************************
''  ZAP SELECTSOURCE CODE FOR PRODUCTION
''*****************************************
'''    TwainPRO.SelectSource

    
    lblStatus.Caption = "LOADING SCANNER SOURCE NAMES..."
    DoEvents
    
    ' Fill the TwainSourceName combo box
    TwainPRO.CloseSession
    TwainPRO.OpenDSM
    For intLoop = 0 To TwainPRO.DataSourceCount - 1
        cmbTwainSourceName.AddItem TwainPRO.DataSourceList(intLoop)
    Next
    'JR-7/19/2004 Added CloseSession here
    TwainPRO.CloseSession
    
    
    ' SET UP VARIABLES
        txtApplicationRECID = frmImaging101BatchList.txtApplicationRECID
        txtApplicationName = frmImaging101BatchList.cmbApplicationList.Text
        txtBatchName = frmImaging101BatchList.txtBatchName
        txtBatchDesc = frmImaging101BatchList.txtBatchDesc
        txtBatchDirectory = frmImaging101BatchList.txtBatchDirectory
    
    
    '***** GET & SET SCANNER SETTINGS From DB
''    Dim test As String
''    test = funcSaveFieldToDB(RegImaging101ConnectionString, "101UserSettings", "UserSettingsUserID='1' and UserSettingsName=''", "BatchDirectory", "C:\TEMP")
''    test = funcGetFieldFromDB(RegImaging101ConnectionString, "Batches", "BatchName='TTC01'", "BatchDirectory")
    
    ' Load list of available ScannerSettings names
    subScannerSettingsLoadList
    
      'If NO Scanner Settings were loaded, get out of here NOW!
    If cmbBatchScanSettingsName.ListCount = 0 Then
        Exit Sub
    End If
    
    
    txtBatchScanUser = gsecUserID

''    '*** Changed code to Save/Load the Scanner Last Selected in the workstations LOCAL REGISTRY
''    cmbBatchScanSettingsName = funcGetSetUserSettings("GET", "BatchScanSettings_" & frmImaging101Winsock.txtComputerName, "")
''    cmbBatchScanSettingsName = GetSetting("Imaging101", "ScannerSettings", "ScanSettingsName", "")
    On Error Resume Next
    cmbBatchScanSettingsName = VBGetPrivateProfileString(RegAppname, "Imaging101ScanMain.cmbBatchScanSettingsName", RegFileName)
    On Error GoTo error_handler
    
  
    lblStatus.Caption = "GETTING SCANNER USER SETTINGS..."
    DoEvents
    subScannerSettingsGetSettings
    lblStatus.Caption = "SCANNER USER SETTINGS LOADED..."
    
    
    
' 8/30/2004 Jacob Disabled Apply because it is now handled by the AdvancedCapabilities procedures
'    lblStatus.Caption = "APPLYING SCANNER USER SETTINGS..."
'    DoEvents
'    subScannerSettingsApply
    
    
    
'''    cmdUpdateImageLayout_Click
'''    subUpdateProperties
'''    subSetScannerResolution cmbTwainResolution

    
''    lblStatus.Caption = "Scanner Settings Applied."
    DoEvents
    
    Screen.MousePointer = vbNormal
    
    'Make main objects visible again
    SSTab1.Visible = True
    frameImageLayout.Visible = True
    cmdScanBegin.Visible = True
    cmbBatchQueue.SetFocus
    
Exit Sub

error_handler:

'    MsgBox "Form_Load ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    GetError
    SSTab1.Visible = True
    frameImageLayout.Visible = True

    Resume Next
    
    
End Sub

Private Sub Form_Resize()
'    If Me.WindowState = vbNormal Then
'        If Screen.Width < Me.Width Then
'            Me.Width = Screen.Width
'        End If
'        If Screen.Height < Me.Height Then
'            Me.Height = Screen.Height
'        End If
'    End If
End Sub

' Make sure the Viewer form closes when the main for closes
Private Sub Form_Terminate()
    Unload Imaging101ScanViewer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdScanStop.Visible = True Then
'        cmdScanStop_Click
        MsgBox "Please STOP the Scanning before exiting!", vbInformation
        Cancel = True
        Exit Sub
    End If
    
'    '*** Changed code to Save/Load the Scanner Last Selected in the workstations LOCAL INI File
    result = WritePrivateProfileString(RegAppname, "Imaging101ScanMain.cmbBatchScanSettingsName", cmbBatchScanSettingsName, RegFileName)
    
    Unload Imaging101ScanViewer
    frmImaging101BatchList.Show
    frmImaging101BatchList.subListBatches
    
End Sub



Private Sub Misc_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub lblStatus_Change()
    
    Open "ScanStatus.log" For Append As #1
    Print #1, lblStatus.Caption
    Close #1
    DoEvents

End Sub

Private Sub sldTwainContrast_Scroll()
    
    'Move in increments of 10
    sldTwainContrast = CInt(sldTwainContrast / 10) * 10
    
End Sub



Private Sub sldTwainIntensity_Scroll()
    
    'Move in increments of 10
    sldTwainIntensity = CInt(sldTwainIntensity / 10) * 10
    
End Sub

Private Sub TwainPRO_PostScan(Cancel As Boolean)

    On Error GoTo error_handler
    

    ' Send the image to the Viewer form and save to file if requested

    Dim temp As Long
    Dim strFileName As String
    Dim strFullBatchDirectory As String
    Dim strBatchPageRECID As Double
    Dim intPageCount As Integer
    
    m_ImageCount = m_ImageCount + 1
    m_PageCount = m_PageCount + 1
    m_ImageSkipCount = m_ImageSkipCount + 1
    
    
    
    ' Save Image if requested
    If ItoB(chkScanPreviewOnly.Value) = False Then
    
        '*** 8/27/2004 Moved Inside the If so empty batches are not created on Preview
        If bolBatchCreated = False Then
            ' ***  CREATE BATCH HEADER RECORD
            subCreateBatchRecord
        End If
    
        
        strFullBatchDirectory = Trim(txtBatchRootDirectory) & "\" & Format(txtBatchRECID, "0000000000")

''        funcCreateDirectoryStructure strFullBatchDirectory
        
        intPageCount = m_PageCount
        strFileName = Format(txtBatchRECID, "000000000") & "-" & Format(intPageCount, "000000000")
        
        Select Case cmbTwainColor
            Case "Black & White"
                strFileName = strFileName & ".TIF"
                TwainPRO.SaveTIFCompression = TWTIF_CCITTFAX4
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "Black & White (Bi-Tonal)"
                strFileName = strFileName & ".TIF"
                TwainPRO.SaveTIFCompression = TWTIF_CCITTFAX4
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "HalfTone"
                TwainPRO.SaveJPGChromFactor = 36
                TwainPRO.SaveJPGLumFactor = 32
                TwainPRO.SaveJPGSubSampling = SS_411
                strFileName = strFileName & ".JPG"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "Grayscale"
                TwainPRO.SaveJPGChromFactor = 36
                TwainPRO.SaveJPGLumFactor = 32
                TwainPRO.SaveJPGSubSampling = SS_411
                strFileName = strFileName & ".JPG"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "8-Bit Gray"
                TwainPRO.SaveJPGChromFactor = 36
                TwainPRO.SaveJPGLumFactor = 32
                TwainPRO.SaveJPGSubSampling = SS_411
                strFileName = strFileName & ".JPG"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "8-Bit Color"
                TwainPRO.SaveJPGChromFactor = 36
                TwainPRO.SaveJPGLumFactor = 32
                TwainPRO.SaveJPGSubSampling = SS_411
                strFileName = strFileName & ".JPG"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "24-Bit Color"
                TwainPRO.SaveJPGChromFactor = 36
                TwainPRO.SaveJPGLumFactor = 32
                TwainPRO.SaveJPGSubSampling = SS_411
                strFileName = strFileName & ".JPG"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "Palette"
                TwainPRO.SaveJPGChromFactor = 36
                TwainPRO.SaveJPGLumFactor = 32
                TwainPRO.SaveJPGSubSampling = SS_411
                strFileName = strFileName & ".JPG"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
            Case "RGB"
                strFileName = strFileName & ".BMP"
                TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
        End Select
    
        '*** If in DUPLEX Mode, Check if this is the REAR Image'
        '***  to see if it is Blank
        'Does the scanner support Duplex?
        TwainPRO.Capability = TWCAP_DUPLEX
        If TwainPRO.CapSupported Then
            TwainPRO.Capability = TWCAP_DUPLEXENABLED
            If TwainPRO.CapSupported Then
                'Is Duplex enabled?
                If TwainPRO.CapValue = 1 Then
    '            If InStr(1, UCase(cmbTwainScanMethod), "DUPLEX") Then
                    'Check for a REMAINDER using the MOD math operator
                    'If there is NO remainder then the page is EVEN,
                    'meaning that it is the REAR image
                    If (m_ImageCount Mod 2) = 0 Then
                        If funcKillFileIfSmallerThan(strFullBatchDirectory & "\" & strFileName, txtMinimumImageSize) = True Then
                            'Decrement the PageCount because we just Killed the page
                            m_PageCount = m_PageCount - 1
                            'Get out of here NOW and Get the Next Page
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        DoEvents
    
        DoEvents
        
        'Create the Batch Page Record - Pass the filename as a parameter
        subCreateBatchPageRecord strFileName, intPageCount
    
    Else
        strFullBatchDirectory = "C:"
        strFileName = "IMGPREVIEWTEMP.TIF"
        TwainPRO.SaveTIFCompression = TWTIF_CCITTFAX4
        TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
    End If


    'Send Image to the Viewer
    If chkScanDisplayImages = vbChecked Then
        If m_PageCount = 1 Or m_ImageSkipCount = CInt(txtScanImageSkipCount) Then
            Imaging101ScanViewer.AddImage strFullBatchDirectory & "\" & strFileName
            m_ImageSkipCount = 0
        End If
    End If

    If ItoB(chkScanPreviewOnly.Value) = True Then
'        Kill strFullBatchDirectory & "\" & strFilename
    End If

Exit Sub

error_handler:

    GetError
    MsgBox "TwainPRO.PostScan ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    bolCancelPendingXfers = True
'''    Resume Next
    
End Sub




Private Sub sldTwainContrast_Change()

    txtTwainContrast = sldTwainContrast
    
    
End Sub

Private Sub sldTwainIntensity_Change()

    
    txtTwainIntensity = sldTwainIntensity
    
End Sub

Private Sub TwainPRO_PreScan(Cancel As Boolean)

    ' Stop Scanning If user clicked the STOP SCANNING button
    '  or if the bolCancelPendingXfers boolean variable was set to TRUE
    If bolCancelPendingXfers Then
        'Save the Advanced Capabilities and set Cancel=True to prevent scanning a page.
        cmdScannerAdvancedCapabilitiesSave_Click
        Cancel = True
        TwainPRO.CancelPendingXfers
        'Reset the CancelPendingXfers flag
        bolCancelPendingXfers = False
    End If
    

End Sub

Private Sub txtTwainContrast_Change()
    On Error Resume Next
    If txtTwainContrast >= sldTwainContrast.Min And txtTwainContrast <= sldTwainContrast.Max Then
        sldTwainContrast = txtTwainContrast
    Else
        txtTwainContrast = sldTwainContrast
    End If
    On Error GoTo 0
End Sub

Private Sub txtTwainIntensity_Change()
    On Error Resume Next
    If txtTwainIntensity >= sldTwainIntensity.Min And txtTwainIntensity <= sldTwainIntensity.Max Then
        sldTwainIntensity = txtTwainIntensity
    Else
        txtTwainIntensity = sldTwainIntensity
    End If
    On Error GoTo 0
End Sub
Private Sub subToggleScanButtons()
    If cmdScanBegin.Visible Then
        cmdScanBegin.Visible = False
        cmdScanStop.Visible = True
        cmdScannerAdvancedCapabilitiesSet.Enabled = False
    Else
        cmdScanBegin.Visible = True
        cmdScanStop.Visible = False
        cmdScannerAdvancedCapabilitiesSet.Enabled = True
    End If
End Sub

Private Sub subCreateBatchRecord()
    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
        On Error GoTo CREATE_BATCH_RECORD_ERROR
        
        txtActionBeforeError = "Open Batches Table"
        Set rsImaging101Batch = New ADODB.Recordset
        rsImaging101Batch.Open "I101Batches", connImaging101Batch, adOpenDynamic, adLockOptimistic
        
        'User Transaction Tracking to prevent partial imports!
        connImaging101Batch.BeginTrans
        
        txtActionBeforeError = "Add New Record"
        rsImaging101Batch.AddNew
        
        txtActionBeforeError = "Assign Variables to Fields"
        
        txtBatchRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchRECID")
        
       If chkBatchAutoName = vbChecked Then
            txtBatchName = ""
            txtBatchDirectory = ""
            If chkBatchAutoUseBatchID = vbChecked Then
                txtBatchName = txtBatchName & Format(txtBatchRECID, "0000000000")
            End If
            If chkBatchAutoUseDateTime = vbChecked Then
                If txtBatchName <> "" Then
                    txtBatchName = txtBatchName & "_"
                End If
                txtBatchName = txtBatchName & Format(Now(), "yyyy-mm-dd_hh-mm-ss")
            End If
        End If
        DoEvents
        
        txtBatchDirectory = Trim(txtBatchRootDirectory) & "\" & Format(txtBatchRECID, "0000000000")
        
        ' CREATE the Directory Structure for storing this Batch
        funcCreateDirectoryStructure txtBatchDirectory
        
        rsImaging101Batch("BatchRECID") = txtBatchRECID
        rsImaging101Batch("ApplicationRECID") = txtApplicationRECID
        rsImaging101Batch("BatchApplication") = ""
        rsImaging101Batch("BatchName") = txtBatchPrefix & txtBatchName & txtBatchSuffix
        rsImaging101Batch("BatchDesc") = txtBatchDesc
        rsImaging101Batch("BatchQueue") = cmbBatchQueue
        rsImaging101Batch("BatchOwner") = cmbBatchOwner
        rsImaging101Batch("BatchStatus") = cmbBatchStatus
        rsImaging101Batch("BatchPriority") = cmbBatchPriority
        rsImaging101Batch("BatchGroup") = cmbBatchGroup
        rsImaging101Batch("BatchScanDate") = Now()
        rsImaging101Batch("BatchDirectory") = txtBatchDirectory
        rsImaging101Batch("BatchNotes") = txtBatchNotes
        rsImaging101Batch("BatchPagesCommitted") = 0
        rsImaging101Batch("BatchPagesNotCommitted") = 0
        rsImaging101Batch("BatchPagesTotal") = 0
        rsImaging101Batch("BatchScanUser") = txtBatchScanUser
        rsImaging101Batch("BatchBoxNumber") = txtBatchBoxNumber
    
        txtActionBeforeError = "Update Values"
        rsImaging101Batch.Update
    
    
''    '*** ADD BATCH PAGES  ***
''    For intLoop = 0 To List1.ListCount - 1
''        '* Using List1 to get the file Extensions in correct Sorted order
''        txtBatchPageFileName = txtecID & "." & Format(List1.List(intLoop), "###")
''        txtBatchPageRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchPageRECID")
''        cmd.CommandText = "INSERT INTO " & "B_" & lblApplicationName & " (BatchPageRECID, BatchRECID, BatchPageFileName, BatchPageOrder) VALUES ('" & txtBatchPageRECID & "', '" & txtBatchRECID & "', '" & txtBatchPageFileName & "', '" & intLoop + 1 & "')"
''
''        txtActionBeforeError = "INSERT Page" & txtBatchPageFileName & " into Imaging101"
''
''        cmd.Execute , , adCmdText
''        txtPagesImported = txtPagesImported + 1
''        DoEvents
''    Next


    ' Commit the successful transactions
    txtActionBeforeError = "COMMIT BATCH TRANSACTION"
    connImaging101Batch.CommitTrans
    
    rsImaging101Batch.Close
    Set rsImaging101Batch = Nothing
'    connImaging101Batch.Close
    
    '*** CREATE BATCH AUDIT RECORD
    funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Scan Batch"
    
    bolBatchCreated = True
    
    Screen.MousePointer = vbDefault

Exit Sub
    
CREATE_BATCH_RECORD_ERROR:
        funcQuickMessage "SHOW", "CREATE_BATCH_RECORD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch NOT Created]"
        
        On Error Resume Next
        connImaging101Batch.RollbackTrans
        
        rsImaging101Batch.Close
        Set rsImaging101Batch = Nothing
'        Set connImaging101Batch = Nothing
        
        Screen.MousePointer = vbDefault

        bolCancelPendingXfers = True
    
End Sub

Private Sub subCreateBatchPageRecord(strFileName As String, intPageCount As Integer)
    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
        On Error GoTo CREATE_BATCH_PAGE_RECORD_ERROR
        
        
        'Position the cursor on the rowset
        txtActionBeforeError = "Open Batches Table"
        Set rsImaging101Batch = New ADODB.Recordset
        '*** Prepare Result Set
        With rsImaging101Batch
            .ActiveConnection = connImaging101Batch
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
        End With
        
        rsImaging101Batch.Source = "SELECT * " & _
                    " FROM I101Batches " & _
                    " WHERE BatchRECID = " & txtBatchRECID
        
        connImaging101Batch.Errors.Clear
        rsImaging101Batch.Open
        rsImaging101Batch.MoveFirst
        
        txtActionBeforeError = "Open Batch Page Table"
        Set rsImaging101BatchPage = New ADODB.Recordset
        rsImaging101BatchPage.Open txtApplicationName & "_BatchPage", connImaging101Batch, adOpenDynamic, adLockOptimistic
        
        'User Transaction Tracking to prevent partial imports!
        connImaging101Batch.BeginTrans
        
        txtActionBeforeError = "Add New Record"
        rsImaging101BatchPage.AddNew
        
        txtActionBeforeError = "Assign Variables to Batch Page Fields"
        rsImaging101BatchPage("BatchPageRECID") = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchPageRECID")
        rsImaging101BatchPage("BatchRECID") = txtBatchRECID
        rsImaging101BatchPage("BatchPageFileName") = strFileName
        rsImaging101BatchPage("BatchPageOrder") = intPageCount
        
        rsImaging101BatchPage("BatchPageIndexed") = ""
        rsImaging101BatchPage("BatchPageIsSeparator") = ""
        rsImaging101BatchPage("BatchPageNote") = ""
        rsImaging101BatchPage("BatchDocDesc") = ""
        rsImaging101BatchPage("BatchPageStatus") = ""
'        rsImaging101BatchPage("BatchPageCommitDate") = ""
        rsImaging101BatchPage("BatchPageCommitUser") = ""
        
        txtActionBeforeError = "Update Batch Page Values"
        rsImaging101BatchPage.Update
        
        ' Set BATCHES field values
        txtActionBeforeError = "Assign Variables to Batch Fields"
        
        Dim intBatchPagesTotal As Integer
        intBatchPagesTotal = rsImaging101Batch("BatchPagesTotal")
        intBatchPagesTotal = intBatchPagesTotal + 1
        
        rsImaging101Batch("BatchPagesTotal") = intBatchPagesTotal
        rsImaging101Batch("BatchPagesQCAppended") = 0
        rsImaging101Batch("BatchPagesQCInserted") = 0
        rsImaging101Batch("BatchPagesQCDeleted") = 0
        rsImaging101Batch("BatchPagesIndexed") = 0
        rsImaging101Batch("BatchPagesCommitted") = 0
        rsImaging101Batch("BatchPagesNotCommitted") = rsImaging101Batch("BatchPagesTotal")
        
        txtActionBeforeError = "Update Batch Values"
        rsImaging101Batch.Update
    

    ' Commit the successful transactions
    txtActionBeforeError = "COMMIT BATCH PAGE TRANSACTION"
    connImaging101Batch.CommitTrans
    
    rsImaging101Batch.Close
    Set rsImaging101Batch = Nothing
    rsImaging101BatchPage.Close
    Set rsImaging101BatchPage = Nothing
'    Set connImaging101Batch = Nothing

    Screen.MousePointer = vbDefault

Exit Sub
    
CREATE_BATCH_PAGE_RECORD_ERROR:
        funcQuickMessage "SHOW", "CREATE_BATCH_PAGE_RECORD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch Page NOT Created]"
        
        On Error Resume Next
        connImaging101Batch.RollbackTrans
        
        rsImaging101Batch.Close
        Set rsImaging101Batch = Nothing
        rsImaging101BatchPage.Close
        Set rsImaging101BatchPage = Nothing
'        Set connImaging101Batch = Nothing
        
        Screen.MousePointer = vbDefault

        bolCancelPendingXfers = True

    
End Sub


Function funcKillFileIfSmallerThan(FullFilePath As String, FileMinimumSize As Long) As Boolean
    
    Dim lngFileSize As Long
    
    'FileLen function returns the file size in BYTES.
    lngFileSize = FileLen(FullFilePath)
    
    funcKillFileIfSmallerThan = False
    
    If lngFileSize < FileMinimumSize Then
'        MsgBox FullFilePath & " Size: " & lngFileSize & " Min: " & FileMinimumSize
        Kill FullFilePath
    
        'Set to true to notify the calling routing that the image WAS deleted
        funcKillFileIfSmallerThan = True

    End If
    
End Function


Function ConvertTwainTypetoName(innum As Integer) As String
  Select Case innum
  Case 0
    ConvertTwainTypetoName = "ONEVALUE "
  Case 1
    ConvertTwainTypetoName = "ENUM "
  Case 2
    ConvertTwainTypetoName = "RANGE "
  Case 3
    ConvertTwainTypetoName = "ARRAY "
  Case Else
    ConvertTwainTypetoName = "UNKNOWN "
  End Select
End Function


Private Sub Save_A_Twain_Cap(CapNum As Integer)
    Dim tempval
    Dim temptype
    Dim i
    Dim arraycount
  
    TwainPRO.Capability = TWCAP_USECAPADVANCED
    TwainPRO.CapAdvanced = CapNum
    If (TwainPRO.CapSupported) Then
      tempval = TwainPRO.CapValue
      temptype = TwainPRO.CapType
    Else
      tempval = "NotSupported"
      temptype = "NotSupported"
    End If
    
      
    If (temptype = 3) Then 'ARRAY
      fsoScannerSettings.WriteLine CapNum
      fsoScannerSettings.WriteLine "ARRAY"
      fsoScannerSettings.WriteLine temptype
      arraycount = TwainPRO.CapNumItems
      fsoScannerSettings.WriteLine arraycount
      For i = 1 To arraycount
        fsoScannerSettings.WriteLine TwainPRO.CapItem(i - 1)
      Next i
        
    Else 'NOT ARRAY
      fsoScannerSettings.WriteLine CapNum
      fsoScannerSettings.WriteLine tempval
      fsoScannerSettings.WriteLine temptype
    End If 'ARRAY

End Sub


Function ConvertTwainNumtoName(innum) As String
    ' This is handy for displaying errors
  Select Case innum
  Case &H1
    ConvertTwainNumtoName = "CAP_XFERCOUNT "
  Case &H100
    ConvertTwainNumtoName = "ICAP_COMPRESSION "
  Case &H101
    ConvertTwainNumtoName = "ICAP_UNITS "
  Case &H103
    ConvertTwainNumtoName = "ICAP_XFERMECH "
  Case &H1000
    ConvertTwainNumtoName = "CAP_AUTHOR "
  Case &H1001
    ConvertTwainNumtoName = "CAP_CAPTION "
  Case &H1002
    ConvertTwainNumtoName = "CAP_FEEDERENABLED "
  Case &H1003
    ConvertTwainNumtoName = "CAP_FEEDERLOADED "
  Case &H1004
    ConvertTwainNumtoName = "CAP_TIMEDATE "
  Case &H1005
    ConvertTwainNumtoName = "CAP_SUPPORTEDCAPS "
  Case &H1006
    ConvertTwainNumtoName = "CAP_EXTENDEDCAPS "
  Case &H1007
    ConvertTwainNumtoName = "CAP_AUTOFEED "
  Case &H1008
    ConvertTwainNumtoName = "CAP_CLEARPAGE "
  Case &H1009
    ConvertTwainNumtoName = "CAP_FEEDPAGE "
  Case &H100A
    ConvertTwainNumtoName = "CAP_REWINDPAGE "
  Case &H100B
    ConvertTwainNumtoName = "CAP_INDICATORS "
  Case &H100C
    ConvertTwainNumtoName = "CAP_SUPPORTEDCAPSEXT "
  Case &H100D
    ConvertTwainNumtoName = "CAP_PAPERDETECTABLE "
  Case &H100E
    ConvertTwainNumtoName = "CAP_UICONTROLLABLE "
  Case &H100F
    ConvertTwainNumtoName = "CAP_DEVICEONLINE "
  Case &H1010 ' 4112
    ConvertTwainNumtoName = "CAP_AUTOSCAN "
  Case &H1011
    ConvertTwainNumtoName = "CAP_THUMBNAILSENABLED "
  Case &H1012
    ConvertTwainNumtoName = "CAP_DUPLEX "
  Case &H1013
    ConvertTwainNumtoName = "CAP_DUPLEXENABLED "
  Case &H1014
    ConvertTwainNumtoName = "CAP_ENABLEDSUIONLY "
  Case &H1015
    ConvertTwainNumtoName = "CAP_CUSTOMDSDATA "
  Case &H1016
    ConvertTwainNumtoName = "CAP_ENDORSER "
  Case &H1017
    ConvertTwainNumtoName = "CAP_JOBCONTROL "
  Case &H1018
    ConvertTwainNumtoName = "CAP_ALARMS "
  Case &H1019
    ConvertTwainNumtoName = "CAP_ALARMVOLUME "
  Case &H101A
    ConvertTwainNumtoName = "CAP_AUTOMATICCAPTURE "
  Case &H101B
    ConvertTwainNumtoName = "CAP_TIMEBEFOREFIRSTCAPTURE "
  Case &H101C
    ConvertTwainNumtoName = "CCAP_TIMEBETWEENCAPTURES "
  Case &H101D
    ConvertTwainNumtoName = "CAP_CLEARBUFFERS "
  Case &H102F
    ConvertTwainNumtoName = "CAP_PAPERBINDING "
  Case &H1030
    ConvertTwainNumtoName = "CAP_REACQUIREALLOWED "
  Case &H1031
    ConvertTwainNumtoName = "CAP_PASSTHRU "
  Case &H1032
    ConvertTwainNumtoName = "CAP_BATTERYMINUTES "
  Case &H1033
    ConvertTwainNumtoName = "CAP_BATTERYPERCENTAGE "
  Case &H1034
    ConvertTwainNumtoName = "CAP_POWERDOWNTIME "
  Case &H1100
    ConvertTwainNumtoName = "ICAP_AUTOBRIGHT "
  Case &H1101
    ConvertTwainNumtoName = "ICAP_BRIGHTNESS "
  Case &H1103
    ConvertTwainNumtoName = "ICAP_CONTRAST "
  Case &H1104 '4356
    ConvertTwainNumtoName = "ICAP_CUSTHALFTONE "
  Case &H1105
    ConvertTwainNumtoName = "ICAP_EXPOSURETIME "
  Case &H1106
    ConvertTwainNumtoName = "ICAP_FILTER "
  Case &H1107
    ConvertTwainNumtoName = "ICAP_FLASHUSED "
  Case &H1108
    ConvertTwainNumtoName = "ICAP_GAMMA "
  Case &H1109 '4361
    ConvertTwainNumtoName = "ICAP_HALFTONES "
  Case &H110A
    ConvertTwainNumtoName = "ICAP_HIGHLIGHT "
  Case &H110C
    ConvertTwainNumtoName = "ICAP_IMAGEFILEFORMAT "
  Case &H110D
    ConvertTwainNumtoName = "ICAP_LAMPSTATE "
  Case &H110E
    ConvertTwainNumtoName = "ICAP_LIGHTSOURCE "
  Case &H1110
    ConvertTwainNumtoName = "ICAP_ORIENTATION "
  Case &H1111
    ConvertTwainNumtoName = "ICAP_PHYSICALWIDTH "
  Case &H1112
    ConvertTwainNumtoName = "ICAP_PHYSICALHEIGHT "
  Case &H1113
    ConvertTwainNumtoName = "ICAP_SHADOW "
  Case &H1114
    ConvertTwainNumtoName = "ICAP_FRAMES "
  Case &H1116
    ConvertTwainNumtoName = "ICAP_XNATIVERESOLUTION "
  Case &H1117
    ConvertTwainNumtoName = "ICAP_YNATIVERESOLUTION "
  Case &H1118
    ConvertTwainNumtoName = "ICAP_XRESOLUTION "
  Case &H1119
    ConvertTwainNumtoName = "ICAP_YRESOLUTION "
  Case &H111A
    ConvertTwainNumtoName = "ICAP_MAXFRAMES "
  Case &H111B
    ConvertTwainNumtoName = "ICAP_TILES "
  Case &H111C
    ConvertTwainNumtoName = "ICAP_BITORDER "
  Case &H111D
    ConvertTwainNumtoName = "ICAP_CCITTKFACTOR "
  Case &H111E
    ConvertTwainNumtoName = "ICAP_LIGHTPATH "
  Case &H111F
    ConvertTwainNumtoName = "ICAP_PIXELFLAVOR "
  Case &H1120
    ConvertTwainNumtoName = "ICAP_PLANARCHUNKY "
  Case &H1121
    ConvertTwainNumtoName = "ICAP_ROTATION "
  Case &H1122
    ConvertTwainNumtoName = "ICAP_SUPPORTEDSIZES "
  Case &H1123
    ConvertTwainNumtoName = "ICAP_THRESHOLD "
  Case &H1124
    ConvertTwainNumtoName = "ICAP_XSCALING "
  Case &H1125
    ConvertTwainNumtoName = "ICAP_YSCALING "
  Case &H1126
    ConvertTwainNumtoName = "ICAP_BITORDERCODES "
  Case &H1127
    ConvertTwainNumtoName = "ICAP_PIXELFLAVORCODES "
  Case &H1128
    ConvertTwainNumtoName = "ICAP_JPEGPIXELTYPE "
  Case &H112A
    ConvertTwainNumtoName = "ICAP_TIMEFILL "
  Case &H112B
    ConvertTwainNumtoName = "ICAP_BITDEPTH "
  Case &H112C
    ConvertTwainNumtoName = "ICAP_BITDEPTHREDUCTION "
  Case &H112D
    ConvertTwainNumtoName = "ICAP_UNDEFINEDIMAGESIZE "
  Case &H112E
    ConvertTwainNumtoName = "ICAP_IMAGEDATASET "
  Case &H112F
    ConvertTwainNumtoName = "ICAP_EXTIMAGEINFO "
  Case &H1130
    ConvertTwainNumtoName = "ICAP_MINIMUMHEIGHT "
  Case &H1131
    ConvertTwainNumtoName = "ICAP_MINIMUMWIDTH "
  Case &H1134
    ConvertTwainNumtoName = "ICAP_AUTODISCARDBLANKPAGES "
  Case &H1136
    ConvertTwainNumtoName = "ICAP_FLIPROTATION "
  Case &H1137
    ConvertTwainNumtoName = "ICAP_BARCODEDETECTIONENABLED "
  Case &H1138
    ConvertTwainNumtoName = "ICAP_SUPPORTEDBARCODETYPES "
  Case &H1139
    ConvertTwainNumtoName = "ICAP_BARCODEMAXSEARCHPRIORITIES "
  Case &H113A
    ConvertTwainNumtoName = "ICAP_BARCODESEARCHPRIORITIES "
  Case &H113B
    ConvertTwainNumtoName = "ICAP_BARCODESEARCHMODE "
  Case &H113C
    ConvertTwainNumtoName = "ICAP_BARCODEMAXRETRIES "
  Case &H113D
    ConvertTwainNumtoName = "ICAP_BARCODETIMEOUT "
  Case &H113E
    ConvertTwainNumtoName = "ICAP_ZOOMFACTOR "
  Case &H113F
    ConvertTwainNumtoName = "ICAP_PATCHCODEDETECTIONENABLED "
  Case &H1140
    ConvertTwainNumtoName = "ICAP_SUPPORTEDPATCHCODETYPES "
  Case &H1141
    ConvertTwainNumtoName = "ICAP_PATCHCODEMAXSEARCHPRIORITIES "
  Case &H1142
    ConvertTwainNumtoName = "ICAP_PATCHCODESEARCHPRIORITIES "
  Case &H1143
    ConvertTwainNumtoName = "ICAP_PATCHCODESEARCHMODE "
  Case &H1144
    ConvertTwainNumtoName = "ICAP_PATCHCODEMAXRETRIES "
  Case &H1145
    ConvertTwainNumtoName = "ICAP_PATCHCODETIMEOUT "
  Case &H1146
    ConvertTwainNumtoName = "ICAP_FLASHUSED2 "
  Case &H1147
    ConvertTwainNumtoName = "ICAP_IMAGEFILTER "
  Case &H1148
    ConvertTwainNumtoName = "ICAP_NOISEFILTER "
  Case &H1149
    ConvertTwainNumtoName = "ICAP_OVERSCAN "
  Case &H1150
    ConvertTwainNumtoName = "ICAP_AUTOMATICBORDERDETECTION "
  Case &H1151
    ConvertTwainNumtoName = "ICAP_AUTOMATICDESKEW "
  Case &H1152
    ConvertTwainNumtoName = "ICAP_AUTOMATICROTATE "
  Case &H1201
    ConvertTwainNumtoName = "ACAP_AUDIOFILEFORMAT  "
  Case &H1202
    ConvertTwainNumtoName = "ACAP_XFERMECH "
  Case Else
    ConvertTwainNumtoName = "UNKNOWN "
  End Select
End Function


Private Sub subSetPickingRectangle()

    Dim l As Single, t As Single, r As Single, B As Single
    Dim ScannerPhysicalWidth As Single
    
    'SET THE PICKING RECTANGLE AREA FOR THE PAGE
    If Not ItoB(chkUseFlatbed) Then
        ScannerPhysicalWidth = 0
        ' Using ADF - must Center scan area
        TwainPRO.Capability = TWCAP_PHYSICALWIDTH
        If TwainPRO.CapSupported Then
            ScannerPhysicalWidth = TwainPRO.CapValue
        End If
        
        If ScannerPhysicalWidth <= 0 Then
            ScannerPhysicalWidth = 8.5
        End If
        
        If txtTwainImageLeft.Text = "" Then
            txtTwainImageLeft.Text = 0
        End If
        
        If txtTwainImageRight.Text = "" Then
            txtTwainImageRight.Text = ScannerPhysicalWidth
        End If
        
        If txtTwainImageTop.Text = "" Then
            txtTwainImageTop.Text = 0
        End If
        
        If txtTwainImageBottom.Text = "" Then
            txtTwainImageBottom.Text = 11
        End If
        
        ' Calculate left & right based on center of Physical Width as read from scanner
'        l = (ScannerPhysicalWidth - (txtTwainImageRight.Text - txtTwainImageLeft.Text)) / 2
        l = txtTwainImageLeft.Text
        t = txtTwainImageTop.Text
        r = l + txtTwainImageRight.Text
        B = txtTwainImageBottom.Text
    
    Else
        ' Using Flatbed - begin from defined left
        l = txtTwainImageLeft.Text
        t = txtTwainImageTop.Text
        r = txtTwainImageRight.Text
        B = txtTwainImageBottom.Text
    
    End If
    lblStatus.Caption = "Set ImageLayout = Left:" & l & " Top:" & t & " Right:" & r & " Bottom:" & B
    
    On Error Resume Next
    TwainPRO.SetImageLayout l, t, r, B
    If Err.Number <> 0 Then
            MsgBox "The Scanner does NOT support the current Picking Rectangle settings.", vbInformation, "Page Size Error"
            TwainPRO.SetImageLayout 0, 0, 8.5, 11
    End If
End Sub
