VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.6#0"; "PIXEZOCX.OCX"
Begin VB.Form Imaging101ScanMainPix 
   BackColor       =   &H00FFFFFF&
   Caption         =   "BATCH SCANNING"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   12105
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10560
      Picture         =   "Imaging101ScanMainPix.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   128
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdScanBegin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Begin Scanning"
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
      Left            =   10080
      Picture         =   "Imaging101ScanMainPix.frx":0693
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   7560
      Width           =   1875
   End
   Begin VB.TextBox txtBatchDirectory 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   78
      Top             =   1800
      Width           =   8175
   End
   Begin VB.CommandButton cmdBatchDirectoryFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtBatchRootDirectory 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   1440
      Width           =   8292
   End
   Begin VB.TextBox cmbBatchScanSettingsDesc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   480
      Width           =   5295
   End
   Begin VB.ComboBox cmbBatchScanSettingsName 
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
      Height          =   360
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   20
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtApplicationName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtScanningStatus 
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton cmdScanStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&End Scanning"
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
      Left            =   8280
      Picture         =   "Imaging101ScanMainPix.frx":0F5D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1875
   End
   Begin VB.Frame frameImageLayout 
      Caption         =   "Image Picking Rectangle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   11895
      Begin VB.ComboBox cmbTwainDocumentSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   6240
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtTwainImageBottom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTwainImageRight 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTwainImageTop 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTwainImageLeft 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkUseFlatBed 
         Alignment       =   1  'Right Justify
         Caption         =   "Use FlatBed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   9360
         MaskColor       =   &H00008000&
         TabIndex        =   8
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1572
      End
      Begin VB.Label Label13 
         Caption         =   "Paper Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "ImageBottom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "ImageRight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "ImageTop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "ImageLeft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Batch Settings"
      TabPicture(0)   =   "Imaging101ScanMainPix.frx":14E7
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblBatchName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label20"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label24"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label29"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label30"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtBatchName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtBatchScanUser"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtBatchPrefix"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBatchSuffix"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBatchNotes"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtBatchDesc"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBatchRECID"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkUpdateCreditsoftFromBarcode"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdLookup"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Scanner Settings"
      TabPicture(1)   =   "Imaging101ScanMainPix.frx":1503
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdScannerSettingsSave"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frameMisc"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frameBatchDefaults"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FrameBatchOptions"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "frameImageQuality"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Image Caption Settings"
      TabPicture(2)   =   "Imaging101ScanMainPix.frx":151F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdScannerCaptionSettingsSave"
      Tab(2).Control(1)=   "FrmCaption"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Advanced Settings"
      TabPicture(3)   =   "Imaging101ScanMainPix.frx":153B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmCaps"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdLookup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Lookup"
         Default         =   -1  'True
         Height          =   495
         Left            =   -71040
         Picture         =   "Imaging101ScanMainPix.frx":1557
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkUpdateCreditsoftFromBarcode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update CreditSoft from Barcode"
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
         Left            =   -73560
         TabIndex        =   127
         Top             =   3120
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Routing Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -68160
         TabIndex        =   114
         Top             =   480
         Width           =   4935
         Begin VB.ComboBox cmbBatchStatus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            ItemData        =   "Imaging101ScanMainPix.frx":1AE1
            Left            =   1680
            List            =   "Imaging101ScanMainPix.frx":1AE3
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   1764
            Width           =   3135
         End
         Begin VB.ComboBox cmbBatchPriority 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   2202
            Width           =   3135
         End
         Begin VB.ComboBox cmbBatchQueue 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   843
            Width           =   3135
         End
         Begin VB.ComboBox cmbBatchOwner 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   1326
            Width           =   3135
         End
         Begin VB.ComboBox cmbBatchGroup 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   312
            ItemData        =   "Imaging101ScanMainPix.frx":1AE5
            Left            =   1680
            List            =   "Imaging101ScanMainPix.frx":1AEF
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox txtBatchBoxNumber 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1680
            TabIndex        =   115
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label lblBatchStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   126
            Top             =   1824
            Width           =   1572
         End
         Begin VB.Label lblBatchPriority 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Priority"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   125
            Top             =   2280
            Width           =   1572
         End
         Begin VB.Label lblBatchQueue 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Route To Queue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   252
            Left            =   120
            TabIndex        =   124
            Top             =   915
            Width           =   1572
         End
         Begin VB.Label lblBatchOwner 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Route To User"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1386
            Width           =   1575
         End
         Begin VB.Label Label35 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   120
            TabIndex        =   122
            Top             =   2760
            Width           =   1572
         End
         Begin VB.Label lblBatchBoxNumber 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Box #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   120
            TabIndex        =   121
            Top             =   468
            Width           =   612
         End
      End
      Begin VB.Frame frameImageQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Image Quality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         TabIndex        =   104
         Top             =   1680
         Width           =   4335
         Begin VB.CommandButton cmdScannerAdvancedCapabilitiesSet 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Scanner Settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            Picture         =   "Imaging101ScanMainPix.frx":1B03
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   240
            Width           =   2295
         End
         Begin PixezocxLib.PixEzScanControl PixEzScanControlFileType 
            Height          =   315
            Left            =   1560
            TabIndex        =   105
            Top             =   1680
            Width           =   2295
            _Version        =   65542
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            ControlType     =   11
            Data            =   0
            Caption         =   ""
            Picture         =   "Imaging101ScanMainPix.frx":208D
            PictureDisabled =   "Imaging101ScanMainPix.frx":20A9
         End
         Begin PixezocxLib.PixEzScanSetting PixEzScanSettingDuplex 
            Height          =   255
            Left            =   1560
            TabIndex        =   106
            Top             =   1080
            Width           =   2295
            _Version        =   65542
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ControlType     =   19
            Data            =   1
            Caption         =   "Duplex (Scan Both Sides)"
            Picture         =   "Imaging101ScanMainPix.frx":20C5
            PictureDisabled =   "Imaging101ScanMainPix.frx":20E1
         End
         Begin PixezocxLib.PixEzScanControl PixEzScanControlColorFormat 
            Height          =   315
            Left            =   1560
            TabIndex        =   108
            Top             =   2040
            Width           =   2295
            _Version        =   65542
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ControlType     =   12
            Data            =   0
            Caption         =   ""
         End
         Begin PixezocxLib.PixEzScanControl PixEzScanControlCompression 
            Height          =   315
            Left            =   1560
            TabIndex        =   109
            Top             =   2400
            Width           =   2295
            _Version        =   65542
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ControlType     =   13
            Data            =   0
            Caption         =   ""
         End
         Begin PixezocxLib.PixEzScanControl PixEzScanControlMultiPage 
            Height          =   315
            Left            =   1560
            TabIndex        =   110
            Top             =   1320
            Visible         =   0   'False
            Width           =   2295
            _Version        =   65542
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ControlType     =   14
            Data            =   0
            Caption         =   "Multi-Page"
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Compression"
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
            Index           =   1
            Left            =   120
            TabIndex        =   113
            Top             =   2460
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Save as File Type"
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
            Index           =   3
            Left            =   120
            TabIndex        =   112
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Color Format"
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
            Index           =   2
            Left            =   120
            TabIndex        =   111
            Top             =   2100
            Width           =   1575
         End
      End
      Begin VB.Frame FrameBatchOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   7320
         TabIndex        =   80
         Top             =   480
         Width           =   4335
         Begin VB.CheckBox chkScanPreviewOnly 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Preview Image Only (NO Save)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   840
            Width           =   2715
         End
         Begin VB.CheckBox chkAutoDetectPaperOut 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Automatically Detect Paper Out"
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
            TabIndex        =   91
            Top             =   600
            Value           =   1  'Checked
            Width           =   2715
         End
         Begin VB.TextBox txtMinimumImageSize 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   88
            Text            =   "1000"
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbImageRotation 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   840
            TabIndex        =   86
            Text            =   "0"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtScanImageSkipCount 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   83
            Text            =   "10"
            Top             =   1680
            Width           =   375
         End
         Begin VB.CheckBox chkScanDisplayImages 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Display Images While Scanning"
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
            TabIndex        =   82
            Top             =   1440
            Width           =   2715
         End
         Begin VB.CheckBox chkScanShowUI 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Manuf. User Interface"
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
            TabIndex        =   81
            Top             =   360
            Width           =   2715
         End
         Begin VB.Label Label37 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bytes"
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
            Left            =   2880
            TabIndex        =   90
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delete Images Smaller Than"
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
            TabIndex        =   89
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label33 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Rotation"
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
            TabIndex        =   87
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Images"
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
            Left            =   2160
            TabIndex        =   85
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Display every"
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
            Left            =   600
            TabIndex        =   84
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.Frame frameBatchDefaults 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch Defaults"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   4680
         TabIndex        =   59
         Top             =   1680
         Width           =   2535
         Begin VB.CheckBox chkBatchBoxNumberRequired 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "BOX # Required"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   92
            Top             =   1920
            Width           =   1875
         End
         Begin VB.CheckBox chkBatchAutoUseDateTime 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   64
            Top             =   1560
            Width           =   1515
         End
         Begin VB.CheckBox chkBatchAutoUseBatchID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use BatchID  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   63
            Top             =   1320
            Width           =   1515
         End
         Begin VB.CheckBox chkBatchAutoName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto-Assign BatchID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   1080
            Width           =   1875
         End
         Begin VB.TextBox txtBatchSettingsPrefix 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   61
            ToolTipText     =   "Characters at the Beginning of the Batch name"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtBatchSettingsSuffix 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   60
            ToolTipText     =   "Characters at the Beginning of the Batch name"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label27 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Batch Prefix"
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
            TabIndex        =   66
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Batch Suffix"
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
            Left            =   1320
            TabIndex        =   65
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbApplicationList 
         Height          =   315
         Left            =   1920
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtBatchRECID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71040
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtBatchDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73320
         TabIndex        =   3
         Top             =   1680
         Width           =   5052
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
         Height          =   765
         Left            =   -73320
         TabIndex        =   4
         Top             =   2040
         Width           =   5052
      End
      Begin VB.TextBox txtBatchSuffix 
         BackColor       =   &H00FEDCC7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69240
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtBatchPrefix 
         BackColor       =   &H00FEDCC7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74760
         TabIndex        =   0
         Top             =   1200
         Width           =   975
      End
      Begin VB.Frame frameMisc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scan Defaults"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtSelectedScannerDriverLevel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txtSelectedScannerDriverVersion 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtSelectedScannerDriverName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtSelectedScannerName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton cmdScannerSelectScanner 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select Scanner"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scanner Driver Level"
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
            Index           =   7
            Left            =   2040
            TabIndex        =   103
            Top             =   990
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scanner Driver Version"
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
            Index           =   6
            Left            =   2040
            TabIndex        =   102
            Top             =   750
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scanner Driver Name"
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
            Index           =   5
            Left            =   2040
            TabIndex        =   101
            Top             =   510
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scanner Name"
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
            Index           =   4
            Left            =   2040
            TabIndex        =   100
            Top             =   270
            Width           =   1815
         End
      End
      Begin VB.TextBox txtBatchScanUser 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73800
         TabIndex        =   54
         Top             =   4080
         Width           =   2892
      End
      Begin VB.TextBox txtBatchName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73800
         TabIndex        =   1
         Top             =   1200
         Width           =   4575
      End
      Begin VB.CommandButton cmdScannerSettingsSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sa&ve Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Frame FrmCaption 
         Caption         =   "Caption to Display on Image"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -73320
         TabIndex        =   36
         Top             =   480
         Width           =   6375
         Begin VB.CheckBox chkCaptionClip 
            Alignment       =   1  'Right Justify
            Caption         =   "Clip Caption"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   45
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbCaptionVerticalAlign 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2565
            Width           =   1635
         End
         Begin VB.ComboBox cmbCaptionHorizontalAlign 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1725
            Width           =   1635
         End
         Begin VB.TextBox txtCaptionWidth 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   42
            Text            =   "0"
            Top             =   2160
            Width           =   675
         End
         Begin VB.TextBox txtCaptionHeight 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   41
            Text            =   "0"
            Top             =   2580
            Width           =   675
         End
         Begin VB.CheckBox chkCaptionShadowText 
            Alignment       =   1  'Right Justify
            Caption         =   "Shadow Text"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2520
            TabIndex        =   40
            Top             =   960
            Width           =   1260
         End
         Begin VB.TextBox txtCaption 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   39
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox txtCaptionLeft 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   38
            Text            =   "0"
            Top             =   1320
            Width           =   675
         End
         Begin VB.TextBox txtCaptionTop 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   37
            Text            =   "0"
            Top             =   1740
            Width           =   675
         End
         Begin VB.Label Label12 
            Caption         =   "Vertical Alignment"
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
            Left            =   2400
            TabIndex        =   52
            Top             =   2340
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Horizontal Alignment"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2400
            TabIndex        =   51
            Top             =   1500
            Width           =   1515
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionHeight:"
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
            Left            =   300
            TabIndex        =   50
            Top             =   2625
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionWidth:"
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
            Left            =   300
            TabIndex        =   49
            Top             =   2205
            Width           =   1035
         End
         Begin VB.Label Label9 
            Caption         =   "Caption:"
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
            Left            =   720
            TabIndex        =   48
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionLeft:"
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
            Left            =   420
            TabIndex        =   47
            Top             =   1365
            Width           =   915
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "CaptionTop:"
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
            Left            =   420
            TabIndex        =   46
            Top             =   1785
            Width           =   915
         End
      End
      Begin VB.Frame FrmCaps 
         Caption         =   "Capabilities"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -72120
         TabIndex        =   26
         Top             =   480
         Width           =   6735
         Begin VB.TextBox txtCapsListIndex 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox CmbCaps 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   480
            TabIndex        =   30
            Top             =   360
            Width           =   3495
         End
         Begin VB.ListBox List1 
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
            Left            =   480
            TabIndex        =   29
            Top             =   840
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.TextBox EdtCurrent 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1740
            TabIndex        =   28
            Top             =   3120
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdateCapability 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   27
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label LblDefault 
            Caption         =   "LblDefault"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   840
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label LblMin 
            Caption         =   "LblMin"
            Height          =   195
            Left            =   480
            TabIndex        =   34
            Top             =   1200
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label LblMax 
            Caption         =   "LblMax"
            Height          =   195
            Left            =   480
            TabIndex        =   33
            Top             =   1560
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Current Value:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   32
            Top             =   3180
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   0
         TabIndex        =   130
         Top             =   360
         Width           =   11895
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix"
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
         Left            =   -74760
         TabIndex        =   72
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Suffix"
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
         Left            =   -69240
         TabIndex        =   71
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Scan User"
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
         Left            =   -74760
         TabIndex        =   70
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Notes"
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
         Left            =   -74760
         TabIndex        =   69
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Description"
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
         Left            =   -74760
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblBatchName 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Name"
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
         Height          =   255
         Left            =   -73800
         TabIndex        =   67
         Top             =   950
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4095
         Left            =   -74880
         Top             =   480
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   10560
      TabIndex        =   129
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Settings Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   79
      Top             =   189
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   74
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Directory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
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
      TabIndex        =   77
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Root Directory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
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
      TabIndex        =   76
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   75
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FEDCC7&
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
      TabIndex        =   73
      Top             =   8160
      Width           =   11895
   End
End
Attribute VB_Name = "Imaging101ScanMainPix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim m_ImageCount As Integer
    Dim m_PageCount As Integer
    Dim m_ImageSkipCount As Integer
    Dim intLoop As Integer
    
    
''''    Dim bolCancelPendingXfers As Boolean
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
    
    Dim strReturn As String
    



Private Sub chkBatchAutoName_Click()

    On Error GoTo ERROR_HANDLER
    
    If ItoB(chkBatchAutoName) = True Then
        chkBatchAutoUseBatchID.Enabled = True
        chkBatchAutoUseDateTime.Enabled = True
    Else
        chkBatchAutoUseBatchID.Enabled = False
        chkBatchAutoUseDateTime.Enabled = False
    End If
    
Exit Sub

ERROR_HANDLER:

'    MsgBox "chkBatchAutoName_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    Resume Next
    
End Sub

Private Sub chkScanDisplayImages_Click()

    'See if the Images should be displayed while scanning.
        'Check if Images should be displayed while scanning
        If chkScanDisplayImages = vbChecked Then
            Imaging101ScanViewer.Visible = True
            Imaging101ScanViewer.Show
        Else
            Imaging101ScanViewer.Visible = False
            Imaging101ScanViewer.Hide
        End If
    
End Sub


Private Sub cmbBatchQueue_Click()

    On Error Resume Next
    
    '2. Save Last Used Batch Queue
    lblStatus.Caption = "Save Last Used Batch Queue"
    funcGetSetUserSettings "SET", "DefaultScanBatchQueue", cmbBatchQueue.Text
    
    txtBatchName.SetFocus

End Sub

Private Sub cmbBatchScanSettingsName_click()

    MousePointer = MousePointerConstants.vbHourglass
    
    On Error GoTo ERROR_HANDLER
    
'    '*** Changed code to Save/Load the Scanner Last Selected in the workstations LOCAL INI File
    lblStatus.Caption = "Save/Load the Scanner Last Selected to the " & RegFileName & " file"
    result = WritePrivateProfileString(RegAppname, "Imaging101ScanMainPix.cmbBatchScanSettingsName", cmbBatchScanSettingsName, RegFileName)
    
    lblStatus.Caption = "CALL subScannerSettingsGetSettings"
    subScannerSettingsGetSettings
    
    If Not bolCancelPendingXfers Then
        lblStatus.Caption = "CALL subShowScannerInfo"
        subShowScannerInfo
    End If
    
    Me.SetFocus

    MousePointer = MousePointerConstants.vbDefault

    
Exit Sub

ERROR_HANDLER:

'    MsgBox "cmbBatchScanSettingsName_click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    Resume Next
End Sub

Private Sub cmbBatchScanSettingsName_DropDown()

    subScannerSettingsLoadList

End Sub







Private Sub cmdBatchDirectoryFind_Click()

    On Error GoTo ERROR_HANDLER
    
    txtBatchRootDirectory = funcGetDirectoryLocation("C:\Workarea\Jacob")
    
Exit Sub

ERROR_HANDLER:

'    MsgBox "cmdBatchDirectoryFind_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    Resume Next
End Sub



Private Sub cmdStopScanning_Click()

    bolCancelPendingXfers = True

End Sub

Private Sub subScannerSettingsGetSettings()

    On Error GoTo ERROR_HANDLER

''     Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbNormal

     Dim txtOvewriteScannerSettings As Integer
     

     '*****************************************************************
     ' Establish the Imaging101 Batch List Connection
     txtActionBeforeError = "CONNECT Imaging101 Batch List Connection"
     lblStatus.Caption = txtActionBeforeError

     Set cmdImaging101Batch.ActiveConnection = connImaging101Batch
     
     ' Open BatchScannerSettings table.
     Dim rsBatchScannerSettings As ADODB.Recordset
     Set rsBatchScannerSettings = New ADODB.Recordset
     rsBatchScannerSettings.CursorType = adOpenDynamic
     rsBatchScannerSettings.LOCKTYPE = adLockReadOnly
     rsBatchScannerSettings.Source = "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'"
     
     txtActionBeforeError = "SELECT * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'"
     lblStatus.Caption = txtActionBeforeError
     
     rsBatchScannerSettings.Open "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'", connImaging101Batch
     
     If rsBatchScannerSettings.EOF Then
        txtOvewriteScannerSettings = MsgBox("Sorry!  I couldn't find Settings for [ " & cmbBatchScanSettingsName & "]... It might have been deleted." & vbCrLf & "Please configure the settings and then Click the [Save Settings] button.", vbOK)
        'Set the focus on the "Scanner Settings" tab.
        SSTab1.Visible = True
        SSTab1.Tab = 1
        Exit Sub
     End If
      
            
     txtActionBeforeError = "Assign Record Values to Variables"
     lblStatus.Caption = txtActionBeforeError
            
    '*** Using the [ & "" ] at the end of each line to Prevent VB Error 94 = Invalid use of Null
    '***      for Chkboxes or numbers must use [ & 0 ]
    cmbBatchScanSettingsName = rsBatchScannerSettings("BatchScannerSettingsName") & ""
    cmbBatchScanSettingsDesc = rsBatchScannerSettings("BatchScannerSettingsDesc") & ""
    txtBatchRootDirectory = rsBatchScannerSettings("BatchRootDirectory") & ""
    
    Imaging101ScanViewer.ezPageView.ScanDriverName = rsBatchScannerSettings("ScanSourceName") & ""
    
    subShowScannerInfo
    
    If Not bolCancelPendingXfers Then
        'Save Pixtran Scan Settings
        Imaging101ScanViewer.ezPageView.ScanDuplex = rsBatchScannerSettings("ScanDuplex") & ""
        PixEzScanControlFileType = rsBatchScannerSettings("ScanFileType") & ""
    '    PixEzScanControlColorFormat = rsBatchScannerSettings("ScanColorFormat") & ""
    '    PixEzScanControlCompression = rsBatchScannerSettings("ScanCompression") & ""
        Imaging101ScanViewer.ezPageView.ScanColorformat = rsBatchScannerSettings("ScanColorFormat") & ""
        Imaging101ScanViewer.ezPageView.ScanCompression = rsBatchScannerSettings("ScanCompression") & ""
        
        'Jacob - 9/18/2009 - DISABLED ScanStateFlush
'        lblStatus.Caption = "FLUSH Scanner Settings"
'        Imaging101ScanViewer.ezPageView.ScanStateFlush
    End If
    
'''    cmbTwainDocumentSize = rsBatchScannerSettings("ScanDocumentSize") & ""
'''    txtTwainImageTop = rsBatchScannerSettings("ScanImageTop") & ""
'''    txtTwainImageLeft = rsBatchScannerSettings("ScanImageLeft") & ""
'''    txtTwainImageRight = rsBatchScannerSettings("ScanImageRight") & ""
'''    txtTwainImageBottom = rsBatchScannerSettings("ScanImageBottom") & ""
    
    
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
    chkUseFlatBed = rsBatchScannerSettings("ScanUseFlatBed") & ""
    
    
    txtCaption = rsBatchScannerSettings("ScanCaption") & ""
    chkCaptionClip = rsBatchScannerSettings("ScanCaptionClip") & "0"
    chkCaptionShadowText = rsBatchScannerSettings("ScanCaptionShadowText") & ""
    txtCaptionLeft = rsBatchScannerSettings("ScanCaptionLeft") & ""
    txtCaptionTop = rsBatchScannerSettings("ScanCaptionTop") & ""
    txtCaptionWidth = rsBatchScannerSettings("ScanCaptionWidth") & ""
    txtCaptionHeight = rsBatchScannerSettings("ScanCaptionHeight") & ""
    cmbCaptionHorizontalAlign.ListIndex = rsBatchScannerSettings("ScanCaptionHorizontalAlign") & ""
    cmbCaptionVerticalAlign.ListIndex = rsBatchScannerSettings("ScanCaptionVerticalAlign") & ""
    
    lblStatus.Caption = "Close "
    
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
    
    '*** LOAD SCANNER SETTINGS
    If Not bolCancelPendingXfers Then
'        subShowScannerInfo
        lblStatus.Caption = "BEGIN: LOAD SCANNER SETTINGS"
        Imaging101ScanViewer.ezPageView.ScanStateRead strScannerSettingsFileName, txtSelectedScannerDriverName
        Call SSTab1_GotFocus
        lblStatus.Caption = "END  : LOAD SCANNER SETTINGS"
        
        'Jacob - 9/18/2009 - DISABLED ScanStateFlush
'        Imaging101ScanViewer.ezPageView.ScanStateFlush
    End If
    

    
Exit Sub

ERROR_HANDLER:

'    MsgBox "subScannerSettingsGetSettings ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError

    MsgBox "SORRY!!!" & vbCrLf & "I CAN'T FIND THE SCANNER!!!", vbCritical
    bolCancelPendingXfers = True
    SSTab1.Visible = True

'    Resume Next
    
End Sub
Private Sub subScannerSettingsLoadList()
    
    On Error GoTo ERROR_HANDLER

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
     rsBatchScannerSettings.LOCKTYPE = adLockReadOnly
     rsBatchScannerSettings.Open "SELECT BatchScannerSettingsName FROM I101BatchScannerSettings WHERE ApplicationRECID=" & txtApplicationRECID & "Order by BatchScannerSettingsName", connImaging101Batch
     
    
     If rsBatchScannerSettings.EOF Then
         txtOvewriteScannerSettings = MsgBox("Sorry!  I couldn't find Settings for [ " & cmbBatchScanSettingsName & "] in Application [ " & _
                                             txtApplicationName & " ]... " & vbCrLf & "It might have been deleted." & vbCrLf & _
                                             "Please create a new one!", vbOK)
        SSTab1.Visible = True
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
    
''    '*** Changed code to Save/Load the Scanner Last Selected in the workstations LOCAL REGISTRY
''    cmbBatchScanSettingsName = funcGetSetUserSettings("GET", "BatchScanSettings_" & frmImaging101Winsock.txtComputerName, "")
''    cmbBatchScanSettingsName = GetSetting("Imaging101", "ScannerSettings", "ScanSettingsName", "")
    On Error Resume Next
    cmbBatchScanSettingsName = VBGetPrivateProfileString(RegAppname, "Imaging101ScanMainPix.cmbBatchScanSettingsName", RegFileName)
    
    
Exit Sub

ERROR_HANDLER:

'    MsgBox "subScannerSettingsLoadList ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    Resume Next
    
End Sub




Private Sub cmdLookup_Click()

    funcWriteToDebugLog Me.name, "ENTERING txtBatchName_GotFocus()"
    funcWriteToDebugLog Me.name, "frmLookupList.Show"
    frmLookupList.Show
    frmLookupList.txtTableLookupField.Text = txtBatchName
    frmLookupList.cmdFind_Click
    
End Sub

Private Sub cmdScanBegin_Click()

    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] *********************************************************************"
    
    '*** Make SURE we have a Batch PATH assigned
    txtBatchRootDirectory = Trim(txtBatchRootDirectory)
    If txtBatchRootDirectory = "" Then
        txtBatchRootDirectory = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID=" & txtApplicationRECID, "RootDirectoryPathForBatches")
        MsgBox "You have NOT assigned a Batch Root Directory!  The DEFAULT directory has been set.  Please check that this is the correct Batch Location and Click [SCAN] again.", vbOKOnly
        Exit Sub
    End If
    
    '*** If Auto-name Batch is selected, set to Auto
    If chkBatchAutoName.Value = vbChecked Then
        txtBatchName.Text = "Auto"
    End If
    
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
    
    If gsecRightsBatchRoute = True _
    And Trim(cmbBatchQueue.Text = "") Then
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
        
        
        'Show Manufacturer User Interface (UI) if selected
        If chkScanShowUI.Value = vbChecked Then
            lblStatus.Caption = "Set Show Manuf. User Interface"
            Imaging101ScanViewer.ezPageView.ScanStateRead strScannerSettingsFileName, txtSelectedScannerDriverName
            Imaging101ScanViewer.ezPageView.ShowScannerSettingsDialog
            Imaging101ScanViewer.ezPageView.ScanStateWrite strScannerSettingsFileName, txtSelectedScannerDriverName
        Else
            lblStatus.Caption = "DON'T Show Manuf. User Interface"
        End If
        
        If ItoB(chkUseFlatBed) Then
            lblStatus.Caption = "Use Flatbed"
            Imaging101ScanViewer.ezPageView.ScanPaperSource = 1 ' Flatbed
            'Scanning with Flatbed
            result = MsgBox("Place paper in Flatbed and Click [OK]... Click [Cancel] when done.", vbOKCancel)
            If result = vbCancel Then
                lblStatus.Caption = "Batch Ended by User."
                txtScanningStatus = "Batch Ended by User."
                subToggleScanButtons
                GoTo EXIT_SUB
            End If
        Else
            
            '*** Prompt for Retry after each Session if chkAutoDetectPaperOut is Unchecked.
            '     This is for scanners like the CANON 3080C Driver that works erratically
            If chkAutoDetectPaperOut = vbUnchecked Then
                    lblStatus.Caption = "Auto Detect Paper Out"
                    result = MsgBox("Place paper in Feeder and Click Retry or Cancel.", vbRetryCancel)
                    If result = vbCancel Then bolCancelPendingXfers = True
                    If bolCancelPendingXfers Then
                        lblStatus.Caption = "Auto Detect Paper Out - Batch Ended by User."
                        txtScanningStatus = "Batch Ended by User."
                        subToggleScanButtons
                        
                        ' 2014-04-25 - Jacob - UN-LOCK the Batch, ignore errors
                        strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)

                        GoTo EXIT_SUB
                    End If
                    
            Else

            End If
        End If
        
        
        txtScanningStatus = ""
        
        ' ************************************************
        ' ***  PREPARE FOR SCANNING
    
        'Check if Images should be displayed while scanning
        If chkScanDisplayImages = vbChecked Then
            Imaging101ScanViewer.Visible = True
            Imaging101ScanViewer.Show
        Else
            Imaging101ScanViewer.Visible = False
            Imaging101ScanViewer.Hide
        End If
    

        'Create the Batch Header Record
        lblStatus.Caption = "Create Batch Record"
        subCreateBatchRecord
                
        
        '****************************************************
        '*** BEGIN SCANNING THE BATCH
        
        lblStatus.Caption = "Scan Batch"
        
        If frmImaging101BatchList.lblApplicationCommitBatchTo = "TTC" Then
            
'            '******************************************************
'            '***  TTC Scanning
'
'            Load frmPixMultiStreamMain
'
'            'Only show scan button if Sys Admin
'            If gsecRightsAdminSystem <> vbChecked Then
'                frmPixMultiStreamMain.CmdScan.Visible = False
'            End If
'
'            'Load and Show the MultiStream form
'            frmPixMultiStreamMain.Show
'            'Now show this form again
'            Me.Hide
'
'            DoEvents
'
'            'Click the Scan button
'            Call frmPixMultiStreamMain.CmdScan_Click
'
''            'Now show this form again
''            Me.Show
'
'            cmdScanStop_Click
'            subToggleScanButtons
'
'            ' 2014-04-25 - Jacob - UN-LOCK the Batch, ignore errors
'            strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)
'
'            Exit Sub
            
        Else
            
            '******************************************************
            '***  NON-TTC Scanning
            
            'Set up the scan parameters
            lblStatus.Caption = "Set up the scan parameters"
            
            Imaging101ScanViewer.ezPageView.Close
            
            Imaging101ScanViewer.ezPageView.ScanStateRead strScannerSettingsFileName, txtSelectedScannerDriverName
            
            funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] ScanStateRead"
            funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] strScannerSettingsFileName     = " & strScannerSettingsFileName
            funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] txtSelectedScannerDriverName   = " & txtSelectedScannerDriverName
            funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] txtSelectedScannerDriverVersion= " & txtSelectedScannerDriverVersion
            funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] txtSelectedScannerDriverLevel  = " & txtSelectedScannerDriverLevel

            'Jacob - 9/18/2009 - DISABLED ScanStateFlush
'            Imaging101ScanViewer.ezPageView.ScanStateFlush
    
            Imaging101ScanViewer.ezPageView.ScanAllowLongNames = 1     ' Allow long names
            Imaging101ScanViewer.ezPageView.ScanInsertMode = 2         ' overwrite
            Imaging101ScanViewer.ezPageView.ImageFileSchemaDetect = 0  ' We are specifying a schema, so do not try to detect
            
            
            PixEzScanControlMultiPage = "Single-Page"
            
            If PixEzScanControlMultiPage = "Multi-Page" Then
                
    '            Const TEMPFILE As String = "c:\tmp\ezapp.tmp"
    '            Const PDFFILE As String = "c:\tmp\app.pdf"
    '            Const SCANDIR As String = "c:\tmp"
    '            Const SCANROOT As String = "root"
    '            Const SCANSCHEMA As String = "%02d%s%05d,b,r,p"
    '            Const SCANEXTPDF As String = ".pdf"
    
                Imaging101ScanViewer.ezPageView.ScanFileExt = funcPixGetFileExt(Imaging101ScanViewer.ezPageView)    ' Set image file extension
                
                Imaging101ScanViewer.ezPageView.ScanFileName = txtBatchName & Imaging101ScanViewer.ezPageView.ScanFileExt
                Imaging101ScanViewer.ezPageView.ScanUseSchema = False
                Imaging101ScanViewer.ezPageView.ScanMultipage = True
            
                'change format to 1,1 for G4 compression
                Imaging101ScanViewer.ezPageView.ScanPackaging = &H100000
                Imaging101ScanViewer.ezPageView.ScanSamplesPerPixel = 1
                Imaging101ScanViewer.ezPageView.ScanBitsPerSample = 1
                Imaging101ScanViewer.ezPageView.ScanCompression = 4
                Imaging101ScanViewer.ezPageView.ScanFileDir = "C:\TEMP"  ' Set directory for images
            
            Else
            
                Imaging101ScanViewer.ezPageView.ScanUseSchema = 1
                Imaging101ScanViewer.ezPageView.ScanFileDir = Trim(txtBatchRootDirectory) & "\" & Format(txtBatchRECID, "0000000000")  ' Set directory for images
                 
        
                Imaging101ScanViewer.ezPageView.ScanFileExt = funcPixGetFileExt(Imaging101ScanViewer.ezPageView)    ' Set image file extension
                
                Imaging101ScanViewer.ezPageView.SavePrecedence = 1   ' 0 = Let Color Format setting determine file types
                
                Imaging101ScanViewer.ezPageView.ScanFileRoot = CStr(Format(txtBatchRECID, "0000000000")) & "-"   ' Set schema root. All files will begin with "PIXEL"
                 
                ' Make sure the ScanFileSchema does not have more than nine (9) pound signs (#'s) otherwise
                '  we get spaces instead of zero's
                Imaging101ScanViewer.ezPageView.ScanFileSchema = "$####;" ' Set schema: root name plus two digits.
            
                funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] ScanPackaging  = " & Imaging101ScanViewer.ezPageView.ScanPackaging
                funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] SaveColorformat= " & Imaging101ScanViewer.ezPageView.SaveColorformat
                funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] SaveCompression= " & Imaging101ScanViewer.ezPageView.SaveCompression
                funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] ScanFileExt    = " & Imaging101ScanViewer.ezPageView.ScanFileExt
                funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] ScanFileRoot   = " & Imaging101ScanViewer.ezPageView.ScanFileRoot
                funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] ScanFileSchema = " & Imaging101ScanViewer.ezPageView.ScanFileSchema
                
            End If
            
            
            Imaging101ScanViewer.ezPageView.ScanBatch
            
            
        
        End If
        
        
        ' See if the Manufacturer UI was enabled and the user Clicked Exit without scanning.
        If m_ImageCount < 1 Then
            lblStatus.Caption = "Image Count < 1"
            
            ' 2014-04-25 - Jacob - UN-LOCK the Batch, ignore errors
'            strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)

            '2017-08-23 - Jacob - Changed from "UNLOCK" to DELETE the Batch with Zero Pages
            funcWriteToDebugLog Me.name, "[cmdScanBegin_Click] - SCAN CANCELLED - ZERO PAGES - DELETE FROM I101Batches WHERE BatchRECID = " & txtBatchRECID
            result = funcRunSQLCommand(RegImaging101BatchListConnectionString, "DELETE FROM I101Batches WHERE BatchRECID = " & txtBatchRECID)
            
            ' Delete Folder
            Dim ofs As Scripting.FileSystemObject
            Set ofs = New Scripting.FileSystemObject
            ofs.DeleteFolder Imaging101ScanViewer.ezPageView.ScanFileDir
            Set ofs = Nothing


            cmdScanStop_Click
            subToggleScanButtons
            SSTab1.Tab = 0
            txtBatchName.SetFocus
            txtBatchName.SelStart = 0          ' Highlight the current entry.
            txtBatchName.SelLength = Len(txtBatchName.Text)

            Exit Sub
        End If
        
    Wend
        
'    'UN-Lock the Batch, ignore return
'    strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)

     
EXIT_SUB:
        
    'UN-Lock the Batch, ignore return
    strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)

     subToggleScanButtons
     
    'RESET FIELDS FOR NEXT SCAN - Except for BatchName if Auto-name selected
    If chkBatchAutoName.Value = vbChecked Then
        txtBatchName.Text = "Auto"
    Else
        txtBatchName.Text = ""
    End If
        
    txtBatchName.Text = ""
    txtBatchDesc.Text = ""
    txtBatchNotes.Text = ""
    SSTab1.Tab = 0
    txtBatchName.SetFocus
    txtBatchName.SelStart = 0          ' Highlight the current entry.
    txtBatchName.SelLength = Len(txtBatchName.Text)

    
Exit Sub


ERROR_HANDLER:

    'UN-Lock the Batch, ignore return
    strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)

'    If Err.Number = 4426 Or Err.Number = 4769 Then
        ' Err 4426 = Scanner Jam
        ' Err 4769 = Command Failed
        result = MsgBox("cmdScanBegin_Click ERROR: " & Err.Number & _
                vbCrLf & vbCrLf & Err.Description & _
                vbCrLf & vbCrLf & "Please check the Scanner and Try Again." & _
                vbCrLf & vbCrLf & "To Retry, Click [OK]" & _
                vbCrLf & vbCrLf & "To Stop scanning, Click [Cancel]", vbOKCancel)
        If result <> vbCancel Then
            bolCancelPendingXfers = False
            Resume
        End If
'    Else
'        MsgBox "cmdScanBegin_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'        GetError
'        Resume Next
'    End If

    subToggleScanButtons
    

End Sub



Private Sub cmdScannerCaptionSettingsSave_Click()
    
    cmdScannerSettingsSave_Click

End Sub



Private Sub cmdScannerSelectScanner_Click()

    On Error GoTo ERROR_HANDLER
    
    Me.Hide
    
    If Not bolCancelPendingXfers Then
        subShowScannerInfo
    End If
    
    'The ScanSelect method ALWAYS reads the SETSCAN.INI File to get the Default Driver
    '   Override by saving the Currently Selected Driver Name to their SETSCAN.INI File
    '   Just in case the user changed the Scanning Profile

    '*** KILL the file if it already exists to make absolutely sure settings from
    '     different scanners don't get mixed.  If we try to load a mixed file,
    '     Pixtran will give strange errors like:
    '          "Run-time error '4433': File Not Found." or
    '          "Error 4433: No document is currently open"
'''    If funcFileExists(strScannerSettingsFileName) Then
''''        Kill strScannerSettingsFileName
'''        Kill "C:\WINDOWS\SETSCAN.INI"
'''    End If
    
'    Imaging101ScanViewer.ezPageView.ScanLoaded = 0
'    result = WritePrivateProfileString("Scanner", "Driver", txtSelectedScannerDriverName, "SETSCAN.INI")
'    Imaging101ScanViewer.ezPageView.ScanLoaded = 1


    'SHOW The Scanner Selection List
    Imaging101ScanViewer.ezPageView.ScanSelect

    If Not bolCancelPendingXfers Then
        subShowScannerInfo
    End If
    
    '*****************************************************************************************
    '***  THIS IS NECESSARY - IT TRICKS THE PIXTRAN ISIS TO SET THE CORRECT SCANNER DRIVER.
    result = WritePrivateProfileString("Scanner", "Driver", txtSelectedScannerDriverName, "SETSCAN.INI")
    
    'Check if the Profile File exists.
    If Not funcFileExists(strScannerSettingsFileName) Then
        'The File Does NOT exist... Configure the Scanner Settings.
        result = MsgBox("Sorry!  I couldn't find Settings file for [" & cmbBatchScanSettingsName & "]... It might have been deleted." & vbCrLf & "Please configure the settings and then Click the [Save Settings] button.", vbOKOnly)
        'Set the focus on the "Scanner Settings" tab.
        SSTab1.Visible = True
        SSTab1.Tab = 1
        cmdScannerAdvancedCapabilitiesSet_Click
    Else
        'File DOES Exist, now check if the Section exists for THIS ScannerDriverName
        Dim strScannerDriverSection As String
        strScannerDriverSection = VBGetPrivateProfileString(txtSelectedScannerDriverName, vbNullString, strScannerSettingsFileName)
        If Trim(strScannerDriverSection) = "" Then
            result = MsgBox("Sorry!  I couldn't find Settings for Scanner Driver [" & txtSelectedScannerDriverName & "]" & vbCrLf & " in Profile [" & cmbBatchScanSettingsName & "]... It might have been deleted." & vbCrLf & "Please configure the settings and then Click the [Save Settings] button.", vbOKOnly)
            cmdScannerAdvancedCapabilitiesSet_Click
        End If
        
    End If

    cmdScannerSettingsSave_Click

    subShowScannerInfo


    Me.Show
    
Exit Sub

ERROR_HANDLER:

    Unload Imaging101ScanViewer
    MsgBox "[cmdScannerSelectScanner_Click] Error Selecting Scanner:  " & Err.Number & " - " & Err.Description
    Me.Show
    
End Sub





Private Sub cmdScannerSettingsSave_Click()
    
    On Error GoTo ERROR_HANDLER
    
    ''     Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbNormal

     ' Establish the Imaging101 Batch List Connection
     txtActionBeforeError = "Establish the Imaging101 Batch List Connection"
     Dim txtOvewriteScannerSettings As Integer
     
     ' Open BatchScannerSettings table.
     Dim rsBatchScannerSettings As ADODB.Recordset
     Set rsBatchScannerSettings = New ADODB.Recordset
     rsBatchScannerSettings.CursorType = adOpenDynamic
     rsBatchScannerSettings.LOCKTYPE = adLockOptimistic
     rsBatchScannerSettings.Source = "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'"
     rsBatchScannerSettings.Open "Select * FROM I101BatchScannerSettings where BatchScannerSettingsName  ='" & cmbBatchScanSettingsName & "'", connImaging101Batch
     
     'User Transaction Tracking to prevent partial imports!
     connImaging101Batch.BeginTrans

     
     If Not rsBatchScannerSettings.EOF Then
         txtOvewriteScannerSettings = MsgBox("Do you wish to Overwrite Batch Scanner Settings for [" & cmbBatchScanSettingsName & "]?", vbYesNo)
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
    
    ' Save Scanner Driver name
    rsBatchScannerSettings("ScanSourceName") = Imaging101ScanViewer.ezPageView.ScanDriverName


    'Save Pixtran Scan Settings
    rsBatchScannerSettings("ScanDuplex") = Imaging101ScanViewer.ezPageView.ScanDuplex
    rsBatchScannerSettings("ScanFileType") = Imaging101ScanMainPix.PixEzScanControlFileType
'    rsBatchScannerSettings("ScanColorFormat") = Imaging101ScanMainPix.PixEzScanControlColorFormat
'    rsBatchScannerSettings("ScanCompression") = Imaging101ScanMainPix.PixEzScanControlCompression
    rsBatchScannerSettings("ScanColorFormat") = Imaging101ScanViewer.ezPageView.ScanColorformat
    rsBatchScannerSettings("ScanCompression") = Imaging101ScanViewer.ezPageView.ScanCompression
    
'    rsBatchScannerSettings("ScanTwainColor") = cmbTwainColor
'    rsBatchScannerSettings("ScanResolution") = cmbTwainResolution
'    rsBatchScannerSettings("ScanContrast") = sldTwainContrast
'    rsBatchScannerSettings("ScanIntensity") = sldTwainIntensity
    
'    rsBatchScannerSettings("ScanDocumentSize") = cmbTwainDocumentSize
'    rsBatchScannerSettings("ScanImageTop") = txtTwainImageTop
'    rsBatchScannerSettings("ScanImageLeft") = txtTwainImageLeft
'    rsBatchScannerSettings("ScanImageRight") = txtTwainImageRight
'    rsBatchScannerSettings("ScanImageBottom") = txtTwainImageBottom
    
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
    rsBatchScannerSettings("ScanUseFlatBed") = chkUseFlatBed
    
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
    
'    '*** Changed code to Save/Load the Scanner Last Selected in the workstations LOCAL INI File
    result = WritePrivateProfileString(RegAppname, "Imaging101ScanMainPix.cmbBatchScanSettingsName", cmbBatchScanSettingsName, RegFileName)
    
    Imaging101ScanViewer.ezPageView.ScanStateWrite strScannerSettingsFileName, txtSelectedScannerDriverName

    
Exit Sub

ERROR_HANDLER:

'    MsgBox "cmdScannerSettingsSave_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    Resume Next
End Sub


Private Sub cmdScanStop_Click()
    
    'Set Flag to stop scanning
    bolCancelPendingXfers = True
    
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





Private Sub cmdScannerAdvancedCapabilitiesSet_Click()
    
    Me.Hide
    
    If Not bolCancelPendingXfers Then
        subShowScannerInfo
    End If
    
'    Imaging101ScanViewer.ezPageView.ScanStateRead strScannerSettingsFileName, txtSelectedScannerDriverName
'
'    Imaging101ScanViewer.ezPageView.ScanStateFlush
    
    '*** Open the ISIS Scanner Settings Dialog
    If frmImaging101BatchList.lblApplicationCommitBatchTo = "TTC" Then
        
'        Load frmPixMultiStreamMain
'
'        'Only show scan button if Sys Admin
'        If gsecRightsAdminSystem <> vbChecked Then
'            frmPixMultiStreamMain.CmdScan.Visible = False
'        End If
'
'        'Load as MODAL so other I101 features are disabled till done scanning or configuring settings
'        frmPixMultiStreamMain.Show 'vbModal, Me
'
'        Exit Sub
'        'Now show this form again
''        Me.Show
        
    Else
        
        Imaging101ScanViewer.ezPageView.ShowScannerSettingsDialog
        
        Imaging101ScanViewer.ezPageView.ScanStateWrite strScannerSettingsFileName, txtSelectedScannerDriverName
        
        'Jacob - 9/18/2009 - DISABLED ScanStateFlush
'        Imaging101ScanViewer.ezPageView.ScanStateFlush
    
    End If
    
    
    If Not bolCancelPendingXfers Then
        subShowScannerInfo
    End If
    
    Me.Show

Exit Sub

ERROR_HANDLER:
'    GetError
    bolCancelPendingXfers = True
    Me.Visible = True
        
End Sub


Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    
        ' Get saved settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "Imaging101ScanMainPix.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "Imaging101ScanMainPix.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "Imaging101ScanMainPix.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "Imaging101ScanMainPix.Height", RegFileName)
    On Error GoTo 0

    'Set Boolean flag
    bolBatchScanningModule = True

    On Error GoTo ERROR_HANDLER

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
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(3) = False
    End If
 

    Me.Show
    
    lblStatus.Caption = "OPENING DATABASES..."
    DoEvents
    
    
    ' SET UP VARIABLES
        txtApplicationRECID = frmImaging101BatchList.txtApplicationRECID
        txtApplicationName = frmImaging101BatchList.cmbApplicationList.Text
        txtBatchName = frmImaging101BatchList.txtBatchName
        txtBatchDesc = frmImaging101BatchList.txtBatchDesc
        txtBatchDirectory = frmImaging101BatchList.txtBatchDirectory
    
        
    Screen.MousePointer = vbHourglass
    
    ' Establish the Imaging101 DB Connections
    txtActionBeforeError = "Prepare Imaging101 DB Connections"
    Set connImaging101 = New ADODB.Connection
    Set cmdImaging101 = New ADODB.Command
    Set rsImaging101 = New ADODB.Recordset
    
    
    connImaging101.ConnectionTimeout = 120
    connImaging101.CommandTimeout = 600
    cmdImaging101.CommandTimeout = 600
    
'''''''    On Error Resume Next
'''''''    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
'''''''    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    connImaging101.ConnectionString = RegImaging101ConnectionString
    connImaging101.mode = adModeReadWrite
    connImaging101.Open
    
    Set cmdImaging101.ActiveConnection = connImaging101
    
    ' Establish the Imaging101Batch DB Connections
    txtActionBeforeError = "Prepare Imaging101Batch DB Connections"
    Set connImaging101Batch = New ADODB.Connection
    Set cmdImaging101Batch = New ADODB.Command
    Set rsImaging101Batch = New ADODB.Recordset
    
    connImaging101Batch.ConnectionTimeout = 120
    connImaging101Batch.CommandTimeout = 600
    cmdImaging101Batch.CommandTimeout = 600
    
'''''''    On Error Resume Next
'''''''    RegImaging101BatchListConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionType", RegFileName)
'''''''    RegImaging101BatchListConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionString." & RegImaging101BatchListConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
    connImaging101Batch.mode = adModeReadWrite
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
    
    rs.Source = "Select * from I101Security ORDER BY UserName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
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
        
    Dim con As ADODB.Connection
    
    Set con = New ADODB.Connection
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "SELECT * from I101BatchQueues " & _
                " WHERE (ApplicationRECID = " & txtApplicationRECID & _
                " OR  ApplicationRECID = 0 OR ApplicationRECID IS NULL)  " & _
                " AND (BatchQueueActive = 'Y') " & _
                " AND (BatchQueueAllowScanInto = '1' )" & _
                " ORDER BY BatchQueue "
                
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
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
    con.Close
    Set con = Nothing

    '*** Set the Top/First item as the default
'    If cmbBatchQueue.ListCount > 0 Then
'        cmbBatchQueue.ListIndex = cmbBatchQueue.TopIndex
'    End If
    
    '***************************************
    '*** LOAD BATCH STATUS LIST DROP-DOWN
        
    Set con = New ADODB.Connection
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchStatus ORDER BY BatchStatus"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
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
    con.Close
    Set con = Nothing

    '****************************
    
    '***************************************
    '*** LOAD BATCH PRIORITY LIST DROP-DOWN
        
    Set con = New ADODB.Connection
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchPriority ORDER BY BatchPriority"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
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
    con.Close
    Set con = Nothing

    '****************************

    '*** Set the Default Values for the DropDown Lists
    funcFindItemInComboBox Me.cmbBatchGroup, "REGULAR"
    funcFindItemInComboBox Me.cmbBatchPriority, "3- LOW"
    funcFindItemInComboBox Me.cmbBatchStatus, "Unassigned"


    '**********************************************************
    '*** ADD SPECIAL OPTIONS TO cmbBatchGroup LIST DROP-DOWN
        
    Select Case frmImaging101BatchList.cmbApplicationList.Text
        Case "TTC"
            cmbBatchGroup.AddItem "TTC PRINTED"
            cmbBatchGroup.AddItem "TTC RECEIVED"
    End Select



    If gsecRightsBatchRoute <> True Then
        lblBatchOwner.Visible = False
        cmbBatchOwner.Visible = False
        lblBatchPriority.Visible = False
        cmbBatchPriority.Visible = False
        lblBatchQueue.Visible = False
        cmbBatchQueue.Visible = False
        lblBatchStatus.Visible = False
        cmbBatchStatus.Visible = False
    End If
    


    
    lblStatus.Caption = "LOADING SCANNER SOURCE NAMES..."
    DoEvents
    
    
    
    ' Load list of available ScannerSettings names
    bolCancelPendingXfers = False
    subScannerSettingsLoadList
    
      'If NO Scanner Settings were loaded, get out of here NOW!
    If cmbBatchScanSettingsName.ListCount = 0 Then
        Exit Sub
    End If
    
    
    txtBatchScanUser = gsecUserID

    On Error GoTo ERROR_HANDLER
    
  
    lblStatus.Caption = "GETTING SCANNER USER SETTINGS..."
    DoEvents
    subScannerSettingsGetSettings
    lblStatus.Caption = "SCANNER USER SETTINGS LOADED..."
    
    
    
    DoEvents
    
    Screen.MousePointer = vbNormal
    
    
''''''    '*** Show Selected Scanner Information
''''''    If Not bolCancelPendingXfers Then
''''''        subShowScannerInfo
''''''    End If
    
    
    '*** SET THE SCANNER SETTINGS PROFILE "DEFAULT BATCH DIRECTORY" from the System Config
    '    if it's blank!
    If Trim(cmbBatchScanSettingsName.Text = "") Then
        lblStatus.Caption = "cmbBatchScanSettingsName.Text = blank : GET DEFAULT BATCH DIRECTORY..."
        txtBatchRootDirectory = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & txtApplicationRECID, "RootDirectoryPathForBatches")
    End If
    
    
    
    'Make main objects visible again
    SSTab1.Visible = True
    frameImageLayout.Visible = False
    
    '1. GET Last Used Batch Queue
    lblStatus.Caption = "GET Last Used Batch Queue..."
    cmbBatchQueue.Text = funcGetSetUserSettings("GET", "DefaultScanBatchQueue", "")

    cmdScanBegin.Visible = True
    If cmbBatchQueue.Visible = True And cmbBatchQueue = "" Then
        lblStatus.Caption = "Set Focus on BATCH QUEUE..."
        cmbBatchQueue.SetFocus
    Else
        '1. Set Focus on BATCH NAME
        lblStatus.Caption = "Set Focus on BATCH NAME..."
        txtBatchName.SetFocus
    End If
    
    
    
Exit Sub

ERROR_HANDLER:

'    MsgBox "Form_Load ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    
    lblStatus.Caption = "ERROR_HANDLER: Error #" & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    
    SSTab1.Visible = True
    frameImageLayout.Visible = False

    Resume Next
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save Form settings to the registry
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanMainPix.Top", Imaging101ScanMainPix.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanMainPix.Left", Imaging101ScanMainPix.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanMainPix.Width", Imaging101ScanMainPix.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanMainPix.Height", Imaging101ScanMainPix.Height, RegFileName)
    End If

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

'''' Make sure the Viewer form closes when the main for closes
'''Private Sub Form_Terminate()
'''    Unload Imaging101ScanViewer
'''End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If cmdScanStop.Visible = True Then
''        cmdScanStop_Click
'        MsgBox "Please STOP the Scanning before exiting!", vbInformation
'        Cancel = True
'        Exit Sub
'    End If
    
    
    If funcIsFormLoaded2("Imaging101ScanViewer") Then
        'Close the ezPageView Control
        Imaging101ScanViewer.ezPageView.Close
        
        Unload Imaging101ScanViewer
        Set Imaging101ScanViewer = Nothing
    End If
    
    If funcIsFormLoaded2("frmLookupList") Then
        Unload frmLookupList
        Set frmLookupList = Nothing
    End If

    
    'Reset Boolean flag
    bolBatchScanningModule = False

    frmImaging101BatchList.Show
    frmImaging101BatchList.subListBatches
    
End Sub



Private Sub Misc_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub lblStatus_Change()
    
    funcWriteToDebugLog Me.name, lblStatus.Caption
    
End Sub

'''Public Sub funcWriteToDebugLog Me.name,(strMessage As String)
'''
'''    If bolDebug Then
'''        Open "ScanStatus.log" For Append As #1
'''        Print #1, Now() & ": " & strMessage
'''        Close #1
'''        DoEvents
'''    End If
'''
'''End Sub







Private Sub SSTab1_GotFocus()

    If Not bolCancelPendingXfers Then
        subShowScannerInfo
    End If

End Sub



Public Sub subPostScan(pixFileName As String)

    On Error GoTo ERROR_HANDLER
    

    ' Send the image to the Viewer form and save to file if requested

    Dim Temp As Long
    Dim strFileName As String
    Dim strFileExtension As String
    Dim strFullBatchDirectory As String
    Dim strBatchPageRECID As Double
    Dim intPageCount As Integer
    
    m_ImageCount = m_ImageCount + 1
    m_PageCount = m_PageCount + 1
    m_ImageSkipCount = m_ImageSkipCount + 1
    
    
    
    ' Save Image if requested
    If ItoB(chkScanPreviewOnly.Value) = False Then
    
''''        '*** 8/27/2004 Moved Inside the If so empty batches are not created on Preview
''''        If bolBatchCreated = False Then
''''            ' ***  CREATE BATCH HEADER RECORD
''''            subCreateBatchRecord
''''        End If
    
        
        strFullBatchDirectory = Trim(txtBatchRootDirectory) & "\" & Format(txtBatchRECID, "0000000000")

''        funcCreateDirectoryStructure strFullBatchDirectory
        
''''        strFilename = Format(txtBatchRECID, "000000000") & "-" & Format(m_PageCount, "000000000")
        intPageCount = m_PageCount
        
        
'        'SPECIAL TWEAK... to handle Multi-Stream TEST scanning
'        If UCase(Left(pixFileName, 8)) = "C:\TEMP\" Then
'            strFullBatchDirectory = "C:\TEMP\"
'        End If
            
            
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  Imaging101ScanViewer.ezPageView.ScanFileDir   = " & Imaging101ScanViewer.ezPageView.ScanFileDir
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  Imaging101ScanViewer.ezPageView.ScanFileName  = " & Imaging101ScanViewer.ezPageView.ScanFileName
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  Imaging101ScanViewer.ezPageView.ScanFileExt   = " & Imaging101ScanViewer.ezPageView.ScanFileExt
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  Imaging101ScanViewer.ezPageView.ScanFileRoot  = " & Imaging101ScanViewer.ezPageView.ScanFileRoot
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  Imaging101ScanViewer.ezPageView.ScanFileSchema= " & Imaging101ScanViewer.ezPageView.ScanFileSchema
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  subPostScan(pixFileName As String)            = " & pixFileName
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  strFullBatchDirectory                         = " & strFullBatchDirectory
        
            
        strFileName = Right(pixFileName, Len(pixFileName) - Len(strFullBatchDirectory) - 1)
        
        strFileExtension = Imaging101ScanViewer.ezPageView.ScanFileExt

        
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  strFileName             = " & strFileName
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  strFileExtension        = " & strFileExtension
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subPostScan -  PixEzScanControlFileType= " & PixEzScanControlFileType
        
        
    
        
        'Create the Batch Page Record - Pass the filename as a parameter
        subCreateBatchPageRecord strFileName, intPageCount
    
    Else
        strFullBatchDirectory = "C:"
        strFileName = "IMGPREVIEWTEMP.TIF"
'''        TwainPRO.SaveTIFCompression = TWTIF_CCITTFAX4
'''        TwainPRO.SaveFile strFullBatchDirectory & "\" & strFileName
    End If



    If ItoB(chkScanPreviewOnly.Value) = True Then
'        Kill strFullBatchDirectory & "\" & strFilename
    End If

Exit Sub

ERROR_HANDLER:

'    MsgBox "chkBatchAutoName ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
'    GetError
    Resume Next
    
End Sub


Private Sub subToggleScanButtons()
    If cmdScanBegin.Visible Then
        cmdScanBegin.Visible = False
        cmdScanStop.Visible = True
        SSTab1.Enabled = False
    Else
        cmdScanBegin.Visible = True
        cmdScanStop.Visible = False
        SSTab1.Enabled = True
    End If
End Sub


Private Sub subCreateBatchRecord()
    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
        On Error GoTo CREATE_BATCH_RECORD_ERROR
        
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subCreateBatchRecord - BEFORE"

        
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
        rsImaging101Batch("BatchInQueueDate") = Now()
        
        '2014-04-25 - Jacob - Added "S" lock to prevent listing in Batch List if Scanning or Splitting
        rsImaging101Batch.Fields("BatchLocked") = "S"
        rsImaging101Batch.Fields("BatchLockedBy") = gsecUserName
        rsImaging101Batch.Fields("BatchLockedDate") = Now()


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

    funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subCreateBatchRecord - AFTER"
    
    '*** CREATE BATCH AUDIT RECORD
    funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Scan Batch"
    
    funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subCreateBatchRecord - AFTER funcCreateBatchAuditRecord"
    
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
        
        funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subCreateBatchPageRecord - BEFORE"
        
        'Position the cursor on the rowset
        txtActionBeforeError = "Open Batches Table"
        Set rsImaging101Batch = New ADODB.Recordset
        '*** Prepare Result Set
        With rsImaging101Batch
            .ActiveConnection = connImaging101Batch
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LOCKTYPE = adLockOptimistic
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

    funcWriteToDebugLog Me.name, "Imaging101ScanMainPix - subCreateBatchPageRecord - AFTER"

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


Private Sub subShowScannerInfo()
    
    lblStatus.Caption = "ENTER subShowScannerInfo"
    
    '*** This Section is only trying to force an error if the scanner is not found!
    bolCancelPendingXfers = False
    
    On Error GoTo NoScanner
    
'    Call Imaging101ScanViewer.ezPageView.ScanStateWrite(App.Path & "\ScanTest.INI", Imaging101ScanViewer.ezPageView.ScanDriverName)
'    Call Imaging101ScanViewer.ezPageView.ScanStateRead(App.Path & "\ScanTest.INI", Imaging101ScanViewer.ezPageView.ScanDriverName)
    '*** END force Scanner error
    
    strScannerSettingsFileName = App.Path & "\ScanSet_" & cmbBatchScanSettingsName & ".INI"

    txtSelectedScannerName.Text = Imaging101ScanViewer.ezPageView.ScanName
    txtSelectedScannerDriverName.Text = Imaging101ScanViewer.ezPageView.ScanDriverName
    txtSelectedScannerDriverVersion.Text = Imaging101ScanViewer.ezPageView.ScanDriverVersion
    txtSelectedScannerDriverLevel.Text = Imaging101ScanViewer.ezPageView.ScanDriverLevel

Exit Sub

NoScanner:
    MsgBox "SORRY!!!" & vbCrLf & "I CAN'T FIND THE SCANNER!!!", vbCritical
    bolCancelPendingXfers = True
    SSTab1.Visible = True
    
End Sub




