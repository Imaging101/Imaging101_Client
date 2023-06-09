VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmImaging101BatchList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Batch List - Imaging101"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmBatchList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   Begin MSMask.MaskEdBox mebDateFrom 
      Height          =   285
      Left            =   1155
      TabIndex        =   73
      Top             =   1410
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm-dd-yyyy"
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin ComctlLib.ListView ListViewColumnOrder 
      Height          =   3555
      Left            =   12930
      TabIndex        =   72
      Top             =   4680
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   6271
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox chkListAllBatches 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List Batches for All Users"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9225
      TabIndex        =   64
      Top             =   600
      Width           =   3015
   End
   Begin VB.CheckBox chkListBatchesCommittedFull 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List Batches Committed/Split -FULL"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9225
      TabIndex        =   63
      Top             =   840
      Width           =   3615
   End
   Begin VB.CheckBox chkOpenBatchesInReadOnlyMode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open Batches as READ-ONLY"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9225
      TabIndex        =   62
      Top             =   1080
      Width           =   3495
   End
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11625
      Picture         =   "frmBatchList.frx":0442
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   60
      Top             =   120
      Width           =   1572
   End
   Begin VB.TextBox txtBatchListSelectedCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8235
      TabIndex        =   46
      Top             =   345
      Width           =   735
   End
   Begin VB.TextBox txtBatchListRecordCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7275
      TabIndex        =   45
      Top             =   345
      Width           =   735
   End
   Begin VB.CommandButton cmdRefreshBatches 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1305
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmBatchList.frx":0AD5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   810
      Width           =   615
   End
   Begin VB.Frame frameButtons 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   33
      Top             =   7680
      Width           =   12735
      Begin VB.CheckBox chkCommitWithLookup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Commit with Lookup"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10800
         TabIndex        =   69
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkShowBatchProperties 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Batch Properties"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10800
         TabIndex        =   43
         Top             =   120
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkCommitAfterBarcode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Commit after Barcode"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10800
         TabIndex        =   42
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdBarcodeSelectedBatches 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ba&rcode Selected  Batches"
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
         Left            =   5340
         Picture         =   "frmBatchList.frx":105F
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdImportBatchFromFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Import Batch From Directory"
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
         Left            =   2670
         Picture         =   "frmBatchList.frx":1929
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdRouteSelectedBatches 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ro&ute Selected Batches"
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
         Left            =   8010
         Picture         =   "frmBatchList.frx":21F3
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdFindNextAvailableBatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Find Next Available Batch"
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
         Left            =   9345
         Picture         =   "frmBatchList.frx":2635
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpenBatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Open / Index Batch"
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
         Height          =   975
         Left            =   1335
         Picture         =   "frmBatchList.frx":38A7
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdScanDocuments 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Scan Documents"
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
         Left            =   4005
         MaskColor       =   &H00E0E9EF&
         Picture         =   "frmBatchList.frx":4171
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdCommitSelectedBatches 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Commit Selected Batches"
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
         Left            =   6675
         Picture         =   "frmBatchList.frx":4A3B
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeleteSelectedBatches 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete Selected  Batches"
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
         Left            =   0
         Picture         =   "frmBatchList.frx":5305
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdImportEcaptureBatch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Import Batch From &eCapture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdResetToDefaults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1905
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmBatchList.frx":5747
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   810
      Width           =   615
   End
   Begin VB.ComboBox cmbBatchListOrder 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmBatchList.frx":5AD1
      Left            =   9225
      List            =   "frmBatchList.frx":5AD3
      TabIndex        =   26
      Top             =   1680
      Width           =   3615
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
      Height          =   345
      Left            =   5520
      TabIndex        =   9
      Top             =   1680
      Width           =   2775
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
      Height          =   345
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ComboBox cmbApplicationList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   105
      TabIndex        =   4
      Top             =   345
      Width           =   4260
   End
   Begin VB.TextBox txtBatchFilter 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   1170
      Width           =   2415
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
      Left            =   3675
      TabIndex        =   1
      Top             =   105
      Width           =   615
   End
   Begin MSAdodcLib.Adodc AdodcImaging101BatchList 
      Height          =   375
      Left            =   3240
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcEcaptureBatchList"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcBatchList 
      Height          =   375
      Left            =   6480
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcBatchList"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcBatchControl 
      Height          =   375
      Left            =   9120
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcBatchControl"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frameBatchDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   0
      TabIndex        =   13
      Top             =   5040
      Width           =   12855
      Begin VB.TextBox txtBatchManager 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtBatchDate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtBatchNotes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   810
         Width           =   5295
      End
      Begin VB.TextBox txtBatchDesc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   525
         Width           =   5295
      End
      Begin VB.TextBox txtBatchName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtBatchStatus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2235
         Width           =   2415
      End
      Begin VB.TextBox txtBatchPriority 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1950
         Width           =   2415
      End
      Begin VB.TextBox txtBatchOwner 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1425
         Width           =   2415
      End
      Begin VB.TextBox txtBatchQueue 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1140
         Width           =   2415
      End
      Begin VB.TextBox txtBatchDirectory 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2040
         Width           =   5295
      End
      Begin VB.TextBox txtBatchCommitStatus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1575
         Width           =   3135
      End
      Begin VB.TextBox txtBatchPagesTotal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtBatchPagesIndexed 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   810
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtBatchPagesCommitted 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   525
         Width           =   975
      End
      Begin VB.TextBox txtBatchRECID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBatchGroup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   525
         Width           =   2415
      End
      Begin VB.TextBox txtBatchBoxNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   810
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   71
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBatchDate 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   68
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblBatchQueue 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Queue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   59
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lblBatchOwner 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Owner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   58
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FEDCC7&
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Committed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   56
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Indexed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   54
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lblBatchStatus 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   53
         Top             =   2235
         Width           =   1215
      End
      Begin VB.Label lblBatchGroup 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   52
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label lblBatchBoxNumber 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Box #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   51
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label lblBatchPriority 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   48
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E3CC6C&
         BackStyle       =   0  'Transparent
         Caption         =   "Commit Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblBatchDirectory 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Directory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmBatchList.frx":5AD5
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         ToolTipText     =   "DOUBLE-CLICK to Open Directory in Windows Explorer"
         Top             =   2040
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   5318
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox mebDateTo 
      Height          =   285
      Left            =   1170
      TabIndex        =   74
      Top             =   1710
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm-dd-yyyy"
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                TO"
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
      Left            =   90
      TabIndex        =   76
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dates FROM"
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
      Left            =   105
      TabIndex        =   75
      Top             =   1470
      Width           =   945
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   11625
      TabIndex        =   61
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected"
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
      Left            =   8265
      TabIndex        =   47
      Top             =   105
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Items"
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
      Left            =   7290
      TabIndex        =   44
      Top             =   105
      Width           =   735
   End
   Begin VB.Label lblBatchListOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch &List Order"
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
      Left            =   9225
      TabIndex        =   27
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblBatchQueueDropDown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch &Queue"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblBatchOwnerDropDown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch &User"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblApplicationCommitBatchTo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4500
      TabIndex        =   7
      Top             =   345
      Width           =   2715
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch &Find"
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
      Left            =   105
      TabIndex        =   6
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label lblSelectApplication 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select &Application"
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
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Commit to:"
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
      Left            =   4530
      TabIndex        =   8
      Top             =   90
      Width           =   1695
   End
End
Attribute VB_Name = "frmImaging101BatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '****************************
    '*** Declarations
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ssql As String
    Dim cmd As ADODB.Command
    Dim strBatchListColumnPositionList As String
    
    




Private Sub About_Click()
    frmAbout.Show
End Sub

Private Sub chkCommitAfterBarcode_Click()

    'Only save the setting if the control is visible... which means the user clicked it
    ' this will avoid commits if the user does not have Commit rights.
    If chkCommitAfterBarcode.Visible Then
        funcGetSetUserSettings "SET", "CommitAfterBarcode", chkCommitAfterBarcode.Value
    End If
    
End Sub

Private Sub chkListAllBatches_Click()

    'Refresh the batch list
    subListBatches

End Sub

Private Sub chkListBatchesCommittedFull_Click()

    'Refresh the batch list
    subListBatches


End Sub

Private Sub chkOpenBatchesInReadOnlyMode_Click()

    '*** 2020-05-22 - Jacob - Added check for chkOpenBatchesInReadOnlyMode to prevent "Commit" of selected batches when checked
    If chkOpenBatchesInReadOnlyMode = vbChecked Then
        cmdCommitSelectedBatches.Enabled = False
        cmdBarcodeSelectedBatches.Enabled = False
    Else
        cmdCommitSelectedBatches.Enabled = True
        cmdBarcodeSelectedBatches.Enabled = True
    End If

End Sub

Private Sub chkShowBatchProperties_Click()

    If chkShowBatchProperties.Value = vbUnchecked Then
        frameBatchDetail.Visible = False
        Form_Resize
    Else
        frameBatchDetail.Visible = True
        Form_Resize
    End If
    
    funcGetSetUserSettings "SET", "ShowBatchProperties", chkShowBatchProperties.Value
    
End Sub

Public Sub cmbApplicationList_Click()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  cmbApplicationList_Click()"


    'Store the selected application
    funcGetSetUserSettings "SET", "Application", cmbApplicationList

    
    ' Get the Application to Commit Batches to
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
        
    rs.Source = "Select * from I101Applications WHERE ApplicationName= '" & cmbApplicationList.Text & "'"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If Not (rs.EOF Or rs.BOF) Then
        txtApplicationRECID = rs!ApplicationRECID
        lblApplicationCommitBatchTo = rs!ApplicationCommitBatchTo & ""
        gAutoAdvanceOnSeparator = rs!ApplicationAutoAdvanceOnSeparator & ""
        gSetUserAsBatchOwnerOnSPLIT = rs!SetUserAsBatchOwnerOnSPLIT & ""
    End If
    
    If lblApplicationCommitBatchTo = "" Then
        MsgBox "No Application has been set to Commit Batches to!", vbInformation, "Missing Commit Batch Setting"
        Exit Sub
    End If
    
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    
    
    '***************************************************************
    '***  GET SECURITY RIGHTS
    
    funcGetSecurityRights gsecSecurityRECID, txtApplicationRECID
    
    subSetBatchButtonSecurity
    
    If bolErrorOccured = True Then
        ListView1.ListItems.Clear
        Exit Sub
    End If
    
    '***************************************************************
    '***  SET THE DEFAULT BATCH QUEUE
    
    cmbBatchQueue.Text = gsecBatchDefaultQueue
    
    
    '****************************************************************
    '*** See if we should display the Barcode Processing button
    cmdBarcodeSelectedBatches.Visible = False
    chkCommitAfterBarcode.Visible = False
    
'    Select Case frmImaging101BatchList.cmbApplicationList.Text
'
'        Case "TTC"
'            cmdBarcodeSelectedBatches.Visible = True
'            chkCommitAfterBarcode.Visible = True
'
'        Case "EVIDENCE"
'            cmdBarcodeSelectedBatches.Visible = True
'            chkCommitAfterBarcode.Visible = True
'
'    End Select

    '*** Validate the Barcode License Key
    bolBarcodeLicenseValidated = funcValidateBarCodeLicense

    If bolBarcodeLicenseValidated Then
        cmdBarcodeSelectedBatches.Visible = True
        chkCommitAfterBarcode.Visible = True
    End If
    

    '*** Only list Batches if NOT here because of an I101AIM Command
    If Not bolAIM_Command Then
        subListBatches
    End If
    
End Sub

Public Sub subListBatches(Optional BatchRECID As String)

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  subListBatches()"


    If Trim(cmbApplicationList.Text) = "" Then
        MsgBox "Please select an APPLICATION !", vbInformation, "No Application Selected"
        Exit Sub
    End If
    

    
    
    '*** 2023-01-03 - Jacob - Moved AutoSize parameters up from List Batches to also handle the "GETTING LIST OF BATCHES" message

    ListView1.ListItems.Clear
    
    Set lstItem = ListView1.ListItems.Add(, , "*** GETTING LIST OF BATCHES.  PLEASE BE PATIENT... THIS MAY TAKE A WHILE. ***")
        
        
    ' AutoSize ALL Columns
    Dim intColumnNumber As Integer
    Dim lparamAutoSize As Long
    
    UseHeader = True
    If UseHeader = False Then
        lparamAutoSize = LVSCW_AUTOSIZE
    Else
        lparamAutoSize = LVSCW_AUTOSIZE_USEHEADER
    End If
    
    For intColumnNumber = 0 To ListView1.ColumnHeaders.Count - 1
        SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, intColumnNumber, ByVal lparamAutoSize
    Next

    DoEvents
        
        
        
     '*** Declarations
''    Dim rs As adodb.Recordset
''    Dim Con As adodb.Connection
''    Dim ssql As String
    Dim strWhere As String

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | subListBatches() | OPEN ADODB.Connection"

    Set con = New ADODB.Connection
    With con
        '*** 2023-01-03 - Jacob - Changed to adXactReadUncommitted
        '                                        and Increased ConnectioTimeout and CommandTimeout to prevent Deadlock errors
        '                                        on slow WAN/LAN connections or with too many Batches.
        .IsolationLevel = adXactCursorStability
        .mode = adModeRead
        .ConnectionTimeout = 600
        .CommandTimeout = 3600
        .Open RegImaging101BatchListConnectionString
    End With
    
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseClient
        '*** 2023-01-03 - Jacob - Changed to adOpenForwardOnly
        .CursorType = adOpenForwardOnly
        .LOCKTYPE = adLockReadOnly
    End With
        
    '4/25/2014 - Jacob - Added  " AND  BatchLocked <> 'S' " to prevent listing Batches in process of being SCANNED
    '*** 2023-01-10 - Jacob - Added  WITH (NOLOCK) to prevent DeadLock Errors when Commits or Splits take too long

    rs.Source = "SELECT BatchRECID, ApplicationRECID, BatchName, BatchScanDate, " & _
                "DATEDIFF(day, BatchScanDate, getdate()) AS AgeDays, " & _
                "DATEDIFF(day, BatchInQueueDate, getdate()) AS DaysInQueue, " & _
                "BatchQueue, BatchGroup, BatchOwner, BatchManager," & _
                "BatchPriority, BatchStatus, " & _
                "BatchPagesTotal, BatchCommitStatus, BatchPagesCommitted, " & _
                "BatchDesc, BatchNotes, " & _
                "BatchDirectory, BatchLocked, BatchLockedBy, BatchLockedDate, " & _
                "BatchApplication, BatchScanUser, BatchQCUser, BatchQCDate, BatchIndexUser, " & _
                "BatchIndexDate, BatchPagesQCAppended, BatchPagesQCInserted, " & _
                "BatchPagesQCDeleted, BatchPagesIndexed, BatchPagesNotCommitted, " & _
                "BatchBoxNumber " & _
                " FROM I101Batches " & _
                " WITH (NOLOCK) " & _
                " WHERE ApplicationRECID = " & txtApplicationRECID & _
                "      AND (BatchLocked <> 'S'  OR  BatchLocked IS NULL) "


            '*** If user entered a Filter, show ONLY items that CONTAIN the phrase
            '    OVERRIDE ALL other WHERE conditions...
            '    Using FILTER will ONLY show FULLY COMMITTED Batches for 24 Hours!
            '    The "DateAdd() function will subtract ONE DAY from TODAY
            
            
            ' 10/10/2013 - Jacob - Modified this section so that the Batch Filter ALSO works if gsecRightsBatchFindRestricted is checked
            If Trim(txtBatchFilter) <> "" Then
            
                strWhere = "  AND  ("
                strWhere = strWhere & "BatchName like '%" & Replace(txtBatchFilter, "*", "%") & "%'  "
                strWhere = strWhere & "            OR  BatchDesc like '%" & Replace(txtBatchFilter, "*", "%") & "%'  "
                strWhere = strWhere & "            OR  BatchNotes like '%" & Replace(txtBatchFilter, "*", "%") & "%'  "
                
                '*** 2020-04-24 - Jacob - Added ability to search by BatchRECID
                If IsNumeric(txtBatchFilter) Then
                    strWhere = strWhere & "        OR  BatchRECID = " & txtBatchFilter
                End If
                
                strWhere = strWhere & " )"

            End If
            
            '*** 2023-01-05 - Jacob - Added Date Range Search
            If IsDate(mebDateFrom) And IsDate(mebDateTo) Then
                    strWhere = strWhere & "  AND ( BatchScanDate BETWEEN '" & mebDateFrom & "' AND '" & mebDateTo & "')"
            End If
                

            
''            If gsecRightsBatchFindRestricted <> vbChecked Then
''
''                ' NO Filters... search ALL Queues and Users
''
''            Else
                
'               '*** CLEAR the Where clause
'               strWhere = ""
'
'                '*** Check if user entered a FILTER
'                If Trim(txtBatchFilter) <> "" Then
'                    strWhere = "  AND  BatchName like '%" & txtBatchFilter & "%' "
'                End If
                    
            
                '*** Check if Find is Restricted
                If gsecRightsBatchFindRestricted = vbChecked Then
            
                    If gsecRightsBatchFindRestrictToQueue = vbChecked Then
                        cmbBatchQueue.Text = gsecBatchDefaultQueue
                    End If
                    
                    If gsecRightsBatchFindRestrictToOwner = vbChecked Then
                        cmbBatchOwner.Text = gsecUserName
                    End If
                
                End If
            
                
                '*** If User did NOT check the List ALL Batches checkbox
                If chkListAllBatches <> vbChecked Then
                    If Trim(cmbBatchOwner) = "" Then
                        '*** Show Batches for the Logged-In User's OR Unassigned Batches
                        strWhere = strWhere & _
                        " AND ( (BatchOwner is NULL) OR (BatchOwner = '') OR (BatchOwner = '" & gsecUserName & "') )"
                    Else
                        '*** Show batches for the Selected user
                        strWhere = strWhere & _
                        " AND ( BatchOwner = '" & Trim(cmbBatchOwner) & "' )"
                    End If
                Else
                    'If a User was selected
                    If Trim(cmbBatchOwner) <> "" Then
                        '*** Show batches for the Selected user
                        strWhere = strWhere & _
                        " AND ( BatchOwner = '" & Trim(cmbBatchOwner) & "' )"
                    End If
                
                    
                End If
                
                '*** If a Queue is selected, show only items in the Selected Queue
                If Trim(cmbBatchQueue) <> "" Then
                    strWhere = strWhere & _
                    " AND (BatchQueue LIKE '" & Trim(cmbBatchQueue) & "') "
                End If
                
                
                
    '            If cmbUserGroupList <> "" Then
    '                rs.Source = rs.Source & _
    '                " AND (BatchGroup like'" & cmbUserGroupList & "%') "
    '            End If
                
                
''            End If
            
            
            
            '*** Check if user Checked the "List Batches Committed-FULL" checkbox
            If chkListBatchesCommittedFull <> vbChecked Then
                '*** Don't show Fully-Committed or Fully-Split Batches
                strWhere = strWhere & " AND ((BatchCommitStatus NOT LIKE 'Committed-FULL%') OR (BatchCommitStatus is null)) "
                strWhere = strWhere & " AND ((BatchCommitStatus + '' NOT LIKE 'Split-FULL%') OR (BatchCommitStatus is null))"
                strWhere = strWhere & " AND ((BatchCommitStatus + '' NOT LIKE 'Updated-FULL%') OR (BatchCommitStatus is null))"
            Else
                '*** Show Fully-Committed or Fully-Split Batches
                strWhere = strWhere & " AND ((BatchCommitStatus  LIKE 'Committed-FULL%') "
                strWhere = strWhere & " OR  (BatchCommitStatus  LIKE 'Split-FULL%')"
                strWhere = strWhere & " OR  (BatchCommitStatus  LIKE 'Updated-FULL%'))"
            End If
            
            '*** If the Optional "BatchRECID" is passed to this routing
            '***    OVERRIDE ALL other WHERE conditions INCLUDING ANY FILTER  ***
            If BatchRECID <> "" Then
                strWhere = " AND BatchRECID = " & BatchRECID
            End If
    
            '*** Set the Default Batch Sort Order if none selected
            If Trim(cmbBatchListOrder.Text) = "" Then
                cmbBatchListOrder.Text = " BatchName "
            End If
    
            rs.Source = rs.Source & strWhere & " ORDER BY " & cmbBatchListOrder.Text
            
 
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenForwardOnly
'    rs.LockType = adLockOptimistic
    
''    On Error GoTo ERROR_TRAP
    
    'sql statement to select items on the drop down list
''    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID
''    rs.Open ssql, Con

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | subListBatches() | OPEN RS | " & rs.Source

    con.Errors.Clear
    rs.Open
        
    
   
    '*** Setup Up ListView properties - BEGIN
    
    '*** 2023-01-03 - Jacob - Added ListView1.ListItems.Clear to clear the "GETTING LIST OF BATCHES" message

    ListView1.ListItems.Clear

    ListView1.Visible = False
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = True

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | subListBatches() | SET COLUMN HEADINGS"
    
        '***  SET COLUMN HEADINGS
        For intListIndex = 0 To rs.Fields.Count - 1
            ListView1.ColumnHeaders.Add , , rs.Fields.item(intListIndex).name, Len(rs.Fields.item(intListIndex).name) * 150, lvwColumnLeft
        Next
                
    On Error Resume Next
    
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    While Not rs.EOF
            For intListIndex = 0 To rs.Fields.Count - 1
                If intListIndex = 0 Then
                    If Not IsNull(rs.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , rs.Fields.item(intListIndex).Value)
                    End If
                Else
            
                        '* This null check is to make sure we don't Skip fields caused by an error.
                        If Not IsNull(rs.Fields.item(intListIndex).Value) Then
                            ' Not null... show value
                            Select Case rs.Fields.item(intListIndex).Type
                                Case adDBTimeStamp
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Format(rs.Fields.item(intListIndex).Value, "yyyy/mm/dd AMPM hh:mm:ss "))
                                Case adInteger
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("      " & Format(rs.Fields.item(intListIndex).Value, "##,###"), 6))
                                Case adNumeric, adDouble
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("          " & Format(rs.Fields.item(intListIndex).Value, "##,###,###"), 10))
                                Case adCurrency
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("              " & Format(rs.Fields.item(intListIndex).Value, "$##,###,###.##"), 14))
                                Case Else
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , rs.Fields.item(intListIndex).Value)
                            End Select
                            
'                            If rs.Fields.item(intListIndex).Type = adDBTimeStamp Then
'                                Set lstSubItem = lstItem.ListSubItems.Add(, , Format(rs.Fields.item(intListIndex).Value, "yyyy/mm/dd AMPM hh:mm:ss "))
'                            Else
'                                Set lstSubItem = lstItem.ListSubItems.Add(, , rs.Fields.item(intListIndex).Value)
'                            End If
                        Else
                            ' Null... show empty string
                            Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                        End If
                        
                        If rs.Fields.item(intListIndex).name = "BatchNotes" Then
                            lstItem.ListSubItems(intListIndex).ForeColor = vbRed
                       End If
                End If
            Next
        rs.MoveNext
    Wend
    On Error GoTo 0
    
    
    
    
    '*** 2023-01-03 - Moved DIM statements and IF logic to top to also handle the "LOADING" message, and added DoEvents
    
    ' AutoSize ALL Columns

    For intColumnNumber = 0 To ListView1.ColumnHeaders.Count - 1
        SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, intColumnNumber, ByVal lparamAutoSize
    Next
    
    DoEvents

    
    
    ' Hide the RecordID's
    ListView1.ColumnHeaders(1).width = 0
    ListView1.ColumnHeaders(2).width = 0

    ' Size the Key fields to a standard size
    '*** 2020-04-24 - Jacob - Commented out the ColumnHeaders(3) default size to allow Auto-Size to work for BatchName
    'ListView1.ColumnHeaders(3).width = 3000
''    ListView1.ColumnHeaders(4).Width = 2000
''    ListView1.ColumnHeaders(5).Width = 1000
''    ListView1.ColumnHeaders(6).Width = 1000
''    ListView1.ColumnHeaders(7).Width = 1000

    
    ListView1.Visible = True
    
    
     'Moved the Refresh to  funcLoadBatchListViewColumnPositions
     'the columns have changed, but the control needs
     'to redisplay its contents in the new order
'      ListView1.Refresh
    
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | subListBatches() | funcLoadBatchListViewColumnPositions"
    
    funcLoadBatchListViewColumnPositions
    
    
    '*** Setup Up ListView properties - END
    
    '*** Set SQL wildcard string
'    RegConnectionWildcard = "%"
        
    
    '*** Connect to DB for Drop Down Lists
    
    txtApplicationRECID = cmbApplicationList.ItemData(cmbApplicationList.ListIndex)
    txtApplicationName = cmbApplicationList.Text
    txtBatchListRecordCount = rs.RecordCount
    
    rs.Close
    con.Close
    
    ' Disable Buttons until at least ONE Batch is selected these and added subSetBatchButtonSecurity
    cmdCommitSelectedBatches.Enabled = False
    cmdBarcodeSelectedBatches.Enabled = False
    cmdDeleteSelectedBatches.Enabled = False
    cmdOpenBatch.Enabled = False
    cmdRouteSelectedBatches.Enabled = False
    

End Sub



Private Sub cmbUserGroupList_Click()

        cmdRefreshBatches_Click

End Sub

Private Sub cmbApplicationList_KeyPress(KeyAscii As Integer)

    '*** Don't want user to key in an invalid value
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0

End Sub

Private Sub cmbBatchListOrder_Click()
    ' Refresh the Batch List
    subListBatches

End Sub

Private Sub cmbBatchOwner_Click()
    ' Refresh the Batch List
    subListBatches

End Sub

Private Sub cmbBatchOwner_DropDown()

    '***************************************
    '*** LOAD USERS LIST DROP-DOWN
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101Security ORDER BY UserName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    rs.Open
    
    txtActionBeforeError = "Populate UserID List"
    
    cmbBatchOwner.Clear
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbBatchOwner.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchOwner.AddItem rs.Fields("UserName")
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

'    '***************************************
'    '*** LOAD USERS LIST DROP-DOWN
'
'    Set Con = New ADODB.Connection
'    Con.Open RegImaging101ConnectionString
'
'    Set rs = New ADODB.Recordset
'    Set rs.ActiveConnection = Con
'
'    rs.Source = "SELECT DISTINCT BatchOwner FROM I101Batches " & _
'                " WHERE BatchOwner IS NOT NULL " & _
'                "   AND BatchOwner <> '' " & _
'                " ORDER BY BatchOwner "
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenDynamic
'    rs.LockType = adLockReadOnly
'
'    Con.Errors.Clear
'    rs.Open
'
'    txtActionBeforeError = "Populate UserID List"
'
'    cmbBatchOwner.Clear
'    'Add a Blank value to allow clearing the BatchOwner
'    cmbBatchOwner.AddItem ""
'    For intIndex = 0 To rs.RecordCount - 1
'        cmbBatchOwner.AddItem rs.Fields("BatchOwner") & ""
'        rs.MoveNext
'    Next
'
'    'Close connection and the recordset
'    rs.Close
'    Set rs = Nothing
'    Con.Close
'    Set Con = Nothing
'
'    '****************************

End Sub

Private Sub cmbBatchOwner_KeyPress(KeyAscii As Integer)
    
    'Carriage Return = Asc(13)
    If KeyAscii = Asc(vbCr) Then
        subListBatches
    End If
   
End Sub

Private Sub cmbBatchQueue_Click()
    
    '****************************
    ' Refresh the Batch List
    subListBatches


End Sub

Private Sub cmbBatchQueue_DropDown()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  cmbBatchQueue_DropDown()"


    '***************************************
    '*** LOAD BATCH QUEUES LIST DROP-DOWN
        
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "SELECT * from I101BatchQueues " & _
                " WHERE ApplicationRECID = " & txtApplicationRECID & _
                " OR (ApplicationRECID = 0 OR ApplicationRECID IS NULL)  " & _
                " AND (BatchQueueActive = 'Y')" & _
                " ORDER BY BatchQueue "
                
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    cmbBatchQueue.Clear

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


'    '***************************************
'    '*** LOAD BATCH QUEUES LIST DROP-DOWN
'
'    Set Con = New ADODB.Connection
'    Con.Open RegImaging101ConnectionString
'
'    Set rs = New ADODB.Recordset
'    Set rs.ActiveConnection = Con
'
'    rs.Source = "SELECT DISTINCT BatchQueue FROM I101Batches " & _
'                " WHERE BatchQueue IS NOT NULL " & _
'                "   AND BatchQueue <> '' " & _
'                " ORDER BY BatchQueue"
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenDynamic
'    rs.LockType = adLockReadOnly
'
'    Con.Errors.Clear
'
'    rs.Open
''    rs.MoveFirst
'
'    cmbBatchQueue.Clear
'    'Add a Blank value to allow clearing the BatchOwner
'    cmbBatchQueue.AddItem ""
'    For intIndex = 0 To rs.RecordCount - 1
'        cmbBatchQueue.AddItem rs.Fields!BatchQueue
'        rs.MoveNext
'    Next
'
'    'Close connection and the recordset
'    rs.Close
'    Set rs = Nothing
'    Con.Close
'    Set Con = Nothing

End Sub

Private Sub cmbBatchQueue_KeyPress(KeyAscii As Integer)

    'Carriage Return = Asc(13)
    If KeyAscii = Asc(vbCr) Then
        subListBatches
    End If

End Sub

Private Sub cmdCommitSelectedBatches_Click()

    If cmdCommitSelectedBatches.Visible = True Then
        cmdCommitSelectedBatches.Enabled = False
    End If

    For i = 1 To ListView1.ListItems.Count
        If frmImaging101BatchList.ListView1.ListItems(i).Selected = True Then
            funcWriteToDebugLog Me.name, frmImaging101BatchList.ListView1.ListItems(i).Text
            frmImaging101BatchList.ListView1.ListItems(i).Selected = True   ' Force item selection
            ' Turn ON Global flag to signal modules that we are committing all selected batches
            blnCommitSelectedBatches = True
            ' Turn ON Global flag to signal modules that we will WAIT till the current Batch is Committed
            blnCommitSelectedBatchesWait = True
            ' Select the Batch to commit
            subGetBatchHeaderInfo
            ' Open the Batch
            cmdOpenBatch_Click
''            ' Execute the Form_Load section
''            frmIndex.Form_Load

            If chkCommitWithLookup = 1 Then
                frmIndex.cmdBookMark_Set True, True  'NoPrompt, BookMarkAllPages
            End If
                
            'Commit the Batch
            frmIndex.cmdCommitBatch_Click
            
            ' Dummy Loop -- Just spin around while the commit finishes
            While blnCommitSelectedBatchesWait = True
                DoEvents
            Wend
            
            ' Unload unnecessary forms
            Unload frmCommitStatus
            Unload frmIndex
        End If
    Next
    ' Refresh the Batch List
    subListBatches
    blnCommitSelectedBatches = False
    
    If cmdCommitSelectedBatches.Visible = False And gsecRightsBatchCommit = vbChecked Then
        cmdCommitSelectedBatches.Enabled = True
    End If


End Sub


Private Sub cmdDeleteSelectedBatches_Click()
        
    ' Exit if the ListView is Empty
    If frmImaging101BatchList.ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    On Error GoTo BATCH_DELETE_ERRORS
    
'    If cmdDeleteSelectedBatches.Visible = True Then
        cmdDeleteSelectedBatches.Enabled = False
'    End If
        
    result = MsgBox("Are you SURE you wish to DELETE ALL of the SELECTED Batches and their related Images?", vbYesNo)
    If result = vbNo Then
        Exit Sub
    End If
    
    
    Dim i As Double
    
    For i = 1 To ListView1.ListItems.Count
    
        
        If frmImaging101BatchList.ListView1.ListItems(i).Selected = True Then
        
            funcWriteToDebugLog Me.name, frmImaging101BatchList.ListView1.ListItems(i).Text
            frmImaging101BatchList.ListView1.ListItems(i).Selected = True   ' Force item selection
            
            ' Select the Batch to commit
            subGetBatchHeaderInfo
            
'            result = MsgBox("Are you SURE you wish to DELETE Batch '" & txtBatchName & "' (BatchID # " & txtBatchRECID & ") and it's related Pages?", vbYesNo)
'
'            If result = vbYes Then
            
    
                '*** CREATE BATCH AUDIT RECORD
                funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Delete Batch"
    
            
                ' LOCK the Batch
                strReturn = frmImaging101Winsock.funcSendData("LOCK BATCH" & "|" & txtBatchRECID)
                
                If Left(strReturn, 5) = "ERROR" Then
                    MsgBox strReturn & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Server Communication Failure"
'                    Exit Sub
                    result = vbNo
                Else
                    'No Errors Locking Batch - Go ahead and Zap the Batch
                    
                    Set conn = New ADODB.Connection
                    Set cmd = New ADODB.Command
                    
                    '4/20/2012 - Jacob - Added CommandTimeout
                    cmd.CommandTimeout = 600
                    
                    txtActionBeforeError = "Prepare to Open Batch DB Connection"
                    With conn
                        .ConnectionString = RegImaging101BatchListConnectionString
                        .ConnectionTimeout = 120
                        .IsolationLevel = adXactReadCommitted
                        .mode = adModeReadWrite
                        .Open
                    End With
                    
                    Set cmd.ActiveConnection = conn
            
                    conn.BeginTrans
                    
                        '''' *** 2004-01-25 - JACOB - Disabled Delete Batch Pages records
                        ''''                          Because together with the BatchAudit Table
                        ''''                          We can create any reports we need.
''''                    ' Delete from Batch Pages Table
''''                    txtActionBeforeError = "Delete from Batch Pages Table"
''''                    cmd.CommandText = "DELETE FROM " & cmbApplicationList.Text & "_BatchPage WHERE BatchRECID = " & txtBatchRECID
''''                    cmd.Execute , , adCmdText
            
                    ' Delete FROM I101Batches Table
                    txtActionBeforeError = "Delete FROM I101Batches Table"
                    cmd.CommandText = "DELETE FROM I101Batches WHERE BatchRECID = " & txtBatchRECID
                    cmd.Execute , , adCmdText
            
                    txtActionBeforeError = "Commit Delete Transactions"
                    conn.CommitTrans
                    
                    conn.Close
                    Set cmd = Nothing
                    Set conn = Nothing
                    
                    ' Ignore Errors related to Directories Not Existing, etc.
                    
                    On Error Resume Next
                    ' Delete Files & Folder if there were NO SQL Errors
                    ' The DeleteFolder method does not distinguish between folders that have contents
                    '    and those that do not.
                    ' The specified folder is deleted regardless of whether or not it has contents.
                    Dim ofs As Scripting.FileSystemObject
                    Set ofs = New Scripting.FileSystemObject
                    ofs.DeleteFolder txtBatchDirectory
                    Set ofs = Nothing
                    
                    On Error GoTo BATCH_DELETE_ERRORS
                    
                    Screen.MousePointer = vbDefault
                
                    ''MsgBox "DELETE of Batch '" & txtBatchName & "' (BatchID # " & txtBatchRECID & ") was SUCCESSFUL!", vbOKOnly
                    
                End If  'Left(strReturn, 5) = "ERROR"
                
'            End If  'result = vbYes - Prompt to Delete Each Batch
        
        End If 'Selected
        
    Next
    
    'Refresh the Batch ListView
    subListBatches
    
'    If cmdDeleteSelectedBatches.Visible = False And gsecRightsDeleteBatches = vbChecked Then
        cmdDeleteSelectedBatches.Enabled = True
'    End If
    
    Exit Sub
    
BATCH_DELETE_ERRORS:
        MsgBox "Batch DELETE ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch NOT Imported]", vbExclamation
        
        If conn.BeginTrans = True Then
            conn.RollbackTrans
        End If

        Screen.MousePointer = vbDefault
        
End Sub

Private Sub cmdFindNextAvailableBatch_Click()

    Dim strReturn As String
        
    strReturn = "0"
        
    'Loop until we find an Available Batch or user Cancels
    While strReturn = "0"
        
        strReturn = frmImaging101Winsock.funcSendData("GET NEXT AVAILABLE BATCH" & "|" & txtApplicationRECID & "|" & cmbBatchQueue & "|" & cmbBatchListOrder)
        
        '*** Changed the Code to INSTR() because there is a problem in the "GET NEXT AVAILABLE BATCH" code
        '    strReturn is set to:
        '      "2/25/2004 2:24:17 AM|I101_Server|PCNJACOB|192.168.1.101|ERROR: | funcParseCommand |0|||"
        '     apparently when the function does NOT find an available batch
        If InStr(1, strReturn, "ERROR:") > 0 Then
            MsgBox strReturn & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Server Communication Failure"
            Exit Sub
'            frmImaging101Winsock.cmdClose_Click
        End If
            
        If Trim(strReturn) = "0" Then
            strReturn = MsgBox("SORRY!  No Batches are Available at this time!", vbRetryCancel, "Find Next Available Batch")
            If strReturn = vbRetry Then
                strReturn = "0"
            Else
                Exit Sub
            End If
         End If
    Wend
        
    ' Re-Submit Search with BatchRECID
    subListBatches strReturn

    ' Select the First Row - We KNOW it exists because we JUST Locked it!
    frmImaging101BatchList.ListView1.ListItems(1).Selected = True   ' Force item selection
    ' Select the Batch to commit
    ListView1_Click
    ' Open the Batch
    cmdOpenBatch_Click
            
        
End Sub

Private Sub cmdImportBatchFromFile_Click()
'''    frmImport.Show
    Me.Hide
    frmImportFilesToBatch.Show
    
End Sub

Private Sub cmdImportEcaptureBatch_Click()

    Me.Hide
    frmEcaptureBatchList.Show
    
End Sub

Private Sub cmdMoveSelectedBatches_Click()

    frmImaging101BatchRouteSelected.Show modal, Me

End Sub

Private Sub cmdBarcodeSelectedBatches_Click()

    If cmdCommitSelectedBatches.Visible = True Then
        cmdCommitSelectedBatches.Enabled = False
        cmdBarcodeSelectedBatches.Enabled = False
    End If

    For i = 1 To ListView1.ListItems.Count
        If frmImaging101BatchList.ListView1.ListItems(i).Selected = True Then
            funcWriteToDebugLog Me.name, frmImaging101BatchList.ListView1.ListItems(i).Text
            frmImaging101BatchList.ListView1.ListItems(i).Selected = True   ' Force item selection
            ' Select the Batch to commit
            subGetBatchHeaderInfo
            ' Open the Batch
            cmdOpenBatch_Click
''            ' Execute the Form_Load section
''            frmIndex.Form_Load
            
            ' Turn ON Global flag to signal modules that we are committing all selected batches
            '  Set flag ON whether we are auto-committing  or NOT
            '  because it stops other modules from processing unnecessary code
            '  that could also interfere with the automatic Barcode processing.
            blnCommitSelectedBatches = True

            ' Turn ON Global flag to signal modules that we will WAIT till the current Batch is Barcoded
            blnBarcodeSelectedBatchesWait = True
            frmIndex.cmdProcessBarcodes_Click
            ' Dummy Loop -- Just spin around while the commit finishes
            While blnBarcodeSelectedBatchesWait = True
                DoEvents
            Wend
            
            'See if the user selected to Commit after processing Barcodes
            If chkCommitAfterBarcode = vbChecked Then
                ' Turn ON Global flag to signal modules that we will WAIT till the current Batch is Committed
                blnCommitSelectedBatchesWait = True
                frmIndex.cmdCommitBatch_Click
                ' Dummy Loop -- Just spin around while the commit finishes
                While blnCommitSelectedBatchesWait = True
                    DoEvents
                Wend
                ' Unload unnecessary forms
                Unload frmCommitStatus
           End If
            
            ' Unload unnecessary forms
            Unload frmIndex
            
        End If
    Next
    ' Refresh the Batch List
    subListBatches
    
    blnCommitSelectedBatches = False
    
    If cmdCommitSelectedBatches.Visible = False And gsecRightsBatchCommit = vbChecked Then
            '*** 2020-05-22 - Jacob - Added check for chkOpenBatchesInReadOnlyMode
            If chkOpenBatchesInReadOnlyMode = vbChecked Then
                cmdCommitSelectedBatches.Enabled = False
                cmdBarcodeSelectedBatches.Enabled = False
            Else
                cmdCommitSelectedBatches.Enabled = True
                cmdBarcodeSelectedBatches.Enabled = True
            End If
    End If



End Sub

Private Sub cmdOpenBatch_Click()

    Dim strBatchModeOption As Variant
    Dim strReturn As String
    
    If Trim(txtBatchRECID) = "" Then
        MsgBox "Please select a Batch first!"
        Exit Sub
    End If
    
    'Open in Read-Only Mode if chkOpenBatchesInReadOnlyMode is checked
    ' or if the Batch Committed or Updated FULL
    If chkOpenBatchesInReadOnlyMode = vbChecked _
    Or InStr(UCase(txtBatchCommitStatus), "-FULL") <> 0 _
    Or gsecRightsBatchIndex = False Then
        gOpenBatchInReadOnlyMode = True
        '*** CREATE BATCH AUDIT RECORD
        funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Batch Opened - ReadOnly"
    Else
'        If gsecRightsBatchIndex = True Then
            ' LOCK the Batch
            strReturn = frmImaging101Winsock.funcSendData("LOCK BATCH" & "|" & txtBatchRECID)
            
            If Left(strReturn, 5) = "ERROR" Then
                'The Batch is Locked - ask if Open in READ-ONLY mode
                strBatchModeOption = MsgBox(strReturn & vbCrLf & vbCrLf & "Would you like to open this Batch in READ-ONLY Mode?", vbYesNo, "Error Locking Batch")
                If strBatchModeOption = vbYes Then
                    gOpenBatchInReadOnlyMode = True
                Else
                    gOpenBatchInReadOnlyMode = False
                    'Don't open the Batch
                    Exit Sub
                End If
            Else
                'NO Errors -- Open Batch for Indexing
                gOpenBatchInReadOnlyMode = False
                'Assign the Current User as the New Owner of the Batch
                funcSaveFieldToDB RegImaging101ConnectionString, "I101Batches", "BatchRECID = " & txtBatchRECID, "BatchOwner", gsecUserName
                 '*** CREATE BATCH AUDIT RECORD
                funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Batch Opened - Index"
               
            End If
'        Else
'            'NO Indexing Rights... Open in Read-Only Mode
'            gOpenBatchInReadOnlyMode = True
'            '*** CREATE BATCH AUDIT RECORD
'            funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Batch Opened - ReadOnly"
'        End If
    End If

    Me.Hide
    
    
    
    
    frmIndex.Show
    
    DoEvents
    
    '*** MUST SET THE FOCUS HERE!!!   OTHERWISE the frmDoctypeList STEALS the Focus!
    '*** IF AutoLookupOnBatchLoad IS ENABLED, then set the focus on the INDEX form
    '    so that the Default Field after Lookup is highlighted.
    '    Otherwise, set the focus on the Lookup List to expedite doing a Lookup.
    
    Dim bolAutoLookupOnBatchLoad As Boolean
    bolAutoLookupOnBatchLoad = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationName='" & cmbApplicationList.Text & "'", "AutoLookupOnBatchLoad")
    
    If bolAutoLookupOnBatchLoad Then
    
        frmIndex.SetFocus
        DoEvents
    
    Else
        
        If funcIsFormLoaded2("frmLookupList") = True Then
            If frmLookupList.Visible = True Then
                
                frmLookupList.SetFocus
                
            End If
        End If
    
    End If
    
    
End Sub


Private Sub Command4_Click()

End Sub



Private Sub cmdRefreshBatches_Click()

    'Refresh the batch list
    subListBatches
    cmdOpenBatch.Default = True

End Sub

Private Sub cmdResetToDefaults_Click()
    
    ListView1.Sorted = False
    subSetBatchQueueDefaults
    subListBatches
    
End Sub

Private Sub cmdRouteSelectedBatches_Click()

    frmImaging101BatchRouteSelected.Show modal, Me
    
    

End Sub

Private Sub cmdScanDocuments_Click()
    '** The following sequence is important
    '   so that the MainMenu doesn't Activate
    Me.Enabled = False
    Imaging101ScanMainPix.Show
    Me.Hide
    Me.Enabled = True
End Sub





Private Sub Form_Initialize()

'    MsgBox "Form Initialize"
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING Form_Initialize()"

End Sub


Private Sub Form_Activate()
    
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING Form_Activate()"
    
    Dim i As Integer
    Dim txtApplication As String
    
    
    
    'If the Application Drop Down is empty
    If cmbApplicationList = "" Then
        'Check for Default Application
        If Trim(gsecBatchDefaultApplication) = "" Then
            ' GET The Application this User used last
            txtApplication = funcGetSetUserSettings("GET", "Application", "")
            
            ' If No application has been saved... select the FIRST item in the List
            If Trim(txtApplication) = "" Then
                cmbApplicationList.ListIndex = 0
            End If
        Else
            ' SET the Default Application Assiged for this user in Security
            txtApplication = gsecBatchDefaultApplication
        End If
    Else
        'Leave the existing application
    End If
    
    ' Walk down the Application list... there was no easier way to set the
    '   ListIndex to the right value to Trigger the "cmbApplicationList_Click" event
    For i = 0 To cmbApplicationList.ListCount - 1
        If txtApplication = cmbApplicationList.List(i) Then
            ' This will Trigger the "cmbApplicationList_Click" event
            '   that will call subListBatches to Load the list of Batches
            cmbApplicationList.ListIndex = i
        End If
    Next i
    
    
    
    
    
     
    '***************************************************************
    '***  AUTO-SELECT PROCESSING
    
    'Allow Bypassing the Batch AutoSelect mode to Find another batch
    If gBypassBatchAutoSelect = True Then
        'Don't find the next available Batch
        'Reset the Bypass flag to revert back to AutoSelect if needed
        gBypassBatchAutoSelect = False
    Else
        'If Batch Mode is set to Auto Select get next available Batch
        If UCase(gsecBatchMode) = "AUTO" Then
            cmdFindNextAvailableBatch_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()
''    Dim RegConnectString As String
''    Dim RegImaging101ConnectionType As String
    
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  FormLoad()"

    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    If bolDebug Then
        Me.Caption = Me.Caption & " - DEBUG"
    End If

    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.Height", RegFileName)
'''    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.Caption", RegFileName)
    On Error GoTo 0
    
    On Error GoTo FORM_LOAD_ERROR
    
    
    cmdImportEcaptureBatch.Enabled = False

   
    
    '***************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | LOAD APPLICATION LIST DROP-DOWN"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    con.Errors.Clear
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
'*** Changed the Load to work with Security
'    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101Applications, I101SecurityApplications"
    rs.Source = rs.Source & " WHERE I101Applications.ApplicationRECID = I101SecurityApplications.ApplicationRECID And I101SecurityApplications.SecurityRECID = " & gsecSecurityRECID
    rs.Source = rs.Source & " ORDER BY ApplicationName"
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    
    rs.Open
    
    If rs.EOF = True Or rs.BOF = True Then
        MsgBox "You (" & gsecUserID & ") have not been granted Security Rights to ANY Applications.!" & vbCrLf & "Please have an administrator grant you rights and try again."
        'Exit the BatchList Module NOW.
        Unload Me
        Exit Sub
    End If
    
    rs.MoveFirst
    
    For intIndex = 0 To rs.RecordCount - 1
        cmbApplicationList.AddItem rs.Fields!ApplicationName
        cmbApplicationList.ItemData(intIndex) = rs.Fields!ApplicationRECID
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
    
    
'    '***************************************
'    '*** LOAD GROUPS LIST DROP-DOWN
'
'    cmbUserGroupList.AddItem ""
'
'    Set Con = New ADODB.Connection
'    Con.Open RegImaging101ConnectionString
'
'    Set rs = New ADODB.Recordset
'    Set rs.ActiveConnection = Con
'
'    rs.Source = "Select GroupName, GroupRECID from I101Groups ORDER BY GroupName"
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenDynamic
'    rs.LockType = adLockReadOnly
'
'    Con.Errors.Clear
'
'    rs.Open
'    rs.MoveFirst
'
'    For intIndex = 0 To rs.RecordCount - 1
'        cmbUserGroupList.AddItem rs.Fields!GroupName
'        cmbUserGroupList.ItemData(intIndex) = rs.Fields!GroupRECID
'        rs.MoveNext
'    Next
'
'    'Close connection and the recordset
'    rs.Close
'    Set rs = Nothing
'    Con.Close
'    Set Con = Nothing
'
'    '****************************
'
    
    '***************************************
    '*** LOAD BATCH LIST ORDER DROP-DOWN
        
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | LOAD BATCH LIST ORDER DROP-DOWN"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchListOrder ORDER BY BatchListOrder"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbBatchListOrder.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchListOrder.AddItem rs.Fields!BatchListOrder
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
    
    
    '***************************************
    '*** Set Batch Queue & User Defaults
    
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | Set Batch Queue & User Defaults"
    
    subSetBatchQueueDefaults
    
    
    cmdDeleteSelectedBatches.Enabled = False
    
    
    '***************************************************************
    '*** Load chkShowBatchProperties setting
    Dim strShowBatchProperties As String
    strShowBatchProperties = funcGetSetUserSettings("GET", "ShowBatchProperties", chkShowBatchProperties.Value) & ""
    If strShowBatchProperties = "0" Then
        chkShowBatchProperties = vbUnchecked
    Else
        chkShowBatchProperties = vbChecked
    End If
    
    
    '***************************************************************
    '*** Load chkCommitAfterBarcode setting
    Dim strCommitAfterBarcode As String
    strCommitAfterBarcode = funcGetSetUserSettings("GET", "CommitAfterBarcode", chkCommitAfterBarcode.Value) & ""
    If strCommitAfterBarcode = "0" Then
        chkCommitAfterBarcode = vbUnchecked
    Else
        chkCommitAfterBarcode = vbChecked
    End If
    


Exit Sub

FORM_LOAD_ERROR:
    result = MsgBox("FORM_LOAD_ERROR: " & Err.Number & " - " & Err.Description, vbOKCancel)
    Err.Clear
    If result = vbOK Then
        'Try again
        Resume
    Else
        Unload Me
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Height", Me.Height, RegFileName)
'''        Result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Caption", Me.Caption, RegFileName)
    End If

    funcSaveBatchListViewColumnPositions
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMainMenu.Show
    frmMainMenu.WindowState = vbNormal
    frmMainMenu.SetFocus
    
    Set frmImaging101BatchList = Nothing
    
End Sub

Private Sub Form_Resize()
  
    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING Form_Resize()"
    
  
  On Error Resume Next
  'This will resize the grid when the form is resized
  If Me.ScaleHeight > 0 Then
  
        frameButtons.width = Me.ScaleWidth
        frameButtons.Top = Me.ScaleHeight - frameButtons.Height
        
        frameBatchDetail.width = Me.ScaleWidth
        frameBatchDetail.Top = Me.ScaleHeight - frameBatchDetail.Height - frameButtons.Height
  
        If chkShowBatchProperties = vbChecked Then
          ListView1.Height = Me.ScaleHeight - ListView1.Top - frameButtons.Height - frameBatchDetail.Height
        Else
          ListView1.Height = Me.ScaleHeight - ListView1.Top - frameButtons.Height - 50
        End If
    
        ListView1.width = Me.ScaleWidth
        
        picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
        lblVersion.Left = picImaging101Logo.Left
        
  End If
'  txtFullPathName.Top = Me.ScaleHeight - 300
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub




Private Sub Label13_Click()
    MsgBox "Test"
End Sub



Private Sub lblBatchDirectory_DblClick()

    'Launch the Directory
    Call shelldoc(txtBatchDirectory)

End Sub

Private Sub lblSelectApplication_Click()
    cmbApplicationList.SetFocus
    
End Sub


Public Sub ListView1_Click()
    
    
    ' Exit if the ListView is Empty
    If frmImaging101BatchList.ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    

    
    On Error GoTo SKIP_LISTITEM_DISPLAY
    
    
    'Count how many items are "Selected"
    Dim i As Double
    Dim j As Double

    For i = 1 To ListView1.ListItems.Count
        If frmImaging101BatchList.ListView1.ListItems(i).Selected = True Then
            frmImaging101BatchList.ListView1.ListItems(i).Selected = True   ' Force item selection
            subGetBatchHeaderInfo
            If txtBatchGroup.Text = "TTC PRINTED" _
            Or txtBatchGroup.Text = "TTC RECEIVED" Then
                'De-Select THIS Item
                frmImaging101BatchList.ListView1.ListItems(i).Selected = False
            Else
                j = j + 1
            End If
        End If
    Next i
    
    'If more than ONE row is selected
''    If (j < 1) Or (j > 1) Then
    If (j > 1) Then
        subClearBatchListFields
        txtBatchName = "*** MULTIPLE SELECTIONS ***"
        txtBatchDesc = "*** " & j & " Items Selected ***"
        txtBatchListSelectedCount = j
        
        '*** 2020-05-22 - Jacob - Added check for chkOpenBatchesInReadOnlyMode
        If chkOpenBatchesInReadOnlyMode = vbChecked Then
            cmdCommitSelectedBatches.Enabled = False
            cmdBarcodeSelectedBatches.Enabled = False
        Else
            cmdCommitSelectedBatches.Enabled = True
            cmdBarcodeSelectedBatches.Enabled = True
        End If
        
        cmdRouteSelectedBatches.Enabled = True
        cmdDeleteSelectedBatches.Enabled = True
        Exit Sub
    Else
        txtBatchListSelectedCount = j
    End If
    
        
        ' Enable / Disable Buttons as Needed
'        If frmImaging101BatchList.ListView1.ListItems.count = 0 Then
        'But make at least ONE VALID item is selected.
        If j = 0 Then
            cmdCommitSelectedBatches.Enabled = False
            cmdBarcodeSelectedBatches.Enabled = False
            cmdDeleteSelectedBatches.Enabled = False
            cmdOpenBatch.Enabled = False
            cmdRouteSelectedBatches.Enabled = False
        Else
            '*** 2020-05-22 - Jacob - Added check for chkOpenBatchesInReadOnlyMode
            If chkOpenBatchesInReadOnlyMode = vbChecked Then
                cmdCommitSelectedBatches.Enabled = False
                cmdBarcodeSelectedBatches.Enabled = False
            Else
                cmdCommitSelectedBatches.Enabled = True
                cmdBarcodeSelectedBatches.Enabled = True
            End If
            cmdRouteSelectedBatches.Enabled = True
            'Assuming the user has rights, enable the delete and open buttons
            cmdDeleteSelectedBatches.Enabled = True
            cmdOpenBatch.Enabled = True
    End If
        


    'NOW set the row as SELECTED
    lstIndex = frmImaging101BatchList.ListView1.SelectedItem.Index
    frmImaging101BatchList.ListView1.ListItems(lstIndex).Selected = True   ' Force item selection
    subGetBatchHeaderInfo
    
Exit Sub

    
SKIP_LISTITEM_DISPLAY:
    On Error GoTo 0
    
    If cmdDeleteSelectedBatches.Visible = True Then
        cmdDeleteSelectedBatches.Enabled = True
    End If
    
End Sub


Private Sub ListView1_DblClick()
    

    If frmImaging101BatchList.ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    cmdOpenBatch_Click
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ' Set the Sort Column
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Sort It!
    ListView1.Sorted = True
End Sub

Private Sub AutoSizeColumns(Listview As Listview, Optional ByVal UseHeader As Boolean = False)
  Dim i As Integer, lparam As Long
  If UseHeader = False Then
      lparam = LVSCW_AUTOSIZE
  Else
      lparam = LVSCW_AUTOSIZE_USEHEADER
  End If
  For i = 0 To Listview.ColumnHeaders.Count - 1
      SendMessage Listview.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lparam
  Next
End Sub



Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    ' Force a ListView1_Click upon mouse up/down
'    ListView1_Click

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

     'Delete Key = Asc(46)
    If KeyCode = vbKeyDelete Then
        'See if user has Rights to Delete Batches
        If gsecRightsDeleteBatches = vbChecked Then
            cmdDeleteSelectedBatches_Click
        End If
    End If

End Sub



Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    ' Exit if the ListView is Empty
    If frmImaging101BatchList.ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    ' Button As Integer.
    '     This is a value represents which mouse button fired the event. '
    '     The value of this parameter is either vbLeftButton, vbRightButton, or vbMiddleButton.
    '     These terms are from the point of view of a right-handed mouse.
    '      vbLeftButton always refers to the primary button, regardless of whether it's physically the left or right button.
    ' Shift As Integer.
    '     This parameter represents an integer that indicates whether an auxiliary key is pressed during the Mouse event.
    '     It contains a value of 0 (none), 1 (Shift), 2 (Ctrl), 4 (Alt), or the sum of any combination of those keys.
    '     For example, if both the Ctrl and Alt key were pressed, the value of the Shift parameter is 6.
    '     You can check for the state of any one of the auxiliary keys with one of the VB constants vbAltMask, vbCtrlMask, or vbShiftMask.
    'The following code illustrates how you could store the state of each auxiliary key in a Boolean variable within the MouseDown or MouseUp event procedure.
    '     The bit-wise representation of 1, 2, or 4 in the Shift parameter is 000000001, 000000010, 00000100.
    '     By doing a logical AND between the Shift parameter and one of the VB Shift-key constants,
    '     you can pick out whether each of the three Shift keys is currently pressed.

    Dim blnIsAlt As Boolean
    Dim blnIsCtrl As Boolean
    Dim blnIsShift As Boolean
    
    blnIsAlt = Shift And vbAltMask
    blnIsCtrl = Shift And vbCtrlMask
    blnIsShift = Shift And vbShiftMask
 
    funcWriteToDebugLog Me.name, "ListIndex=" & ListView1.SelectedItem.Index
    
    
    
    '********************************************************************************
    '*** UN-COMMIT SELECTED BATCHES - CTRL+RIGHT-CLICK
    
    

    If Button = vbRightButton And blnIsCtrl Then
        

        'Only ADMINS can do this
        If gsecRightsAdminSystem = vbChecked Then
            
            
                result = MsgBox("Are you SURE you wish to UNCOMMIT the SELECTED Batch(es)?", vbYesNo)
                If result <> vbYes Then
                    Exit Sub
                End If
                
                If txtBatchName = "*** MULTIPLE SELECTIONS ***" Then
                    result = MsgBox("Are you ABSOLUTELY SURE you wish to UNCOMMIT the SELECTED Batches)?", vbYesNo)
                    If result <> vbYes Then
                        Exit Sub
                    End If
                End If
                
                
                Dim i As Double
                
                For i = 1 To ListView1.ListItems.Count
                
                    
                    If frmImaging101BatchList.ListView1.ListItems(i).Selected = True Then
                    
                        funcWriteToDebugLog Me.name, frmImaging101BatchList.ListView1.ListItems(i).Text
                        frmImaging101BatchList.ListView1.ListItems(i).Selected = True   ' Force item selection
                        
                        ' Select the Batch to commit
                        subGetBatchHeaderInfo
                        
                        If (InStr(txtBatchCommitStatus, "Committed") > 0) Then
                    
                            funcUncommitBatch txtBatchRECID, cmbApplicationList.Text, txtBatchCommitStatus.Text
                            
                        End If
                        
                    End If
                    
                Next
        
        End If
        
        subListBatches
        
        Exit Sub

    End If
    
        
        '********************************************************************************
        '*** RESET BATCH LIST COLUMN ORDER - ALT+RIGHT-CLICK
        If Button = vbRightButton And blnIsAlt Then
            
            'Only ADMINS can do this
            If gsecRightsAdminSystem = vbChecked Then
                
                
                    result = MsgBox("Are you SURE you wish to RESET the Batch List COLUMN ORDER?", vbYesNo)
                    If result <> vbYes Then
                        Exit Sub
                    End If
                        
                    funcResetBatchListViewColumnPositions
            End If
              
            Exit Sub

            
        End If
    
    
    '************************************************************************************************
     '*** RIGHT Mouse-Button Clicked
     If Button = vbRightButton And txtBatchName <> "*** MULTIPLE SELECTIONS ***" Then
     
        'Select the Item to Make sure we populate the proper fields
        '  otherwise a Right-Click may return null values.
        ListView1_Click

'        If gsecRightsBatchIndex = True Then  ' Disabled 8/6/2009 Jacob

            'If Batch is in Read-Only Mode... Allow opening the Properties Form
            ' the Properties Form will handle the required logic to limit editing capabilities
            If gOpenBatchInReadOnlyMode <> True Then
                ' LOCK the Batch
                strReturn = frmImaging101Winsock.funcSendData("LOCK BATCH" & "|" & txtBatchRECID)
                
                If Left(strReturn, 5) = "ERROR" Then
                    If gOpenBatchInReadOnlyMode <> True Then
                        'The Batch is Locked - CANNOT Edit Properties
                        strBatchModeOption = MsgBox(strReturn & vbCrLf & vbCrLf & " Properties will be opened to modify Notes ONLY!", vbOKCancel, "Error Locking Batch")
                        If strBatchModeOption = vbCancel Then
                            Exit Sub
                        End If
                        gOpenBatchInReadOnlyMode = True
                    End If
                End If
            End If
                
            txtCurrentModule = "frmImaging101BatchList"
            frmImaging101BatchProperties.Show
            
            Exit Sub

        End If
        
        
        
End Sub

Private Sub MenuName_Click(Index As Integer)

End Sub



Private Sub txtBatchCommitStatus_Click()
    
    'Make sure the user CANNOT Edit this field
    ListView1.SetFocus
    
End Sub

Private Sub txtBatchFilter_Change()

    cmdRefreshBatches.Default = True

End Sub

Private Sub subSetBatchQueueDefaults()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  subSetBatchQueueDefaults()"

'    'Only set the BatchQueue if in Auto Select Mode
'    If UCase(gsecBatchMode) = "AUTO" Then
        cmbBatchQueue.Text = gsecBatchDefaultQueue
'    End If
    
    cmbBatchListOrder.Text = gsecBatchListOrder
    txtBatchFilter.Text = ""
    cmbBatchOwner.Text = ""
    
End Sub




Private Sub txtBatchNotes_Change()

    If Trim(txtBatchNotes) = "" Then
        txtBatchNotes.BackColor = vbWhite
        txtBatchNotes.ForeColor = vbBlack
    Else
        txtBatchNotes.BackColor = vbYellow
        txtBatchNotes.ForeColor = vbRed
    End If
End Sub

Private Sub txtBatchNotes_GotFocus()

    'Make sure the user CANNOT Edit this field
    ListView1.SetFocus
    
    'Now show the Expanded Notes form
    frmImaging101BatchNotes.Show
    frmImaging101BatchNotes.txtBatchNotes = Me.txtBatchNotes
    
End Sub

Public Sub subClearBatchListFields()

        txtBatchRECID = ""
        txtBatchName = ""
        txtBatchDate = ""
        txtBatchQueue = ""
        txtBatchOwner = ""
        txtBatchPriority = ""
        txtBatchStatus = ""
        txtBatchPagesTotal = ""
        txtBatchCommitStatus = ""
        txtBatchPagesCommitted = ""
        txtBatchDesc = ""
        txtBatchNotes = ""
        txtBatchGroup = ""
        txtBatchDirectory = ""
        txtBatchPagesIndexed = ""
        txtBatchBoxNumber = ""

End Sub

Public Sub subGetBatchHeaderInfo()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  subGetBatchHeaderInfo()"

        '*** GET BATCH HEADER INFORMATION
        ' Set index of selected Row
        
        lstIndex = frmImaging101BatchList.ListView1.SelectedItem.Index
        ' Get Main Item
        txtBatchRECID = frmImaging101BatchList.ListView1.ListItems(lstIndex).Text
        ' Get Sub-Items
    
        txtBatchName = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(2).Text
        txtBatchDate = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(3).Text
        txtBatchQueue = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(6).Text
        txtBatchGroup = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(7).Text
        txtBatchOwner = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(8).Text
        txtBatchManager = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(9).Text
        txtBatchPriority = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(10).Text
        txtBatchStatus = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(11).Text
        txtBatchPagesTotal = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(12).Text
        txtBatchCommitStatus = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(13).Text
        txtBatchPagesCommitted = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(14).Text
        txtBatchDesc = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(15).Text
        txtBatchNotes = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(16).Text
        txtBatchDirectory = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(17).Text
        txtBatchPagesIndexed = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(29).Text
        txtBatchBoxNumber = frmImaging101BatchList.ListView1.ListItems(lstIndex).ListSubItems(32).Text

End Sub

Private Sub subSetBatchButtonSecurity()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  subSetBatchButtonSecurity()"

    '***************************************************************
    '***  BUTTON SECURITY SECTION
    
    If gsecRightsAdminSystem = vbChecked Then
        txtBatchDirectory.Visible = True
        lblBatchDirectory.Visible = True
    Else
        txtBatchDirectory.Visible = False
        lblBatchDirectory.Visible = False
    End If
    
    If gsecRightsBatchAdministration = vbChecked Then
        chkListAllBatches.Visible = True
        chkListBatchesCommittedFull.Visible = True
    Else
        chkListAllBatches.Visible = False
        chkListBatchesCommittedFull.Visible = False
    End If
    
    If gsecRightsBatchScan = vbChecked Then
        cmdScanDocuments.Visible = True
    Else
        cmdScanDocuments.Visible = False
    End If
 
'    If gsecRightsImportFromEcapture = vbChecked Then
'        cmdImportEcaptureBatch.Visible = True
'    Else
'        cmdImportEcaptureBatch.Visible = False
'    End If
    
    If gsecRightsImportFromFile = vbChecked Then
        cmdImportBatchFromFile.Visible = True
    Else
        cmdImportBatchFromFile.Visible = False
    End If
    

    If gsecRightsBatchIndex = vbChecked _
    Or gsecRightsBatchView = vbChecked Then
        cmdOpenBatch.Visible = True
        cmdFindNextAvailableBatch.Visible = True
        cmdBarcodeSelectedBatches.Visible = True
    Else
        cmdOpenBatch.Visible = False
        cmdFindNextAvailableBatch.Visible = False
        cmdBarcodeSelectedBatches.Visible = False
    End If
    
    
    If gsecRightsBatchCommit = vbChecked Then
        cmdCommitSelectedBatches.Visible = True
        chkCommitAfterBarcode.Visible = True
    Else
        cmdCommitSelectedBatches.Visible = False
        cmdBarcodeSelectedBatches.Visible = False
       'NO Commit Rights... UnCheck to not allow committing after barcode
        chkCommitAfterBarcode = vbUnchecked
        chkCommitAfterBarcode.Visible = False
    End If
    
    
    If gsecRightsBatchRoute = vbChecked Then
        cmdRouteSelectedBatches.Visible = True
    Else
        cmdRouteSelectedBatches.Visible = False
    End If

'        lblBatchQueue.Visible = False
'        txtBatchQueue.Visible = False
'        lblBatchOwner.Visible = False
'        txtBatchOwner.Visible = False
'        lblBatchPriority.Visible = False
'        txtBatchPriority.Visible = False
'        lblBatchStatus.Visible = False
'        txtBatchStatus.Visible = False


    If gsecRightsBatchChangeOwner = vbChecked Then
        lblBatchOwnerDropDown.Visible = True
        cmbBatchOwner.Visible = True
    Else
        lblBatchOwnerDropDown.Visible = False
        cmbBatchOwner.Visible = False
    End If
    
    If gsecRightsBatchChangeQueue = vbChecked Then
        lblBatchQueueDropDown.Visible = True
        cmbBatchQueue.Visible = True
    Else
        lblBatchQueueDropDown.Visible = False
        cmbBatchQueue.Visible = False
    End If
    
    
    If gsecRightsDeleteBatches = vbChecked Then
        cmdDeleteSelectedBatches.Visible = True
        cmdDeleteSelectedBatches.Enabled = True
    Else
        cmdDeleteSelectedBatches.Visible = False
        cmdDeleteSelectedBatches.Enabled = False
    End If
    
    If gsecRightsBatchChangeOrder = vbChecked Then
        cmbBatchListOrder.Enabled = True
        cmbBatchListOrder.Visible = True
        lblBatchListOrder.Visible = True
    Else
        cmbBatchListOrder.Enabled = False
        cmbBatchListOrder.Visible = False
        lblBatchListOrder.Visible = False
    End If
    
    '*** 2020-05-22 - Jacob - Added check for chkOpenBatchesInReadOnlyMode to prevent "Commit" of selected batches when checked
    If chkOpenBatchesInReadOnlyMode = vbChecked Then
        cmdCommitSelectedBatches.Enabled = False
        cmdBarcodeSelectedBatches.Enabled = False
    Else
        cmdCommitSelectedBatches.Enabled = True
        cmdBarcodeSelectedBatches.Enabled = True
    End If
    
    
End Sub




Private Function funcSaveBatchListViewColumnPositions()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  funcSaveBatchListViewColumnPositions()"
    
    Dim strNewBatchListColumnPositionList As String
    
    For i = 1 To ListView1.ColumnHeaders.Count
        strNewBatchListColumnPositionList = strNewBatchListColumnPositionList & ListView1.ColumnHeaders(i).Position & "|"
    Next
    
'    If strNewBatchListColumnPositionList <> strBatchListColumnPositionList Then
        'SAVE the list
         result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.strBatchListColumnPositionList", strNewBatchListColumnPositionList, RegFileName)
'    End If

End Function

Private Function funcLoadBatchListViewColumnPositions()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  funcLoadBatchListViewColumnPositions()"
    
    Dim strBatchListColumnPositionList As String
    Dim strBatchListColumnPositionArray() As String
        
    strBatchListColumnPositionList = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.strBatchListColumnPositionList", RegFileName)
    
    If Trim(strBatchListColumnPositionList) <> "" Then
    
        strBatchListColumnPositionArray = Split(strBatchListColumnPositionList, "|")
    
    '*** 2020-05-17 - Jacob - Modified Load to use a Sorted ListViewColumnOrder
        ListViewColumnOrder.ListItems.Clear
        For i = 0 To ListView1.ColumnHeaders.Count - 1
            'Add items to the ListViewColumnOrder with a Format lf "pp-ii" where pp=Position and ii=Index with LEADING ZERO's
            ListViewColumnOrder.ListItems.Add = Format(strBatchListColumnPositionArray(i), "00") & "~" & Format(i, "00")
       Next
        
        'Ignore Errors
        On Error Resume Next
        
        ListView1.Visible = False
        
        '*** 2020-05-17 - Jacob - Modified Load to use a Sorted ListViewColumnOrder
        '                                             Set Column Position in REVERSE  ORDER
        For i = ListView1.ColumnHeaders.Count To 1 Step -1
                Dim strPositionData() As String
                strPositionData = Split(ListViewColumnOrder.ListItems(i), "~")
                ListView1.ColumnHeaders(CInt(strPositionData(1)) + 1).Position = CInt(CInt(strPositionData(0)))
       Next
        
        
    End If
    
    ListView1.Refresh
    ListView1.Visible = True
    
End Function

Private Function funcResetBatchListViewColumnPositions()

    funcWriteToDebugLog Me.name, "frmImaging101BatchLIst() | *** ENTERING  funcResetBatchListViewColumnPositions()"

    'Ignore Errors
    On Error Resume Next
    
    Debug.Print " "
    Debug.Print "funcResetBatchListViewColumnPositions()"

   ListView1.Visible = False


    For i = 1 To ListView1.ColumnHeaders.Count
        ListView1.ColumnHeaders(i).Position = i
         Debug.Print i & "|" & ListView1.ColumnHeaders(i).Position & "|" & ListView1.ColumnHeaders(i).Text

        DoEvents
    Next
    
    ListView1.Refresh
    ListView1.Visible = True

    '*** 2020-05-17 - Jacob Disabled the Save after a Reset
'    funcSaveBatchListViewColumnPositions
    
End Function
