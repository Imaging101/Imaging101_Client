VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{895CDC7A-8837-11D1-8109-020701190C00}#8.0#0"; "docctrl.ocx"
Begin VB.Form frmIndex 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Indexing - Imaging101"
   ClientHeight    =   6525
   ClientLeft      =   3705
   ClientTop       =   1845
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   5070
   Begin ComctlLib.Toolbar toolbarBottom 
      Align           =   2  'Align Bottom
      Height          =   990
      Left            =   0
      TabIndex        =   23
      Top             =   5535
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   1746
      ButtonHeight    =   1587
      Appearance      =   1
      _Version        =   327682
      Begin VB.CommandButton cmdUpdatePrintedStatus 
         BackColor       =   &H00E3CC6C&
         Caption         =   "Update &Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2745
         Picture         =   "frmIndex.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdFindUncommitted 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Uncommited"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdFindQuestionable 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Questionable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdGotoImage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Image #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtFind 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E3CC6C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   61
         Text            =   "FIND"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdSplitBatch 
         BackColor       =   &H00E3CC6C&
         Caption         =   "Split B&atch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2895
         Picture         =   "frmIndex.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdRotateImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Rotate Image"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Picture         =   "frmIndex.frx":0B16
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearFieldValues 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear Fields"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Picture         =   "frmIndex.frx":10A0
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdNextImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next Page"
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
         Height          =   495
         Left            =   3720
         Picture         =   "frmIndex.frx":11EA
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdPreviousImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Previous Page"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Picture         =   "frmIndex.frx":1574
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdBookMark 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&BookMark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Picture         =   "frmIndex.frx":18FE
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCommitBatch 
         BackColor       =   &H00E3CC6C&
         Caption         =   "Co&mmit Batch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Picture         =   "frmIndex.frx":1E88
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancelCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel Copy"
      Height          =   300
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdMakeCopies 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3CC6C&
      Caption         =   "Make Copies"
      Height          =   300
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCopyItems 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   780
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   77
      Text            =   "frmIndex.frx":1FD2
      Top             =   2715
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtFieldRouteToBatchManager 
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   76
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtFieldRouteToBatchUser 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   75
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtFieldTableLookupOverridesDefault 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCommitViaFTP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsForOutputOnly 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchDocDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtIndexValues 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmIndex.frx":1FDF
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtFieldDefaultForBarcodeOnly 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdProcessBarcodes 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Process Barcodes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      Picture         =   "frmIndex.frx":1FEE
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin SPICERDOCUMENTLib.SpicerDoc SpicerDoc1 
      Left            =   4680
      Top             =   3720
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.FileListBox AnnotationFileListBox 
      Height          =   285
      Left            =   2640
      Pattern         =   "NOFILE.PATTERN"
      TabIndex        =   57
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFieldRouteToBatchQueue 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchPageRotation 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFieldSplitBatches 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsRequiredForSplit 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchDesc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
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
      Height          =   2295
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdDeleteSelectedPage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete Page"
         Height          =   350
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteSelectedPageIcon 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   3480
         Picture         =   "frmIndex.frx":2578
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBatchPagesTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Expected # of Pages for This Document"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtBatchOwner 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtBatchGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditBatchProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edi&t Batch"
         Height          =   735
         Left            =   4440
         Picture         =   "frmIndex.frx":26C2
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtBatchQueue 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtBatchStatus 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdFindBatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find Batch"
         Height          =   735
         Left            =   4440
         Picture         =   "frmIndex.frx":280C
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtBatchCommitStatus 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtBatchPageStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   195
         Left            =   1440
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtApplicationName 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtBatchName 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblCommitViaFTP 
         BackStyle       =   0  'Transparent
         Caption         =   "FTP"
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
         Left            =   4080
         TabIndex        =   72
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pages"
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
         Left            =   3480
         TabIndex        =   71
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner"
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
         Left            =   120
         TabIndex        =   63
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Group"
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
         Left            =   120
         TabIndex        =   51
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Queue"
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
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Status"
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
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Page Status"
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
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label LabelBatchCommitStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Commit Status"
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
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Application"
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
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Name"
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
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAcceptValues 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save Values"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtBatchRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtBatchPageRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtBatchPageFileName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtBatchDirectory 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtFieldIsSticky 
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdBatchPageList 
      Caption         =   "&List"
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
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtFieldLowValue 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldHighValue 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldDefaultValue 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldsRECID 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchFieldsRECID 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldType 
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldSize 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldName 
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsRequiredForCommit 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   330
      HelpContextID   =   1
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   2385
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CheckBox chkProcessBarcodesEntireBatch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Entire Batch"
      Height          =   432
      Left            =   1200
      TabIndex        =   59
      Top             =   4920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page Desc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   65
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblFieldDescription 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   840
      TabIndex        =   14
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Descr"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPageFilename 
      BackStyle       =   0  'Transparent
      Caption         =   "Page Filename"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Directory"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Dim intHoldFocusIndex As Integer
    Dim intListView1CurrentItem As Integer
    Dim intBatchPageCount As Integer
    
    Dim FormUnloadMode As Variant
    Dim txtBatchSplitRECID As Double
    
    
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ssql As String
    
    Dim connImaging101 As ADODB.Connection
    Dim connImaging101Batch As ADODB.Connection
    
    Dim rsImaging101Batch As ADODB.Recordset
    Dim rsImaging101BatchPage As ADODB.Recordset
    Dim rsImaging101BatchSplit As ADODB.Recordset
    Dim rsImaging101Document As ADODB.Recordset
    Dim rsImaging101DocumentDetail As ADODB.Recordset
    
    Dim ssqlImaging101Batch As String
    Dim ssqlImaging101BatchPage As String
    
    Dim connTTC(2) As ADODB.Connection
    Dim cmdTTC(2) As ADODB.Command
    Dim rsTTC(2) As ADODB.Recordset
    
    Dim connHA As ADODB.Connection
    Dim cmdHA As ADODB.Command
    Dim rsHA As ADODB.Recordset
    
'    bolIndexFormLoadComplete Defined in VariableDeclarations Module As Boolean
    
    Dim bolErrorOccured As Boolean
    Dim bolProcessingBarcodes As Boolean
    Dim bolDropLeadingZeroes As Boolean
    Dim bolUseBarcodeAsDocumentHeader As Boolean
    
    'Barcode Clip Variables
    Dim intBarcodeClipBeginPosition As Integer
    Dim intBarcodeClipNumberOfCharacters As Integer
    
    'Field Spacing
    Dim intFieldSpacing As Integer
    
    Dim txtApplicationCommitBatchTo As String
    Dim txtApplicationCommitBatchOption As String
    
    'DocType Fields
    Dim strDOCGROUP As String
    Dim strDOCTYPE As String
    Dim strDOCSUBTYPE As String
    

    


    



Private Sub cmdAcceptValues_Click()
    
        subSaveBatchPageValues
    
End Sub


Private Sub cmdBatchPageList_Click()

    '*** 2023-02-23 - Jacob - Commented SetFocus and Added subLoadPagesIntoListView
'    ListView1.SetFocus
    subLoadPagesIntoListView

End Sub

Private Sub cmdBookMark_Click()

    cmdBookMark_Set

End Sub
    
Public Sub cmdBookMark_Set(Optional NoPrompt As Boolean, Optional BookMarkAllPages As Boolean)

    Dim intPageIndex As Integer
    Dim intIndex As Integer
    Dim intBookMarkOfFieldsToCopy As Integer
    Dim intBookMarkSwitch As Integer
    
    
    '2014-09-16 - Auto-Bookmark Option
    If BookMarkAllPages = True Then
        'Bypass the Set Bookmark
        blnBookMark = True
        intBookMarkBegin = 1
        intBookMarkEnd = frmIndex.ListView1.ListItems.Count
    End If
    
    
    If blnBookMark = False Then
        'SET the BookMark
        blnBookMark = True
        cmdBookMark.Caption = "&CopyValues"
        intBookMarkBegin = ListView1.SelectedItem.Index
        ListView1.SelectedItem.ForeColor = vbGreen
        'Disable the Commit  & ProccessBarcodes buttons
        cmdCommitBatch.Enabled = False
        cmdProcessBarcodes.Enabled = False
    
    Else
        
        'Handle the CopyValues Click OR BookMarkAllPages
        If BookMarkAllPages = True Then
            'Move right along
        Else
        
            intBookMarkOfFieldsToCopy = ListView1.SelectedItem.Index
            intBookMarkEnd = intBookMarkOfFieldsToCopy
            
    
            If intBookMarkEnd < intBookMarkBegin Then
                intBookMarkSwitch = intBookMarkEnd
                intBookMarkEnd = intBookMarkBegin
                intBookMarkBegin = intBookMarkSwitch
                frmIndex.ListView1.ListItems.item(intBookMarkBegin).ForeColor = vbGreen
                frmIndex.ListView1.ListItems.item(intBookMarkEnd).ForeColor = vbBlack
            End If
            
        End If
        
        'If the NoPrompt flag is NOT set, Verify the user wants to copy the values
        If NoPrompt Then
            result = vbOK
        
        Else
            result = MsgBox("Are you sure you wish to Copy " & vbCrLf & "the Current values to pages " & intBookMarkBegin & " To " & intBookMarkEnd & " ? ", vbOKCancel, "Bookmark Copy Values")
        End If
        
        If result = vbOK Then
        
            '*** Save Values for Current Image
            subSaveBatchPageValues
            
            frmIndex.ListView1.ListItems.item(intBookMarkEnd).ForeColor = vbRed

            If intBookMarkEnd > intBookMarkBegin Then
            
                '*** Walk from BookMark Begin to End
                '* Loop Through Pages
                For intPageIndex = intBookMarkBegin To intBookMarkEnd
                    
                    frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
                    
                    '*** Load field values from "Bookmarked" page by passing
                    '    the intBookMarkOfFieldsToCopy as a parameter
                    '    to the subGetBatchFieldValues subroutine
                    subGetBatchFieldValues intBookMarkOfFieldsToCopy
                    
                        ' Save the loaded values
                        subSaveBatchPageValues
                    
                    'Color the pages in between Begin and End Bookmarks
                    If intPageIndex > intBookMarkBegin And intPageIndex < intBookMarkEnd Then
                        ListView1.SelectedItem.ForeColor = vbYellow
                    End If
                    
                Next
            
            End If
            
'            MsgBox "Copy Indexes Complete... " & intBookMarkBegin & " To " & intBookMarkEnd, vbInformation
        Else
            'Reset the Bookmark and Re-select the Bookmark Begin page
            frmIndex.ListView1.ListItems.item(intBookMarkBegin).Selected = True
            ListView1.SelectedItem.ForeColor = vbBlack
            ListView1_Click
        End If
        
        blnBookMark = False
        cmdBookMark.Caption = "&BookMark"
        'Re-enable the Commit & Process Barcodes buttons
        cmdCommitBatch.Enabled = True
        cmdProcessBarcodes.Enabled = True

    End If
End Sub

Private Sub cmdCancelCopy_Click()

                '*** 2023-02-20 - Jacob - Added to Handle Copying Pages
                cmdCancelCopy.Visible = False
                txtCopyItems.Visible = False
                cmdMakeCopies.Visible = False
                cmdMakeCopies.Default = False
                

End Sub

Public Sub cmdCommitBatch_Click()

    
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "-------------------------------------------"
    funcWriteToDebugLog Me.name, "ENTERING: cmdCommitBatch_Click()"
    
    'Unload Annotations Form if loaded
    If funcIsFormLoaded2("frmAnnotate") Then
        funcWriteToDebugLog Me.name, "frmAnnotate is Loaded... UNLOAD it."
        Unload frmAnnotate
'        Set frmAnnotate = Nothing
    End If

    DoEvents
    
    '*** Only Save the Indexes if NOT in Read-Only Mode
    If (gOpenBatchInReadOnlyMode <> True) Then
        'Save Field Values
        funcWriteToDebugLog Me.name, "subSaveBatchPageValues"
        subSaveBatchPageValues
    End If
    
    DoEvents
    
    'Make the user was in the middle of a Bookmark operation
    If blnBookMark = True Then
        result = MsgBox("Please complete the 'Bookmark' operation and try again.", vbInformation, "Bookmark Enabled")
        Exit Sub
    End If
        
    ' If committing Selected Batches, set Result to yes and let it rip!
    If blnCommitSelectedBatches = True Then
        funcWriteToDebugLog Me.name, "blnCommitSelectedBatches = True"
        result = vbYes
    Else
        funcWriteToDebugLog Me.name, "blnCommitSelectedBatches = False"
        result = MsgBox("Are you sure you wish to Commit/Release the Batch?", vbYesNo, "Batch Commit")
    End If
    
    DoEvents
    
    If result = vbYes Then
        ' COMMIT / RELEASE the Batch!
        Dim intPageIndex As Integer
        Dim intIndex As Integer
        Dim strOutputLine As String
        Dim bolSkipPage As Boolean
        Dim strOutputFileName As String
        Dim intAppendFieldIndex As Integer
        
        '*** SET GLOBAL FLAG FOR COMMITTING BATCH PAGES
        '     This way we Won't display the Images as we process them!
        blnCommittingBatchPages = True
        
        '*** HIDE the Indexing Forms for Speed & to Unclutter the desktop!
        '     works faster if video doesn't display detail.
        funcWriteToDebugLog Me.name, "HIDE the Indexing Forms "
        frmLookupList.Visible = False
        frmDocTypeList.Visible = False
        MainMDIForm.Visible = False
        Me.Visible = False
        
        '*** SHOW the CommitStatus Window & Disable Buttons
        funcWriteToDebugLog Me.name, "SHOW the CommitStatus Window & Disable Buttons "
        frmCommitStatus.Show
        frmCommitStatus.cmdCloseBatch.Enabled = False
        frmCommitStatus.cmdStayOnBatch.Enabled = False
        funcMakeTopMost frmCommitStatus, True
        
        DoEvents

        
        
        '*** RESET Counter Variables
        frmCommitStatus.txtPagesProcessed = 0
        frmCommitStatus.txtPagesCommitted = 0
        frmCommitStatus.txtPagesPreviouslyCommitted = 0
        frmCommitStatus.txtPagesSeparator = 0
        frmCommitStatus.txtPagesQuestionable = 0
        frmCommitStatus.txtPagesDoNotFile = 0
        frmCommitStatus.txtPagesRequiredButEmpty = 0
        frmCommitStatus.txtPagesFailedValidation = 0
        frmCommitStatus.txtPagesTotalSkipped = 0
        
        
        
        '*********************************************************************
        '*** Commit to System Selected by User in Application Configuration
        
        Select Case frmImaging101BatchList.lblApplicationCommitBatchTo
             
             Case "Imaging101"
                'Commit Batch to Imaging101
                funcWriteToDebugLog Me.name, "Case 'Imaging101' - Commit Batch to Imaging101 "
                CommitBatchToImaging101
                
             Case "Imaging101AutoImport"
                '*** 2021-06-14 - Jacob - Added Commit Batch to Imaging101AutoImport
                funcWriteToDebugLog Me.name, "Case 'Imaging101AutoImport' - Commit Batch to Imaging101AutoImport "
                CommitBatchToImaging101AutoImport
                
           Case "TTC"
                'Commit the Batch to TTC
                CommitBatchToTTC
                
            Case "HMIS Software"
                'Commit Batch to HMIS Software
                CommitBatchToHMISSoftware
                
            Case "ISAC"
                'Commit Batch to Imaging101
                CommitBatchToISAC

            Case "Oracle IPM"
                'Commit Batch to Oracle IPM (Oracle/Stellent)
                funcWriteToDebugLog Me.name, "Case 'Oracle IPM' - Commit Batch to Oracle IPM "
                CommitBatchToOracleIPM

        End Select
                    
        DoEvents

        '*** DISABLE GLOBAL FLAG FOR COMMITTING BATCH PAGES
        '     To resume displaying the Images as we index them!
        blnCommittingBatchPages = False
        
'        Me.WindowState = vbNormal
'        frmDocTypeList.WindowState = vbNormal
'        frmLookupList.WindowState = vbNormal
        
        Me.Visible = True
        frmDocTypeList.Visible = True
        frmLookupList.Visible = True

        
        
        If blnCommitSelectedBatches = True Then
            '*** CREATE BATCH AUDIT RECORD
            funcWriteToDebugLog Me.name, "CREATE BATCH AUDIT RECORD - Commit Batch - Selected "
            funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Commit Batch - Selected"
        Else
            '*** CREATE BATCH AUDIT RECORD
            funcWriteToDebugLog Me.name, "CREATE BATCH AUDIT RECORD - Commit Batch - Single "
            funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Commit Batch - Single"
        End If
    

        DoEvents

        
       ' Set the flag back to False to END the dummy loop in
        '   "Private Sub cmdCommitSelectedBatches_Click()" of frmImaging101BatchList
        If blnCommitSelectedBatchesWait = True Then
            blnCommitSelectedBatchesWait = False
        End If
                
        frmCommitStatus.SetFocus
        
        frmCommitStatus.cmdCloseBatch.Enabled = True
        frmCommitStatus.cmdStayOnBatch.Enabled = True

    End If

    funcWriteToDebugLog Me.name, "EXITING: cmdCommitBatch_Click()"
    funcWriteToDebugLog Me.name, "-------------------------------------------"
    funcWriteToDebugLog Me.name, ""


End Sub

Private Sub CommitBatchToOracleIPM()

        Dim intPageIndex As Integer
        Dim intIndex As Integer
        Dim strBatchInputFileDirectory As String
        Dim strBatchInputFileExtension As String
        Dim strBatchOracleInputFileExtension As String
        Dim strOutputFileName As String
        Dim strOracleOutputFileName As String
        Dim strRootDirectoryPathForHtmlSource As String
        
        'Get the Input File Directory & Extension
        strBatchInputFileDirectory = txtBatchDirectory
        strBatchInputFileExtension = ".TMP"
        strBatchOracleInputFileExtension = ".TXT"
        
        'Make sure we have an InputFile Directory & Extension
        If Trim(strBatchInputFileDirectory) = "" Or Trim(strBatchInputFileExtension) = "" Then
            MsgBox "Please set the" & txtApplicationName & ".BatchInputFileDirectory AND " & txtApplicationName & ".BatchInputFileExtension PARAMETERS in the " & RegFileName
            blnBatchError = True
            Exit Sub
        End If
        
        'Define the Output Directory & File Name
        strRootDirectoryPathForHtmlSource = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationName = '" & txtApplicationName & "'", "RootDirectoryPathForHtmlSource") & ""
'        strOutputFileName = strRootDirectoryPathForHtmlSource & "\CLI_I101_" & txtBatchName & "_" & Format(Now, "yymmddhhmmss") & strBatchInputFileExtension
        strOutputFileName = strRootDirectoryPathForHtmlSource & "\CLI_I101_" & Trim(txtBatchName) & "_" & Format(Now, "mmss") & strBatchInputFileExtension
        strOracleOutputFileName = strRootDirectoryPathForHtmlSource & "\CLI_I101_" & Trim(txtBatchName) & "_" & Format(Now, "mmss") & strBatchOracleInputFileExtension
       
        '***********************************
        '* Loop Through Batch Pages
        '***********************************
        For intPageIndex = 1 To ListView1.ListItems.Count
        
           frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            bolSkipPage = False
            
            '* Select the Page Item on the List and Get Fields for Selected Item
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            subGetBatchFieldValues frmIndex.ListView1.SelectedItem.Index
            
'            ' Locate the Record
'            frmIndex.ListView1.SetFocus
'            ListView1_Click

                
            '* Loop Through Fields
            If txtBatchPageStatus <> "Committed" Then
            
            
               '*******************************************************************
               '*** FIRST Loop to see if this record should be skipped
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable _
                   Or txtIndexValues(intIndex).Text = txtQuestionable _
                   Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator _
                   Or txtIndexValues(intIndex).Text = txtSeparator _
                   Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile _
                   Or txtIndexValues(intIndex).Text = txtDoNotFile _
                   Then
                        frmCommitStatus.txtPagesDoNotFile = frmCommitStatus.txtPagesDoNotFile + 1
                        bolSkipPage = True
                        Exit For
                   End If
               Next
               '*** END Loop to see if this record should be skipped
               '*******************************************************************

               
               '****************************************************************************
               '*** Only check for Fields Required or Valid if Not flagged for Skip
               If bolSkipPage <> True Then
               
                    For intIndex = 0 To mebIndexValues.Count - 1
                        ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                        If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "" And txtIndexValues(intIndex).Text = "") Then
                             frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                             bolSkipPage = True
                             Exit For
                        End If
                        
                         '*** VALIDATE DATE FIELD!
                         funcValidateDate intHoldFocusIndex
                         If blnDateError = True Then
                             frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                             bolSkipPage = True
                             Exit For
                         End If
                     
                    Next
               End If   'bolSkipPage <> True
               
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            '*** END check for Fields Required or Valid if Not flagged for Skip
            '****************************************************************************
           
           
            '****************************************************************************
            '*** NOW CHECK IF PAGE SHOULD BE SKIPPED
            If bolSkipPage <> True Then
                
                    '* Begin the Transaction
''                    connImaging101Batch.BeginTrans
'' JR 2/6/03 Disabled Transaction logging to speed it up because this only deals with ONE table and ONE update!

                    
                    '**********************************************************************************
                    ' Set up the FilePath
                    
                    strOutputLine = "APPEND PAGE|" & txtBatchDirectory & "\" & txtBatchPageFileName


                    '**********************************************************************************
                    ' Loop and append all the Application Fields -- Skip the Application System Fields
                    
                    '* Loop Through Fields
                    For intIndex = 0 To mebIndexValues.Count - 1
                    
                         ' Add field to the OutputLine since there were no exceptions!
                         If txtFieldType(intIndex) = "Date" Then
                            
                            If IsDate(mebIndexValues(intIndex).FormattedText) Then
                                'If it's a VALID "Date" then reformat it for Oracle IPM
                                strOutputLine = strOutputLine & "|" & Format(mebIndexValues(intIndex).FormattedText, "yyyy-MM-dd")
                            Else
                                'If NOT a VALID date... simply use a blank.
                                strOutputLine = strOutputLine & "|" & ""
                            End If
                         Else
                            strOutputLine = strOutputLine & "|" & Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format)
                         End If
                         
                         
                         
                    Next
                       
                 ''***********************************
                
                
                '****************************************************************************
                '*** WRITE OUTPUT LINE
                
                Open strOutputFileName For Append As #1
                Print #1, strOutputLine
                Close #1
                
                
                
                        '****************************************************************************
                        '*** ESTABLISH DATABASE CONNECTIONS
                        
                            '**************************************************************
                            '*** Establish BATCH DB Connection
                            Set connImaging101Batch = New ADODB.Connection
                            connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
                            connImaging101Batch.ConnectionTimeout = 120
                            connImaging101Batch.mode = adModeReadWrite
                            connImaging101Batch.Open
                            connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
            
                            '**************************************************************
                            '*** CONNECT to Batch DB RESULT SET
                            Set rsImaging101Batch = New ADODB.Recordset
                            txtActionBeforeError = "Connect to Batch DB"
                            Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                            rsImaging101Batch.CursorLocation = adUseServer
                            rsImaging101Batch.CursorType = adOpenDynamic
                            rsImaging101Batch.LOCKTYPE = adLockOptimistic
                            
                            '**************************************************************
                            '*** FIND BATCH RECORD
                            
                            txtActionBeforeError = "Open Batch DB"
                            rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                            rsImaging101Batch.Open
                            
                            '**************************************************************
                            '*** CONNECT to Batch Page DB RESULT SET
                            Set rsImaging101BatchPage = New ADODB.Recordset
                            txtActionBeforeError = "Connect to Batch Pages DB"
                            Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                            rsImaging101BatchPage.CursorLocation = adUseServer
                            rsImaging101BatchPage.CursorType = adOpenDynamic
                            rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                            
                            '**************************************************************
                            '*** FIND BATCHPAGE RECORD
                    
                            txtActionBeforeError = "Open Batch Page DB"
                            rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                            rsImaging101BatchPage.Open

                    rsImaging101BatchPage.MoveFirst
                    
                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                    txtActionBeforeError = "Update BatchPageStatus in Batch Pages Table"
                    rsImaging101BatchPage.Fields!BatchPageStatus = "Committed"
                    rsImaging101BatchPage.Update
    ''                    connImaging101BatchPage.CommitTrans
                    rsImaging101BatchPage.Close
                    Set rsImaging101BatchPage = Nothing
                    
                frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
            
            Else
            
                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
            
            DoEvents
        
        Next
        
        
                funcWriteToDebugLog Me.name, "END: Loop Through Batch Pages"
        
        '*********************************************
        '* END: Loop Through Batch Pages
        '*********************************************
        
        
        subCommitTransactions
        
        
        
        '******************************************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '******************************************************************************************
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
        
        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.Open
        
        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH COMMITTED - FULL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH COMMITTED - FULL"
            funcWriteToDebugLog Me.name, "BATCH COMMITTED - FULL"

        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH COMMITTED - PARTIAL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH COMMITTED - PARTIAL"
            funcWriteToDebugLog Me.name, "BATCH COMMITTED - PARTIAL"
        End If
        
        rsImaging101Batch.Update
        
        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        
        Set rsImaging101Batch = Nothing
        
        Set connImaging101Batch = Nothing
        
        'NOW RENAME THE OUTPUT FILE TO HAND IT OVER TO ORACLE
        Dim fso As New FileSystemObject
        Set fso = New Scripting.FileSystemObject
        On Error Resume Next
        Err.Clear
        fso.MoveFile strOutputFileName, strOracleOutputFileName
        Set fso = Nothing
        
        '*** Bring the CommitStatus Window to the front
        '*** Allow user to select to CLOSE or STAY
        frmCommitStatus.SetFocus
        
        
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "EXITING: CommitBatchToOracle"
    funcWriteToDebugLog Me.name, "----------------------------------------------------------------------------------"
        
End Sub

Private Sub CommitBatchToTTC()

    Dim strTempFilePath As String
    
    '************************************
    '*** Get FTP Information
    Dim txtFTPSite(2) As String
    Dim txtFTPUserID(2) As String
    Dim txtFTPPassword(2) As String
    Dim txtTTCConnectionString(2) As String
     
    Dim txtTTCUserID As Integer
    
    Dim intSiteIdIndex As Integer
    Dim dblCaseIdCutoff As Double
    
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Applications where ApplicationRECID = " & txtApplicationRECID
    rs.Open ssql, con
    
        'Get the
        txtFTPSite(0) = rs.Fields!FTPSite
        txtFTPUserID(0) = rs.Fields!FTPUserID
        txtFTPPassword(0) = rs.Fields!FTPPassword
        txtTTCConnectionString(0) = rs!LookupDBConnectionString

        txtFTPSite(1) = rs.Fields!FTPSite_B
        txtFTPUserID(1) = rs.Fields!FTPUserID_B
        txtFTPPassword(1) = rs.Fields!FTPPassword_B
        txtTTCConnectionString(1) = rs!LookupDBConnectionString_B
        
        dblCaseIdCutoff = CDbl(rs.Fields!CaseIdCutoff)
       
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    
    For intSiteIdIndex = 0 To 1
    
        '**********************************
        '*** TTC DB Connection Setup
        
        Set connTTC(intSiteIdIndex) = New ADODB.Connection
        Set cmdTTC(intSiteIdIndex) = New ADODB.Command
        Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
    
        '**************************************************************
        '*** Establish TTC DB Connections
        
        txtActionBeforeError = "Open DB ConnectionString - Site = : " & CStr(intSiteIdIndex) & ": " & txtTTCConnectionString(intSiteIdIndex)
        
        connTTC(intSiteIdIndex).ConnectionString = txtTTCConnectionString(intSiteIdIndex)
        connTTC(intSiteIdIndex).ConnectionTimeout = 120
        connTTC(intSiteIdIndex).mode = adModeReadWrite
        connTTC(intSiteIdIndex).Open
    
        Set cmdTTC(intSiteIdIndex).ActiveConnection = connTTC(intSiteIdIndex)
    
    Next


    On Error GoTo UPDATE_TTC_fields_names_TABLE_ERROR
    
    
    
    
    '**************************************************************
    '*** Get TTC Login info
    
    If txtBatchGroup = "TTC RECEIVED" Then
    
        bolTTCUserFound = False
        While Not bolTTCUserFound
            frmLoginTTC.Show vbModal, Me
            
            If frmLoginTTC.bolTTCLoginClickedLogin = False Then
                Unload frmLoginTTC
                Exit Sub
            End If
            
            'FORCE to Site A
            intSiteIdIndex = 1
            
            ' Validate USER and make sure user is "Active"
            Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
            txtSource = "select id, username, password, active from users where username = '" & frmLoginTTC.txtUserID & "' AND password = '" & frmLoginTTC.txtPassword & "' AND active = 1"
            rsTTC(intSiteIdIndex).Open txtSource, connTTC(intSiteIdIndex), adOpenDynamic, adLockOptimistic
            If rsTTC(intSiteIdIndex).EOF Or rsTTC(intSiteIdIndex).BOF Then
                 MsgBox "Invalid User name or Password!" & vbCrLf & "Please try again...", vbOKOnly, "TTCLoginFailed"
            Else
                bolTTCUserFound = True
                'SAVE UserID to update RECEIVED fields.
                txtTTCUserID = rsTTC(intSiteIdIndex).Fields!ID
                Unload frmLoginTTC
                'This command will immediately "Close" the rsTTC after executing
            End If
         Wend
         
        
    End If
    
    
    
    '**************************************************************
    '*** Establish BATCH DB Connection
    Set connImaging101Batch = New ADODB.Connection
    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
    connImaging101Batch.ConnectionTimeout = 120
    connImaging101Batch.mode = adModeReadWrite
    connImaging101Batch.Open
    connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"

    
    
        '***********************************
        '* Loop Through Batch Pages
        '***********************************
        txtActionBeforeError = "*** LOOP Through Batch Pages"
        funcWriteToDebugLog Me.name, txtActionBeforeError & " intPageIndex=" & intPageIndex
        
        For intPageIndex = 1 To ListView1.ListItems.Count
            
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            txtActionBeforeError = "Call subGetBatchFieldValues"
            funcWriteToDebugLog Me.name, txtActionBeforeError & " intPageIndex=" & intPageIndex
            subGetBatchFieldValues 0
            
            ' Locate the Record
            frmIndex.ListView1.SetFocus
            txtActionBeforeError = "Call ListView1_Click"
            funcWriteToDebugLog Me.name, txtActionBeforeError & " intPageIndex=" & intPageIndex
            ListView1_Click
                
            If txtBatchPageStatus <> "Committed" Then
            
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                   If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "") Then
                        frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                    '*** VALIDATE DATE FIELD!
                    funcValidateDate intHoldFocusIndex
                    If blnDateError = True Then
                        frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                        bolSkipPage = True
                        Exit For
                    End If
                    
               Next
                   
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            
            If bolSkipPage <> True Then
                
                '* Begin the Transaction
'                connImaging101Batch.BeginTrans
'                connTTC.BeginTrans

                    Set rsImaging101BatchPage = New ADODB.Recordset
                    Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                    
                    rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                    
                    rsImaging101BatchPage.CursorLocation = adUseServer
                    rsImaging101BatchPage.CursorType = adOpenDynamic
                    rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                    
''                    rsImaging101BatchPage.Errors.Clear
                    rsImaging101BatchPage.Open
                    
                    
                    
                    'DEFAULT to Site A
                    intSiteIdIndex = 0
                
                    If CDbl(rsImaging101BatchPage.Fields("CaseID")) >= dblCaseIdCutoff Then
                        intSiteIdIndex = 1
                    End If
    
                    
                    
                    '*** Setup FTP File Directory and Name

                    Dim ttcFileName As String
                    Dim ttcFileDirectory As String
                    
'                    ' Use the BACK-SLASH (\) for INTEGER Division to get the Destination/Target Directory
'                    txtActionBeforeError = "Define ttcFileDirectory"
'                    funcWriteToDebugLog Me.name, txtActionBeforeError
'                    ttcFileDirectory = ".\" & CStr(CDbl(rsImaging101BatchPage.Fields("CaseID")) \ 1000)
                    
                    txtActionBeforeError = "Define ttcFileName"
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    
                    If intSiteIdIndex = 0 Then
                        ttcFileName = rsImaging101BatchPage.Fields("CaseNumber") & "_" & Format(Now, "yymmdd") & "_" & rsImaging101BatchPage.Fields("DocType") & "_" & gsecUserID & "_" & txtBatchPageRECID & ".JPG"
                    Else
                        ttcFileName = rsImaging101BatchPage.Fields("CaseID") & "_" & Format(Now, "yymmdd") & "_" & rsImaging101BatchPage.Fields("DocType") & "_" & gsecUserID & "_" & txtBatchPageRECID & ".JPG"
                    End If
                    
'                    ttcFileName = "./" & ttcFileDirectory & "/" & rsImaging101BatchPage.Fields("CaseID") & "/" & rsImaging101BatchPage.Fields("CaseNumber") & "_" & Format(Now, "yymmdd") & "_" & rsImaging101BatchPage.Fields("DocType") & "_" & gsecUserID & "_" & txtBatchPageRECID & ".JPG"
                    funcWriteToDebugLog Me.name, txtActionBeforeError & ttcFileName
                    
                    
                    
                    blnFTPError = False

                    
                    
                    '*** CONVERT Image to Lower Resolution
''''' DISABLED... Should NOT be necessary anymore due to Multi-stream scanning
''''                    strTempFilePath = funcExportTTCImageToJPG

                    strTempFilePath = txtBatchDirectory & "\" & txtBatchPageFileName
                    
'                    'Change Current Working Directory (CWD)
'                    Dim strTTCDestinationDirectory As String
'                    strTTCDestinationDirectory = "TTC-Docs"
'
'                    txtActionBeforeError = "Change Current Working Directory (CWD) to " & strTTCDestinationDirectory
'                    funcWriteToDebugLog Me.name, txtActionBeforeError
'                    frmFTP.FTPFile txtFTPSite, "CWD", txtFTPUserID, txtFTPPassword, strTTCDestinationDirectory, ttcFileName, False
                    
                    '*** TRANSFER File Via FTP
'''                    frmFTP.FTPFile "my.ticketclinic.com", "PUT", "pcn", "pcn00", strTempFilePath, ttcFileName, False
                    txtActionBeforeError = "PUT " & txtFTPUserID(intSiteIdIndex) & ", " & txtFTPPassword(intSiteIdIndex) & ", " & strTempFilePath & ", " & ttcFileName & ", " & False
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    frmFTP.FTPFile txtFTPSite(intSiteIdIndex), "PUT", txtFTPUserID(intSiteIdIndex), txtFTPPassword(intSiteIdIndex), strTempFilePath, ttcFileName, False
                    
                    ' Insert record into SQL only if FTP transfer did NOT return an error.
                    If blnFTPError = False Then
                    
                        '*** UPDATE TTC FIELDS
                        txtActionBeforeError = "Update Fields in 'Files101' Table in TTC Database"
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
                        
''                        txtSource = "SELECT MAX(Fileid) FROM Files"
''                        rsTTC.Open txtSource, connTTC, adOpenDynamic, adLockOptimistic
                        
                        Dim dblTTCFileID As Double
                        
''                        dblTTCFileID = rsTTC.Fields("fileid") + 1
''                        rsTTC.Close
                        
'                        txtSource = "insert into files (fileid, name, caseid, clientid, viewed ) Values (" & dblTTCFileID & ", '" & ttcFileName & "', " & rsImaging101BatchPage.Fields("CaseID") & ",  0, 0)"
                        
                        
                        txtSource = "insert into Files101 (name, caseid, clientid, viewed, missing, archive_missing, verified_missing ) Values ('" & ttcFileName & "', " & rsImaging101BatchPage.Fields("CaseID") & ",  0, 0, 0, 0, 0)"
                        
                        rsTTC(intSiteIdIndex).Open txtSource, connTTC(intSiteIdIndex), adOpenDynamic, adLockOptimistic
                        'This command will immediately "Close" the rsTTC after executing
                        
                        
                        If txtBatchGroup = "TTC RECEIVED" Then
                            ' Update the CASE with the Date the NOA was received.
                            txtActionBeforeError = "Update the CASE with the Date the NOA was received."
                            funcWriteToDebugLog Me.name, txtActionBeforeError

                           Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
                           '* Modified 2/2/2012 to work ONLY with the CASES table instead of the viewI101Cases
                            '  txtSource = "UPDATE viewI101Cases SET cases_stampedreceived = 1, cases_stampedreceiveduserid = '" & txtTTCUserID & "', cases_stampedreceiveddatetime = '" & Format(Now(), "yyyy-mm-dd HH:mm:dd") & "' WHERE cases_id = " & rsImaging101BatchPage.Fields("CaseID")
                            txtSource = "UPDATE Cases SET stamped_received = 1, stamped_received_user_id = '" & txtTTCUserID & "', stamped_received_datetime = '" & Format(Now(), "yyyy-mm-dd HH:mm:dd") & "' WHERE BarcodeId = " & rsImaging101BatchPage.Fields("CaseID")
                            rsTTC(intSiteIdIndex).Open txtSource, connTTC(intSiteIdIndex), adOpenDynamic, adLockOptimistic
                            'This command will immediately "Close" the rsTTC after executing
                        End If
                        
                        

                        
                    '****************************************************************************
                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                        rsImaging101BatchPage.Fields!BatchPageStatus = "Committed"
                        rsImaging101BatchPage.Fields!BatchPageCommitDate = Now()
                        rsImaging101BatchPage.Fields!BatchPageCommitUser = gsecUserID


                        '**************************************************************
                        '*** UPDATE THE ORIGINAL BATCH
                        '*** FLAG BATCH RECORD as Committed, set counters and Update
                        '**************************************************************
                        '*** CONNECT to Batch DB
                        Set rsImaging101Batch = New ADODB.Recordset
                        txtActionBeforeError = "Connect to Batch DB"
                        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                        rsImaging101Batch.CursorLocation = adUseServer
                        rsImaging101Batch.CursorType = adOpenDynamic
                        rsImaging101Batch.LOCKTYPE = adLockOptimistic
                        txtActionBeforeError = "Open Batch DB"
                        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                        rsImaging101Batch.Open
    
                        
                        If intPageIndex = ListView1.ListItems.Count _
                        And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
                            'If all pages are processed AND no pages requiring action are left
                            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
                        Else
                            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
                        End If
                        
                        rsImaging101Batch.Fields!BatchCommitDate = Now()
                        rsImaging101Batch.Fields!BatchCommitUser = gsecUserID
                        
                        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
                        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
                        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
                        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
                        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
                        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
            
                        
                        '****************************************************************************
                        '*** UPDATE TRANSACTIONS AND CLOSE RECORD SETS
                        
                        txtActionBeforeError = "Update BATCH PAGE " & txtBatchPageRECID
                        rsImaging101BatchPage.Update
    
                        txtActionBeforeError = "Update BATCH " & txtBatchRECID
                        rsImaging101Batch.Update
                        
                        frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
                        
    
                    Else
                        MsgBox "An error occured while Commiting this Image... I will NOT commit the rest of the images...  Please call technical support.   They may request the information displayed in the 'FTP File Transfer Status' window.", vbCritical
                        Exit Sub
                    End If
                 ''***********************************
                
''''                strOutputFileName = "C:\" & txtBatchName & ".txt"
''''                Open strOutputFileName For Append As #1
''''                Print #1, strOutputLine
''''                Close #1
            
            Else
            
                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
            
            DoEvents
        
        Next
        
        
        '**************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '**************************************************************
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
        
        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.Open
        
        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
        End If
        
        rsImaging101Batch.Update
        
        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        
        For intSiteIdIndex = 0 To 1
            connTTC(intSiteIdIndex).Close
            Set rsTTC(intSiteIdIndex) = Nothing
            Set connTTC(intSiteIdIndex) = Nothing
        Next
        
        Set rsImaging101Batch = Nothing
        Set connImaging101Batch = Nothing
        
        '*** Bring the CommitStatus Window to the front
        frmCommitStatus.SetFocus
        
Exit Sub

UPDATE_TTC_fields_names_TABLE_ERROR:
        
        txtActionBeforeError = "UPDATE_TTC_fields_names_TABLE_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Data NOT Imported]"
        funcWriteToDebugLog Me.name, txtActionBeforeError
        MsgBox txtActionBeforeError, vbExclamation
        
'        If connImaging101Batch.BeginTrans = True Then
'            connImaging101Batch.RollbackTrans
'        End If
'        If connTTC.BeginTrans = True Then
'            connTTC.RollbackTrans
'        End If
        Screen.MousePointer = vbDefault

End Sub






Private Sub CommitBatchToImaging101()
    
    On Error GoTo ERROR_HANDLER
'    On Error GoTo 0
        
        
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "----------------------------------------------------------------------------------"
    funcWriteToDebugLog Me.name, "ENTERING: CommitBatchToImaging101"
            
    
    bolErrorOccured = False
    
    Dim txtFilterStatement As String
    Dim txtFieldsList As String
    Dim txtValuesList As String
    Dim txtSystemFieldsList As String
    Dim txtSystemValuesList As String

    Dim txtOrderByList As String
    Dim txtFieldNameHold As String
                    
    Dim txtValuesListHold As String
    Dim txtDocumentRECID As Double
    Dim txtDetailRECID As Double
    Dim intDetailOrder As Integer
    Dim intPageCount As Integer
    Dim txtDestinationFilename As String
    Dim txtDestinationFileType As String
    
    Dim intPositionOfLastPeriod As Integer
    
    Dim intCopyFileRetryCount As Integer
    
    
    MainMDIForm.ActiveForm.txtChildFormMessage.Visible = True
    MainMDIForm.ActiveForm.txtChildFormMessage.Text = "COMMITTING BATCH!"
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "COMMITTING BATCH"
    MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
    MainMDIForm.ActiveForm.lstPageList.Visible = False
    MainMDIForm.ActiveForm.SpicerView1.Visible = False
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).AutoSize = sbrContents
    
    '*********************************************************
    '*** CLOSE THE IMAGE IN THE VIEWER
    '*** TO PREVENT "Runtime Error 75: File/Path Access Error"
    '*** WHEN PROCESSING SINGLE-PAGE PDF's
    '*** WHICH SEEM TO REMAIN "IN-USE" WHEN OPEN.
    'Close the document to release it
    
    funcWriteToDebugLog Me.name, "MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument"
    
     '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True

    MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument False
    
    
    
    
    
    '*************************************************************************************
    '*** IF FTP UPLOAD -- GET FTP CONNECTION INFO - BEGIN
    
    txtApplicationCommitBatchTo = txtApplicationName
    txtApplicationCommitBatchOption = funcGetFieldFromDB(RegImaging101ConnectionString, "i101Applications", "ApplicationName = '" & txtApplicationCommitBatchTo & "'", "ApplicationCommitBatchOption")
    
    If txtApplicationCommitBatchOption = "FTP Only" _
    Or txtApplicationCommitBatchOption = "Application & FTP" Then
    
    
        Dim strTempFilePath As String
        
        '************************************
        '*** Get FTP Information
        Dim txtFTPSite As String
        Dim intFTPport As Integer
        Dim txtFTPUserID As String
        Dim txtFTPPassword As String
        
        Dim txtTTCUserID As Integer
        
        
        Dim rs As ADODB.Recordset
        Dim con As ADODB.Connection
        Dim ssql As String
    
        Set con = New ADODB.Connection
        Set rs = New ADODB.Recordset
        con.Open RegImaging101ConnectionString
        
        'sql statement to select items on the drop down list
        ssql = "Select * from I101Applications where ApplicationRECID = " & txtApplicationRECID
        rs.Open ssql, con
        
            txtFTPSite = rs.Fields!FTPSite
            intFTPport = CInt(rs.Fields!ftpport)
            txtFTPUserID = rs.Fields!FTPUserID
            txtFTPPassword = rs.Fields!FTPPassword
    
            RegTTCConnectionString = rs!LookupDBConnectionString
    
    
        'Close connection and the recordset
        rs.Close
        Set rs = Nothing
        con.Close
        Set con = Nothing
        
    End If
        
    '*** IF FTP UPLOAD -- GET FTP CONNECTION INFO - END
    '*************************************************************************************
    
    
    
         '*************************************************************************************
        '* BEGIN: Loop Through Batch Pages
         '*************************************************************************************
         
        funcWriteToDebugLog Me.name, "'*************************************************************************************"
       funcWriteToDebugLog Me.name, "* BEGIN LOOP Through Batch Pages | Total Pages = " & ListView1.ListItems.Count
        
        For intPageIndex = 1 To ListView1.ListItems.Count
            
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            
            funcWriteToDebugLog Me.name, "  "
            funcWriteToDebugLog Me.name, " <<<----------------------------------------------------->>>"
            funcWriteToDebugLog Me.name, " <<<   BEGIN PROCESSING Page #  = " & intPageIndex
            
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
''''            subGetBatchFieldValues 0   ' 2013-05-08 - Jacob - Disabled because ListView1_Click below does this.
            
            ' Locate the Record
''''            frmIndex.ListView1.SetFocus
            ListView1_Click
                
            If txtBatchPageStatus <> "Committed" Then
            
                funcWriteToDebugLog Me.name, "    txtBatchPageStatus <> 'Committed' "

               '*******************************************************************
               '*** FIRST Loop to see if this record should be skipped
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable _
                   Or txtIndexValues(intIndex).Text = txtQuestionable _
                   Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator _
                   Or txtIndexValues(intIndex).Text = txtSeparator _
                   Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile _
                   Or txtIndexValues(intIndex).Text = txtDoNotFile _
                   Then
                        frmCommitStatus.txtPagesDoNotFile = frmCommitStatus.txtPagesDoNotFile + 1
                        bolSkipPage = True
                        Exit For
                   End If
               Next
               '*** END Loop to see if this record should be skipped
               '*******************************************************************

               
               '****************************************************************************
               '*** Only check for Fields Required or Valid if Not flagged for Skip
               If bolSkipPage <> True Then
               
                    funcWriteToDebugLog Me.name, "    bolSkipPage <> True | Check for Fields Required or Valid "

                    For intIndex = 0 To mebIndexValues.Count - 1
                        ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                        If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "" And txtIndexValues(intIndex).Text = "") Then
                             frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                             bolSkipPage = True
                             Exit For
                        End If
                        
                         '*** VALIDATE DATE FIELD!
                         funcValidateDate intHoldFocusIndex
                         If blnDateError = True Then
                             frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                             bolSkipPage = True
                             Exit For
                         End If
                     
                    Next
               End If   'bolSkipPage <> True
               
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            '*** END check for Fields Required or Valid if Not flagged for Skip
            '****************************************************************************
           
           
            '****************************************************************************
            '*** NOW CHECK IF PAGE SHOULD BE SKIPPED
            If bolSkipPage <> True Then
                
                    funcWriteToDebugLog Me.name, "    NOW CHECK IF PAGE SHOULD BE SKIPPED | bolSkipPage <> True "

                    '**************************************************************
                    '*** Clear variables
                    txtFilterStatement = ""
                    txtFieldsList = ""
                    txtValuesList = ""
                    txtOrderByList = ""
                    txtFieldNameHold = ""
                   
                        
                    '****************************************************************************
                    '*** Prepare the List of Fields to Compare with the Previous Image's Values
                    For intIndex = 0 To lblFieldDescription.Count - 1
            
                        '*** 2020-04-24 - Jacob - Added code to IGNORE the DOCNOTES field for the comparison
                        If UCase(txtFieldName(intIndex)) <> "DOCNOTES" Then
                        
                                If txtFieldType(intIndex).Text = "LongText" Then
                                    txtValuesList = txtValuesList & txtIndexValues(intIndex) & "|"
                                Else
                                    txtValuesList = txtValuesList & mebIndexValues(intIndex) & "|"
                                End If
                                
                        End If
                        
                    Next
                    
                    
RETRY_TRANSACTION:
                    
                    
                    '****************************************************************************
                    '*** Insert DOCUMENT record into SQL
                    '***    only if the Index Values are Different from the previous record
                    '*** OTHERWISE FIND THE EXISTING DOCUMENT RECORD
                    
                    Dim txtFullPathName As String
                    Dim strFTPUploadSourceFileName As String
                    
                    
                    
                    
                    '***********************************************************************************
                    '*** CHECK IF THIS RECORD BELONGS TO THE SAME DOCUMENT
                    If txtValuesList <> txtValuesListHold Then
                        
                        funcWriteToDebugLog Me.name, "    CHECK IF THIS RECORD BELONGS TO THE SAME DOCUMENT | txtValuesList <> txtValuesListHold "

                        '****************************************************************************
                        '*** FTP DOCUMENT SECTION - BEGIN
                        
                        If txtValuesListHold <> "" And bolCommitViaFTP = True Then
                                'This means it's the beginning of a NEW Document
                                'So let's do the FTP Routine
                                bolErrorOccured = funcFtpUploadFile(txtFTPSite, intFTPport, txtFTPUserID, txtFTPPassword, strFTPUploadSourceFileName)
                                 
                                 If bolErrorOccured Then
                                        Err.Raise 10101, "CommitBatchToImaging101", "ERROR Occured transfering file via FTP." & vbCrLf & "Page #" & intPageIndex
                                End If
                                
                        End If
                        
                        '*** FTP DOCUMENT SECTION - END
                        '****************************************************************************
                        
                        '****************************************************************************
                        '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                        '*** but do not commit the first time around
                        If txtValuesListHold <> "" Then
                        
                            funcWriteToDebugLog Me.name, "    COMMIT TRANSACTIONS AND CLOSE RECORD SETS |  txtValuesListHold <> '' "

                            subCommitTransactions
                            
                        End If
                        
                        
                        '****************************************************************************
                        '*** ESTABLISH DATABASE CONNECTIONS
                        
                            funcWriteToDebugLog Me.name, "    ESTABLISH DATABASE CONNECTIONS"

                            '**************************************************************
                            '*** Establish BATCH DB Connection
                            Set connImaging101Batch = New ADODB.Connection
                            connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
                            connImaging101Batch.ConnectionTimeout = 120
                            connImaging101Batch.mode = adModeReadWrite
                            connImaging101Batch.Open
                            connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
            
                            '**************************************************************
                            '*** CONNECT to Batch DB RESULT SET
                            Set rsImaging101Batch = New ADODB.Recordset
                            txtActionBeforeError = "Connect to Batch DB"
                            Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                            rsImaging101Batch.CursorLocation = adUseServer
                            rsImaging101Batch.CursorType = adOpenDynamic
                            rsImaging101Batch.LOCKTYPE = adLockOptimistic
                            
                            '**************************************************************
                            '*** FIND BATCH RECORD
                            
                            txtActionBeforeError = "Open Batch DB"
                            rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                            rsImaging101Batch.Open
                            
                            '**************************************************************
                            '*** CONNECT to Batch Page DB RESULT SET
                            Set rsImaging101BatchPage = New ADODB.Recordset
                            txtActionBeforeError = "Connect to Batch Pages DB"
                            Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                            rsImaging101BatchPage.CursorLocation = adUseServer
                            rsImaging101BatchPage.CursorType = adOpenDynamic
                            rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                            
                            '**************************************************************
                            '*** FIND BATCHPAGE RECORD
                    
                            txtActionBeforeError = "Open Batch Page DB"
                            funcWriteToDebugLog Me.name, "    Open Batch Page DB | FIND BATCHPAGE RECORD"

                            rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                            rsImaging101BatchPage.Open

                            
                            '**************************************************************
                            '*** Imaging101 DB Connection Setup
                           funcWriteToDebugLog Me.name, "    Imaging101 DB Connection Setup"

                            Set connImaging101 = New ADODB.Connection
                            Set cmdImaging101 = New ADODB.Command
                            connImaging101.ConnectionString = RegImaging101ConnectionString
                            connImaging101.ConnectionTimeout = 120
                            connImaging101.mode = adModeReadWrite
                            connImaging101.Open
                            connImaging101.Execute "SET LOCK_TIMEOUT -1"
                            
                            Set cmdImaging101.ActiveConnection = connImaging101
                        
                            
        
                            
                        '****************************************************************************
                        '*** PREPARE TO INSERT DOCUMENT AND DETAIL RECORDS
                           
                                 
                        '****************************************************************************
                        '* BEGIN TRANSACTIONS
                        connImaging101Batch.BeginTrans
                        connImaging101.BeginTrans
                        
                        '****************************************************************************
                        '*** CREATE NEW DOCUMENT RECORD
                       
                        txtActionBeforeError = "Get Next Control DocumentRECID"
                        funcWriteToDebugLog Me.name, txtActionBeforeError

                        txtDocumentRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "DocumentRECID")
                        
                        '*** Establish the Imaging101 Document Recordset
                        Set rsImaging101Document = New ADODB.Recordset
                       
                        ' Open the Imaging Application Document Table
                        txtActionBeforeError = "Open " & txtApplicationName
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        rsImaging101Document.Open txtApplicationName.Text, connImaging101, adOpenDynamic, adLockOptimistic, adCmdTable

                        'Add NEW Imaging101 Doocument Record
                        txtActionBeforeError = "Add New Document Record " & txtApplicationName
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        rsImaging101Document.AddNew
                        
                        'Set Field Values
                        txtActionBeforeError = "Assign Document System Field Values " & txtApplicationName
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        
                        rsImaging101Document.Fields("DocumentRECID") = txtDocumentRECID
                        rsImaging101Document.Fields("DocumentScanUserID") = rsImaging101Batch.Fields("BatchScanUser")
                        rsImaging101Document.Fields("DocumentScanDate") = rsImaging101Batch.Fields("BatchScanDate")
'                        rsImaging101Document.Fields("DocumentIndexUserID") = rsImaging101Batch.Fields("BatchIndexUser")
'                        rsImaging101Document.Fields("DocumentIndexDate") = rsImaging101Batch.Fields("BatchIndexDate")
                        rsImaging101Document.Fields("DocumentIndexUserID") = rsImaging101BatchPage.Fields("BatchPageIndexUser")
                        rsImaging101Document.Fields("DocumentIndexDate") = rsImaging101BatchPage.Fields("BatchPageIndexDate")
                        rsImaging101Document.Fields("DocumentCommitUserID") = gsecUserID
                        rsImaging101Document.Fields("DocumentCommitDate") = Now()
                        rsImaging101Document.Fields("DocumentBatchRECID") = rsImaging101Batch.Fields("BatchRECID")
                        rsImaging101Document.Fields("DocumentBatchName") = rsImaging101Batch.Fields("BatchName")
                        rsImaging101Document.Fields("BatchBoxNumber") = rsImaging101Batch.Fields("BatchBoxNumber")
                        
                        rsImaging101Document.Fields("DocumentPages") = 0
                        rsImaging101Document.Fields("DocumentImages") = 0
                        rsImaging101Document.Fields("DocumentNotes") = ""
                        rsImaging101Document.Fields("DocumentLockedBy") = ""
                        rsImaging101Document.Fields("DocumentLockedDate") = Null
                        rsImaging101Document.Fields("DocumentLockExpDate") = Null

                        '*** Update Application Fields based on Values from the BatchPage record
                        For intIndex = 0 To lblFieldDescription.Count - 1
                            txtActionBeforeError = "Assign Document User-defined Field Values " & vbCrLf & _
                                                    "Application = " & txtApplicationName & vbCrLf & _
                                                    "FieldName = " & txtFieldName(intIndex).Text & vbCrLf & _
                                                    "Value = " & rsImaging101BatchPage.Fields("" & txtFieldName(intIndex) & "")
                            
                            
                            
                            rsImaging101Document.Fields("" & txtFieldName(intIndex).Text & "") = rsImaging101BatchPage.Fields("" & txtFieldName(intIndex) & "")
                            
                        Next

                   

                        'Reset Detail Order counter
                        intDetailOrder = 0
                        intPageCount = 0
                        'Hold List of Index values to compare with next record
                        txtValuesListHold = txtValuesList
                        
                    Else
                        
                       '******************************************************************************
                       '*** ALL Field Values are the SAME
                       '*** FIND THE DOCUMENT RECORD TO APPEND PAGE TO
                        '*** Establish the Imaging101 Document Recordset
                        Set rsImaging101Document = New ADODB.Recordset
                        
                        ' Open the Imaging Application Document Table
                        Set rsImaging101Document.ActiveConnection = connImaging101
                        rsImaging101Document.CursorLocation = adUseServer
                        rsImaging101Document.CursorType = adOpenDynamic
                        rsImaging101Document.LOCKTYPE = adLockOptimistic
                        txtActionBeforeError = "Find Existing Document Record to APPEND PAGE to."
                        funcWriteToDebugLog Me.name, txtActionBeforeError

                        rsImaging101Document.Source = "SELECT * FROM " & txtApplicationName & " WHERE DocumentRECID = " & txtDocumentRECID
                        rsImaging101Document.Open
                    
                            '**************************************************************
                            '*** FIND BATCHPAGE RECORD
                    
                            txtActionBeforeError = "Close Batch Page DB"
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            rsImaging101BatchPage.Close
                            
                            subSaveBatchPageValues connImaging101Batch, rsImaging101BatchPage
                                                        
                            txtActionBeforeError = "Open Batch Page DB"
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                            rsImaging101BatchPage.Open
                   
                    End If
                    
                    
                    
                    '*****************************************************************************************************
                    '*** NOW LET'S GET THE SHOW ON THE ROAD
                  
                        
                    '**************************************************************
                    '*** Increment PageCount
                    txtActionBeforeError = "Increment PageCount."
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                   If IsNull(rsImaging101BatchPage!BatchPagePageCount) _
                    Or Not IsNumeric(rsImaging101BatchPage!BatchPagePageCount) Then
                        intPageCount = 0
                    Else
                        txtActionBeforeError = txtActionBeforeError + " BatchPageCount = " & rsImaging101BatchPage!BatchPagePageCount
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        intPageCount = intPageCount + rsImaging101BatchPage!BatchPagePageCount
                    End If
                    
                        
                            
                    '****************************************************************************
                    '****************************************************************************
                    '*** GET DETAIL SUBDIRECTORY STRUCTURE AND CREATE IT
                    
                    txtActionBeforeError = "Get Next Control DetailRECID " & txtApplicationName
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    txtDetailRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "DetailRECID")
                    
                    txtDocumentDirectoryStructure = RegRootDirToStoreObjects & "\" & _
                                                    Format(CStr(txtApplicationRECID), "0000") & _
                                                    funcGetDetailSubdirectoryString(txtDetailRECID)
                    
                    txtActionBeforeError = "Create Directory Structure: " & txtDocumentDirectoryStructure
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    funcCreateDirectoryStructure txtDocumentDirectoryStructure & ""
                    
                    
                    
                    '****************************************************************************
                    '*** DEFINE DESTINATION FILE NAME AND FILE TYPE
                
                    intPositionOfLastPeriod = InStrRev(txtBatchPageFileName, ".")
                    
                    If intPositionOfLastPeriod = 0 Then
                        txtDestinationFileType = ""
                    Else
                        txtDestinationFileType = Trim(UCase(Right(txtBatchPageFileName, Len(txtBatchPageFileName) - intPositionOfLastPeriod)))
                    End If
                    
                    
                    txtDestinationFilename = Format(CStr(txtDetailRECID), "0000000000") & "." & txtDestinationFileType
                    
                    
                    
                    '****************************************************************************
                    '*** Insert DOCUMENT DETAIL record into SQL
                
                    'Increment order counter - the counter gets reset every time a new Batch is created
                    intDetailOrder = intDetailOrder + 1

                    '*** Establish the Imaging101 Document Recordset
                    Set rsImaging101DocumentDetail = New ADODB.Recordset
                    
                    ' Open the Imaging Application Document Table
                    txtActionBeforeError = "Open " & txtApplicationName & "_Detail"
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101DocumentDetail.Open txtApplicationName.Text & "_Detail", connImaging101, adOpenDynamic, adLockOptimistic, adCmdTable

                    'Add NEW Imaging101 Doocument Record
                    txtActionBeforeError = "Add New Document DETAIL Record " & txtApplicationName & "_Detail ( " & txtDetailRECID & " )"
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101DocumentDetail.AddNew
                        
                    rsImaging101DocumentDetail.Fields("DetailRECID") = txtDetailRECID
                    rsImaging101DocumentDetail.Fields("DocumentRECID") = txtDocumentRECID
                    rsImaging101DocumentDetail.Fields("DetailOrder") = intDetailOrder
                    rsImaging101DocumentDetail.Fields("DetailCreatedDate") = Now()
                    rsImaging101DocumentDetail.Fields("DetailSubdirectory") = txtDocumentDirectoryStructure
                    rsImaging101DocumentDetail.Fields("DetailFileName") = txtDestinationFilename
                    rsImaging101DocumentDetail.Fields("DetailFileType") = txtDestinationFileType
                    rsImaging101DocumentDetail.Fields("DetailRotation") = txtBatchPageRotation
                    
                    
                    '****************************************************************************
                    '*** UPDATE the DOCUMENT record with Page & Image Counts
                    
                    rsImaging101Document.Fields("DocumentPages") = intPageCount
                    rsImaging101Document.Fields("DocumentImages") = intDetailOrder
                    
                    
                    '****************************************************************************
                    '*** COPY the file to the Storage Destination
                    Dim strSourceFile As String
                    Dim strDestinationFile As String
                
                    strSourceFile = txtBatchDirectory & "\" & txtBatchPageFileName
                    strDestinationFile = txtDocumentDirectoryStructure & "\" & txtDestinationFilename
                    
'                    txtActionBeforeError = "FileCopy [" & strSourceFile & "] TO [" & strDestinationFile & "]"
'                    FileCopy strSourceFile, strDestinationFile

                    intCopyFileRetryCount = 1
                    
                    
'*TO-DO: See if we can speed up the CopyFile
                    ''*** 2021-02-12 - Jacob - Commented FSO to use Copy File Via 32-Bit Windows API
'                    txtActionBeforeError = "CopyFile [" & strSourceFile & "] TO [" & strDestinationFile & "]"
'                    funcWriteToDebugLog Me.name, txtActionBeforeError
'                    With New FileSystemObject
'                       .CopyFile strSourceFile, strDestinationFile, True
'                    End With

                    
                    
                    ''*** 2021-05-27 - Jacob - Re-Activated Copy File Via 32-Bit Windows API
                    txtActionBeforeError = "APIFileCopy [" & strSourceFile & "] TO [" & strDestinationFile & "]"
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    Dim bSuccess As Boolean
                    bSuccess = APIFileCopy(strSourceFile, strDestinationFile, False)



                    '*** 2021-02-16 - Jacob - Implemented CopyFileByChunk
                    '*** 2021-05-27 - Jacob - Tried to use ShellExecute, because CopyFileByChunk does NOT
'                    txtActionBeforeError = "CopyFileByChunk [" & strSourceFile & "] TO [" & strDestinationFile & "]"
'                    funcWriteToDebugLog Me.name, txtActionBeforeError
'
'                    Dim lngPageSuccess As Long
'                    lngPageSuccess = CopyFileByChunk(strSourceFile, strDestinationFile, 4096)
'
'
'                    txtActionBeforeError = "*  CopyFileByChunk COMPLETE !"
'                    funcWriteToDebugLog Me.name, txtActionBeforeError

                    
                    '*** 2021-05-27 - Jacob - Tried to use ShellExecute.  Worked on MY laptop, but FAILED at ISE
'                    Dim strShellCmd As String
'                    strShellCmd = " /S /Q /Y /F " & """" & strSourceFile & """" & " " & """" & strDestinationFile & "*"""
'
'                    ShellExecute 0, vbNullString, """xcopy""", strShellCmd, vbNullString, vbMinimizedNoFocus
'
'                    Do Until funcFileExists(strDestinationFile) = True
'                        Debug.Print "Waiting..."
'                        DoEvents
'                    Loop
                    
                    '*** Now Check if the  File Exists...
                    If funcFileExists(strDestinationFile) Then
                        lngPageSuccess = 1
                        txtActionBeforeError = "*  Page File [" & strDestinationFile & "] Exists at Destination."
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                    Else
                        txtActionBeforeError = "* ERROR:  Page File [" & strDestinationFile & "] DOES NOT Exist at Destination."
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        lngPageSuccess = -1
                    End If
                    
                    
                    
                    '****************************************************************************
                    '*** Now COPY the Annotation files, if any, to the Storage Destination
                    
                    ' *** 2023-02-20 - Jacob - Re-Wrote the ANNOTATION FILE COPY Section using a loop based on Dir()
                    '                                          because the AnnotationFileListBox was causing Commit Delays of 6 to 10 MINUTES loading the filenames
                    
                    
                    txtActionBeforeError = "*  CHECK if Page has ANNOTATIONS!"
                    funcWriteToDebugLog Me.name, txtActionBeforeError

                    Dim strSourceAnnotationFile As String
                    Dim strDestinationAnnotationFile As String
                    Dim strPageNumber As String
                    
                    Dim strAnnotationFilePattern As String
                    Dim strAnnotationFileName As String
                    
                    intPositionOfLastPeriod = InStrRev(txtBatchPageFileName, ".")

                    If intPositionOfLastPeriod = 0 Then
                        strAnnotationFilePattern = txtBatchDirectory & "\" & txtBatchPageFileName & _
                                                "_*.ANN"
                    Else
                        strAnnotationFilePattern = txtBatchDirectory & "\" & Left(txtBatchPageFileName, InStrRev(txtBatchPageFileName, ".") - 1) & _
                                                "_*.ANN"
                    End If
                    
                    
                    '*** 2023-02-20 - Jacob - LOOP THROUGH ANNOTATION FILES IN DIRECTORY
                    
                    'Get the First FilePath
                    strAnnotationFileName = Dir(strAnnotationFilePattern, vbDirectory)
                    
                    ' Start looping around found filenames
                    Do While strAnnotationFileName <> ""
                    
                       ' Ignore the current directory and the encompassing directory.
                       If strAnnotationFileName <> "." And strAnnotationFileName <> ".." Then
                            Debug.Print strAnnotationFileName
                            strSourceAnnotationFile = txtBatchDirectory & "\" & strAnnotationFileName
                            Debug.Print strSourceAnnotationFile

                            DoEvents
                    
                        txtActionBeforeError = "*  ANNOTATION BATCH FileName = " & strSourceAnnotationFile
                        funcWriteToDebugLog Me.name, txtActionBeforeError

            
                         '*** Now Check if the Annotation File Exists... if NOT then DON'T try to Load/Import the Layer.
                        ' *** 2021-02-22 - Jacob - Added Check for lngPageSuccess, to make sure Page file was copied successfuly
    
                        If funcFileExists(strSourceAnnotationFile) And lngPageSuccess > 0 Then
    
                        
            '                    If AnnotationFileListBox.ListCount > 0 Then
                                    
                                    txtActionBeforeError = "*  ANNOTATIONS Detected - Prepare to Copy ANNOTATIONS!"
                                    funcWriteToDebugLog Me.name, txtActionBeforeError
            
                                        ' Build the Annotation FilePath
                                        strFullDirectoryPathForAnnotation = funcGetFullPathForAnnotation(txtApplicationRECID, txtDetailRECID)
                                        
                                        txtActionBeforeError = "Create Directory Structure: " & txtDocumentDirectoryStructure
                                        funcWriteToDebugLog Me.name, txtActionBeforeError
                                        
                                        'Extract the Page # from the Source Filename
                                        strPageNumber = Mid(AnnotationFileListBox.FileName, InStrRev(AnnotationFileListBox.FileName, "_") + 1, 6)
                                         
                                        'Create the directory if needed.
                                        funcCreateDirectoryStructure strFullDirectoryPathForAnnotation & ""
                                        
                                        strDestinationAnnotationFile = strFullDirectoryPathForAnnotation & "\" & _
                                                                Format(CStr(txtDetailRECID), "0000000000") & _
                                                                "_" & _
                                                                Format(CStr(intDetailOrder), "000000") & _
                                                                ".ANN"
                                        
            '                            txtActionBeforeError = "FileCopy [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"
            '                            FileCopy strSourceAnnotationFile, strDestinationAnnotationFile
                                        
                                        intCopyFileRetryCount = 1
                                        
                                        
                                        '*** 2021-02-16 - Jacob - Replaced FSO CopyFile with CopyFileByChunk
            '                            txtActionBeforeError = "FileCopy [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"
            '                            funcWriteToDebugLog Me.name, txtActionBeforeError
            '                            With New FileSystemObject
            '                               .CopyFile strSourceAnnotationFile, strDestinationAnnotationFile, True
            '                            End With
                                        
            '                            txtActionBeforeError = "CopyFileByChunk [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"
            '                            funcWriteToDebugLog Me.name, txtActionBeforeError
            '
            '                            Dim lngAnnotationSuccess As Long
            '                            lngAnnotationSuccess = CopyFileByChunk(strSourceAnnotationFile, strDestinationAnnotationFile, 4096)
            '
            '                            txtActionBeforeError = "*  CopyFileByChunk COMPLETE !"
            '                            funcWriteToDebugLog Me.name, txtActionBeforeError
            
            
                                        ''*** 2021-05-27 - Jacob - Re-Activated Copy File Via 32-Bit Windows API
                                        'Dim bSuccess As Boolean
                                        txtActionBeforeError = "APIFileCopy [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"
                                        funcWriteToDebugLog Me.name, txtActionBeforeError
                                        bSuccess = APIFileCopy(strSourceAnnotationFile, strDestinationAnnotationFile, False)
                    
                                        '*** Now Check if the  File Exists...
                                        If funcFileExists(strDestinationFile) Then
                                            lngPageSuccess = 1
                                            txtActionBeforeError = "*  Annotation File [" & strDestinationAnnotationFile & "] Exists at Destination."
                                            funcWriteToDebugLog Me.name, txtActionBeforeError
                                        Else
                                            txtActionBeforeError = "* ERROR:  Annotation File [" & strDestinationAnnotationFile & "] DOES NOT Exist at Destination."
                                            funcWriteToDebugLog Me.name, txtActionBeforeError
                                            lngPageSuccess = -1
                                        End If
            
            
            '                        Next
                                    
                                End If
                                
                       End If
                       
                       
                            'Get Next Annotation File Entry
                            strAnnotationFileName = Dir
                    Loop
                
                    
                    '*** Check to Make SURE the Files were copied properly!
                    txtActionBeforeError = "*  Check to Make SURE the Files were copied properly!"
                    funcWriteToDebugLog Me.name, txtActionBeforeError

       
                    '*** 20231-02-16 - Jacob - Changed to use "bSuccess" instead of funcFileExists()
'                    If Not funcFileExists(strDestinationFile) Then
                    'lngPageSuccess should be the File Size returned by CopyFileByChunk()
                    If lngPageSuccess < 1 Then
                    
                        connImaging101Batch.RollbackTrans
                        connImaging101.RollbackTrans
                        
                        txtActionBeforeError = "ERROR Occured During File Copy after Action (" & txtActionBeforeError & ")... TRANSACTION ROLLED BACK!"
                        funcWriteToDebugLog Me.name, txtActionBeforeError

                        
                        funcQuickMessage "SHOW", txtActionBeforeError
                    
                        If rsImaging101Batch.State = adStateOpen Then
                            rsImaging101Batch.Close
                        End If
                        Set rsImaging101Batch = Nothing
                        
                        If rsImaging101BatchPage.State = adStateOpen Then
                            rsImaging101BatchPage.Close
                        End If
                        Set rsImaging101BatchPage = Nothing
                        
                        Set rsImaging101 = Nothing
                        
                        Set connImaging101 = Nothing
                        Set cmdImaging101 = Nothing
                        
                        Set connImaging101Batch = Nothing
                        
                        Exit Sub
                    End If
                        
                    
                    
                    '****************************************************************************
                    '*** IF FTP - IMPORT PAGE TO SPICERDOC2 - BEGIN
                    
                   If txtApplicationCommitBatchOption = "Application & FTP" _
                   And txtCommitViaFTP = vbChecked Then
'                    Or txtApplicationCommitBatchOption = "FTP Only" Then
                    
                            'Set flag to Commit this Doc via FTP
                            bolCommitViaFTP = True
                            
                            '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
                            MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True

                            '**********************************************
                            '*** Import Page into Second Doc Control
                            Set docContents = MainMDIForm.ActiveForm.SpicerDoc2.object
                            txtFullPathName = txtBatchDirectory & "\" & txtBatchPageFileName
                            docContents.ImportPage 0, 0, IN_NEWPAGE_END, txtImageNumber, txtFullPathName
                            Set docContents = Nothing
                            
                            '**********************************************
                            '*** Build the FTP Upload File Name
                            strFTPUploadSourceFileName = funcBuildFTPUploadFileName(txtApplicationRECID, Me)
                            
                    Else
                        
                             bolCommitViaFTP = False
                            
                    End If
                    
                    '*** IF FTP - IMPORT PAGE TO SPICERDOC2 - END
                    '****************************************************************************
                
                    '****************************************************************************
                    '*** FTP UPLOAD DOCUMENT SECTION - BEGIN
                    
                    'If we have just processed the LAST PAGE... see if doc should be transferred via FTP.
                    '     It is CRITICAL that we do this BEFORE the COMMIT'S
                    '     in case we get an error during the transfer.
                    
                    If intPageIndex = ListView1.ListItems.Count Then
                        
                        If bolCommitViaFTP = True Then
                        
                            'This means it's the beginning of a NEW Document
                            'So let's do the FTP Routine
                            
                            bolErrorOccured = funcFtpUploadFile(txtFTPSite, intFTPport, txtFTPUserID, txtFTPPassword, strFTPUploadSourceFileName)
            
                            If bolErrorOccured Then
                                    Err.Raise 10101, "CommitBatchToImaging101", "ERROR Occured transfering file via FTP." & vbCrLf & "Page #" & intPageIndex
                            End If
                            
                        End If
                   
                   End If
                   
                     '*** FTP DOCUMENT SECTION - END
                    '****************************************************************************
                   
                    
                    '****************************************************************************
                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                    rsImaging101BatchPage.Fields!BatchPageStatus = "Committed"
                    rsImaging101BatchPage.Fields!BatchPageCommitDate = Now()
                    rsImaging101BatchPage.Fields!BatchPageCommitUser = gsecUserID

                    
                    '**************************************************************
                    '*** UPDATE THE ORIGINAL BATCH
                    '*** FLAG BATCH RECORD as Committed, set counters and Update
                    If intPageIndex = ListView1.ListItems.Count _
                    And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
                        'If all pages are processed AND no pages requiring action are left
                        rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
                    Else
                        rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
                    End If
                    
                    rsImaging101Batch.Fields!BatchCommitDate = Now()
                    rsImaging101Batch.Fields!BatchCommitUser = gsecUserID
                    

                    frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
                    
                    'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
                    rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
                    rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
                    rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
                    rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
                    rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
                    
                    
                   
                    
                    '****************************************************************************
                    '*** UPDATE TRANSACTIONS
                    
                    txtActionBeforeError = "Update DOCUMENT " & txtDocumentRECID
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101Document.Update
                    
                    txtActionBeforeError = "Update DOCUMENT DETAIL " & txtDetailRECID
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101DocumentDetail.Update
                    
                    txtActionBeforeError = "Update BATCH PAGE " & txtBatchPageRECID
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101BatchPage.Update

                    txtActionBeforeError = "Update BATCH " & txtBatchRECID
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101Batch.Update
                    
                    
                    
                            
                    '****************************************************************************
                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                    '*** IF ON LAST PAGE
                    If intPageIndex = ListView1.ListItems.Count Then
                            
                            txtActionBeforeError = "LAST PAGE | COMMIT TRANSACTIONS | intPageIndex = " & ListView1.ListItems.Count
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            
                            subCommitTransactions
                        
                            txtActionBeforeError = "COMMIT TRANSACTIONS SUCCESSFUL"
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            
                             txtActionBeforeError = "'************************************************************************************* "
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                       
                             txtActionBeforeError = " "
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            
                    End If
                    
                    
                    
                    
                    
                    '****************************************************************************
                    '****************************************************************************
                
''''                strOutputFileName = "C:\" & txtBatchName & ".txt"
''''                Open strOutputFileName For Append As #1
''''                Print #1, strOutputLine
''''                Close #1
            
            Else
                                            
                txtActionBeforeError = "SKIP PAGE | frmCommitStatus.txtPagesTotalSkipped = " & frmCommitStatus.txtPagesTotalSkipped + 1

                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
            
            DoEvents
        
        Next
        
        txtActionBeforeError = "END: Loop Through Batch Pages"
        funcWriteToDebugLog Me.name, txtActionBeforeError
        

        
        '*********************************************
        '* END: Loop Through Batch Pages
        '*********************************************
        
        txtActionBeforeError = "subCommitTransactions"
        funcWriteToDebugLog Me.name, txtActionBeforeError

        subCommitTransactions
        
         txtActionBeforeError = "COMMIT TRANSACTIONS for ALL Batch Pages SUCCESSFUL"
         funcWriteToDebugLog Me.name, txtActionBeforeError
         
          txtActionBeforeError = "'************************************************************************************* "
         funcWriteToDebugLog Me.name, txtActionBeforeError
    
          txtActionBeforeError = " "
         funcWriteToDebugLog Me.name, txtActionBeforeError
        
        
        '******************************************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '******************************************************************************************
        
        txtActionBeforeError = "***   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***"
        funcWriteToDebugLog Me.name, txtActionBeforeError
         
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
        
        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.Open
        
        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH COMMITTED - FULL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH COMMITTED - FULL"
            funcWriteToDebugLog Me.name, "BATCH COMMITTED - FULL"

        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH COMMITTED - PARTIAL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH COMMITTED - PARTIAL"
            funcWriteToDebugLog Me.name, "BATCH COMMITTED - PARTIAL"
        End If
        
        txtActionBeforeError = " rsImaging101Batch.Update"
        funcWriteToDebugLog Me.name, txtActionBeforeError
        
        rsImaging101Batch.Update
        
        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        
        Set rsImaging101Batch = Nothing
        
        Set connImaging101Batch = Nothing
        
        '*** Bring the CommitStatus Window to the front
        '*** Allow user to select to CLOSE or STAY
        frmCommitStatus.SetFocus
        
        
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "EXITING: CommitBatchToImaging101"
    funcWriteToDebugLog Me.name, "----------------------------------------------------------------------------------"
    funcWriteToDebugLog Me.name, " "
    
Exit Sub

ERROR_HANDLER:
        
        On Error Resume Next
        
        Dim dbErr As ADODB.Error
        Dim strErrMsg As String
        
        If (connImaging101.Errors.Count > 0) Or (connImaging101Batch.Errors.Count > 0) Then
            If (connImaging101.Errors(0).SQLState = "40001") Or (connImaging101.Errors(0).SQLState = "40001") Then
                'Handle transaction commit failure - Serialization Error
                funcQuickMessage "SHOW", "Commit Failure DURING ACTION: (" & txtActionBeforeError & ") - RETRYING TRANSATCION"
                connImaging101.RollbackTrans
                connImaging101Batch.RollbackTrans
                Resume RETRY_TRANSACTION   'Automatically retry in this example
            End If
        
            'Loop through the Errors collection for the Connection
            strErrMsg = ""
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
        
            'Loop through the Errors collection for the Connection
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101Batch" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
            
        Else
            strErrMsg = "Description: " & Err.Description & vbNewLine & _
                           "Error:       " & Err.Number & vbNewLine & vbNewLine
        End If
        
        bolErrorOccured = True
        
        ' On Error 75 - Path/File access error, allow three tries
        If (Err.Number = 75) And (intCopyFileRetryCount < 3) Then
                    intCopyFileRetryCount = intCopyFileRetryCount + 1
                    txtActionBeforeError = "CopyFile (TRY #" & intCopyFileRetryCount & ") [" & strSourceFile & "] TO [" & strDestinationFile & "]"
                    ' Try the copy again!!!
                    Resume
        End If
                        
        frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted - 1
                        
        funcQuickMessage "SHOW", "CommitBatchToImaging101 ERROR: " & strErrMsg & vbNewLine & vbNewLine & _
                        "  DURING ACTION: (" & txtActionBeforeError & ") " & vbNewLine & vbNewLine & _
                        "  [Transaction Rolled Back - Document NOT Committed]" & vbNewLine & vbNewLine & _
                        "  BatchRECID     = " & txtBatchRECID & vbNewLine & _
                        "  BatchPageRECID = " & txtBatchPageRECID & vbNewLine & _
                        "  Batch Page #   = " & intPageIndex
                        
        
        Set rsImaging101Batch = Nothing
        Set rsImaging101BatchPage = Nothing
        Set rsImaging101Document = Nothing
        Set rsImaging101DocumentDetail = Nothing
        
        Set connImaging101 = Nothing
        Set connImaging101Batch = Nothing
                    
        Set cmdImaging101 = Nothing
        
        Screen.MousePointer = vbDefault

End Sub

Private Sub CommitBatchToImaging101AutoImport()
    
    '*** 2021-06-14 - Jacob - Added CommitBatchToImaging101AutoImport() Sub
    
    On Error GoTo ERROR_HANDLER
'    On Error GoTo 0
        
        
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "----------------------------------------------------------------------------------"
    funcWriteToDebugLog Me.name, "ENTERING: CommitBatchToImaging101AutoImport()"
            
    
    bolErrorOccured = False
    
    Dim txtFilterStatement As String
    Dim txtFieldsList As String
    Dim txtValuesList As String
    Dim txtSystemFieldsList As String
    Dim txtSystemValuesList As String

    Dim txtOrderByList As String
    Dim txtFieldNameHold As String
                    
    Dim txtValuesListHold As String
    Dim txtDocumentRECID As Double
    Dim txtDetailRECID As Double
    Dim intDetailOrder As Integer
    Dim intPageCount As Integer
    Dim txtDestinationFilename As String
    Dim txtDestinationFileType As String
    
    Dim intPositionOfLastPeriod As Integer
    
    Dim intCopyFileRetryCount As Integer
    

    
    
    MainMDIForm.ActiveForm.txtChildFormMessage.Visible = True
    MainMDIForm.ActiveForm.txtChildFormMessage.Text = "COMMITTING BATCH!"
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "COMMITTING BATCH"
    MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
    MainMDIForm.ActiveForm.lstPageList.Visible = False
    MainMDIForm.ActiveForm.SpicerView1.Visible = False
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).AutoSize = sbrContents
    
    '*********************************************************
    '*** CLOSE THE IMAGE IN THE VIEWER
    '*** TO PREVENT "Runtime Error 75: File/Path Access Error"
    '*** WHEN PROCESSING SINGLE-PAGE PDF's
    '*** WHICH SEEM TO REMAIN "IN-USE" WHEN OPEN.
    'Close the document to release it
    
    funcWriteToDebugLog Me.name, "MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument"
    
     '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True

    MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument False
    
    
    
        '*************************************************************************************
        'Define the Output Directory & File Name
        
        Dim strRootDirectoryPathForHtmlSource As String
        Dim strOutputFileTempName As String
        Dim strOutputFileName As String
        
        'Get the Input File Directory & Extension
        strBatchInputFileDirectory = txtBatchDirectory
        strBatchInputFileTempExtension = ".TMP"
        strBatchInputFileExtension = ".I101"
    
        strRootDirectoryPathForHtmlSource = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationName = '" & txtApplicationName & "'", "RootDirectoryPathForHtmlSource") & ""
        If Trim(strRootDirectoryPathForHtmlSource) = "" Then
                funcQuickMessage "SHOW", "Please set the  'Full Directory Path of HTML Source Code for EXPORT'" & vbCrLf & "for Application '" & txtApplicationName & "' " & vbCrLf & "and try again."
                Exit Sub
        End If
        
        '*** 2021-11-30 - Jacob - Corrected strOutputFileName from "101Commit" to "I101Commit"
        strOutputFileTempName = strRootDirectoryPathForHtmlSource & "\I101Commit_" & Trim(txtBatchRECID) & "_" & Format(Now, "yyyy-MM-dd_hhmmss") & strBatchInputFileTempExtension
        strOutputFileName = strRootDirectoryPathForHtmlSource & "\I101Commit_" & Trim(txtBatchRECID) & "_" & Format(Now, "yyyy-MM-dd_hhmmss") & strBatchInputFileExtension

        If Not funcDirectoryExists(strRootDirectoryPathForHtmlSource) Then
                funcCreateDirectoryStructure strRootDirectoryPathForHtmlSource
        End If
    
         '*************************************************************************************
        '* BEGIN: Loop Through Batch Pages
         '*************************************************************************************
         
        funcWriteToDebugLog Me.name, "'*************************************************************************************"
       funcWriteToDebugLog Me.name, "* BEGIN LOOP Through Batch Pages | Total Pages = " & ListView1.ListItems.Count
        
        For intPageIndex = 1 To ListView1.ListItems.Count
            
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            
            funcWriteToDebugLog Me.name, "  "
            funcWriteToDebugLog Me.name, " <<<----------------------------------------------------->>>"
            funcWriteToDebugLog Me.name, " <<<   BEGIN PROCESSING Page #  = " & intPageIndex
            
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
''''            subGetBatchFieldValues 0   ' 2013-05-08 - Jacob - Disabled because ListView1_Click below does this.
            
            ' Locate the Record
''''            frmIndex.ListView1.SetFocus
            ListView1_Click
                
            If txtBatchPageStatus <> "Committed" Then
            
                funcWriteToDebugLog Me.name, "    txtBatchPageStatus <> 'Committed' "

               '*******************************************************************
               '*** FIRST Loop to see if this record should be skipped
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable _
                   Or txtIndexValues(intIndex).Text = txtQuestionable _
                   Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator _
                   Or txtIndexValues(intIndex).Text = txtSeparator _
                   Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile _
                   Or txtIndexValues(intIndex).Text = txtDoNotFile _
                   Then
                        frmCommitStatus.txtPagesDoNotFile = frmCommitStatus.txtPagesDoNotFile + 1
                        bolSkipPage = True
                        Exit For
                   End If
               Next
               '*** END Loop to see if this record should be skipped
               '*******************************************************************

               
               '****************************************************************************
               '*** Only check for Fields Required or Valid if Not flagged for Skip
               If bolSkipPage <> True Then
               
                    funcWriteToDebugLog Me.name, "    bolSkipPage <> True | Check for Fields Required or Valid "

                    For intIndex = 0 To mebIndexValues.Count - 1
                        ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                        If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "" And txtIndexValues(intIndex).Text = "") Then
                             frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                             bolSkipPage = True
                             Exit For
                        End If
                        
                         '*** VALIDATE DATE FIELD!
                         funcValidateDate intHoldFocusIndex
                         If blnDateError = True Then
                             frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                             bolSkipPage = True
                             Exit For
                         End If
                     
                    Next
               End If   'bolSkipPage <> True
               
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            '*** END check for Fields Required or Valid if Not flagged for Skip
            '****************************************************************************
           
           
            '****************************************************************************
            '*** NOW CHECK IF PAGE SHOULD BE SKIPPED
            If bolSkipPage <> True Then
                
                    funcWriteToDebugLog Me.name, "    NOW CHECK IF PAGE SHOULD BE SKIPPED | bolSkipPage <> True "

                    '**************************************************************
                    '*** Clear variables
                    txtFilterStatement = ""
                    txtFieldsList = ""
                    txtValuesList = ""
                    txtOrderByList = ""
                    txtFieldNameHold = ""
                   
                        
'''                    '****************************************************************************
'''                    '*** Prepare the List of Fields to Compare with the Previous Image's Values
'''                    For intIndex = 0 To lblFieldDescription.Count - 1
'''
'''                        '*** 2020-04-24 - Jacob - Added code to IGNORE the DOCNOTES field for the comparison
'''                        If UCase(txtFieldName(intIndex)) <> "DOCNOTES" Then
'''
'''                                If txtFieldType(intIndex).Text = "LongText" Then
'''                                    txtValuesList = txtValuesList & txtIndexValues(intIndex) & "|"
'''                                Else
'''                                    txtValuesList = txtValuesList & mebIndexValues(intIndex) & "|"
'''                                End If
'''
'''                        End If
'''
'''                    Next
                    
                    
RETRY_TRANSACTION:
                    

                        
                        
                        '****************************************************************************
                        '*** ESTABLISH DATABASE CONNECTIONS
                        
                            funcWriteToDebugLog Me.name, "    ESTABLISH DATABASE CONNECTIONS"

                            '**************************************************************
                            '*** Establish BATCH DB Connection
                            Set connImaging101Batch = New ADODB.Connection
                            connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
                            connImaging101Batch.ConnectionTimeout = 120
                            connImaging101Batch.mode = adModeReadWrite
                            connImaging101Batch.Open
                            connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
            
                            '**************************************************************
                            '*** CONNECT to Batch DB RESULT SET
                            Set rsImaging101Batch = New ADODB.Recordset
                            txtActionBeforeError = "Connect to Batch DB"
                            Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                            rsImaging101Batch.CursorLocation = adUseServer
                            rsImaging101Batch.CursorType = adOpenDynamic
                            rsImaging101Batch.LOCKTYPE = adLockOptimistic
                            
                            '**************************************************************
                            '*** FIND BATCH RECORD
                            
                            txtActionBeforeError = "Open Batch DB"
                            rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                            rsImaging101Batch.Open
                            
                            '**************************************************************
                            '*** CONNECT to Batch Page DB RESULT SET
                            Set rsImaging101BatchPage = New ADODB.Recordset
                            txtActionBeforeError = "Connect to Batch Pages DB"
                            Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                            rsImaging101BatchPage.CursorLocation = adUseServer
                            rsImaging101BatchPage.CursorType = adOpenDynamic
                            rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                            
                            '**************************************************************
                            '*** FIND BATCHPAGE RECORD
                    
                            txtActionBeforeError = "Open Batch Page DB"
                            funcWriteToDebugLog Me.name, "    Open Batch Page DB | FIND BATCHPAGE RECORD"

                            rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                            rsImaging101BatchPage.Open

                           
                                 
                        '****************************************************************************
                        '* BEGIN TRANSACTIONS
                        connImaging101Batch.BeginTrans

                        
                    
                    
                    '*****************************************************************************************************
                    '*** NOW LET'S GET THE SHOW ON THE ROAD
                  
                    
                    '**********************************************************************************
                    '*** Set up the Output Line
                    
                     Dim strOutputLine As String
                     
                   'Prepend "I101Commit|" to notify AutoImport that we will be adding the BatchRECID, UserID's and Dates
                    strOutputLine = "I101Commit" & "|" & txtApplicationRECID.Text & "|" & gsecUserID & "|" & txtBatchDirectory & "\" & txtBatchPageFileName

                    '*** TODO - ADD BATCHRECID AND BATCHNAME
                    strOutputLine = strOutputLine & "|" & rsImaging101Batch.Fields("BatchScanUser")
                    strOutputLine = strOutputLine & "|" & rsImaging101Batch.Fields("BatchScanDate")
                    strOutputLine = strOutputLine & "|" & rsImaging101BatchPage.Fields("BatchPageIndexUser")
                    strOutputLine = strOutputLine & "|" & rsImaging101BatchPage.Fields("BatchPageIndexDate")
                    strOutputLine = strOutputLine & "|" & gsecUserID
                    strOutputLine = strOutputLine & "|" & Now()
                    strOutputLine = strOutputLine & "|" & rsImaging101Batch.Fields("BatchRECID")
                    strOutputLine = strOutputLine & "|" & rsImaging101Batch.Fields("BatchName")
                    strOutputLine = strOutputLine & "|" & rsImaging101Batch.Fields("BatchBoxNumber")
                        
                        
                    '* Loop Through Fields
                    For intIndex = 0 To mebIndexValues.Count - 1
                    
                         ' Add field to the OutputLine since there were no exceptions!
                         If txtFieldType(intIndex) = "Date" Then
                            
                            If IsDate(mebIndexValues(intIndex).FormattedText) Then
                                    'If it's a VALID "Date" then reformat it for Oracle IPM
                                    strOutputLine = strOutputLine & "|" & Format(mebIndexValues(intIndex).FormattedText, "yyyy-MM-dd")
                            Else
                                    'If NOT a VALID date... simply use a blank.
                                    strOutputLine = strOutputLine & "|" & ""
                            End If
                            
                         Else
                            
                            strOutputLine = strOutputLine & "|" & Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format)
                           
                         End If
                         
                    Next
                       
                
                    '***********************************************************************************
                    '*** WRITE OUTPUT LINE
                    
                    Open strOutputFileTempName For Append As #1
                    Print #1, strOutputLine
                    Close #1
                    
                    '***********************************************************************************
                  
                  
                        
                    '**************************************************************
                    '*** Increment PageCount
                    txtActionBeforeError = "Increment PageCount."
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                   If IsNull(rsImaging101BatchPage!BatchPagePageCount) _
                    Or Not IsNumeric(rsImaging101BatchPage!BatchPagePageCount) Then
                        intPageCount = 0
                    Else
                        txtActionBeforeError = txtActionBeforeError + " BatchPageCount = " & rsImaging101BatchPage!BatchPagePageCount
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        intPageCount = intPageCount + rsImaging101BatchPage!BatchPagePageCount
                    End If
                    
                        

                   
                    
                    '****************************************************************************
                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                    rsImaging101BatchPage.Fields!BatchPageStatus = "Committed"
                    rsImaging101BatchPage.Fields!BatchPageCommitDate = Now()
                    rsImaging101BatchPage.Fields!BatchPageCommitUser = gsecUserID






                    
                    '**************************************************************
                    '*** UPDATE THE ORIGINAL BATCH
                    '*** FLAG BATCH RECORD as Committed, set counters and Update
                    If intPageIndex = ListView1.ListItems.Count _
                    And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
                        'If all pages are processed AND no pages requiring action are left
                        rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
                    Else
                        rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
                    End If
                    
                    rsImaging101Batch.Fields!BatchCommitDate = Now()
                    rsImaging101Batch.Fields!BatchCommitUser = gsecUserID
                    

                    frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
                    
                    'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
                    rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
                    rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
                    rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
                    rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
                    rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
                    
                    
                   
                    
                    '****************************************************************************
                    '*** UPDATE TRANSACTIONS
                    
                    txtActionBeforeError = "Update BATCH PAGE " & txtBatchPageRECID
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101BatchPage.Update

                    txtActionBeforeError = "Update BATCH " & txtBatchRECID
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    rsImaging101Batch.Update
                    
                    
                    
                            
                    '****************************************************************************
                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                    '*** IF ON LAST PAGE
                    If intPageIndex = ListView1.ListItems.Count Then
                            
                            txtActionBeforeError = "LAST PAGE | COMMIT TRANSACTIONS | intPageIndex = " & ListView1.ListItems.Count
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            
                            subCommitTransactions
                        
                            txtActionBeforeError = "COMMIT TRANSACTIONS SUCCESSFUL"
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            
                             txtActionBeforeError = "'************************************************************************************* "
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                       
                             txtActionBeforeError = " "
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            
                    End If
                    
                    
                    
                    
                    
                    '****************************************************************************
                    '****************************************************************************
                
''''                strOutputFileName = "C:\" & txtBatchName & ".txt"
''''                Open strOutputFileName For Append As #1
''''                Print #1, strOutputLine
''''                Close #1
            
            Else
                                            
                txtActionBeforeError = "SKIP PAGE | frmCommitStatus.txtPagesTotalSkipped = " & frmCommitStatus.txtPagesTotalSkipped + 1

                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
            
            DoEvents
        
        Next
        
        txtActionBeforeError = "END: Loop Through Batch Pages"
        funcWriteToDebugLog Me.name, txtActionBeforeError
        

        
        '*********************************************
        '* END: Loop Through Batch Pages
        '*********************************************
        
        txtActionBeforeError = "subCommitTransactions"
        funcWriteToDebugLog Me.name, txtActionBeforeError

        subCommitTransactions
        
         txtActionBeforeError = "COMMIT TRANSACTIONS for ALL Batch Pages SUCCESSFUL"
         funcWriteToDebugLog Me.name, txtActionBeforeError
         
          txtActionBeforeError = "'************************************************************************************* "
         funcWriteToDebugLog Me.name, txtActionBeforeError
    
          txtActionBeforeError = " "
         funcWriteToDebugLog Me.name, txtActionBeforeError
        
        
        '******************************************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '******************************************************************************************
        
        txtActionBeforeError = "***   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***"
        funcWriteToDebugLog Me.name, txtActionBeforeError
         
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
        
        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.Open
        
        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH COMMITTED - FULL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH COMMITTED - FULL"
            funcWriteToDebugLog Me.name, "BATCH COMMITTED - FULL"

        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH COMMITTED - PARTIAL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH COMMITTED - PARTIAL"
            funcWriteToDebugLog Me.name, "BATCH COMMITTED - PARTIAL"
        End If
        
        txtActionBeforeError = " rsImaging101Batch.Update"
        funcWriteToDebugLog Me.name, txtActionBeforeError
        
        rsImaging101Batch.Update
        
        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        
        Set rsImaging101Batch = Nothing
        
        Set connImaging101Batch = Nothing
        
        
        
    '*** RENAME the TEMP File to hand it off to Imaging101AutoImport
    funcWriteToDebugLog Me.name, "RENAME " & strOutputFileTempName & " As " & strOutputFileName
    Name strOutputFileTempName As strOutputFileName
    
    
        
        '*** Bring the CommitStatus Window to the front
        '*** Allow user to select to CLOSE or STAY
        frmCommitStatus.SetFocus
        
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "EXITING: CommitBatchToImaging101"
    funcWriteToDebugLog Me.name, "----------------------------------------------------------------------------------"
    funcWriteToDebugLog Me.name, " "
    
Exit Sub

ERROR_HANDLER:
        
        On Error Resume Next
        
        Dim dbErr As ADODB.Error
        Dim strErrMsg As String
        
        If (connImaging101.Errors.Count > 0) Or (connImaging101Batch.Errors.Count > 0) Then
            If (connImaging101.Errors(0).SQLState = "40001") Or (connImaging101.Errors(0).SQLState = "40001") Then
                'Handle transaction commit failure - Serialization Error
                funcQuickMessage "SHOW", "Commit Failure DURING ACTION: (" & txtActionBeforeError & ") - RETRYING TRANSATCION"
                connImaging101.RollbackTrans
                connImaging101Batch.RollbackTrans
                Resume RETRY_TRANSACTION   'Automatically retry in this example
            End If
        
            'Loop through the Errors collection for the Connection
            strErrMsg = ""
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
        
            'Loop through the Errors collection for the Connection
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101Batch" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
            
        Else
            strErrMsg = "Description: " & Err.Description & vbNewLine & _
                           "Error:       " & Err.Number & vbNewLine & vbNewLine
        End If
        
        bolErrorOccured = True
        
        ' On Error 75 - Path/File access error, allow three tries
        If (Err.Number = 75) And (intCopyFileRetryCount < 3) Then
                    intCopyFileRetryCount = intCopyFileRetryCount + 1
                    txtActionBeforeError = "CopyFile (TRY #" & intCopyFileRetryCount & ") [" & strSourceFile & "] TO [" & strDestinationFile & "]"
                    ' Try the copy again!!!
                    Resume
        End If
                        
        frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted - 1
                        
        funcQuickMessage "SHOW", "CommitBatchToImaging101 ERROR: " & strErrMsg & vbNewLine & vbNewLine & _
                        "  DURING ACTION: (" & txtActionBeforeError & ") " & vbNewLine & vbNewLine & _
                        "  [Transaction Rolled Back - Document NOT Committed]" & vbNewLine & vbNewLine & _
                        "  BatchRECID     = " & txtBatchRECID & vbNewLine & _
                        "  BatchPageRECID = " & txtBatchPageRECID & vbNewLine & _
                        "  Batch Page #   = " & intPageIndex
                        
        
        Set rsImaging101Batch = Nothing
        Set rsImaging101BatchPage = Nothing
        Set rsImaging101Document = Nothing
        Set rsImaging101DocumentDetail = Nothing
        
        Set connImaging101 = Nothing
        Set connImaging101Batch = Nothing
                    
        Set cmdImaging101 = Nothing
        
        Screen.MousePointer = vbDefault

End Sub






Private Sub CommitBatchToHMISSoftware()
    
    On Error GoTo ERROR_HANDLER
'    On Error GoTo 0
        
    bolErrorOccured = False
    
    Dim txtFilterStatement As String
    Dim txtFieldsList As String
    Dim txtValuesList As String
    Dim txtSystemFieldsList As String
    Dim txtSystemValuesList As String

    Dim txtOrderByList As String
    Dim txtFieldNameHold As String
                    
    Dim txtValuesListHold As String
    Dim txtDocumentRECID As Double
    Dim txtDetailRECID As Double
    Dim intDetailOrder As Integer
    Dim intPageCount As Integer
    Dim txtDestinationFilename As String
    Dim txtDestinationFileType As String
    
    
        '***********************************
        '* Loop Through Batch Pages
        '***********************************
        For intPageIndex = 1 To ListView1.ListItems.Count
            
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            subGetBatchFieldValues 0
            
            ' Locate the Record
            frmIndex.ListView1.SetFocus
            ListView1_Click
                
            If txtBatchPageStatus <> "Committed" Then
            
               'FIRST Loop to see if this record should be skipped
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile Then
                        frmCommitStatus.txtPagesDoNotFile = frmCommitStatus.txtPagesDoNotFile + 1
                        bolSkipPage = True
                        Exit For
                   End If
               Next
                
               'Only check for Fields Required or Valid if Not flagged for Skip
               If bolSkipPage <> True Then
               
                    For intIndex = 0 To mebIndexValues.Count - 1
                        ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                        If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "") Then
                             frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                             bolSkipPage = True
                             Exit For
                        End If
                        
                         '*** VALIDATE DATE FIELD!
                         funcValidateDate intHoldFocusIndex
                         If blnDateError = True Then
                             frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                             bolSkipPage = True
                             Exit For
                         End If
                     
                    Next
               End If   'bolSkipPage <> True
               
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            
            If bolSkipPage <> True Then
                

                    '**************************************************************
                    '*** Establish BATCH DB Connection
                    Set connImaging101Batch = New ADODB.Connection
                    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
                    connImaging101Batch.ConnectionTimeout = 120
                    connImaging101Batch.mode = adModeReadWrite
                    connImaging101Batch.Open
                    connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
    
                    '**************************************************************
                    '*** CONNECT to Batch DB
                    Set rsImaging101Batch = New ADODB.Recordset
                    txtActionBeforeError = "Connect to Batch DB"
                    Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                    rsImaging101Batch.CursorLocation = adUseServer
                    rsImaging101Batch.CursorType = adOpenDynamic
                    rsImaging101Batch.LOCKTYPE = adLockOptimistic
                    txtActionBeforeError = "Open Batch DB"
                    rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                    rsImaging101Batch.Open
                    
                    '**************************************************************
                    '*** CONNECT to Batch Page DB
                    Set rsImaging101BatchPage = New ADODB.Recordset
                    txtActionBeforeError = "Connect to Batch Pages DB"
                    Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                    rsImaging101BatchPage.CursorLocation = adUseServer
                    rsImaging101BatchPage.CursorType = adOpenDynamic
                    rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                    txtActionBeforeError = "Open Batch Page DB"
                    rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                    rsImaging101BatchPage.Open
                    
                    
                    '**********************************
                    '*** Imaging101 DB Connection Setup
                    Set connImaging101 = New ADODB.Connection
                    Set cmdImaging101 = New ADODB.Command
                    connImaging101.ConnectionString = RegImaging101ConnectionString
                    connImaging101.ConnectionTimeout = 120
                    connImaging101.mode = adModeReadWrite
                    connImaging101.Open
                    connImaging101.Execute "SET LOCK_TIMEOUT -1"
                    
                    Set cmdImaging101.ActiveConnection = connImaging101
                
                    

                    
                    '**************************************************************
                    '*** PREPARE TO INSERT DOCUMENT AND DETAIL RECORDS
                   
                    '*** Clear variables
                    txtFilterStatement = ""
                    txtFieldsList = ""
                    txtValuesList = ""
                    txtOrderByList = ""
                    txtFieldNameHold = ""
                    
                        
                    '****************************************************************************
                    '*** Prepare the List of Fields to Compare with the Previous Image's Values
                    For intIndex = 0 To lblFieldDescription.Count - 1
                        txtValuesList = txtValuesList & mebIndexValues(intIndex) & "|"
                    Next
                    
                    '****************************************************************************
                    '* BEGIN TRANSACTIONS
RETRY_TRANSACTION:
                    connImaging101Batch.BeginTrans
                    connImaging101.BeginTrans
                    
                    '****************************************************************************
                    '*** Close the Spicer Document Control to discard previous pages, and Add the current page
                    '***    only if the Index Values are Different from the previous record
                    
                    '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
                    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True
                    
                    If ((txtValuesList <> txtValuesListHold) _
                        And (intPageIndex > 1)) _
                        Or (intPageIndex = ListView1.ListItems.Count) Then
                        
                            If intPageIndex = ListView1.ListItems.Count Then
                                '****************************************************************************
                                '*** IMPORT THE FILE TO THE DOCUMENT - To Make sure we Get the LAST Image
                                subImportFileToDoc txtBatchDirectory, txtBatchPageFileName
                                txtValuesListHold = txtValuesList
                            End If
                            
                            Dim txtDestinationDirectory As String
                            txtDestinationDirectory = RegRootDirToStoreObjects & "\" & _
                                                        txtApplicationName & "\"
                            
                            funcCreateDirectoryStructure txtDestinationDirectory
                    
                            Dim txtValuesArray() As String
                            txtValuesArray = Split(txtValuesListHold, "|")
                            
                                
                            If InStr(1, UCase(txtApplicationName), "LOT") > 0 Then
                                Dim txtSpacesArray() As String
                                Dim iSpacesLoop As Integer
                                
                                txtSpacesArray = Split(txtValuesArray(1), "+")
                                

                                For iSpacesLoop = 0 To UBound(txtSpacesArray)
                                    'Format = contract-LC-space
                                    txtDestinationFilename = txtDestinationDirectory & "\" & txtValuesArray(0) + "-" & txtValuesArray(2) & "-" & txtSpacesArray(iSpacesLoop) & ".TIF"
                                    SpicerDoc1.Save 0, False, API_MPAGE_TIFF, txtDestinationFilename, txtDestinationFilename
                                Next
                                
                            Else
                                'Format = contract-doctype
                                txtDestinationFilename = txtDestinationDirectory & "\" & txtValuesArray(0) + "-" + txtValuesArray(1) & ".TIF"
                                SpicerDoc1.Save 0, False, API_MPAGE_TIFF, txtDestinationFilename, txtDestinationFilename
                            End If
                            
                        SpicerDoc1.CloseDocument False
                            
                            
                        'Reset Detail Order counter
                        intDetailOrder = 0
                        intPageCount = 0
                        'Hold List of Index values to compare with next record
                    
                    End If
                                    
                            
                    
                    '****************************************************************************
                    '*** IMPORT THE FILE TO THE DOCUMENT
                    
                    subImportFileToDoc txtBatchDirectory, txtBatchPageFileName
                    txtValuesListHold = txtValuesList
                    
                   
                        
                    '**************************************************************
                    '*** Increment PageCount
                    If IsNull(rsImaging101BatchPage!BatchPagePageCount) Then
                        intPageCount = 0
                    Else
                        intPageCount = intPageCount + rsImaging101BatchPage!BatchPagePageCount
                    End If
                    
                        
                    '****************************************************************************
'                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
'                    rsImaging101BatchPage.Fields!BatchPageStatus = "Exported"
'                    rsImaging101BatchPage.Fields!BatchPageCommitDate = Now()
'                    rsImaging101BatchPage.Fields!BatchPageCommitUser = gsecUserID
'
'
'                    '**************************************************************
'                    '*** UPDATE THE ORIGINAL BATCH
'                    '*** FLAG BATCH RECORD as Committed, set counters and Update
'                    If intPageIndex = ListView1.ListItems.Count _
'                    And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
'                        'If all pages are processed AND no pages requiring action are left
'                        rsImaging101Batch.Fields!BatchCommitStatus = "Exported-FULL"
'                    Else
'                        rsImaging101Batch.Fields!BatchCommitStatus = "Exported-PARTIAL"
'                    End If
'
'                    rsImaging101Batch.Fields!BatchCommitDate = Now()
'                    rsImaging101Batch.Fields!BatchCommitUser = gsecUserID
'
'                    'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
'                    rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
'                    rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
'                    rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
'                    rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
'                    rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
'
'
'                    '****************************************************************************
'                    '*** UPDATE TRANSACTIONS AND CLOSE RECORD SETS
'
'                    txtActionBeforeError = "Update BATCH PAGE " & txtBatchPageRECID
'                    rsImaging101BatchPage.Update
'
'                    txtActionBeforeError = "Update BATCH " & txtBatchRECID
'                    rsImaging101Batch.Update
'
'                    frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
'
'
'
'                    '****************************************************************************
'                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
'                    connImaging101Batch.CommitTrans
'                    connImaging101.CommitTrans
'
'                    If rsImaging101Batch.State = adStateOpen Then
'                        rsImaging101Batch.Close
'                    End If
'                    Set rsImaging101Batch = Nothing
'
'                    If rsImaging101BatchPage.State = adStateOpen Then
'                        rsImaging101BatchPage.Close
'                    End If
'                    Set rsImaging101BatchPage = Nothing
'
'
'                    Set cmdImaging101 = Nothing
'
'                    Set connImaging101 = Nothing
'                    Set connImaging101Batch = Nothing
'
'
'
'                    '****************************************************************************
'                    '****************************************************************************
'
'''''                strOutputFileName = "C:\" & txtBatchName & ".txt"
'''''                Open strOutputFileName For Append As #1
'''''                Print #1, strOutputLine
'''''                Close #1
'
            Else

                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
'
'            '*** Update Application Fields
'            txtActionBeforeError = "Assign Document User-defined Field Values " & txtApplicationName
'            For intIndex = 0 To lblFieldDescription.Count - 1
'                rsImaging101Document.Fields("" & txtFieldName(intIndex).Text & "") = rsImaging101BatchPage.Fields("" & txtFieldName(intIndex) & "")
'            Next
'            DoEvents
'
        Next
        
        '**************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '**************************************************************
        
        '**************************************************************
        '*** Establish BATCH DB Connection
'        Set connImaging101Batch = New ADODB.Connection
'        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
'        connImaging101Batch.ConnectionTimeout = 30
'        connImaging101Batch.Mode = adModeReadWrite
'        connImaging101Batch.Open
'        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
'
'        '*** CONNECT to Batch DB
'        Set rsImaging101Batch = New ADODB.Recordset
'        txtActionBeforeError = "Connect to Batch DB"
'        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
'        rsImaging101Batch.CursorLocation = adUseServer
'        rsImaging101Batch.CursorType = adOpenDynamic
'        rsImaging101Batch.LockType = adLockOptimistic
'        txtActionBeforeError = "Open Batch DB"
'        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
'        rsImaging101Batch.Open
'
'        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
'        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
'        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
'        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
'        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
'        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
'
'        '*** FLAG BATCH RECORD as Committed, set counters and Update
'        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
'            'If all pages are processed AND no pages requiring action are left
'            rsImaging101Batch.Fields!BatchCommitStatus = "Exported-FULL"
'        Else
'            rsImaging101Batch.Fields!BatchCommitStatus = "Exported-PARTIAL"
'        End If
'
'        rsImaging101Batch.Update
'
'        If rsImaging101Batch.State = adStateOpen Then
'            rsImaging101Batch.Close
'        End If
'        Set rsImaging101Batch = Nothing
'
'        Set connImaging101Batch = Nothing

'''        'Close the Spicer Document Control - discard loaded images
        
        '*** Bring the CommitStatus Window to the front
        frmCommitStatus.SetFocus
        
Exit Sub

ERROR_HANDLER:
        
        
        Dim dbErr As ADODB.Error
        Dim strErrMsg As String
        
        If (connImaging101.Errors.Count > 0) Or (connImaging101Batch.Errors.Count > 0) Then
            If (connImaging101.Errors(0).SQLState = "40001") Or (connImaging101.Errors(0).SQLState = "40001") Then
                'Handle transaction commit failure - Serialization Error
                funcQuickMessage "SHOW", "Commit Failure DURING ACTION: (" & txtActionBeforeError & ") - RETRYING TRANSATCION"
                connImaging101.RollbackTrans
                connImaging101Batch.RollbackTrans
                Resume RETRY_TRANSACTION   'Automatically retry in this example
            End If
        
            'Loop through the Errors collection for the Connection
            strErrMsg = ""
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
        
            'Loop through the Errors collection for the Connection
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101Batch" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
            
        Else
            strErrMsg = "Description: " & Err.Description & vbNewLine & _
                           "Error:       " & Err.Number & vbNewLine & vbNewLine
        End If
        
        
        bolErrorOccured = True
        
        funcQuickMessage "SHOW", "CommitBatchToImaging101 ERROR: " & strErrMsg & vbNewLine & vbNewLine & _
                        "  DURING ACTION: (" & txtActionBeforeError & ") " & vbNewLine & vbNewLine & _
                        "  [Transaction Rolled Back - Page NOT Committed]" & vbNewLine & vbNewLine & _
                        "  BatchRECID     = " & txtBatchRECID & vbNewLine & _
                        "  BatchPageRECID = " & txtBatchPageRECID & vbNewLine & _
                        "  Batch Page #   = " & intPageIndex
                        
        On Error Resume Next
        
        Set rsImaging101Batch = Nothing
        Set rsImaging101BatchPage = Nothing
        Set rsImaging101Document = Nothing
        Set rsImaging101DocumentDetail = Nothing
        
        Set connImaging101 = Nothing
        Set connImaging101Batch = Nothing
                    
        Set cmdImaging101 = Nothing
        
        Screen.MousePointer = vbDefault

End Sub

Private Sub CommitBatchToISAC()
    
    On Error GoTo ERROR_HANDLER
'    On Error GoTo 0
        
    bolErrorOccured = False
    
    Dim txtFilterStatement As String
    Dim txtFieldsList As String
    Dim txtValuesList As String
    Dim txtSystemFieldsList As String
    Dim txtSystemValuesList As String

    Dim txtOrderByList As String
    Dim txtFieldNameHold As String
                    
    Dim txtValuesListHold As String
    Dim txtDocumentRECID As Double
    Dim txtDetailRECID As Double
    Dim intDetailOrder As Integer
    Dim intPageCount As Integer
    Dim txtDestinationFilename As String
    Dim txtDestinationFileType As String
    
    
        '***********************************
        '* Loop Through Batch Pages
        '***********************************
        For intPageIndex = 1 To ListView1.ListItems.Count
            
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            subGetBatchFieldValues 0
            
            ' Locate the Record
            frmIndex.ListView1.SetFocus
            ListView1_Click
                
            If txtBatchPageStatus <> "Committed" Then
            
               'FIRST Loop to see if this record should be skipped
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile Then
                        frmCommitStatus.txtPagesDoNotFile = frmCommitStatus.txtPagesDoNotFile + 1
                        bolSkipPage = True
                        Exit For
                   End If
               Next
                
               'Only check for Fields Required or Valid if Not flagged for Skip
               If bolSkipPage <> True Then
               
                    For intIndex = 0 To mebIndexValues.Count - 1
                        ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                        If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "") Then
                             frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                             bolSkipPage = True
                             Exit For
                        End If
                        
                         '*** VALIDATE DATE FIELD!
                         funcValidateDate intHoldFocusIndex
                         If blnDateError = True Then
                             frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                             bolSkipPage = True
                             Exit For
                         End If
                     
                    Next
               End If   'bolSkipPage <> True
               
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            
            If bolSkipPage <> True Then
                

                    '**************************************************************
                    '*** Establish BATCH DB Connection
                    Set connImaging101Batch = New ADODB.Connection
                    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
                    connImaging101Batch.ConnectionTimeout = 120
                    connImaging101Batch.mode = adModeReadWrite
                    connImaging101Batch.Open
                    connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
    
                    '**************************************************************
                    '*** CONNECT to Batch DB
                    Set rsImaging101Batch = New ADODB.Recordset
                    txtActionBeforeError = "Connect to Batch DB"
                    Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                    rsImaging101Batch.CursorLocation = adUseServer
                    rsImaging101Batch.CursorType = adOpenDynamic
                    rsImaging101Batch.LOCKTYPE = adLockOptimistic
                    txtActionBeforeError = "Open Batch DB"
                    rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                    rsImaging101Batch.Open
                    
                    '**************************************************************
                    '*** CONNECT to Batch Page DB
                    Set rsImaging101BatchPage = New ADODB.Recordset
                    txtActionBeforeError = "Connect to Batch Pages DB"
                    Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                    rsImaging101BatchPage.CursorLocation = adUseServer
                    rsImaging101BatchPage.CursorType = adOpenDynamic
                    rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                    txtActionBeforeError = "Open Batch Page DB"
                    rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                    rsImaging101BatchPage.Open
                    
                    
                    '**********************************
                    '*** Imaging101 DB Connection Setup
                    Set connImaging101 = New ADODB.Connection
                    Set cmdImaging101 = New ADODB.Command
                    connImaging101.ConnectionString = RegImaging101ConnectionString
                    connImaging101.ConnectionTimeout = 120
                    connImaging101.mode = adModeReadWrite
                    connImaging101.Open
                    connImaging101.Execute "SET LOCK_TIMEOUT -1"
                    
                    Set cmdImaging101.ActiveConnection = connImaging101
                
                    

                    
                    '**************************************************************
                    '*** PREPARE TO INSERT DOCUMENT AND DETAIL RECORDS
                   
                    '*** Clear variables
                    txtFilterStatement = ""
                    txtFieldsList = ""
                    txtValuesList = ""
                    txtOrderByList = ""
                    txtFieldNameHold = ""
                    
                        
                    '****************************************************************************
                    '*** Prepare the List of Fields to Compare with the Previous Image's Values
                    For intIndex = 0 To lblFieldDescription.Count - 1
                        txtValuesList = txtValuesList & mebIndexValues(intIndex) & "|"
                    Next
                    
                    '****************************************************************************
                    '* BEGIN TRANSACTIONS
RETRY_TRANSACTION:
                    connImaging101Batch.BeginTrans
                    connImaging101.BeginTrans
                    
                    '****************************************************************************
                    '*** Close the Spicer Document Control to discard previous pages, and Add the current page
                    '***    only if the Index Values are Different from the previous record
                    
                    If ((txtValuesList <> txtValuesListHold) _
                        And (intPageIndex > 1)) _
                        Or (intPageIndex = ListView1.ListItems.Count) Then
                        
                            If intPageIndex = ListView1.ListItems.Count Then
                                '****************************************************************************
                                '*** IMPORT THE FILE TO THE DOCUMENT - To Make sure we Get the LAST Image
                                subImportFileToDoc txtBatchDirectory, txtBatchPageFileName
                                txtValuesListHold = txtValuesList
                            End If
                            
                    
                            '***********************************************************
                            '*** Break up each of the Index Values into an Array
                            '    to allow handling them individually!
                            
                            Dim txtValuesArray() As String
                            txtValuesArray = Split(txtValuesListHold, "|")
                            
                            '*** Define Field Variables :: HARD-CODED
                            Dim txtClaimNumber As String
                            Dim txtDocGroup As String
                            Dim txtDocType As String
                            Dim txtDocDate As String
                            Dim intUniqueID As Integer
                            Dim txtUniqueID As String
                           
                            '*** Assign Values to Variables :: HARD-CODED
                            txtClaimNumber = Trim(txtValuesArray(0))
                            txtDocGroup = Trim(txtValuesArray(1))
                            txtDocType = Trim(txtValuesArray(2))
                            txtDocDate = Trim(txtValuesArray(3))
                            
                            
                            '**********************************************************
                            '*** Generate a Destination Directory
                            
                            Dim txtDestinationDirectory As String
                            txtDestinationDirectory = RegRootDirToStoreObjects
                            
                            funcCreateDirectoryStructure txtDestinationDirectory
                            
                            
                            '**********************************************************
                            '*** Generate a Unique File Name
                            
                            intUniqueID = 0
                            'Loop until we find a unique filename
                            Do Until Not funcFileExists(txtDestinationFilename)
                                intUniqueID = intUniqueID + 1
                                txtUniqueID = Format(intUniqueID, "00")
                                txtDestinationFilename = txtDestinationDirectory & "\" & txtClaimNumber & txtDocGroup & txtDocType & txtUniqueID & ".PDF"
                            Loop
                   
                            '**********************************************************
                            '*** SAVE the File based on whether it is a MultiPage
                            '*** or Single-Page document
                            
                            Dim docSave As IDocSave
                        
                            '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
                            MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True

                            ' Save the modified pages in the Spicer Document format
                            If intPageCount > 1 Then
'                                '*** Rasterize the Pages before sending
'                        '         me.subRasterizeBatch
'                                Me.subRasterizeBatchEX
                                ' Set the object variable for the IDocSave interface to the Document Control object
                                ' that was saved by the Rasterize sub
                                Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
                        '        docSave.SaveAsDialog False
                        '        docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, txtAttachmentFileName
                                'PDF - Flate, 619, R/W, Raster only Multipage PDF subset - bilevel, 8 bit, 24 bit
                                docSave.Save 0, False, 619, txtDestinationFilename, txtDestinationFilename
                                ' De-initialize the object variable
                                Set docSave = Nothing
                            Else
'                                '*** Rasterize the Pages before sending
'                                 Me.subRasterizeBatchEX
'                                ' Set the object variable for the IDocSave interface to the Document Control object
'                                ' that was saved by the Rasterize sub
                                Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
                        '        docSave.SaveAsDialog False
                        '        docSave.Save 0, False, API_FF_TIFFM, txtAttachmentFileName, txtAttachmentFileName
                                'PDF - Flate, 101, R/W, Raster only PDF subset - bilevel, 8 bit, 24 bit
                                docSave.Save 0, False, 101, txtDestinationFilename, txtDestinationFilename
                            End If
                                                    
                                                    
                            
                            
                            
                            
                            
                        SpicerDoc1.CloseDocument False
                            
                            
                        'Reset Detail Order counter
                        intDetailOrder = 0
                        intPageCount = 0
                        'Hold List of Index values to compare with next record
                    
                    End If
                                    
                            
                    
                    '****************************************************************************
                    '*** IMPORT THE FILE TO THE DOCUMENT
                    
                    subImportFileToDoc txtBatchDirectory, txtBatchPageFileName
                    txtValuesListHold = txtValuesList
                    
                   
                        
                    '**************************************************************
                    '*** Increment PageCount
                    If IsNull(rsImaging101BatchPage!BatchPagePageCount) Then
                        intPageCount = 0
                    Else
                        intPageCount = intPageCount + rsImaging101BatchPage!BatchPagePageCount
                    End If
                    
                        
                    '****************************************************************************
                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                    rsImaging101BatchPage.Fields!BatchPageStatus = "Exported"
                    rsImaging101BatchPage.Fields!BatchPageCommitDate = Now()
                    rsImaging101BatchPage.Fields!BatchPageCommitUser = gsecUserID


                    '**************************************************************
                    '*** UPDATE THE ORIGINAL BATCH
                    '*** FLAG BATCH RECORD as Committed, set counters and Update
                    If intPageIndex = ListView1.ListItems.Count _
                    And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
                        'If all pages are processed AND no pages requiring action are left
                        rsImaging101Batch.Fields!BatchCommitStatus = "Exported-FULL"
                    Else
                        rsImaging101Batch.Fields!BatchCommitStatus = "Exported-PARTIAL"
                    End If

                    rsImaging101Batch.Fields!BatchCommitDate = Now()
                    rsImaging101Batch.Fields!BatchCommitUser = gsecUserID

                    'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
                    rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
                    rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
                    rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
                    rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
                    rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)


                    '****************************************************************************
                    '*** UPDATE TRANSACTIONS AND CLOSE RECORD SETS

                    txtActionBeforeError = "Update BATCH PAGE " & txtBatchPageRECID
                    rsImaging101BatchPage.Update

                    txtActionBeforeError = "Update BATCH " & txtBatchRECID
                    rsImaging101Batch.Update

                    frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1



                    '****************************************************************************
                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                    connImaging101Batch.CommitTrans
                    connImaging101.CommitTrans

                    If rsImaging101Batch.State = adStateOpen Then
                        rsImaging101Batch.Close
                    End If
                    Set rsImaging101Batch = Nothing

                    If rsImaging101BatchPage.State = adStateOpen Then
                        rsImaging101BatchPage.Close
                    End If
                    Set rsImaging101BatchPage = Nothing


                    Set cmdImaging101 = Nothing

                    Set connImaging101 = Nothing
                    Set connImaging101Batch = Nothing



                    '****************************************************************************
'                    '****************************************************************************
'
'''''                strOutputFileName = "C:\" & txtBatchName & ".txt"
'''''                Open strOutputFileName For Append As #1
'''''                Print #1, strOutputLine
'''''                Close #1
'
            Else

                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
'
'            '*** Update Application Fields
'            txtActionBeforeError = "Assign Document User-defined Field Values " & txtApplicationName
'            For intIndex = 0 To lblFieldDescription.Count - 1
'                rsImaging101Document.Fields("" & txtFieldName(intIndex).Text & "") = rsImaging101BatchPage.Fields("" & txtFieldName(intIndex) & "")
'            Next
'            DoEvents
'
        Next
        
        '**************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '**************************************************************
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"

        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.Open

        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)

        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-FULL"
        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Committed-PARTIAL"
        End If

        rsImaging101Batch.Update

        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        Set rsImaging101Batch = Nothing

        Set connImaging101Batch = Nothing

'''        'Close the Spicer Document Control - discard loaded images
        
        '*** Bring the CommitStatus Window to the front
        frmCommitStatus.SetFocus
        
Exit Sub

ERROR_HANDLER:
        
        
        Dim dbErr As ADODB.Error
        Dim strErrMsg As String
        
        If (connImaging101.Errors.Count > 0) Or (connImaging101Batch.Errors.Count > 0) Then
            If (connImaging101.Errors(0).SQLState = "40001") Or (connImaging101.Errors(0).SQLState = "40001") Then
                'Handle transaction commit failure - Serialization Error
                funcQuickMessage "SHOW", "Commit Failure DURING ACTION: (" & txtActionBeforeError & ") - RETRYING TRANSATCION"
                connImaging101.RollbackTrans
                connImaging101Batch.RollbackTrans
                Resume RETRY_TRANSACTION   'Automatically retry in this example
            End If
        
            'Loop through the Errors collection for the Connection
            strErrMsg = ""
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
        
            'Loop through the Errors collection for the Connection
            For Each dbErr In connImaging101.Errors
                strErrMsg = strErrMsg & _
                         "Connection   " & "connImaging101Batch" & vbNewLine & _
                         "Source:      " & dbErr.Source & vbNewLine & _
                         "Description: " & dbErr.Description & vbNewLine & _
                         "SQL State:   " & dbErr.SQLState & vbNewLine & _
                         "NativeError: " & dbErr.NativeError & vbNewLine & _
                         "Number:      " & dbErr.Number & vbNewLine & vbNewLine
            Next
            
        Else
            strErrMsg = "Description: " & Err.Description & vbNewLine & _
                           "Error:       " & Err.Number & vbNewLine & vbNewLine
        End If
        
        
        bolErrorOccured = True
        
        funcQuickMessage "SHOW", "CommitBatchToImaging101 ERROR: " & strErrMsg & vbNewLine & vbNewLine & _
                        "  DURING ACTION: (" & txtActionBeforeError & ") " & vbNewLine & vbNewLine & _
                        "  [Transaction Rolled Back - Page NOT Committed]" & vbNewLine & vbNewLine & _
                        "  BatchRECID     = " & txtBatchRECID & vbNewLine & _
                        "  BatchPageRECID = " & txtBatchPageRECID & vbNewLine & _
                        "  Batch Page #   = " & intPageIndex
                        
        On Error Resume Next
        
        Set rsImaging101Batch = Nothing
        Set rsImaging101BatchPage = Nothing
        Set rsImaging101Document = Nothing
        Set rsImaging101DocumentDetail = Nothing
        
        Set connImaging101 = Nothing
        Set connImaging101Batch = Nothing
                    
        Set cmdImaging101 = Nothing
        
        Screen.MousePointer = vbDefault

End Sub

Private Sub cmdDeleteSelectedPage_Click()

    MainMDIForm.mnuBatchDeletePage_Click

End Sub

Private Sub cmdDeleteSelectedPageIcon_Click()

    MainMDIForm.mnuBatchDeletePage_Click

End Sub

Private Sub cmdEditBatchProperties_Click()

    'Edit Batch Properties
    txtCurrentModule = "frmIndex"
    frmImaging101BatchProperties.Show
    
End Sub

Private Sub cmdFindBatch_Click()
    
    'Allow Bypassing the Batch AutoSelect mode to Find another batch
    gBypassBatchAutoSelect = True
    Unload Me

End Sub

Private Sub cmdFindQuestionable_Click()

    Dim intPageIndex As Integer
    Dim intIndex As Integer
    
    '* Loop Through Pages
    For intPageIndex = 1 To ListView1.ListItems.Count
        '* Loop Through Fields
        frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
        frmIndex.ListView1.ListItems(intPageIndex).EnsureVisible
        subGetBatchFieldValues 0
        For intIndex = 0 To mebIndexValues.Count - 1
            ' Stop on Questionable
            If mebIndexValues(intIndex).Text = txtQuestionable Then
                frmIndex.ListView1.SetFocus
                ListView1_Click
               Exit Sub
            End If
        Next
    Next
    
End Sub

Private Sub cmdFindUncommitted_Click()

    Dim intPageIndex As Integer
    Dim intIndex As Integer
    Dim blnSkipThisPage As Boolean
    Dim intHoldSelectedItem As Integer
    Dim intNextSelectedItem As Integer
    
    intHoldSelectedItem = ListView1.SelectedItem.Index
    
    If intHoldSelectedItem < ListView1.ListItems.Count Then
        intNextSelectedItem = intHoldSelectedItem + 1
    Else
        intNextSelectedItem = 1
    End If
    
    '* Loop Through Pages
    For intPageIndex = intNextSelectedItem To ListView1.ListItems.Count
        '* Loop Through Fields
        frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
        frmIndex.ListView1.ListItems(intPageIndex).EnsureVisible
        cmdClearFieldValues_Click
        subGetBatchFieldValues 0
        
        ' Skip Committed Pages
        blnSkipThisPage = False
        If txtBatchPageStatus <> "Committed" Then
'            For intIndex = 0 To mebIndexValues.Count - 1
'                'Loop through ALL FIELDS to see if Any is flagged as Separator or DoNotFile
'                ' and Skip The Page
'                If mebIndexValues(intIndex).Text = txtSeparator _
'                Or mebIndexValues(intIndex).Text = txtDoNotFile Then
'                    blnSkipThisPage = True
'                    Exit For
'                End If
'            Next
            
            ' Stop on pages NOT Committed, only if they are not Separators or DoNotFile pages
            If blnSkipThisPage <> True Then
                frmIndex.ListView1.SetFocus
                ListView1_Click
                Exit Sub
            End If
            
        End If
    Next
    ' Check that Last Page is NOT a Separator Page, if so go back to the Hold item
    If blnSkipThisPage = True Then
        frmIndex.ListView1.ListItems.item(intHoldSelectedItem).Selected = True
        frmIndex.ListView1.ListItems(intHoldSelectedItem).EnsureVisible
        frmIndex.ListView1.SetFocus
        ListView1_Click
        Exit Sub
    End If

End Sub

Private Sub cmdGotoImage_Click()
    
    On Error Resume Next
    ' Prompt for the page number to display
    Dim iPageNum As Integer
    
    iPageNum = InputBox("Which page number do you want to display?", "Go To Page")
    If iPageNum < 1 Or iPageNum > frmIndex.ListView1.ListItems.Count Then
        result = MsgBox("Please Select a number between 1 and " & frmIndex.ListView1.ListItems.Count, vbOKOnly)
    Else
        frmIndex.ListView1.SelectedItem.Bold = False
        frmIndex.ListView1.SetFocus
        frmIndex.ListView1.ListItems.item(iPageNum).Selected = True
        frmIndex.ListView1.ListItems(iPageNum).EnsureVisible
        frmIndex.ListView1.SelectedItem.Bold = True
        ListView1_Click
    End If
    On Error GoTo 0
End Sub

Private Sub cmdHoldAndFind_Click()

End Sub

Private Sub cmdMakeCopies_Click()


    '*** 2023-02-20 - Jacob - Begin process of creating multiple copies
    '                                         of the current document, using the current values for all fields
        cmdMakeCopies.Enabled = False
        cmdCancelCopy.Enabled = False
        
        Screen.MousePointer = vbHourglass
    
        Dim intItemsToCopy As Integer
        Dim strItemsArray() As String
        Dim intColumnsCounter As Integer
        Dim intBatchPagesTotal As Integer
        Dim intBatchPagesNotCommitted As Integer
        Dim strColumnsList As String
        Dim strColumnsListValues As String
        Dim intItemsCounter As Integer
        Dim strSqlCommand As String
        Dim dblNextBatchPageRECID As Double
        Dim strBatchPageTableName As String
        
        'Split the items entered into an Array
        strItemsArray() = Split(txtCopyItems.Text, vbNewLine)
        
        
        'Loop through Items Entered
        For intItemsCounter = 0 To UBound(strItemsArray)
                 'Only Count if the Array Item is Not Blank
                 If Trim(strItemsArray(intItemsCounter)) <> "" Then
                        intItemsToCopy = intItemsToCopy + 1
                 End If
        Next
        
        result = MsgBox("Copy the current values to these " & intItemsToCopy & " items you entered?", vbYesNo)
        
         
        If result = vbNo Then
                Exit Sub
        End If
        
        
        
        'Set the Selected Field to the FIRST vaue to copy
        If txtFieldType(txtCopyItems.Tag) = "LongText" Then
                txtIndexValues(txtCopyItems.Tag) = strItemsArray(0)
        Else
                mebIndexValues(txtCopyItems.Tag) = strItemsArray(0)
        End If
        
        
        'Save the Field Values Entered
        subSaveBatchPageValues
        
        'Prepare BatchPage Table Name
        strBatchPageTableName = txtApplicationName & "_BatchPage"
        
        
        'MsgBox "Will make " & UBound(Items()) + 1 & " copies. "
        
        On Error GoTo MAKECOPIES_ERROR
        

        '**************************************************************
        '*** Establish BATCH DB Connection
        Dim connImaging101Batch As ADODB.Connection
        Dim rsImaging101Batch As ADODB.Recordset
        
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.CommandTimeout = 600
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
        
        
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
                    " WHERE BatchRECID = " & frmIndex.txtBatchRECID
        
        connImaging101Batch.Errors.Clear
        rsImaging101Batch.Open
        rsImaging101Batch.MoveFirst

        intBatchPagesTotal = rsImaging101Batch("BatchPagesTotal")
        intBatchPagesNotCommitted = rsImaging101Batch("BatchPagesNotCommitted")
        
        'User Transaction Tracking to make sure the Batch and BatchPage tables are updated together!
        connImaging101Batch.BeginTrans

        '**************************************************************
        '*** CONNECT to Batch Page DB
        Set rsImaging101BatchPage = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch Pages DB"
        Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
        rsImaging101BatchPage.CursorLocation = adUseServer
        rsImaging101BatchPage.CursorType = adOpenDynamic
        rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
        
        'Get List of Columns
        strSqlCommand = ""
        strSqlCommand = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & strBatchPageTableName & "'"
        txtActionBeforeError = "Get List of BatchPage Columns"
        rsImaging101BatchPage.Source = strSqlCommand
        rsImaging101BatchPage.Open
        
        rsImaging101BatchPage.MoveFirst
        
        'Prepare String with list of Columns
        strColumnsList = ""
        While Not rsImaging101BatchPage.EOF
                    strColumnsList = strColumnsList & rsImaging101BatchPage("COLUMN_NAME")
                    rsImaging101BatchPage.MoveNext
                    'Add Comma if NOT the last Column
                    If Not rsImaging101BatchPage.EOF Then
                        strColumnsList = strColumnsList & ","
                    End If
        Wend
            
        rsImaging101BatchPage.Close
         
        'Loop through Items Entered
        For intItemsCounter = 1 To UBound(strItemsArray)
             
                 'Only copy if the Array Item is Not Blank
                 If Trim(strItemsArray(intItemsCounter)) <> "" Then
                 
                        intBatchPagesTotal = intBatchPagesTotal + 1
                        intBatchPagesNotCommitted = intBatchPagesNotCommitted + 1
                        
                        dblNextBatchPageRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchPageRECID")
                        
                        strSqlCommand = ""
                        strSqlCommand = strSqlCommand & " INSERT INTO " & strBatchPageTableName & " ("
                        strSqlCommand = strSqlCommand & strColumnsList & ") "
    
                        
                        strSqlCommand = strSqlCommand & " SELECT "
                        'Replace key values in the Column List
                        strColumnsListValues = strColumnsList
                        strColumnsListValues = Replace(strColumnsListValues, "BatchPageRECID", dblNextBatchPageRECID)
                        strColumnsListValues = Replace(strColumnsListValues, txtFieldName(txtCopyItems.Tag), strItemsArray(intItemsCounter))
                        strColumnsListValues = Replace(strColumnsListValues, "BatchPageOrder", intBatchPagesTotal)
                        'Now concatenate
                        strSqlCommand = strSqlCommand & strColumnsListValues
                   
                        strSqlCommand = strSqlCommand & " FROM " & txtApplicationName & "_BatchPage "
                        
                        strSqlCommand = strSqlCommand & " WHERE BatchPageRECID = " & txtBatchPageRECID
                 
                        'COPY ROW
                         
                         txtActionBeforeError = "COPY ROW | " & strSqlCommand
                         rsImaging101BatchPage.Source = strSqlCommand
                         rsImaging101BatchPage.Open
                 
                
                End If
             
             
        Next
        
        rsImaging101Batch("BatchPagesTotal") = intBatchPagesTotal
        rsImaging101Batch("BatchPagesNotCommitted") = intBatchPagesNotCommitted
        
        txtActionBeforeError = "Update Batch Values"
        rsImaging101Batch.Update
        
        connImaging101Batch.CommitTrans
        
        txtBatchPagesTotal = intBatchPagesTotal
        
        On Error Resume Next
        
'        rsImaging101Batch.Close
        Set rsImaging101Batch = Nothing
'        rsImaging101BatchPage.Close
        Set rsImaging101BatchPage = Nothing
'        Set connImaging101Batch = Nothing
        
        cmdMakeCopies.Enabled = True
        cmdCancelCopy.Enabled = True
        
        subLoadPagesIntoListView
        
        
        txtCopyItems.Visible = False
        cmdMakeCopies.Visible = False
        cmdCancelCopy.Visible = False
        
        Screen.MousePointer = vbDefault
                
Exit Sub

MAKECOPIES_ERROR:
        
        On Error Resume Next

        funcQuickMessage "SHOW", "subDeleteBatchPageRecord ERROR: " & Err.Number & " - " & Err.Description & _
        vbCrLf & "DURING ACTION: (" & txtActionBeforeError & ")" & _
        vbCrLf & "WILL ROLL-BACK THIS TRANSACTION"
        
        connImaging101Batch.RollbackTrans
        
        
        
'        rsImaging101Batch.Close
        Set rsImaging101Batch = Nothing
'        rsImaging101BatchPage.Close
        Set rsImaging101BatchPage = Nothing
'        Set connImaging101Batch = Nothing
        
        Screen.MousePointer = vbDefault
        
        cmdMakeCopies.Enabled = True
        cmdCancelCopy.Enabled = True


End Sub

Public Sub cmdNextImage_Click()
    
    '*** VALIDATE THE FIELD - Mainly for valid Date!
    '    *** THERE SEEMS TO BE A BUG IN VB
    '        If you TAB out of a field, it will execute the Validate sub
    '        If you Press [ENTER]... the Default Key takes precedence and
    '            BYPASSES the Validate sub.
    
    txtActionBeforeError = "mebIndexValues_Validate intHoldFocusIndex, False"
    mebIndexValues_Validate intHoldFocusIndex, False
    
    If bolErrorOccured Then
        Exit Sub
    End If
    
    On Error GoTo cmdNextImage_Click_ERROR
    
    '*** Only Save the Indexes if NOT in Read-Only Mode
    If (gOpenBatchInReadOnlyMode <> True) Then
        'Save Field Values
        txtActionBeforeError = "subSaveBatchPageValues"
        subSaveBatchPageValues
    End If
    
    '*** CHECK if Annotations Need to be SAVED
    '    Placed here to allow Annotations EVEN if in Read-Only Mode
    txtActionBeforeError = "MainMDIForm.ActiveForm.subAnnotationLayerSaveCheck"
    MainMDIForm.ActiveForm.subAnnotationLayerSaveCheck
    
    txtActionBeforeError = "If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then"
    If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
        If gOpenBatchInReadOnlyMode = True Then
            Exit Sub   ' Bail out of here, don't try to move to next item
        Else
            '08/17/2006 - Check if user has Batch Commit Rights
            If gsecRightsBatchCommit <> True Then
                result = MsgBox("This is the Last image!  CLOSE BATCH?", vbYesNo)
                Me.SetFocus
                If result = vbYes Then
                    Unload Me
                Else
                    Exit Sub   ' Bail out of here, don't try to move to next item
                End If
                Exit Sub  'Bail out NOW... User does not have rights to COMMIT!
            End If
            
            '*** DETERMINE WHAT TO DO BASED ON BATCH GROUP (TYPE)
            Select Case txtBatchGroup.Text
            
                Case "REGULAR"
                    result = MsgBox("This is the Last image!  COMMIT this BATCH?", vbYesNo)
                    If result = vbYes Then
                        txtActionBeforeError = "cmdCommitBatch_Click"
                        cmdCommitBatch_Click
    '                    Unload Me
                        Exit Sub   ' Bail out of here, don't try to move to next item
                    End If
                    
                Case "SPLIT" ' txtBatchGroup = "SPLIT"
                    result = MsgBox("This is the Last image!  SPLIT this BATCH?", vbYesNo)
                    If result = vbYes Then
                         txtActionBeforeError = "cmdSplitBatch_Click"
                        cmdSplitBatch_Click
    '                    Unload Me
                        Exit Sub   ' Bail out of here, don't try to move to next item
                    End If
                    
                Case "TTC PRINTED"
                    result = MsgBox("This is the Last image!  UPDATE the PRINTED Status?", vbYesNo)
                    If result = vbYes Then
                        txtActionBeforeError = "cmdSplitBatch_Click"
                        cmdUpdatePrintedStatus_Click
    '                    Unload Me
                        Exit Sub   ' Bail out of here, don't try to move to next item
                    End If
                    
                Case Else
                    'Either NOT REGULAR or SPLIT or answered "NO"
                    Exit Sub   ' Bail out of here, don't try to move to next item
                    
            End Select
            
        End If
    End If
    
    '*** 5/11/2011 - Jacob - ADDED Ability to COMPARE DocTypes to see if
    '                                  the CommitViaFTP flag is kept or reset
    'Prepare to Compare DocTypes
    Dim strPreviousDocTypeString As String
    Dim strCurrentDocTypeString As String
    Dim strCommitViaFTPHold As String
    
    strCommitViaFTPHold = txtCommitViaFTP.Text
    
    txtActionBeforeError = "strPreviousDocTypeString = funcBuildDocTypeString()"
    strPreviousDocTypeString = funcBuildDocTypeString()
    
    'Move to next item
    ListView1.SelectedItem.Bold = False
    frmIndex.ListView1.SetFocus
    If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
        txtActionBeforeError = "frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index + 1).Selected = True"
        frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index + 1).Selected = True
    End If
    ' Hold the Current Item's Index value to use when user Clicks and "Jumps" to another item
    intListView1CurrentItem = ListView1.SelectedItem.Index
    ListView1.SelectedItem.Bold = True
    
    '*** GET the Field Values for the currently selected Page
    txtActionBeforeError = "subGetBatchFieldValues 0"
    subGetBatchFieldValues 0

    
    txtActionBeforeError = "strCurrentDocTypeString = funcBuildDocTypeString()"
    strCurrentDocTypeString = funcBuildDocTypeString()

    
    '*** SHOW the Page
    txtActionBeforeError = "MainMDIForm.Show"
    MainMDIForm.Show
    
    Dim strBatchPageRECID As String
    Dim strBatchPageFileName As String
    
    strBatchPageRECID = frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(1).Text
    strBatchPageFileName = frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(2).Text
    txtBatchPageRotation = frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(4).Text
    
    'COMPARE the DocType Fields of the CURRENT page to the PREVIOUS page
    'If same DocType... then carry the DocType flag to this page... otherwise use what was loaded
    'from the DB
    If strPreviousDocTypeString = strCurrentDocTypeString Then
        txtActionBeforeError = " txtCommitViaFTP = strCommitViaFTPHold"
        txtCommitViaFTP = strCommitViaFTPHold
    Else
        txtActionBeforeError = " txtCommitViaFTP = frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(5).Text"
        txtCommitViaFTP = frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(5).Text
    End If
    
    frmIndex.txtBatchPageRECID = strBatchPageRECID
    
    Dim txtCaption As String
    txtCaption = "BATCH: " & txtBatchName & "   Page: " & frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).Text
    
    txtActionBeforeError = " result = MainMDIForm.funcShowImage(txtBatchDirectory, txtBatchPageFileName, 1, 1, txtCaption, 1, 1, txtBatchPageRotation, "", "", "", gI101ModuleIndex)"
    result = MainMDIForm.funcShowImage(txtBatchDirectory, txtBatchPageFileName, 1, strBatchPageRECID, txtCaption, 1, 1, txtBatchPageRotation, "", "", "", gI101ModuleIndex)
    
    '*** 2021-10-20 - Jacob - Added test for Object Launched.
    If bolObjectLaunched = False Then
            txtActionBeforeError = "MainMDIForm.ActiveForm.subInitializeChildForm"
            MainMDIForm.ActiveForm.subInitializeChildForm
            txtActionBeforeError = "MainMDIForm.ActiveForm.subSetCurrentPage"
            MainMDIForm.ActiveForm.subSetCurrentPage
    End If
    
'    frmIndex.ListView1.SetFocus
'    frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).Selected = True

    
    ' Set focus to list item.
    txtActionBeforeError = "Set focus to list item."
    frmIndex.ListView1.SetFocus
    frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).Selected = True
    frmIndex.ListView1.ListItems(ListView1.SelectedItem.Index).EnsureVisible

    
    '*** If NOT in Read-Only Mode, set the Focus
    If (gOpenBatchInReadOnlyMode <> True) Then
        
        If funcIsFormLoaded2("frmLookupList") Then
                If frmLookupList.chkHighlightLookpFieldAfterNextPage = vbChecked Then
                    ' Set focus to the Lookup List
                    txtActionBeforeError = "frmLookupList.txtTableLookupField.SetFocus"
                    frmLookupList.txtTableLookupField.SetFocus
                Else
                    '********************************************************
                    '*** CHECK HERE FOR FieldToSelectAfterNextPageClick
                    '********************************************************
                    txtActionBeforeError = "Get strFIELDAFTERNEXTPAGECLICK"
                    strFIELDAFTERNEXTPAGECLICK = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToSelectAfterNextPageClick") & ""
    
                    If strFIELDAFTERNEXTPAGECLICK = "" Then
                        ' Set focus to the field we were on when we hit the Next Page button.
                        
                        '*** Check if the txtIndexValues TextBox control is VISIBLE...
                        '    this will handle saving the value of the the TextBox control
                        '    instead of the mebIndexValues Masked Edit control as needed.
                        '    This is because the Masked Edit control has a MAX size of 64 Char.
                        If txtIndexValues(intHoldFocusIndex).Visible = True Then
                            'Use the TEXT Control
                            txtActionBeforeError = "strFIELDAFTERNEXTPAGECLICK = "" , Use the TEXT Control"
                            txtIndexValues(intHoldFocusIndex).SetFocus
                        Else
                            txtActionBeforeError = "strFIELDAFTERNEXTPAGECLICK = "",  Use the MEB Control"
                            mebIndexValues(intHoldFocusIndex).SetFocus
                        End If
                        
                    Else  'strFIELDAFTERNEXTPAGECLICK = ""
                    
                         ' ***** Find the Field Specified in the INI file as
                         ' *****   FieldToSelectAfterNextPageClick and set the focus to it.
                         txtActionBeforeError = "Find the Field Specified in the INI file as FieldToSelectAfterNextPageClick and set Focus"
                        
                         For intIndex = 0 To frmIndex.lblFieldDescription.Count - 1
                             Select Case Trim(frmIndex.lblFieldDescription.item(intIndex).Caption)
                                 Case Trim(strFIELDAFTERNEXTPAGECLICK)
                                     frmIndex.SetFocus
                                     '*** Check if the txtIndexValues TextBox control is VISIBLE...
                                     '    this will handle saving the value of the the TextBox control
                                     '    instead of the mebIndexValues Masked Edit control as needed.
                                     '    This is because the Masked Edit control has a MAX size of 64 Char.
                                     'IGNORE ERROR if field flagged as "Do Not Allow Manual Input"
                                     On Error Resume Next
                                     If txtIndexValues(intIndex).Visible = True Then
                                         'Use the TEXT Control
                                          txtActionBeforeError = "strFIELDAFTERNEXTPAGECLICK <> "",  Use the TXT Control"
                                        txtIndexValues(intHoldFocusIndex).SetFocus
                                     Else
                                         txtActionBeforeError = "strFIELDAFTERNEXTPAGECLICK <> "",  Use the MEB Control"
                                         mebIndexValues(intIndex).SetFocus
                                     End If
                                     On Error GoTo cmdNextImage_Click_ERROR
                             End Select
                        Next
                        
                    End If  'strFIELDAFTERNEXTPAGECLICK = ""
                    
            End If ' frmLookupList.chkHighlightLookpFieldAfterNextPage = vbChecked
            
        Else  '  funcIsFormLoaded2("frmLookupList")
        
            ' Set focus to the field we were on when we hit the Next Page button.
            
            '*** Check if the txtIndexValues TextBox control is VISIBLE...
            '    this will handle saving the value of the the TextBox control
            '    instead of the mebIndexValues Masked Edit control as needed.
            '    This is because the Masked Edit control has a MAX size of 64 Char.
            If txtIndexValues(intHoldFocusIndex).Visible = True Then
                  'Use the TEXT Control
                txtActionBeforeError = "txtIndexValues(intHoldFocusIndex).Visible = True ,  Use the TXT Control"
                txtIndexValues(intHoldFocusIndex).SetFocus
            Else
                txtActionBeforeError = "txtIndexValues(intHoldFocusIndex).Visible = False ,  Use the MEB Control"
                mebIndexValues(intHoldFocusIndex).SetFocus
            End If
                
        End If   ' funcIsFormLoaded2("frmLookupList")
    
    End If   ' (gOpenBatchInReadOnlyMode <> True)
    
    DoEvents
    
Exit Sub

cmdNextImage_Click_ERROR:
    
        funcQuickMessage "SHOW", "cmdNextImage_Click_ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & " DURING ACTION: (" & txtActionBeforeError & ")"
        
        Screen.MousePointer = vbDefault


End Sub

Private Sub cmdPreviousImage_Click()
    
    If ListView1.SelectedItem.Index = 1 Then
        '*** 2021-10-20 - Jacob - Commented out this section... NOT necessary.
'        result = MsgBox("This is the FIRST image!", vbOKOnly)
'        frmIndex.ListView1.SetFocus
'        frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).Selected = True
'        ListView1_Click
    Else
        'Move to Previous item
        ListView1.SelectedItem.Bold = False
        frmIndex.ListView1.SetFocus
        frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index - 1).Selected = True
        ListView1.SelectedItem.Bold = True
        ListView1_Click

    End If
    
    DoEvents
 
End Sub

Private Sub cmdRevertToSaved_Click()
    subLoadFieldDefinitions
End Sub





Private Sub cmdBarcodeBatch_Click()

End Sub

Public Sub cmdProcessBarcodes_Click()

    '***********************************************************************************************************************************************
    '1/3/2015 - Jacob - DISABLED THE BARCODE DUE TO A PROBLEM WITH THE PEGASUS BARCODE CONTROL
    '                           JUST DISPLAY MESSAGE AND GET OUT !!!
    MsgBox "* * * * * * * * * * *   S O R R Y   ! ! !   * * * * * * * * * * * *" & vbCrLf & _
                "THE BARCODE FEATURE IS NOT OPERATIONAL" & vbCrLf & _
                "IN THIS VERSION... " & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                "PLEASE USE A PREVIOUS VERSION OF IMAGING101" & vbCrLf & _
                "TO PROCESS BARCODES" & vbCrLf & _
                "UNTIL WE ARE ABLE TO RESOLVE THIS ISSUE."
    Exit Sub
    '***********************************************************************************************************************************************
    
    bolProcessingBarcodes = True
    
    On Error Resume Next
    intBarcodeClipBeginPosition = CInt(VBGetPrivateProfileString(RegAppname, "frmConfig.txtBarcodeClipBeginPosition", RegFileName))
    intBarcodeClipNumberOfCharacters = CInt(VBGetPrivateProfileString(RegAppname, "frmConfig.txtBarcodeClipNumberOfCharacters", RegFileName))
    On Error GoTo 0

    If chkProcessBarcodesEntireBatch = vbChecked Then
    
        '*** Walk from BookMark Begin to End
        '* Loop Through Pages
        For intPageIndex = 1 To frmIndex.ListView1.ListItems.Count
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            frmIndex.ListView1_Click
            'Ensure the Line Item being processed is visible in the listview... Auto-Scroll the List View
            frmIndex.ListView1.ListItems.item(intPageIndex).EnsureVisible
            
            subProcessBarCodes (intPageIndex)
        Next
        
    Else
            subProcessBarCodes (intPageIndex)
                
    End If

'            frmIndex.ListView1.ListItems.item(frmIndex.ListView1.ListItems.Count).Selected = True
'            frmIndex.ListView1_Click

 
    'Skip the Message if Committing Selected Batches
    If blnBarcodeSelectedBatchesWait = False And chkProcessBarcodesEntireBatch = True Then
        MsgBox "Barcode Processing COMPLETE for this Batch!", vbInformation, "Barcode Processing Complete"
    End If
    
    'Flag that Barcode is complete
    bolProcessingBarcodes = False

    blnBarcodeSelectedBatchesWait = False
    
End Sub

Private Sub subProcessBarCodes(intPageIndex As Integer)

    Dim result As String
        '***************************************
        '*** PROCESS BAR-CODE
        
        Me.MousePointer = vbHourglass
        
        txtBatchPageFileName = frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(2).Text
        
        '*** TEST Using the the Barcode Form frmBarcodeHandler
'        frmBarcodeHandler.Show
'        Exit Sub
        
        'Scan the Image for a Barcode
        result = frmBarcodeHandler.funcProcessBarcode(txtBatchDirectory & "\" & txtBatchPageFileName)
    
    
        If result <> "0" Then
            
            '*** FOUND BARCODE ***
            
            'Select the Current Item to Clear Fields
            frmIndex.ListView1_Click
            subGetBatchFieldValues intPageIndex
            
            
            '*** Clip the Barcode if Applicable
            If intBarcodeClipBeginPosition > 0 And intBarcodeClipNumberOfCharacters > 0 Then
                result = Mid(result, intBarcodeClipBeginPosition, intBarcodeClipNumberOfCharacters)
            End If
            
            
            '*** Check if Drop Leading Zeroes was selected
            If bolDropLeadingZeroes Then
                intLoop = 1
                While Mid(result, intLoop, 1) = 0
                    intLoop = intLoop + 1
                Wend
                result = Right(result, Len(result) - intLoop + 1)
            End If
            
            DoEvents
            
            'Only attempt a Table Lookup if the Table Lookup form is Loaded AND Visible
            If funcIsFormLoaded2("frmLookupList") And frmLookupList.Visible Then
                frmLookupList.txtTableLookupField.Text = result
                frmLookupList.cmdFind_Click
            End If
            
            DoEvents
            
            'HARD CODED DOCTYPE
            Select Case frmImaging101BatchList.cmbApplicationList.Text
                Case "TTC"
                    frmIndex.mebIndexValues(3) = txtBatchName
                    If frmIndex.mebIndexValues(0) = "" Then
                        frmIndex.mebIndexValues(5) = result
                    End If
                        
                Case "EVIDENCE"
                    frmIndex.mebIndexValues(5) = "Evidence"
                    frmIndex.mebIndexValues(6) = txtBatchName
                    If frmIndex.mebIndexValues(0) = "" Then
                        frmIndex.mebIndexValues(0) = result
                    End If
                    
                Case "PET"
                        frmIndex.mebIndexValues(0) = result
                
                Case "CLIENT_FILES"
                    If txtBatchQueue = "Initial Client Packets" Then
                        
                        Dim strSQL  As String
                        
                        '************************************
                        '*** UPDATE DocumentHistory
                        strSQL = ""
                        strSQL = strSQL & "UPDATE DocumentHistory SET "
                        strSQL = strSQL & "DateReceived = '" & Date & "', "
                        strSQL = strSQL & "ReceivedStatus = 'RECEIVED', "
                        strSQL = strSQL & "UserReceived = 'I101'"
                        strSQL = strSQL & "WHERE DocumentHistoryID = " & result
    
                        RegLookupListConnectionString = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & txtApplicationRECID, "LookupDBConnectionString") & ""
                        funcRunSQLCommand RegLookupListConnectionString, strSQL

                
                        '************************************
                        '*** GET Client Name
                        
                        strSQL = ""
                        strSQL = strSQL & "SELECT LeadClient.ClientID, LeadClient.Wholename "
                        strSQL = strSQL & " FROM DocumentHistory, LeadClient "
                        strSQL = strSQL & " WHERE DocumentHistoryID = " & result & " "
                        strSQL = strSQL & " AND DocumentHistory.ClientID = LeadClient.ClientID"
                        
        
                        Dim conn As ADODB.Connection
                        Dim rs As ADODB.Recordset
                        Set conn = New ADODB.Connection
                        Set rs = New ADODB.Recordset
                        
                        conn.ConnectionString = RegLookupListConnectionString
                        conn.ConnectionTimeout = 120
                        conn.mode = adModeRead
                        conn.Open
                        
                        With rs
                            .ActiveConnection = conn
                            .CursorLocation = adUseServer
                            .CursorType = adOpenDynamic
                            .LOCKTYPE = adLockOptimistic
                            .Source = strSQL
                        End With
                
                        rs.Open
                        
                        'Return the Found Value
                        frmIndex.mebIndexValues(0) = rs.Fields("ClientID")
                        frmIndex.mebIndexValues(1) = rs.Fields("Wholename")
                        frmIndex.mebIndexValues(2) = "Customer Service"
                        frmIndex.mebIndexValues(3) = "Initial Client Packets"

                        
                        rs.Close
                        conn.Close
                        Set rs = Nothing
                        Set conn = Nothing
                        '********************************
                
                        '************************************
                        '*** UPDATE CLIENT_FILES
                        strSQL = ""
                        strSQL = strSQL & "UPDATE " & txtApplicationName & "_BatchPage SET "
                        strSQL = strSQL & "CLIENTNO = '" & frmIndex.mebIndexValues(0) & "', "
                        strSQL = strSQL & "CLIENTNAME = '" & frmIndex.mebIndexValues(1) & "', "
                        strSQL = strSQL & "DOCDATE = '" & Date & "', "
                        strSQL = strSQL & "DOCGROUP = " & "'Customer Service', "
                        strSQL = strSQL & "DOCTYPE = " & "'Initial Client Packets' "
                        strSQL = strSQL & "WHERE BATCHRECID = " & txtBatchRECID
    
                        funcRunSQLCommand RegImaging101BatchListConnectionString, strSQL
                        '************************************
                        
                        
                        '************************************
                        '*** UPDATE BATCH
                        strSQL = ""
                        strSQL = strSQL & "UPDATE I101BATCHES  SET "
                        strSQL = strSQL & "BATCHNAME = '" & frmIndex.mebIndexValues(0) & "' "
                        strSQL = strSQL & "WHERE BATCHRECID = " & txtBatchRECID
    
                        funcRunSQLCommand RegImaging101BatchListConnectionString, strSQL
                        
                        txtBatchName = frmIndex.mebIndexValues(0)
                        '************************************

                        
                        
                        frmIndex.ListView1.ListItems.item(frmIndex.ListView1.ListItems.Count).Selected = True
                        ListView1_Click
                        DoEvents
                

            
'            '*** Walk from BookMark Begin to End
'            '* Loop Through Pages
'            For intPageIndex = 1 To frmIndex.ListView1.ListItems.count
'
'                frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
'
'                '*** Load field values from "Bookmarked" page by passing
'                '    the intBookMarkOfFieldsToCopy as a parameter
'                '    to the subGetBatchFieldValues subroutine
'                subGetBatchFieldValues intPageIndex
'
'                ' Save the loaded values
'                subSaveBatchPageValues
'
'                'Color the pages in between Begin and End Bookmarks
'                If intPageIndex > intBookMarkBegin And intPageIndex < intBookMarkEnd Then
'                    ListView1.SelectedItem.ForeColor = vbYellow
'                End If
'
'            Next
'
'            frmIndex.ListView1.ListItems.item(frmIndex.ListView1.ListItems.count).Selected = True
'            ListView1_Click
'            DoEvents
            
                
                    End If
                    
                Case Else

            End Select
            
            'Color the pages in between Begin and End Bookmarks
            ListView1.SelectedItem.ForeColor = vbMagenta
            DoEvents
            
            subSaveBatchPageValues
            DoEvents
            
        Else
        
            '***  DID NOT FIND BARCODE ***
            
            If bolUseBarcodeAsDocumentHeader Then
                'Color the pages in between Begin and End Bookmarks
                ListView1.SelectedItem.ForeColor = vbBlue
                DoEvents
                
                subSaveBatchPageValues
                DoEvents
            Else
                
                'DO NOT SAVE the Index Info for THIS Page
                'Each page must have it's own valid Barcode
                'Clear the Field Values so users don't get confused
                cmdClearFieldValues_Click
                
            End If
                
        End If
        

        
        Me.MousePointer = vbDefault
    
End Sub


Private Sub cmdRotateImage_Click()
    MainMDIForm.cmdImageRotateRight_Click
End Sub



Private Sub cmdSplitBatch_Click()

    ' If committing Selected Batches, set Result to yes and let it rip!
    If blnCommitSelectedBatches = True Then
        result = vbYes
    Else
        result = MsgBox("Are you sure you wish to SPLIT the Batch?", vbYesNo, "Batch SPLIT Verification")
    End If
    
    If result = vbYes Then
    
        'Unload Annotations Form if loaded
        If funcIsFormLoaded2("frmAnnotate") Then
            Unload frmAnnotate
            Set frmAnnotate = Nothing
        End If
    
        ' COMMIT / RELEASE the Batch!
        Dim intPageIndex As Integer
        Dim intIndex As Integer
        Dim strOutputLine As String
        Dim bolSkipPage As Boolean
        Dim strOutputFileName As String
        Dim intAppendFieldIndex As Integer
        Dim dblHoldOldBatchRECID As Double
        
        '*** SET GLOBAL FLAG FOR COMMITTING BATCH PAGES
        '     This way we Won't display the Images as we process them!
        blnCommittingBatchPages = True
        


        '*** HIDE the Indexing Forms for Speed & to Unclutter the desktop!
        '     works faster if video doesn't display detail.

        frmLookupList.Visible = False
        frmDocTypeList.Visible = False
        MainMDIForm.Visible = False
        Me.Visible = False
        
        '*** SHOW the CommitStatus Window & Disable Buttons
        frmCommitStatus.Show
        frmCommitStatus.cmdCloseBatch.Enabled = False
        frmCommitStatus.cmdStayOnBatch.Enabled = False
        funcMakeTopMost frmCommitStatus, True
        
        DoEvents


        
        '*** RESET Counter Variables
        frmCommitStatus.txtPagesProcessed = 0
        frmCommitStatus.txtPagesCommitted = 0
        frmCommitStatus.txtPagesPreviouslyCommitted = 0
        frmCommitStatus.txtPagesSeparator = 0
        frmCommitStatus.txtPagesQuestionable = 0
        frmCommitStatus.txtPagesDoNotFile = 0
        frmCommitStatus.txtPagesRequiredButEmpty = 0
        frmCommitStatus.txtPagesFailedValidation = 0
        frmCommitStatus.txtPagesTotalSkipped = 0
        
        
        
        '*********************************************************************
        '*** Go Ahead and Split the Batch
        subSplitBatchToMultipleBatches
        
        
        '*** DISABLE GLOBAL FLAG FOR COMMITTING BATCH PAGES
        '     To resume displaying the Images as we index them!
        blnCommittingBatchPages = False
        
''''        Me.WindowState = vbNormal
''''        frmDocTypeList.WindowState = vbNormal
''''        frmLookupList.WindowState = vbNormal

''        frmCommitStatus.MakeTopMost
        
        
        Set connImaging101Batch = New ADODB.Connection
        
       ' Set the flag back to False to END the dummy loop in
        '   "Private Sub cmdCommitSelectedBatches_Click()" of frmImaging101BatchList
        If blnCommitSelectedBatches = True Then
            blnCommitSelectedBatches = False
        End If
                
        frmCommitStatus.SetFocus
        frmCommitStatus.cmdCloseBatch.Enabled = True
        frmCommitStatus.cmdStayOnBatch.Enabled = True

    End If

End Sub



Private Sub cmdUpdatePrintedStatus_Click()

    Dim intSiteIdIndex As Integer
    Dim txtCaseIdCutoff As Double
    Dim txtTTCConnectionString(2) As String

    Dim con As ADODB.Connection
    Dim ssql As String

    Dim txtTTCUserID As Integer
    
        'Unload Annotations Form if loaded
        If funcIsFormLoaded2("frmAnnotate") Then
            Unload frmAnnotate
            Set frmAnnotate = Nothing
        End If
        
        
        '*** Show the Commit Status Form & Make it Stay on top
        frmCommitStatus.Show modal, Me
        frmCommitStatus.cmdCloseBatch.Enabled = False
        frmCommitStatus.cmdStayOnBatch.Enabled = False

''        frmCommitStatus.MakeTopMost

        
        '*** RESET Counter Variables
        frmCommitStatus.txtPagesProcessed = 0
        frmCommitStatus.txtPagesCommitted = 0
        frmCommitStatus.txtPagesPreviouslyCommitted = 0
        frmCommitStatus.txtPagesSeparator = 0
        frmCommitStatus.txtPagesQuestionable = 0
        frmCommitStatus.txtPagesDoNotFile = 0
        frmCommitStatus.txtPagesRequiredButEmpty = 0
        frmCommitStatus.txtPagesFailedValidation = 0
        frmCommitStatus.txtPagesTotalSkipped = 0
        
        

    On Error GoTo cmdUpdatePrintedStatus_Click_ERROR
    
    
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Applications where ApplicationRECID = " & txtApplicationRECID
    rs.Open ssql, con
    
        txtTTCConnectionString(0) = rs!LookupDBConnectionString
        txtTTCConnectionString(1) = rs!LookupDBConnectionString_B

        txtCaseIdCutoff = rs!CaseIdCutoff
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    
    
    
    For intSiteIdIndex = 0 To 1
    
        '**************************************************************
        '*** Establish TTC DB Connections
        
        Set connTTC(intSiteIdIndex) = New ADODB.Connection
        Set cmdTTC(intSiteIdIndex) = New ADODB.Command
        Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
    
        txtActionBeforeError = "Open DB ConnectionString: " & txtTTCConnectionString(intSiteIdIndex)
        
        connTTC(intSiteIdIndex).ConnectionString = txtTTCConnectionString(intSiteIdIndex)
        connTTC(intSiteIdIndex).ConnectionTimeout = 120
        connTTC(intSiteIdIndex).mode = adModeReadWrite
        connTTC(intSiteIdIndex).Open
    
        Set cmdTTC(intSiteIdIndex).ActiveConnection = connTTC(intSiteIdIndex)
    
    Next
    
    
    'DEFAULT to Site A
    intSiteIdIndex = 0

    If CDbl(txtTableLookupField) >= dblCaseIdCutoff Then
        intSiteIdIndex = 1
    End If

    
    '**************************************************************
    '*** Get TTC Login info
    
    If txtBatchGroup = "TTC PRINTED" Then
    
        bolTTCUserFound = False
        While Not bolTTCUserFound
            frmLoginTTC.Show vbModal, Me
            
            If frmLoginTTC.bolTTCLoginClickedLogin = False Then
                Unload frmLoginTTC
                Exit Sub
            End If
            
            ' Validate USER and make sure user is "Active"
            Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
            txtSource = "select id, username, password, active from users where username = '" & frmLoginTTC.txtUserID & "' AND password = '" & frmLoginTTC.txtPassword & "' AND active = 1"
            rsTTC(intSiteIdIndex).Open txtSource, connTTC(intSiteIdIndex), adOpenDynamic, adLockOptimistic
            If rsTTC(intSiteIdIndex).EOF Or rsTTC(intSiteIdIndex).BOF Then
                 MsgBox "Invalid User name or Password!" & vbCrLf & "Please try again...", vbOKOnly, "TTCLoginFailed"
            Else
                bolTTCUserFound = True
                'SAVE UserID to update RECEIVED fields.
                txtTTCUserID = rsTTC(intSiteIdIndex).Fields!ID
                Unload frmLoginTTC
                'This command will immediately "Close" the rsTTC after executing
            End If
         Wend
         
        
    End If
    
    
    
    
    
    '**************************************************************
    '*** Establish BATCH DB Connection
    Set connImaging101Batch = New ADODB.Connection
    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
    connImaging101Batch.ConnectionTimeout = 120
    connImaging101Batch.mode = adModeReadWrite
    connImaging101Batch.Open
    connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"

    
    
        '***********************************
        '* Loop Through Batch Pages
        '***********************************
        txtActionBeforeError = "Loop Through Batch Pages"
        
        For intPageIndex = 1 To ListView1.ListItems.Count
            
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            subGetBatchFieldValues 0
            
            ' Locate the Record
            frmIndex.ListView1.SetFocus
            ListView1_Click
                
            If txtBatchPageStatus <> "Updated" Then
            
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                   If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "") Then
                        frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                    '*** VALIDATE DATE FIELD!
                    funcValidateDate intHoldFocusIndex
                    If blnDateError = True Then
                        frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                        bolSkipPage = True
                        Exit For
                    End If
                    
               Next
                   
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            
            If bolSkipPage <> True Then
                
                '* Begin the Transaction
'                connImaging101Batch.BeginTrans
'                connTTC.BeginTrans

                    Set rsImaging101BatchPage = New ADODB.Recordset
                    Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                    
                    rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                    
                    rsImaging101BatchPage.CursorLocation = adUseServer
                    rsImaging101BatchPage.CursorType = adOpenDynamic
                    rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                    
''                    rsImaging101BatchPage.Errors.Clear
                    rsImaging101BatchPage.Open
                    
                    
                    '*******************************************************
                    '*** UPDATE TTC FIELDS
                    
                    txtActionBeforeError = "Update Fields in 'viewI101Cases'  in TTC Database - Site =  " & intSiteIdIndex

                    Dim ttcFileName As String
                    Dim ttcFileDirectory As String
                    
                    
                        Set rsTTC(intSiteIdIndex) = New ADODB.Recordset
                        txtSource = "UPDATE viewI101Cases SET cases_print_confirm = 1, cases_print_confirm_user_id = '" & txtTTCUserID & "', cases_print_confirm_datetime = '" & Format(Now(), "yyyy-mm-dd HH:mm:dd") & "' WHERE cases_id = " & rsImaging101BatchPage.Fields("CaseID")
                        rsTTC(intSiteIdIndex).Open txtSource, connTTC(intSiteIdIndex), adOpenDynamic, adLockOptimistic
                        
                        
                        
                        
                    '****************************************************************************
                    '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                        rsImaging101BatchPage.Fields!BatchPageStatus = "Updated"
                        rsImaging101BatchPage.Fields!BatchPageCommitDate = Now()
                        rsImaging101BatchPage.Fields!BatchPageCommitUser = gsecUserID


                        '**************************************************************
                        '*** UPDATE THE ORIGINAL BATCH
                        '*** FLAG BATCH RECORD as Committed, set counters and Update
                        '**************************************************************
                        '*** CONNECT to Batch DB
                        Set rsImaging101Batch = New ADODB.Recordset
                        txtActionBeforeError = "Connect to Batch DB"
                        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                        rsImaging101Batch.CursorLocation = adUseServer
                        rsImaging101Batch.CursorType = adOpenDynamic
                        rsImaging101Batch.LOCKTYPE = adLockOptimistic
                        txtActionBeforeError = "Open Batch DB"
                        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                        rsImaging101Batch.Open
    
                        
                        If intPageIndex = ListView1.ListItems.Count _
                        And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
                            'If all pages are processed AND no pages requiring action are left
                            rsImaging101Batch.Fields!BatchCommitStatus = "Updated-FULL"
                        Else
                            rsImaging101Batch.Fields!BatchCommitStatus = "Updated-PARTIAL"
                        End If
                        
                        rsImaging101Batch.Fields!BatchCommitDate = Now()
                        rsImaging101Batch.Fields!BatchCommitUser = gsecUserID
                        
                        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
                        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
                        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
                        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
                        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
                        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
            
                        
                        '****************************************************************************
                        '*** UPDATE TRANSACTIONS AND CLOSE RECORD SETS
                        
                        txtActionBeforeError = "Update BATCH PAGE " & txtBatchPageRECID
                        rsImaging101BatchPage.Update
    
                        txtActionBeforeError = "Update BATCH " & txtBatchRECID
                        rsImaging101Batch.Update
                        
                        frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
                        
    
                 ''***********************************
                
''''                strOutputFileName = "C:\" & txtBatchName & ".txt"
''''                Open strOutputFileName For Append As #1
''''                Print #1, strOutputLine
''''                Close #1
            
            Else
            
                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
            
            DoEvents
        
        Next
        
        
        '**************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '**************************************************************
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
        
        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.Open
        
        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Updated-FULL"
        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Updated-PARTIAL"
        End If
        
        rsImaging101Batch.Update
        
        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        Set rsImaging101Batch = Nothing
        
        Set connImaging101Batch = Nothing
        
        '*** Bring the CommitStatus Window to the front
        frmCommitStatus.SetFocus
        
        frmCommitStatus.cmdCloseBatch.Enabled = True
        frmCommitStatus.cmdStayOnBatch.Enabled = True
        
Exit Sub

cmdUpdatePrintedStatus_Click_ERROR:
    
        MsgBox "cmdUpdatePrintedStatus_Click_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Data NOT Imported]", vbExclamation
        
        Screen.MousePointer = vbDefault


End Sub

Private Sub Form_GotFocus()

    txtCurrentModule = "frmIndex"
    If bolIndexFormLoadComplete = True Then
        subDisplayOrHideForms
    End If
    
End Sub

Public Sub Form_Load()

    'Initialize the Index Form Load Complete global variable
    bolIndexFormLoadComplete = False
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    If bolDebug Then
        Me.Caption = Me.Caption & " - DEBUG"
    End If

    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    funcWriteToDebugLog Me.name, "ENTERING frmIndex.Form_Load"

    '*** Disable Notifications
    frmPopUpNotifyForm.timNoficationTimer.Enabled = False


'    Erase gFormArrayIndex


    
    ' Get saved settings from the registry
    On Error Resume Next
    frmIndex.Top = VBGetPrivateProfileString(RegAppname, "frmIndex.Top", RegFileName)
    frmIndex.Left = VBGetPrivateProfileString(RegAppname, "frmIndex.Left", RegFileName)
    frmIndex.width = VBGetPrivateProfileString(RegAppname, "frmIndex.Width", RegFileName)
    frmIndex.Height = VBGetPrivateProfileString(RegAppname, "frmIndex.Height", RegFileName)

    bolDropLeadingZeroes = CInt(VBGetPrivateProfileString(RegAppname, "frmConfig.chkDropLeadingZeroes", RegFileName))
    bolUseBarcodeAsDocumentHeader = CInt(VBGetPrivateProfileString(RegAppname, "frmConfig.chkUseBarcodeAsDocumentHeader", RegFileName))
    

    
    '*** GET BATCH HEADER INFORMATION
    ' Set index of selected Row
    lstIndex = frmImaging101BatchList.ListView1.SelectedItem.Index
    ' Get Main Item
    txtBatchRECID = frmImaging101BatchList.txtBatchRECID
    ' Get Sub-Items
    txtApplicationRECID = frmImaging101BatchList.txtApplicationRECID
    txtApplicationName = frmImaging101BatchList.cmbApplicationList.Text
    txtBatchName = frmImaging101BatchList.txtBatchName
    txtBatchQueue = frmImaging101BatchList.txtBatchQueue
    txtBatchStatus = frmImaging101BatchList.txtBatchStatus
    txtBatchDesc = frmImaging101BatchList.txtBatchDesc
    txtBatchDirectory = frmImaging101BatchList.txtBatchDirectory
    txtBatchCommitStatus = frmImaging101BatchList.txtBatchCommitStatus
    txtBatchGroup = frmImaging101BatchList.txtBatchGroup
    txtBatchOwner = frmImaging101BatchList.txtBatchOwner
    txtBatchPagesTotal = Trim(frmImaging101BatchList.txtBatchPagesTotal)
    
    If txtBatchPagesTotal = "" Then
        result = funcQuickMessage("SHOW", "WARNING!!!" & vbCrLf & vbCrLf & _
                                                                "This Batch [ " & txtBatchName & "] Has NO PAGES...")
    End If

        
    
    '*****************************************************************************
    '*** CHECK BATCH OPEN MODE
    
    funcWriteToDebugLog Me.name, "CheckBatchOpenMode"
    CheckBatchOpenMode
    DoEvents
    
    '*****************************************************************************
    '*** INITIALIZE LookupFieldsAvailable Global Flag

    funcWriteToDebugLog Me.name, "Set gNoLookupFieldsAvailable = False"
    gNoLookupFieldsAvailable = False

    '*****************************************************************************
    '*** LOAD FIELD DEFINITIONS

    funcWriteToDebugLog Me.name, "subLoadFieldDefinitions"
    subLoadFieldDefinitions
    DoEvents

    On Error GoTo FRMINDEX_FORMLOAD_ERROR

    '***************************************************
    '*** LOAD PAGES INTO LISTVIEW

    ListView1.Visible = False
    funcWriteToDebugLog Me.name, "subLoadPagesIntoListView"
    subLoadPagesIntoListView
    DoEvents

    '***************************************************
    '*** Set the ListView1 Image List to the First item
    '*** and display image if there is at least one item.
    If ListView1.ListItems.Count > 0 Then
'        frmIndex.ListView1.SetFocus
        funcWriteToDebugLog Me.name, "frmIndex.ListView1.ListItems.item(1).Selected = True"
        frmIndex.ListView1.ListItems.item(1).Selected = True
        DoEvents
        funcWriteToDebugLog Me.name, "ListView1_Click"
        ListView1_Click
        DoEvents
    End If

    funcWriteToDebugLog Me.name, "ListView1.Visible = True"
    ListView1.Visible = True

    '*************************************************
    '*** Get the Root Directory to Store Objects
    funcWriteToDebugLog Me.name, "Get the Root Directory to Store Objects"
    RegRootDirToStoreObjects = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "RootDirectoryPathForImageArchive") & ""


    '********************************************************************
    '*** GET DocType Fields
    strDOCGROUP = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToAssignDocumentGroup") & ""
    strDOCTYPE = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToAssignDocumentType") & ""
    strDOCSUBTYPE = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmIndex.txtApplicationRECID, "FieldToAssignDocumentSubType") & ""
    

    DoEvents

    '****************************************************************
    '*** See if we should display the Barcode Processing button
    cmdProcessBarcodes.Visible = False
    chkProcessBarcodesEntireBatch.Visible = False


    '*** Validate the Barcode License Key
    bolBarcodeLicenseValidated = funcValidateBarCodeLicense

    If bolBarcodeLicenseValidated Then
        cmdProcessBarcodes.Visible = True
        chkProcessBarcodesEntireBatch.Visible = True
    End If


    '****************************************************************
    '*** See if we should display the Delete Page button
    If gsecRightsBatchIndex = vbChecked Then
        cmdDeleteSelectedPage.Visible = True
        cmdDeleteSelectedPageIcon.Visible = True
    Else
        cmdDeleteSelectedPage.Visible = False
        cmdDeleteSelectedPageIcon.Visible = False
    End If



    
    '******************************************************************
    '*** Load DocTypeList and Table Lookup only if NOT in Read-Only Mode
    If (gOpenBatchInReadOnlyMode = False) Then
        frmDocTypeList.Show
        frmLookupList.Show
    End If
    
    
    
    'Now see if we should Display or Hide the remaining forms
    funcWriteToDebugLog Me.name, "subDisplayOrHideForms"
    subDisplayOrHideForms
    
    
    
    

    '*************************************************
    '*** NOW the Form Load is Complete
    funcWriteToDebugLog Me.name, "bolIndexFormLoadComplete = True"
    bolIndexFormLoadComplete = True


    
Exit Sub

FRMINDEX_FORMLOAD_ERROR:
    result = MsgBox("FRMINDEX_FORMLOAD_ERROR: " & Err.Number & " - " & Err.Description, vbOKCancel)
    Err.Clear
    If result = vbOK Then
        'Try again
        Resume
    Else
        Unload Me
    End If
    

End Sub

Public Sub CheckBatchOpenMode()

    '*****************************************************************************
    '*** If Batch is opened in READ-ONLY Mode
    '***    Disable Both the Commit and Split Buttons
    If (gOpenBatchInReadOnlyMode = True) Then
        cmdBookMark.Visible = False
        cmdClearFieldValues.Visible = False
        cmdCommitBatch.Visible = False
        'If Batch is in Read-Only Mode... Allow opening the Properties Form
        ' the Properties Form will handle the required logic to limit editing capabilities
        cmdEditBatchProperties.Visible = True
        cmdSplitBatch.Visible = False
    Else
        '*** Now check if user has Rights to Commit
        If gsecRightsBatchCommit = vbChecked Then
            '*** See if we should display the Commit or Split buttons
            If txtBatchGroup.Text = "SPLIT" Then
                cmdSplitBatch.Top = cmdCommitBatch.Top
                cmdSplitBatch.Left = cmdCommitBatch.Left
                cmdCommitBatch.Visible = False
                cmdSplitBatch.Visible = True
            Else
                'NOT a SPLIT Batch
                Select Case frmImaging101BatchList.cmbApplicationList.Text
            
                    Case "TTC"
                        If txtBatchGroup.Text = "TTC PRINTED" Then
                            'If TTC PRINTED... DON'T Allow Commit OR Split!!!
                            cmdSplitBatch.Visible = False
                            cmdCommitBatch.Visible = False
                            cmdUpdatePrintedStatus.Top = cmdCommitBatch.Top
                            cmdUpdatePrintedStatus.Left = cmdCommitBatch.Left
                            cmdUpdatePrintedStatus.Visible = True
                        Else
                            cmdSplitBatch.Visible = False
                            cmdUpdatePrintedStatus.Visible = False
                            cmdCommitBatch.Visible = True
                        End If
                        
                    Case Else
                        'Not TTC... Allow Commit as "REGULAR" Batch
                        cmdSplitBatch.Visible = False
                        cmdCommitBatch.Visible = True
                        'Do not show the Update button
                        cmdUpdatePrintedStatus.Visible = False
           
                End Select

            End If
        Else
            cmdCommitBatch.Visible = False
            cmdSplitBatch.Visible = False
        End If
    
    End If
        
        

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Do NOT allow this form to be unloaded while it's still loading.
    If Not bolIndexFormLoadComplete Then
        Cancel = True
        Exit Sub
    End If
    
    
    'Save Form settings to the registry
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "frmIndex.Top", frmIndex.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmIndex.Left", frmIndex.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmIndex.Width", frmIndex.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmIndex.Height", frmIndex.Height, RegFileName)
    End If
    
    FormUnloadMode = UnloadMode
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If bolIndexFormLoadComplete = True Then
        subDisplayOrHideForms
    End If
            
            
    If Me.WindowState <> vbMinimized Then
        
        Frame1.width = Me.ScaleWidth
        
    End If
    
    ListView1.Height = toolbarBottom.Top - ListView1.Top
    
    cmdProcessBarcodes.Top = toolbarBottom.Top - cmdProcessBarcodes.Height
    chkProcessBarcodesEntireBatch.Top = cmdProcessBarcodes.Top
    
    '*** 2020-04-24 - Jacob - Added RESIZE for All Fields
    For intIndex = 0 To lblFieldDescription.Count - 1

                If txtFieldType(intIndex).Text = "LongText" Then
                    txtIndexValues(intIndex).width = Me.ScaleWidth - txtIndexValues(intIndex).Left - 10
                Else
                    mebIndexValues(intIndex).width = Me.ScaleWidth - mebIndexValues(intIndex).Left - 10
                End If
       
    Next

    If txtCopyItems.Visible = True Then
                txtCopyItems.Top = mebIndexValues(Index).Top
                txtCopyItems.Left = mebIndexValues(Index).Left
                txtCopyItems.width = mebIndexValues(Index).width
                txtCopyItems.Height = mebIndexValues(Index).Height * 10
    End If
    
''    MsgBox "resize"
End Sub

Public Sub subDisplayOrHideForms()
            
    On Error Resume Next
    
    If bolIndexFormLoadComplete = False Then
        Exit Sub
    End If
    
    If frmIndex.WindowState = vbMinimized Then
        If funcIsFormLoaded2("frmDocTypeList") Then
            frmDocTypeList.Visible = False
        End If
        If funcIsFormLoaded2("frmLookupList") Then
            frmLookupList.Visible = False
        End If
        If funcIsFormLoaded2("MainMDIForm") Then
            MainMDIForm.WindowState = vbMinimized
        End If
        
    Else
    
        If bolIndexFormLoadComplete = True Then
        
            ' Only load the DocTypes List If
            ' the Batch is NOT in READ-ONLY Mode
            ' and NOT committing Selected Batches
            If (gOpenBatchInReadOnlyMode <> True) Then
            
                '* LOAD the QuickFields Forms
                If (blnCommitSelectedBatches <> True) Then
                        If funcIsFormLoaded2("frmDocTypeList") Then
                            If frmDocTypeList.Visible = False Then
                                frmDocTypeList.Show
                                DoEvents
                                frmDocTypeList.Visible = True
                            End If
                        End If
                End If

                ' Only load the DocTypes List If
                ' the Batch is NOT in READ-ONLY Mode
                ' and either NOT committing Selected Batches or are Commiting Selected Batches With Lookup
                If (blnCommitSelectedBatches <> True) _
                Or (blnCommitSelectedBatches = True And chkCommitWithLookup = True) Then
                
                    If funcIsFormLoaded2("frmLookupList") Then
                        If frmLookupList.Visible = False And gNoLookupFieldsAvailable = False Then
                            frmLookupList.Show
                            DoEvents
                            frmLookupList.Visible = True
                        End If
                    End If
                    
                End If

            End If
            
            If funcIsFormLoaded2("MainMDIForm") Then
                If MainMDIForm.WindowState <> vbNormal Or MainMDIForm.Visible = False Then
                    MainMDIForm.Show
                    DoEvents
                    MainMDIForm.Visible = True
                    MainMDIForm.WindowState = vbNormal
                    DoEvents
                End If
            End If
        End If
    End If
    
    Frame1.width = Me.ScaleWidth
    If Me.ScaleHeight > 0 Then
        ListView1.Height = Me.ScaleHeight - ListView1.Top - toolbarBottom.Height
    End If


    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    funcWriteToDebugLog Me.name, "[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[["
    funcWriteToDebugLog Me.name, "ENTERING:  frmIndex.Form_Unload()"
    
    '*****************************************************************************
    '*** Only UNLOCK If Batch is NOT open in READ-ONLY Mode
    If (gOpenBatchInReadOnlyMode = False) Then   'And (gsecRightsBatchIndex = vbChecked)
        funcWriteToDebugLog Me.name, "Send UNLOCK Request"

        strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)
        If Left(strReturn, 5) = "ERROR" Then
            funcWriteToDebugLog Me.name, strReturn & " - UNLOCK Failed"
'            funcWriteToDebugLog Me.name, strReturn & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server." & " - Server Communication Failure"
'            MsgBox strReturn & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Server Communication Failure"
'            Exit Sub
'            frmImaging101Winsock.cmdClose_Click
        Else
            funcWriteToDebugLog Me.name, strReturn & " - UNLOCK Successful"

        End If
    Else
    
        funcWriteToDebugLog Me.name, "Batch Read-Only - Skip UNLOCK"

    End If
        
    On Error GoTo ERROR_HANDLER
        
    
    '*** UNLOAD and DESTROY ALL FORMS
    
    If funcIsFormLoaded2("frmCommitStatus") Then
        funcWriteToDebugLog Me.name, "Unload frmCommitStatus"
        Unload frmCommitStatus
        funcWriteToDebugLog Me.name, "Set frmCommitStatus = Nothing"
        Set frmCommitStatus = Nothing
    End If
    
    If funcIsFormLoaded2("frmAnnotate") Then
        funcWriteToDebugLog Me.name, "Unload frmAnnotate"
        Unload frmAnnotate
        funcWriteToDebugLog Me.name, "Set frmAnnotate = Nothing"
        Set frmAnnotate = Nothing
    End If
    
    If funcIsFormLoaded2("frmThumb") Then
        funcWriteToDebugLog Me.name, "Unload frmThumb"
        Unload frmThumb
        funcWriteToDebugLog Me.name, "Set frmThumb = Nothing"
        Set frmThumb = Nothing
    End If
    
    If funcIsFormLoaded2("frmFTP") Then
        funcWriteToDebugLog Me.name, "Unload frmFTP"
        Unload frmFTP
        funcWriteToDebugLog Me.name, "Set frmFTP = Nothing"
        Set frmFTP = Nothing
    End If
    
    If funcIsFormLoaded2("frmDocTypeList") Then
        funcWriteToDebugLog Me.name, "Unload frmDocTypeList"
        Unload frmDocTypeList
        funcWriteToDebugLog Me.name, "Set frmDocTypeList = Nothing"
        Set frmDocTypeList = Nothing
    End If
    
    If funcIsFormLoaded2("frmLookupList") Then
        funcWriteToDebugLog Me.name, "Unload frmLookupList"
        Unload frmLookupList
        funcWriteToDebugLog Me.name, "Set frmLookupList = Nothing"
        Set frmLookupList = Nothing
    End If
    
    
    '*** 2022-07-28 - Jacob - Only Unload MainMDIForm if gFormArrayRetrieve is empty (No Retrieval documents loaded)
    If IsArrayEmpty(gFormArrayRetrieve) Then
            If funcIsFormLoaded2("MainMDIForm") Then
                funcWriteToDebugLog Me.name, "Unload MainMDIForm"
                Unload MainMDIForm
                DoEvents
                funcWriteToDebugLog Me.name, "Set MainMDIForm = Nothing"
                Set MainMDIForm = Nothing
            End If
    Else
    
            '*** 2022-07-28 - Jacob - UNLOAD ONLY THE BATCH CHILD FORM
            'If Not IsArrayEmpty(gFormArrayRetrieve) Then
                    Unload gFormArrayIndex(0)
            'End If
            
    End If
    
    

    
    If (FormUnloadMode = vbFormControlMenu) And (UCase(gsecBatchMode) = "AUTO") Then
        'If user is in "Auto Batch Select Mode" and clicked the X to close the form,
        'then get back to the MainMenu
        funcWriteToDebugLog Me.name, "frmMainMenu.Show"
        frmMainMenu.Show
    Else
        'Otherwise go back to the BatchList
        funcWriteToDebugLog Me.name, "frmImaging101BatchList.Show"
        frmImaging101BatchList.Show
        
        'DON'T refresh Batches if in Read-Only mode or if Committing Selected Batches
        If gOpenBatchInReadOnlyMode = False And blnCommitSelectedBatches = False Then
            funcWriteToDebugLog Me.name, "frmImaging101BatchList.subListBatches"
            frmImaging101BatchList.subListBatches
        End If
    End If
    
    funcWriteToDebugLog Me.name, "txtCurrentModule = ''"
    txtCurrentModule = ""

    
    
    '*** Re-Enable the Notification Timer
    funcWriteToDebugLog Me.name, "frmPopUpNotifyForm.timNoficationTimer.enabled = True"
    frmPopUpNotifyForm.timNoficationTimer.Enabled = True
    
    'Flag the Form as NOT loaded
    funcWriteToDebugLog Me.name, "bolIndexFormLoadComplete = False"
    bolIndexFormLoadComplete = False
    
    
    'SAFE way of saying: Set Me = Nothing
    funcWriteToDebugLog Me.name, "BEGIN SAFE UNLOAD of frmIndex"
    
    Dim Form As Form
    For Each Form In Forms
            If Form Is Me Then
                    funcWriteToDebugLog Me.name, "frmIndex = Nothing"
                    Set Form = Nothing
                    funcWriteToDebugLog Me.name, "Exit For"
                    Exit For
            End If
    Next Form
    
    
    funcWriteToDebugLog Me.name, "EXIT SUB - frmIndex"
    funcWriteToDebugLog Me.name, "]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]"
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    funcWriteToDebugLog Me.name, ""
    
Exit Sub

ERROR_HANDLER:
    
    funcWriteToDebugLog Me.name, "##################################################"
    funcWriteToDebugLog Me.name, "ENTERING: frmIndex.Unload ERROR_HANDLER"
    funcWriteToDebugLog Me.name, "frmIndex.Unload ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "frmIndex.Unload ERROR: " & Err.Number & " - " & Err.Description, vbInformation
    funcWriteToDebugLog Me.name, "Resume Next"
    funcWriteToDebugLog Me.name, "##################################################"
    
    Resume Next
    

End Sub

Function funcGetApplicationFieldValue(strFieldValue As String, intRECID As Integer)
    On Error GoTo 0
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select " & strFieldValue & ", ApplicationRECID from I101Applications where ApplicationRECID=" & txtApplicationRECID
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    rs.Open
    
    funcGetApplicationFieldValue = rs.Fields(strFieldValue)
        
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    

End Function
Sub subLoadFieldDefinitions()
    
    '*** THIS SUBROUTINE LOADS ALL THE APPLICATION FIELD DEFINITION INFORMATION
    '***  INCLUDING FIELD FORMAT VALUES INTO AN ARRAY.
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    '1/19/2011 - Jacob - Changed to handle "Prevent Manual Indexing".  Instead of NOT showing field, prevent INPUT into it.
    'rs.Source = "Select * from I101Fields WHERE ApplicationRECID = " & txtApplicationRECID & " AND FieldIsForOutputOnly <> '1' " & " ORDER BY FieldOrderBatch"
    '2020-09-18 - Jacob - Added HideForSearchIndex <> '1' or NULL -- To ignore fields flagged as HideForSearchIndex
    rs.Source = "Select * from I101Fields WHERE ApplicationRECID = " & txtApplicationRECID & " AND (HideForSearchIndex <> '1'  OR HideForSearchIndex is NULL)  ORDER BY FieldOrderBatch"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LOCKTYPE = adLockReadOnly
    
    On Error GoTo ERROR_TRAP
    
    intFieldSpacing = 40

    
    'sql statement to select items on the drop down list
''    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID
''    rs.Open ssql, Con
    con.Errors.Clear
    rs.Open
    

''    funcWriteToDebugLog Me.Name, rs.PageCount
''    funcWriteToDebugLog Me.Name, rs.RecordCount
''    funcWriteToDebugLog Me.Name, rs.AbsolutePage
''    funcWriteToDebugLog Me.Name, rs.AbsolutePosition
    
    
    rs.MoveFirst
    
    For intIndex = 0 To rs.RecordCount - 1

        '*****************************************************************************
        '* Create Field Objects - BEGIN
        
            Dim intFieldTop As Integer
            
        '* Create Field Objects - BEGIN
            If intIndex > 0 Then
                Load lblFieldDescription(intIndex)
                Load mebIndexValues(intIndex)
                Load txtIndexValues(intIndex)
                Load txtBatchFieldsRECID(intIndex)
                Load txtFieldsRECID(intIndex)
                Load txtFieldDefaultValue(intIndex)
                Load txtFieldLowValue(intIndex)
                Load txtFieldHighValue(intIndex)
                Load txtFieldIsSticky(intIndex)
                Load txtFieldType(intIndex)
                Load txtFieldSize(intIndex)
                Load txtFieldName(intIndex)
                Load txtFieldIsRequiredForCommit(intIndex)
                Load txtFieldIsRequiredForSplit(intIndex)
                Load txtFieldSplitBatches(intIndex)
                Load txtFieldRouteToBatchQueue(intIndex)
                Load txtFieldRouteToBatchUser(intIndex)
                Load txtFieldRouteToBatchManager(intIndex)
                Load txtFieldDefaultForBarcodeOnly(intIndex)
                Load txtFieldIsForOutputOnly(intIndex)
                Load txtFieldTableLookupOverridesDefault(intIndex)
                
                'Set the top to slightly below the previous field
                intFieldTop = lblFieldDescription(intIndex - 1).Top + lblFieldDescription(intIndex - 1).Height + intFieldSpacing
            Else
                'Set top to where the first field is
                intFieldTop = lblFieldDescription(intIndex).Top
            End If
        '* Create Field Objects - END
            
            
'                Set lblFieldDescription(intIndex).Container = Frame2
            lblFieldDescription(intIndex).Top = intFieldTop
            lblFieldDescription(intIndex).Enabled = True
            lblFieldDescription(intIndex).Visible = True
            lblFieldDescription(intIndex).Caption = ""
            
'                Set mebIndexValues(intIndex).Container = Frame2
            mebIndexValues(intIndex).Top = intFieldTop
            mebIndexValues(intIndex).Enabled = True
            mebIndexValues(intIndex).Visible = False
            mebIndexValues(intIndex).TabIndex = intIndex + 1
            mebIndexValues(intIndex).Text = ""
            
'                Set txtIndexValues(intIndex).Container = Frame2
            txtIndexValues(intIndex).Top = intFieldTop
            txtIndexValues(intIndex).Enabled = True
            txtIndexValues(intIndex).Visible = False
            txtIndexValues(intIndex).TabIndex = intIndex + 1
            txtIndexValues(intIndex).Text = ""
                
'                Set txtBatchFieldsRECID(intIndex).Container = Frame2
            txtBatchFieldsRECID(intIndex).Enabled = True
            txtBatchFieldsRECID(intIndex).Visible = False

'                Set txtFieldsRECID(intIndex).Container = Frame2
            txtFieldsRECID(intIndex).Enabled = True
            txtFieldsRECID(intIndex).Visible = False

'                Set txtFieldDefaultValue(intIndex).Container = Frame2
            txtFieldDefaultValue(intIndex).Enabled = True
            txtFieldDefaultValue(intIndex).Visible = False
            txtFieldDefaultValue(intIndex).Text = ""
            
'                Set txtFieldLowValue(intIndex).Container = Frame2
            txtFieldLowValue(intIndex).Enabled = True
            txtFieldLowValue(intIndex).Visible = False
            txtFieldLowValue(intIndex).Text = ""
        
'                Set txtFieldHighValue(intIndex).Container = Frame2
            txtFieldHighValue(intIndex).Enabled = True
            txtFieldHighValue(intIndex).Visible = False
            txtFieldHighValue(intIndex).Text = ""
            
'                Set txtFieldIsSticky(intIndex).Container = Frame2
            txtFieldIsSticky(intIndex).Enabled = True
            txtFieldIsSticky(intIndex).Visible = False
            txtFieldIsSticky(intIndex).Text = ""
        
'                Set txtFieldType(intIndex).Container = Frame2
            txtFieldType(intIndex).Enabled = True
            txtFieldType(intIndex).Visible = False
            txtFieldType(intIndex).Text = ""
        
'                Set txtFieldSize(intIndex).Container = Frame2
            txtFieldSize(intIndex).Enabled = True
            txtFieldSize(intIndex).Visible = False
            txtFieldSize(intIndex).Text = ""
            
'                Set txtFieldName(intIndex).Container = Frame2
            txtFieldName(intIndex).Enabled = True
            txtFieldName(intIndex).Visible = False
            txtFieldName(intIndex).Text = ""
        
'                Set txtFieldIsRequiredForCommit(intIndex).Container = Frame2
            txtFieldIsRequiredForCommit(intIndex).Enabled = True
            txtFieldIsRequiredForCommit(intIndex).Visible = False
            txtFieldIsRequiredForCommit(intIndex).Text = ""
        
'                Set txtFieldIsRequiredForSplit(intIndex).Container = Frame2
            txtFieldIsRequiredForSplit(intIndex).Enabled = True
            txtFieldIsRequiredForSplit(intIndex).Visible = False
            txtFieldIsRequiredForSplit(intIndex).Text = ""
        
'                Set txtFieldSplitBatches(intIndex).Container = Frame2
            txtFieldSplitBatches(intIndex).Enabled = True
            txtFieldSplitBatches(intIndex).Visible = False
            txtFieldSplitBatches(intIndex).Text = ""
        
'                Set txtFieldSplitBatches(intIndex).Container = Frame2
            txtFieldRouteToBatchQueue(intIndex).Enabled = True
            txtFieldRouteToBatchQueue(intIndex).Visible = False
            txtFieldRouteToBatchQueue(intIndex).Text = ""
            
'                Set txtFieldSplitBatches(intIndex).Container = Frame2
            txtFieldRouteToBatchUser(intIndex).Enabled = True
            txtFieldRouteToBatchUser(intIndex).Visible = False
            txtFieldRouteToBatchUser(intIndex).Text = ""

'                Set txtFieldSplitBatches(intIndex).Container = Frame2
            txtFieldRouteToBatchManager(intIndex).Enabled = True
            txtFieldRouteToBatchManager(intIndex).Visible = False
            txtFieldRouteToBatchManager(intIndex).Text = ""

        
'                Set txtFieldSplitBatches(intIndex).Container = Frame2
            txtFieldDefaultForBarcodeOnly(intIndex).Enabled = True
            txtFieldDefaultForBarcodeOnly(intIndex).Visible = False
            txtFieldDefaultForBarcodeOnly(intIndex).Text = ""

            '1/19/2011 - Jacob
'                Set txtFieldSplitBatches(intIndex).Container = Frame2
            txtFieldIsForOutputOnly(intIndex).Enabled = True
            txtFieldIsForOutputOnly(intIndex).Visible = False
            txtFieldIsForOutputOnly(intIndex).Text = ""


        
        

        '*****************************************************************************
        '*** Check if Batch should be opened in  READ-ONLY Mode
        '     Either condition will disable modifying the Index values
        If (gOpenBatchInReadOnlyMode = True) Then  ' Or (gsecRightsBatchIndex <> vbChecked)
            mebIndexValues(intIndex).Enabled = False
            txtIndexValues(intIndex).Enabled = False
        End If
                
        
        '*****************************************************************************
        '*** Assign Field Values
        
        txtFieldsRECID(intIndex) = rs.Fields!FieldsRECID
        If (IsNull(rs.Fields!FieldNameForInput)) Or (rs.Fields!FieldNameForInput <> "") Then
            lblFieldDescription(intIndex) = rs.Fields!FieldNameForInput
        Else
            lblFieldDescription(intIndex) = rs.Fields!FieldName
        End If
        
        If Not IsNull(rs.Fields!FieldMask) Then mebIndexValues(intIndex).Mask = rs.Fields!FieldMask
        If Not IsNull(rs.Fields!FieldFormat) Then mebIndexValues(intIndex).Format = rs.Fields!FieldFormat
        
        
        If Not IsNull(rs.Fields!FieldDefaultValue) Then txtFieldDefaultValue(intIndex) = rs.Fields!FieldDefaultValue
        If Not IsNull(rs.Fields!FieldLowValue) Then txtFieldLowValue(intIndex) = rs.Fields!FieldLowValue
        If Not IsNull(rs.Fields!FieldHighValue) Then txtFieldHighValue(intIndex) = rs.Fields!FieldHighValue
        If Not IsNull(rs.Fields!FieldIsSticky) Then txtFieldIsSticky(intIndex) = rs.Fields!FieldIsSticky
        If Not IsNull(rs.Fields!FieldType) Then txtFieldType(intIndex) = rs.Fields!FieldType
        If Not IsNull(rs.Fields!FieldSize) Then txtFieldSize(intIndex) = rs.Fields!FieldSize
        If Not IsNull(rs.Fields!FieldName) Then txtFieldName(intIndex) = rs.Fields!FieldName
        
        If Not IsNull(rs.Fields!FieldIsRequiredForCommit) Then txtFieldIsRequiredForCommit(intIndex) = rs.Fields!FieldIsRequiredForCommit
        If Not IsNull(rs.Fields!FieldIsRequiredForSplit) Then txtFieldIsRequiredForSplit(intIndex) = rs.Fields!FieldIsRequiredForSplit
        If Not IsNull(rs.Fields!FieldSplitBatches) Then txtFieldSplitBatches(intIndex) = rs.Fields!FieldSplitBatches
        If Not IsNull(rs.Fields!FieldRouteToBatchQueue) Then txtFieldRouteToBatchQueue(intIndex) = rs.Fields!FieldRouteToBatchQueue
        If Not IsNull(rs.Fields!FieldRouteToBatchUser) Then txtFieldRouteToBatchUser(intIndex) = rs.Fields!FieldRouteToBatchUser
        If Not IsNull(rs.Fields!FieldRouteToBatchManager) Then txtFieldRouteToBatchManager(intIndex) = rs.Fields!FieldRouteToBatchManager
        If Not IsNull(rs.Fields!FieldDefaultForBarcodeOnly) Then txtFieldDefaultForBarcodeOnly(intIndex) = rs.Fields!FieldDefaultForBarcodeOnly
        If Not IsNull(rs.Fields!FieldTableLookupOverridesDefault) Then txtFieldTableLookupOverridesDefault(intIndex) = rs.Fields!FieldTableLookupOverridesDefault
        
         '1/19/2011 - Jacob - FieldIsForOutputOnly handles the "Prevent Manual Indexing" application field option
         If Not IsNull(rs.Fields!FieldIsForOutputOnly) Then txtFieldIsForOutputOnly(intIndex) = rs.Fields!FieldIsForOutputOnly
            
            
        '10/7/2012 - Jacob - Moved this down so txtFieldType test will work
        '*** Determine whether to use the Text or Masked Edit Control
        If txtFieldType(intIndex) = "LongText" Then
            mebIndexValues(intIndex).TabStop = False
            txtIndexValues(intIndex).TabStop = True
            txtIndexValues(intIndex).Visible = True
            txtIndexValues(intIndex).TabIndex = intIndex + 1
        Else
            txtIndexValues(intIndex).TabStop = False
            mebIndexValues(intIndex).TabStop = True
            mebIndexValues(intIndex).Visible = True
            mebIndexValues(intIndex).TabIndex = intIndex + 1
        End If
        
        
        
        If txtFieldIsForOutputOnly(intIndex) = vbChecked Then
            mebIndexValues(intIndex).Enabled = False
            txtIndexValues(intIndex).Enabled = False
        End If
        
        
        'Highlight the Appropriate fields if REGULAR or SPLIT Batch
        If txtBatchGroup = "SPLIT" Then
            If txtFieldIsRequiredForSplit(intIndex) = vbChecked Then
                'Show that field is required!
                lblFieldDescription(intIndex) = lblFieldDescription(intIndex)
                lblFieldDescription(intIndex).ForeColor = vbRed
            Else
                lblFieldDescription(intIndex).ForeColor = vbNormal
            End If
        Else ' txtBatchGroup = "REGULAR"
            If txtFieldIsRequiredForCommit(intIndex) = vbChecked Then
                'Show that field is required!
                lblFieldDescription(intIndex) = lblFieldDescription(intIndex)
                lblFieldDescription(intIndex).ForeColor = vbRed
            Else
                lblFieldDescription(intIndex).ForeColor = vbNormal
            End If
        End If
        
        rs.MoveNext
    Next
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    Me.Show
    
    DoEvents
    
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    result = MsgBox("LoadFieldFormats - Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
End Sub
    


Private Sub subGetBatchFieldValues(intListViewSelectedItemIndex As Integer)

    Dim intFieldIndex As Integer
    Dim bolCloseConnection As Boolean
    
    On Error Resume Next
    
    If intListViewSelectedItemIndex > 0 Then
        ' If parameter was passed, select the fields for the page requested using "intListViewSelectedItemIndex"
        txtBatchPageRECID = ListView1.ListItems(intListViewSelectedItemIndex).ListSubItems(1).Text
        txtBatchPageFileName = ListView1.ListItems(intListViewSelectedItemIndex).ListSubItems(2).Text
        
        txtBatchDocDesc = ListView1.ListItems.item(intListViewSelectedItemIndex).ListSubItems(3).Text
        txtBatchPageRotation = ListView1.ListItems.item(intListViewSelectedItemIndex).ListSubItems(4).Text
        txtCommitViaFTP = ListView1.ListItems.item(intListViewSelectedItemIndex).ListSubItems(5).Text
        
    Else
        ' Load the Currently selected page fields using "SelectedItem.Index"
        txtBatchPageRECID = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
        txtBatchPageFileName = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
        
        txtBatchDocDesc = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(3).Text
        txtBatchPageRotation = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(4).Text
        txtCommitViaFTP = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(5).Text
    End If
    
    ' *** Set Connection Properties only if no previous connection was established.
    '      If a connection was previously established... USE IT...
    '      othewise a SQL TimeOut / Deadlock occurs if an UPDATE was performed with this connection
    If connImaging101Batch Is Nothing Then
        'Do nothing
    Else
        If connImaging101Batch = "" Then
            Set connImaging101Batch = Nothing
        End If
    End If
    
    If connImaging101Batch Is Nothing Then
        bolCloseConnection = True
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.Open RegImaging101BatchListConnectionString
    Else
            'Connection was pre-established... Don't close it when done with this sub
            bolCloseConnection = False
    End If
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = connImaging101Batch
    
    
    connImaging101Batch.Errors.Clear
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LOCKTYPE = adLockBatchOptimistic
    
    '5/25/2011 - Jacob - Removed "BatchRECID = " & txtBatchRECID " because BatchPageRECID is UNIQUE
'    rs.Source = "Select * from " & txtApplicationName & "_BatchPage" & " WHERE BatchRECID = " & txtBatchRECID & " AND BatchPageRECID = " & txtBatchPageRECID
    
    rs.Source = "Select * from " & txtApplicationName & "_BatchPage" & " WHERE  BatchPageRECID = " & txtBatchPageRECID
    rs.Open
    
    rs.MoveFirst
    
    '*** Display Field Values
    For intIndex = 0 To mebIndexValues.Count - 1
            
            ' Cycle through Fields to Set the Field Index to the field we are on now
            For intFieldIndex = 0 To rs.Fields.Count
                If rs.Fields(intFieldIndex).name = txtFieldName(intIndex) Then
                    Exit For
                End If
            Next
            

            '***  LOAD SAVED VALUES FROM DATABASE
            '*** Check if the txtFieldType TextBox control is LongText...
            '    this will handle saving the value of the the TextBox control
            '    instead of the mebIndexValues Masked Edit control as needed.
            '    This is because the Masked Edit control has a MAX size of 64 Char.
            If txtFieldType(intIndex).Text = "LongText" Then
            
                '*** Use the TEXT Control
                
                If (Not IsNull(rs.Fields(intFieldIndex))) And (rs.Fields(intFieldIndex) <> "") Then
                    ' If the field is not empty, Set field value

                    txtIndexValues(intIndex).Text = rs.Fields(intFieldIndex) & ""
                    
                Else
                
                    ' Set Default Value ONLY if Field is NOT Sticky and NO Value has been saved.
                     If txtFieldIsSticky(intIndex) = "1" And txtIndexValues(intIndex).Text <> "" Then
                         ' Leave the existing Field Value
                         ' because the value entered by the indexer
                         ' overrides the Default Value
                     Else
                         '* If field is NOT flagged as "Sticky" , Clear it.
                         txtIndexValues(intIndex).Text = ""
                         ' *** SET DEFAULT VALUE IF APPLICABLE
                         subSetDefaultFieldValues
                     End If
                
                    
                End If
                
                '* Highlight in RED if field flagged as Questionable, Separator or DoNotFile
                If txtIndexValues(intIndex).Text = txtQuestionable _
                Or txtIndexValues(intIndex).Text = txtSeparator _
                Or txtIndexValues(intIndex).Text = txtDoNotFile Then
                    txtIndexValues(intIndex).Font.Bold = True
                    txtIndexValues(intIndex).ForeColor = vbRed
                Else
                    txtIndexValues(intIndex).Font.Bold = False
                    txtIndexValues(intIndex).ForeColor = vbNormal
                End If
            
            Else
                
                '*** Use the Masked Edit Control
                
                If (Not IsNull(rs.Fields(intFieldIndex))) And (rs.Fields(intFieldIndex) <> "") Then
                    ' If the field is not empty, Set field value
                    
                    If txtFieldType(intIndex) = "Date" Then
                        Dim strCaseDateFormat As String
                        Select Case mebIndexValues(intIndex).Mask
                            Case "##-##-##"
                                strCaseDateFormat = "mm-dd-yy"
                            Case "##-##-####"
                                strCaseDateFormat = "mm-dd-yyyy"
                            Case "##/##/##"
                                strCaseDateFormat = "mm/dd/yy"
                            Case "##/##/####"
                                strCaseDateFormat = "mm/dd/yyyy"
                            Case Else
                        End Select
                        mebIndexValues(intIndex) = Format(rs.Fields(intFieldIndex), strCaseDateFormat)
                    Else
                        If Trim(mebIndexValues(intIndex).Format) <> "" Then
                            'Use the FieldMask value to format the data
                            mebIndexValues(intIndex) = Format(rs.Fields(intFieldIndex), mebIndexValues(intIndex).Format)
                        Else
                            'No Format... just display the raw field value
                            mebIndexValues(intIndex) = rs.Fields(intFieldIndex) & ""
                        End If
                    End If
                    
                Else
                
                    ' Set Default Value ONLY if Field is NOT Sticky and NO Value has been saved.
                     If txtFieldIsSticky(intIndex) = "1" And mebIndexValues(intIndex).Text <> "" Then
                         ' Leave the existing Field Value
                         ' because the value entered by the indexer
                         ' overrides the Default Value
                     Else
                         '* If field is NOT flagged as "Sticky" , Clear it.
                         '  VB Mask Glitch - Must Remove/Clear the Mask before Clearing the Field Value
                         strHoldFieldMask = mebIndexValues(intIndex).Mask
                         mebIndexValues(intIndex).Mask = ""
                         mebIndexValues(intIndex).Text = ""
                         ' *** SET DEFAULT VALUE IF APPLICABLE
                         subSetDefaultFieldValues
                         ' Restore the Mask
                         mebIndexValues(intIndex).Mask = strHoldFieldMask
                     End If

                End If
                    
                '* Highlight in RED if field flagged as Questionable, Separator or DoNotFile
                If mebIndexValues(intIndex).Text = txtQuestionable _
                Or mebIndexValues(intIndex).Text = txtSeparator _
                Or mebIndexValues(intIndex).Text = txtDoNotFile Then
                    mebIndexValues(intIndex).Font.Bold = True
                    mebIndexValues(intIndex).ForeColor = vbRed
                Else
                    mebIndexValues(intIndex).Font.Bold = False
                    mebIndexValues(intIndex).ForeColor = vbNormal
                End If
                
            
            End If ' txtIndexValues(intIndex).Visible = True
            
            
            '*** MOVED LOGIC FROM NEXT IMAGE CLICK - SET DEFAULT OR STICKY VALUES
            '*** 2013-05-09 - Jacob - Moved to do this AFTER Loading Values from the DB
            '                                   because DB was overriding the defaults, even when blank
            
''            For intIndex = 0 To mebIndexValues.Count - 1
        
                
''            Next
            
    Next
    
    If Not IsNull(rs.Fields!BatchDocDesc) Then
        frmIndex.txtBatchDocDesc = rs.Fields!BatchDocDesc
    Else
        frmIndex.txtBatchDocDesc = ""
    End If
    
    If Not IsNull(rs.Fields!BatchPageStatus) Then
        frmIndex.txtBatchPageStatus = rs.Fields!BatchPageStatus
    Else
        frmIndex.txtBatchPageStatus = ""
    End If
    
'    If Not IsNull(rs.Fields!BatchPageRotation) Then
'        frmIndex.txtBatchPageRotation = rs.Fields!BatchPageRotation
'    Else
'        frmIndex.txtBatchPageRotation = "0"
'    End If
'
'
'    frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(3).Text = frmIndex.txtBatchDocDesc
'
    
    
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    
    On Error Resume Next
    
    If bolCloseConnection = True Then
        'Close connection
        connImaging101Batch.Close
        Set connImaging101Batch = Nothing
    End If
    

      
End Sub






Private Sub LabelBatchCommitStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        
        funcUncommitBatch txtBatchRECID, txtApplicationName.Text, txtBatchCommitStatus.Text
        
    End If

End Sub

Private Sub lblFieldDescription_Click(Index As Integer)

    '*** 2023-02-20 - Jacob - Moved to blFieldDescription_DblClick to allow handling the Ctrl+RightClick for lblFieldDescription_MouseUp event

        
End Sub



Private Sub lblFieldDescription_DblClick(Index As Integer)

    '*** 2023-02-20 - Jacob - Moved from blFieldDescription_Click
    '                                       If User DoubleClicks the FieldDescription label
    '                                        and the mebIndexValues field is NOT Empty, then we open the Search Form to search for items
    '                                        matching this value in the mebIndexValues field.

    bolSearchFormLoadComplete = False

    If Trim(Replace(mebIndexValues(Index).Text, mebIndexValues(Index).PromptChar, "")) <> "" Then
        frmImaging101Search.Show
        
            'Setting the cmbApplicationList.Text DropDown value and calling the Click event
            '  works much more consistently than the funcFindItemInComboBox.
            '  If the user had closed the Retrieve window and you try to re-launch I101AIM
            '  the cmbApplicationList_Click Event would not fire if the Text value was the same!
            
            ' Walk down the Application list... there was no easier way to set the
            '   ListIndex to the right value to Trigger the "cmbApplicationList_Click" event
            j = Len(Me.txtApplicationName) - 1
            bolFoundApplication = False
            For i = 0 To frmImaging101Search.cmbApplicationList.ListCount - 1
                If Left(UCase(Me.txtApplicationName), j) = Left(UCase(frmImaging101Search.cmbApplicationList.List(i)), j) Then
                    ' Get the case-sensitive version from the List
                    '   if any to take any required actions
                    Me.txtApplicationName = frmImaging101Search.cmbApplicationList.List(i)
                    bolFoundApplication = True
                    Exit For
                End If
            Next i
            
            'Check if we found the Application.
            If bolFoundApplication = True Then
                frmImaging101Search.cmbApplicationList.Text = Me.txtApplicationName
                frmImaging101Search.cmbApplicationList_Click
            Else
                MsgBox "Sorry... The requested application [" & Me.txtApplicationName & "] does not exist!", vbCritical, "Application Not Found"
                Exit Sub
            End If
        
            While Not bolSearchFormLoadComplete
                DoEvents
            Wend
        
         frmImaging101Search.subSetFieldValue Index, mebIndexValues(Index)

    End If

End Sub

Private Sub lblFieldDescription_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

            '*** 2023-02-20 - Jacob - Catch Ctrl + RightClick on lblFieldDescription to begin process of creating multiple copies
            '                                         of the current document, using the current values for all fields

            'Catch Ctrl + RightClick
            If (Shift = vbCtrlMask) And (Button = vbRightButton) Then
            
                'Save the Field Index in the "Tag" for when user clicks the [Make Copies] button.
                txtCopyItems.Tag = Index
                
                'funcQuickMessage "SHOW", "Ctrl + Mouse RighButton for field " & lblFieldDescription(Index)
                txtCopyItems.Top = mebIndexValues(Index).Top
                txtCopyItems.Left = mebIndexValues(Index).Left
                txtCopyItems.width = mebIndexValues(Index).width
                txtCopyItems.Height = mebIndexValues(Index).Height * 10
                
                cmdCancelCopy.Top = txtCopyItems.Top + txtCopyItems.Height + 10
                cmdCancelCopy.Left = mebIndexValues(Index).Left
                cmdCancelCopy.width = mebIndexValues(Index).width / 2

                
                cmdMakeCopies.Top = txtCopyItems.Top + txtCopyItems.Height + 10
                cmdMakeCopies.Left = cmdCancelCopy.Left + cmdCancelCopy.width
                cmdMakeCopies.width = mebIndexValues(Index).width / 2
                
                txtCopyItems.Text = "Enter List of " & lblFieldDescription(Index) & " items to Copy this Page to." & vbCrLf & "You can Copy+Paste." & vbCrLf & "Use 'Ctrl+Enter' to insert a NewLine"
                
                
                cmdCancelCopy.Visible = True
                txtCopyItems.Visible = True
                cmdMakeCopies.Visible = True
                
                'Disable DEFAULT button for ENTER Key
                cmdMakeCopies.Default = True
                cmdMakeCopies.Default = False
                
            End If

End Sub




Public Sub ListView1_Click()

    On Error GoTo ERROR_HANDLER
    
    Dim strBatchPageRECID As String
    Dim strBatchPageFileName As String
    
    funcWriteToDebugLog Me.name, "********************************************"
    funcWriteToDebugLog Me.name, "*** ENTERING LISTVIEW1_CLICK()"
    
    
    funcWriteToDebugLog Me.name, "CHECK if MainMDIForm is Loaded to see if Annotations Need to be SAVED"

    If funcIsFormLoaded2("MainMDIForm") Then
        '*** CHECK if Annotations Need to be SAVED
        '    BEFORE loading the New page's info...
        
        MainMDIForm.ActiveForm.subAnnotationLayerSaveCheck
    End If
    
    funcWriteToDebugLog Me.name, "Set Field Values for ListView1.SelectedItem.Index= " & ListView1.SelectedItem.Index

    strBatchPageRECID = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(1).Text
    strBatchPageFileName = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(2).Text
    txtBatchDocDesc = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(3).Text
    txtBatchPageRotation = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(4).Text
    txtCommitViaFTP = ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(5).Text

    funcWriteToDebugLog Me.name, "Un-Bold the previously selected item... it begins being an empty variable"

    ' Un-Bold the previously selected item... it begins being an empty variable
    If intListView1CurrentItem > 0 Then
        frmIndex.ListView1.ListItems.item(intListView1CurrentItem).Bold = False
    End If
    
    
    '* Clear the fields and Get the FieldValues ONLY if the BookMark was NOT Set
    '  or if we are NOT processing Barcodes.
    If blnBookMark = False And bolProcessingBarcodes = False Then
        funcWriteToDebugLog Me.name, "Clear the fields and Get the FieldValues."
        cmdClearFieldValues_Click
        
        subGetBatchFieldValues 0
        
    Else
        
        '*** 2020-05-05 - Jacob - Added ELSE to Get DocNotes Field Only, if we're in BookMark mode.
        '*** 2020-05-28 - Jacob - Moved "txtIndexValues(i) =" INSIDE the "If".  Was causing Out of Bounds errors.
        funcWriteToDebugLog Me.name, "Set field to value of DocNotes."
        
        For i = 0 To txtFieldName.UBound
            If txtFieldName(i) = "DocNotes" Then
                txtIndexValues(i) = funcGetFieldFromDB(RegImaging101BatchListConnectionString, txtApplicationName & "_BatchPage", "BatchPageRECID = " & strBatchPageRECID, "DocNotes")
                Exit For
            End If
        Next
        
        
    End If
        
        

    
    '*** ONLY SHOW IMAGE IF NOT COMMITTING BATCH PAGES... Unless it's for TTC!
    '    The CommitToTTC logic Rasterizes the image instead of moving the actual file,
    '      so the image MUST be Loaded for the Rasterize to work.
    If blnCommittingBatchPages <> True _
        Or frmImaging101BatchList.lblApplicationCommitBatchTo = "TTC" _
        Then
        
        '***********************************************************
        '*** Check if file Exists
        Dim txtFullPathName As String
        txtFullPathName = txtBatchDirectory & "\" & strBatchPageFileName
        If Not funcFileExists(txtFullPathName) Then
            result = MsgBox("SORRY! I can't find file:" + vbNewLine + txtFullPathName + vbNewLine + "PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR!", vbCritical)
    '        txtLOGOutputFilePath = Form1.txtOutputFilePath + "\" + txtFullPathName + ".LOG"
    '        Open txtLOGOutputFilePath For Append As #4
    '        Print #4, "Pass2 - Could Not Open either Original or Fixed file:  " + txtFullPathName
    '        Close #4
            Exit Sub
        End If
    
        funcWriteToDebugLog Me.name, "MainMDIForm.Show"

        MainMDIForm.Show
        txtCurrentModule = "frmIndex"
        
        '*** 2022-07-28 - Jacob - SHOW THE BATCH CHILD FORM TO MAKE IT THE "ACTIVE" FORM
        If Not IsArrayEmpty(gFormArrayRetrieve) Then
            'This prevents an error upon loading the first page, because the gFormArrayIndex has not been initialized yet
            If Not IsArrayEmpty(gFormArrayIndex) Then
                    gFormArrayIndex(0).SetFocus
            End If
        End If
        
        
        Dim txtCaption As String
        txtCaption = "BATCH: " & txtBatchName & "   Page: " & frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).Text
        
'        result = MainMDIForm.funcShowImage(txtBatchDirectory, txtBatchPageFileName, vbNull, vbNull, vbNull, 1, 1)
        
        funcWriteToDebugLog Me.name, "MainMDIForm.funcShowImage - " & txtCaption

        result = MainMDIForm.funcShowImage(txtBatchDirectory, strBatchPageFileName, 1, strBatchPageRECID, txtCaption, 1, 1, txtBatchPageRotation, "", "", "", gI101ModuleIndex)
        
        
        If result = "-1" Then
                    'MainMDIForm.funcShowImage returned an error (-1), when trying to Show the image, get out now.
                    Exit Sub
        End If
        
        bolObjectLaunched = False
        'If Object was NOT Launched... Initialize the Child Form
        If bolObjectLaunched = False Then
            
            funcWriteToDebugLog Me.name, "MainMDIForm.ActiveForm.subInitializeChildForm"
            MainMDIForm.ActiveForm.subInitializeChildForm
            
'            If UCase(Right(txtFullPathName, 4)) = "XLSX" Then
'                'Skip subSetCurrentPage
'            Else
                 MainMDIForm.ActiveForm.subSetCurrentPage
'            End If

        Else
            
            '*** Object was Launched!  Set basic page info.
            funcWriteToDebugLog Me.name, "*** Object Launched ***"

            MainMDIForm.ActiveForm.txtPageCount = 1
            MainMDIForm.ActiveForm.txtPageNumber = 1
            
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "Object Launched"
            MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
            MainMDIForm.ActiveForm.SpicerView1.Visible = False
            MainMDIForm.ActiveForm.lstPageList.Visible = False
            MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
            MainMDIForm.ActiveForm.txtChildFormMessage.Visible = True
        End If
   End If
    
   If Not blnCommittingBatchPages Then
        frmIndex.ListView1.Visible = True
        frmIndex.ListView1.SetFocus
    End If
    
    frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).Selected = True
    frmIndex.txtBatchPageRECID = strBatchPageRECID
    ' Hold the Current Item's Index value to use when user Clicks and "Jumps" to another item
    intListView1CurrentItem = ListView1.SelectedItem.Index
    ListView1.SelectedItem.Bold = True
    
    
Exit Sub

ERROR_HANDLER:

        MsgBox "ListView1_Click: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Data NOT Imported]", vbExclamation
Resume Next
'        If connImaging101Batch.BeginTrans = True Then
'            connImaging101Batch.RollbackTrans
'        End If
'        If connImaging101.BeginTrans = True Then
'            connImaging101.RollbackTrans
'        End If
        Screen.MousePointer = vbDefault

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




Private Sub ListView1_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = Asc("[") And frmIndex.Visible = True Then
        frmIndex.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("]") And frmLookupList.Visible = True Then
        frmLookupList.SetFocus
        frmLookupList.txtTableLookupField.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("\") And MainMDIForm.Visible = True Then
        MainMDIForm.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            ListView1_Click
    End Select

End Sub

Private Sub mebIndexValues_Change(Index As Integer)
    
    'Remove the Prompt Character and Trim the field
    'Underline if there is a Value in the field to signal the enabled Hyperlink
    If Trim(Replace(mebIndexValues(Index).Text, mebIndexValues(Index).PromptChar, "")) <> "" Then
'        lblFieldDescription(Index).ForeColor = vbBlue
        lblFieldDescription(Index).FontUnderline = True
    Else
'        lblFieldDescription(Index).ForeColor = vbBlack
        lblFieldDescription(Index).FontUnderline = False
    End If
    
    
End Sub

Private Sub mebIndexValues_GotFocus(Index As Integer)
        
        mebIndexValues(Index).SelStart = 0
        mebIndexValues(Index).SelLength = mebIndexValues(Index).MaxLength
        ' Hold the field with the current focus
        intHoldFocusIndex = Index
        DoEvents

End Sub



Private Sub mebIndexValues_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = Asc("[") And frmIndex.Visible = True Then
        frmIndex.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If

    If KeyAscii = Asc("]") And frmLookupList.Visible = True Then
        frmLookupList.SetFocus
        frmLookupList.txtTableLookupField.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If

    If KeyAscii = Asc("\") And MainMDIForm.Visible = True Then
        MainMDIForm.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
       KeyAscii = 0
    End If

End Sub

Private Sub mebIndexValues_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

            'Catch Alt + RightClick
            If (Shift = vbAltMask) And (KeyCode = vbKeyC) Then
            
                funcQuickMessage "SHOW", "Key Press = Alt+C for field " & lblFieldDescription(Index)
                
            End If


End Sub

Private Sub mebIndexValues_LostFocus(Index As Integer)
    
    'Force Validation
    mebIndexValues_Validate Index, False
    
End Sub

Public Sub mebIndexValues_Validate(Index As Integer, Cancel As Boolean)
    
    bolErrorOccured = False
    
    On Error GoTo ERROR_TRAP
    
    If (txtFieldType(Index) = "Date") And (Trim(mebIndexValues(Index).Text) <> "") Then
    
        strDateFormatted = Replace(Trim(mebIndexValues(Index).FormattedText), "_", "")
'        strDateFormatted = CDate(strDateFormatted)
        
        If IsDate(strDateFormatted) = False Then
            MsgBox "The DATE you entered is NOT VALID!" & vbCrLf & vbCrLf & _
                            "Please CORRECT and Try Again.", vbOKOnly, _
                            "Date Validation Error"
                            
            Me.SetFocus
            mebIndexValues(Index).Text = vbNullString
            mebIndexValues(Index).SetFocus
            
            bolErrorOccured = True
            Cancel = True
            Exit Sub
        End If
        
        '* Remove the "Prompt" characters
    End If
            
            
            
            
            
Exit Sub

ERROR_TRAP:

    If Err.Number = 13 Then
        result = MsgBox("Field Format Error: " & Err.Number & " - " & Err.Description & vbCrLf & "PLEASE CHECK YOUR INPUT.", vbOK)
        Me.SetFocus
        mebIndexValues(Index).SetFocus
        bolErrorOccured = True
        Err.Clear
        'Prevent moving to the next field
        Cancel = True
        'Force the Field to Highlight
        mebIndexValues_GotFocus (Index)
'        Exit Sub
    End If


End Sub






Private Sub txtIndexValues_GotFocus(Index As Integer)
        
     'Re-size the Text field.
'    If txtIndexValues(Index).Height = mebIndexValues(Index).Height Then
        txtIndexValues(Index).BackColor = vbYellow
        txtIndexValues(Index).Refresh
        txtIndexValues(Index).Height = txtIndexValues(Index).Height * 3
'        txtIndexValues(Index).SelStart = 0
'        txtIndexValues(Index).SelLength = 0
'    Else
'        txtIndexValues(Index).BackColor = vbWhite
'        txtIndexValues(Index).Height = mebIndexValues(Index).Height
'        txtIndexValues(Index).SelStart = 0
'        txtIndexValues(Index).SelLength = 0
'    End If
    
'        txtIndexValues(Index).SelStart = 0
'        txtIndexValues(Index).SelLength = txtIndexValues(Index).MaxLength
        ' Hold the field with the current focus
        intHoldFocusIndex = Index

End Sub

Private Sub txtIndexValues_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = Asc("[") And frmIndex.Visible = True Then
        frmIndex.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("]") And frmLookupList.Visible = True Then
        frmLookupList.SetFocus
        frmLookupList.txtTableLookupField.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("\") And MainMDIForm.Visible = True Then
        MainMDIForm.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If

End Sub



Public Function funcValidateDate(Index As Integer)

'    Dim strHoldDate As Date
    blnDateError = False
    If txtFieldType(Index) = "Date" Then
        On Error Resume Next
'        strHoldDate = CDate(Left(mebIndexValues(Index), 2) & "/" & Mid(mebIndexValues(Index), 3, 2) & "/" & Right(mebIndexValues(Index), 2))
        If mebIndexValues(Index).Text <> "" And IsDate(mebIndexValues(Index).FormattedText) = False Then
            blnDateError = True
        End If
    End If
    On Error GoTo 0

End Function

Private Sub SSTab1_GotFocus()
''    fldIndexDescription(0).SetFocus
End Sub


Private Sub cmdClearFieldValues_Click()
    Dim strHoldFieldMask As String
    For intIndex = 0 To mebIndexValues.Count - 1
        '* If field is NOT flagged as "Sticky" , Clear it.
''        If txtFieldIsSticky(intIndex) <> "1" Then
            ' VB Mask Glitch - Must Remove/Clear the Mask before Clearing the Field Value
            strHoldFieldMask = mebIndexValues(intIndex).Mask
            mebIndexValues(intIndex).Mask = ""
            mebIndexValues(intIndex).Text = ""
            mebIndexValues(intIndex).Mask = strHoldFieldMask
            
            txtIndexValues(intIndex).Text = ""

''        End If
    Next

    txtCommitViaFTP = "0"

    If gOpenBatchInReadOnlyMode = False Then
        If frmLookupList.Visible = True Then
            frmLookupList.txtTableLookupField.SetFocus
            frmLookupList.txtTableLookupField.SelStart = 0
            frmLookupList.txtTableLookupField.SelLength = Len(frmLookupList.txtTableLookupField)
            DoEvents
        End If
    End If
    
End Sub

Public Sub SetAccountNumber(PatientAccountNumber As String)

    If frmLookupList.Visible = True Then
        frmLookupList.txtTableLookupField = Mid(PatientAccountNumber, InStr(1, PatientAccountNumber, "-") + 1, Len(PatientAccountNumber) - InStr(1, PatientAccountNumber, "-"))
    End If
    
End Sub

Private Sub subSaveBatchPageValues(Optional ByRef connImaging101Batch As ADODB.Connection, Optional ByRef rsImaging101BatchPage As ADODB.Recordset)
    
    Dim intFieldIndex As Integer
    Dim bolConnectionNotPassed As Boolean
    Dim bolResultSetNotPassed As Boolean
    
    
    On Error Resume Next
    
    If intListViewSelectedItemIndex > 0 Then
        ' If parameter was passed, select the fields for the page requested
        txtBatchPageRECID = ListView1.ListItems(intListViewSelectedItemIndex).ListSubItems(1).Text
        txtBatchPageFileName = ListView1.ListItems(intListViewSelectedItemIndex).ListSubItems(2).Text
        frmIndex.ListView1.ListItems.item(intListViewSelectedItemIndex).ListSubItems(3).Text = frmIndex.txtBatchDocDesc
        frmIndex.ListView1.ListItems.item(intListViewSelectedItemIndex).ListSubItems(5).Text = frmIndex.txtCommitViaFTP
    Else
        ' Load the Currently selected page fields
        txtBatchPageRECID = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
        txtBatchPageFileName = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
        'Assign the values to the Page List
        frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(3).Text = frmIndex.txtBatchDocDesc
        frmIndex.ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(5).Text = frmIndex.txtCommitViaFTP
    End If
    


   
    If connImaging101Batch Is Nothing Then
         ' Set Connection Properties
        bolConnectionNotPassed = True
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.Open RegImaging101BatchListConnectionString
    End If
    
    If rsImaging101BatchPage Is Nothing Then
        bolResultSetNotPassed = True
        Set rsImaging101BatchPage = New ADODB.Recordset
    End If

    Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
    
    connImaging101Batch.Errors.Clear
    
    '*  Begin Transaction
'    connImaging101Batch.BeginTrans
    
    rsImaging101BatchPage.CursorLocation = adUseServer
    rsImaging101BatchPage.CursorType = adOpenDynamic
    rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
    
    rsImaging101BatchPage.Source = "Select * from " & txtApplicationName & "_BatchPage" & " WHERE BatchRECID = " & txtBatchRECID & " AND BatchPageRECID = " & txtBatchPageRECID
    rsImaging101BatchPage.Open
    
    
    rsImaging101BatchPage.MoveFirst
    
    '*** Cycle through Field Values
    For intIndex = 0 To mebIndexValues.Count - 1
            
            ' Get Field Index by comparing the Form fieldname with the DB result set fieldname
            For intFieldIndex = 0 To rsImaging101BatchPage.Fields.Count - 1
                If rsImaging101BatchPage.Fields(intFieldIndex).name = txtFieldName(intIndex) Then
                    Exit For
                End If
            Next
            
            '*** 2020-05-05 - Jacob - Added code to prevent overwriting the DocNotes Field, because it may contain the FilePath of the Imported file.
            '*** 2020-05-27 - Jacob - Changed from <> ""  to if first character is a BackSlash "\"
           Dim bolSkipThisField As Boolean
            If rsImaging101BatchPage.Fields(intFieldIndex).name = "DocNotes" _
            And Left(rsImaging101BatchPage.Fields(intFieldIndex), 1) = "\" Then
                    bolSkipThisField = True
            Else
                    bolSkipThisField = False
            End If
            
            '*** 2020-05-05 - Jacob - Added code to prevent overwriting the DocNotes Field,
            If ((Not IsNull(mebIndexValues(intIndex))) _
                    Or (mebIndexValues(intIndex) <> "") _
                    Or (Not IsNull(txtIndexValues(intIndex))) _
                    Or (txtIndexValues(intIndex) <> "")) _
                    And bolSkipThisField = False _
                    Then
                
                
                '*** Check if the txtFieldType is set to LongText...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If txtFieldType(intIndex) = "LongText" Then
                    'Save the TEXT Control Value
                    If txtIndexValues(intIndex).Text = "" Then
                        rsImaging101BatchPage.Fields(intFieldIndex) = Null
                    Else
                        rsImaging101BatchPage.Fields(intFieldIndex) = txtIndexValues(intIndex).Text
                    End If
                    
                Else
                    'Save the MASKED EDIT Control Value
                
                    ' If the field is empty, Set to Null value
'                    If mebIndexValues(intIndex) = "" Then

'                        rsImaging101BatchPage.Fields(intFieldIndex) = Null

                    ' If the field is Date, Format as Date value
                    If txtFieldType(intIndex) = "Date" Then
                        '7/25/2017 - Jacob - Added If test for mebIndexValues(intIndex).Text because it was NOT saving Blank Date Values
                        '                             This also affected the "Clear Values" functionality.
                        If mebIndexValues(intIndex).Text = "" Then
                            rsImaging101BatchPage.Fields(intFieldIndex) = Null
                        Else
                            rsImaging101BatchPage.Fields(intFieldIndex) = Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format)
                        End If
                        If Err.Number <> 0 Then
                            mebIndexValues(intIndex).SetFocus
                            mebIndexValues(intIndex).ForeColor = vbRed
                        End If
                    
                    ' If the field is Numeric, convert to Double
                    ElseIf txtFieldType(intIndex) = "Numeric" Then
                        If Trim(mebIndexValues(intIndex).FormattedText) = "" Then
                            mebIndexValues(intIndex).Text = 0
                        End If

                        rsImaging101BatchPage.Fields(intFieldIndex) = CDbl(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
                    
                    ' If the field is Currency, convert to Currency
                    ElseIf txtFieldType(intIndex) = "Currency" Then
                       '*** 2012-05-28 - Jacob - Added Trim() because "FormattedText" returns spaces
                        If Trim(mebIndexValues(intIndex).FormattedText) = "" Then
                            mebIndexValues(intIndex).Text = 0
                        End If
    
                        rsImaging101BatchPage.Fields(intFieldIndex) = CCur(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
                        
                    Else
                        If mebIndexValues(intIndex).Text = "" Then
                            rsImaging101BatchPage.Fields(intFieldIndex) = Null
                        Else
                            'Use the FieldMask value to format the data
                            '  also Trim it save only up to the defined length of the field
                            '*** 2012-05-28 - Jacob - Added Trim() because "FormattedText" returns spaces
                            If Trim(mebIndexValues(intIndex).Format) = "" Then
                                'NO Format -- Simply Trim and Clip the Field
                                rsImaging101BatchPage.Fields(intFieldIndex) = Left(mebIndexValues(intIndex).Text, txtFieldSize(intIndex))
                            Else
                                rsImaging101BatchPage.Fields(intFieldIndex) = Left(Format(mebIndexValues(intIndex).Text, mebIndexValues(intIndex).Format), txtFieldSize(intIndex))
                            End If
                        End If
                    End If
                
                End If '  txtFieldType(intIndex) = "LongText"
                    
            End If
            
            '* If field flagged as questionable, flag as red
            If mebIndexValues(intIndex).Text = txtQuestionable Or txtIndexValues(intIndex).Text = txtQuestionable Then
                mebIndexValues(intIndex).Font.Bold = True
                mebIndexValues(intIndex).ForeColor = vbRed
                txtIndexValues(intIndex).Font.Bold = True
                txtIndexValues(intIndex).ForeColor = vbRed
            Else
                mebIndexValues(intIndex).Font.Bold = False
                mebIndexValues(intIndex).ForeColor = vbNormal
                txtIndexValues(intIndex).Font.Bold = False
                txtIndexValues(intIndex).ForeColor = vbNormal
            End If
        
            Debug.Print rsImaging101BatchPage.Fields(intFieldIndex).name & " = " & rsImaging101BatchPage.Fields(intFieldIndex)
            
    Next
    
    rsImaging101BatchPage.Fields!BatchDocDesc = txtBatchDocDesc
    
    '*** 2020-05-05 - Jacob - Added count of Pages INDEXED if the BatchPageIndexDate is NULL
    If IsNull(rsImaging101BatchPage.Fields!BatchPageIndexDate) Then
            Set rsImaging101Batch = New ADODB.Recordset
            Set rsImaging101Batch.ActiveConnection = connImaging101Batch
            connImaging101Batch.Errors.Clear
            rsImaging101Batch.CursorLocation = adUseServer
            rsImaging101Batch.CursorType = adOpenDynamic
            rsImaging101Batch.LOCKTYPE = adLockOptimistic
            
            rsImaging101Batch.Source = "Select BatchPagesIndexed from I101Batches" & " WHERE BatchRECID = " & txtBatchRECID
            rsImaging101Batch.Open
            
            rsImaging101Batch.MoveFirst

            rsImaging101Batch!BatchPagesIndexed = rsImaging101Batch!BatchPagesIndexed + 1
            rsImaging101Batch.Update
            'Close connection and the recordset
            rsImaging101Batch.Close
            Set rsImaging101Batch = Nothing
    End If
    
    rsImaging101BatchPage.Fields!BatchPageIndexDate = Now()
    rsImaging101BatchPage.Fields!BatchPageIndexUser = gsecUserID
    rsImaging101BatchPage.Fields!BatchPagePageCount = MainMDIForm.ActiveForm.txtPageCount
    rsImaging101BatchPage.Fields!CommitViaFTP = txtCommitViaFTP
    
    ' Update and Commit the Transactions
    rsImaging101BatchPage.Update
'    connImaging101Batch.CommitTrans
    
   
    'Close connection and the recordset
    rsImaging101BatchPage.Close
    
    If bolResultSetNotPassed = True Then
        Set rsImaging101BatchPage = Nothing
    End If
    
    If bolConnectionNotPassed Then
        connImaging101Batch.Close
        Set connImaging101Batch = Nothing
    End If

End Sub


Private Sub subSplitBatchToMultipleBatches()

    '**********************************
    '*** Imaging101 DB Connection Setup
    
    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER:  subSplitBatchToMultipleBatches()"

    
    Dim txtFilterStatement As String
    Dim txtFieldsList As String
    Dim txtValuesList As String
    Dim txtOrderByList As String
    Dim txtFieldNameHold As String
    
    Dim txtSplitFieldName As String
    Dim txtSplitFieldValue As String
    
    Dim txtRouteQueueFieldName As String
    Dim txtRouteQueueFieldValue As String
    
    Dim txtRouteUserFieldName As String
    Dim txtRouteUserFieldValue As String

    Dim txtRouteManagerFieldName As String
    Dim txtRouteManagerFieldValue As String

                    
    Dim txtValuesListHold As String
    Dim txtDocumentRECID As Double
    Dim txtDetailRECID As Double
    Dim intDetailOrder As Integer
    Dim txtDestinationFilename As String
    Dim txtDestinationFileType As String
    
    
    txtBatchSplitRECID = 0
    
     '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True
    
    
    MainMDIForm.ActiveForm.txtChildFormMessage.Visible = True
    MainMDIForm.ActiveForm.txtChildFormMessage.Text = "SPLITTING BATCH!"
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "SPLITTING BATCH"
    MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
    MainMDIForm.ActiveForm.lstPageList.Visible = False
    MainMDIForm.ActiveForm.SpicerView1.Visible = False
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).AutoSize = sbrContents
    
    
    '*********************************************************
    '*** CLOSE THE IMAGE IN THE VIEWER
    '*** TO PREVENT "Runtime Error 75: File/Path Access Error"
    '*** WHEN PROCESSING SINGLE-PAGE PDF's
    '*** WHICH SEEM TO REMAIN "IN-USE" WHEN OPEN.
    'Close the document to release it
    
    funcWriteToDebugLog Me.name, "MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument"
    
    MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument False
    
                    
                    
    intBatchPageCount = 0
    
    
        '***********************************
        '* Loop Through Batch Pages
        '***********************************
        
        funcWriteToDebugLog Me.name, "    Loop Through Batch Pages"

        For intPageIndex = 1 To ListView1.ListItems.Count
        
            funcWriteToDebugLog Me.name, "    Batch Pages " & intPageIndex
        
            frmCommitStatus.txtPagesProcessed = frmCommitStatus.txtPagesProcessed + 1
            bolSkipPage = False
            
            '* Loop Through Fields
            frmIndex.ListView1.ListItems.item(intPageIndex).Selected = True
            subGetBatchFieldValues 0
            
            ' Locate the Record
''''            frmIndex.ListView1.SetFocus
            ListView1_Click
                
            
            If txtBatchPageStatus <> "Split" Then
            
               'FIRST Loop to see if this record should be skipped
               For intIndex = 0 To mebIndexValues.Count - 1
                   
                   ' If ANY Field is "Questionable" - Skip this record
                   If mebIndexValues(intIndex).Text = txtQuestionable _
                   Or txtIndexValues(intIndex).Text = txtQuestionable _
                   Then
                        frmCommitStatus.txtPagesQuestionable = frmCommitStatus.txtPagesQuestionable + 1
                        bolSkipPage = True
                        Exit For
                   End If
                       
                   ' If ANY Field is flagged as "Separator" - Skip this record
                   If mebIndexValues(intIndex).Text = txtSeparator _
                   Or txtIndexValues(intIndex).Text = txtSeparator _
                   Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
                   
                   ' If ANY Field is flagged as "*DO NOT FILE*" - Skip this record
                   If mebIndexValues(intIndex).Text = txtDoNotFile _
                   Or txtIndexValues(intIndex).Text = txtDoNotFile _
                   Then
                        frmCommitStatus.txtPagesSeparator = frmCommitStatus.txtPagesSeparator + 1
                        bolSkipPage = True
                        Exit For
                   End If
               
               Next
                        
               'Only check for Fields Required or Valid if Not flagged for Skip
               If bolSkipPage <> True Then
                    For intIndex = 0 To mebIndexValues.Count - 1
                     
                        ' If ANY Field is flagged as "Required" but is Empty - Skip this record
                        If (txtFieldIsRequiredForSplit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "" And txtIndexValues(intIndex).Text = "") Then
                             frmCommitStatus.txtPagesRequiredButEmpty = frmCommitStatus.txtPagesRequiredButEmpty + 1
                             bolSkipPage = True
                             Exit For
                        End If
                        
                         '*** VALIDATE DATE FIELD!
                         funcValidateDate intHoldFocusIndex
                         If blnDateError = True Then
                             frmCommitStatus.txtPagesFailedValidation = frmCommitStatus.txtPagesFailedValidation + 1
                             bolSkipPage = True
                             Exit For
                         End If
                         
                    Next
               End If   'bolSkipPage <> True
                   
            Else
                frmCommitStatus.txtPagesPreviouslyCommitted = frmCommitStatus.txtPagesPreviouslyCommitted + 1
                bolSkipPage = True
            End If
            
            If bolSkipPage <> True Then
                
                    funcWriteToDebugLog Me.name, "    bolSkipPage = FALSE... Establish DB Connections"

                    '**************************************************************
                    '*** Establish BATCH DB Connection
                    Set connImaging101Batch = New ADODB.Connection
                    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
                    connImaging101Batch.ConnectionTimeout = 120
                    connImaging101Batch.mode = adModeReadWrite
                    connImaging101Batch.Open
                    connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"
                            
                    '**************************************************************
                    '*** CONNECT to Batch DB
                    Set rsImaging101Batch = New ADODB.Recordset
                    txtActionBeforeError = "Connect to Batch DB"
                    Set rsImaging101Batch.ActiveConnection = connImaging101Batch
                    rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
                    rsImaging101Batch.CursorLocation = adUseServer
                    rsImaging101Batch.CursorType = adOpenDynamic
                    rsImaging101Batch.LOCKTYPE = adLockOptimistic
                    txtActionBeforeError = "Open Batch DB"
                    rsImaging101Batch.Open
                    
                    '**************************************************************
                    '*** CONNECT to Batch Page DB
                    Set rsImaging101BatchPage = New ADODB.Recordset
                    txtActionBeforeError = "Connect to Batch Pages DB"
                    Set rsImaging101BatchPage.ActiveConnection = connImaging101Batch
                    rsImaging101BatchPage.Source = "SELECT * FROM " & txtApplicationName & "_BatchPage" & " WHERE BatchPageRECID= " & txtBatchPageRECID
                    rsImaging101BatchPage.CursorLocation = adUseServer
                    rsImaging101BatchPage.CursorType = adOpenDynamic
                    rsImaging101BatchPage.LOCKTYPE = adLockOptimistic
                    txtActionBeforeError = "Open Batch Page DB"
                    rsImaging101BatchPage.Open
                    

                
                    '****************************************************************************
                    '*** Begin Transaction
                    connImaging101Batch.BeginTrans
                            
                    
                    '**************************************************************
                    '*** CREATE the Directory Structure for storing this Batch
                    On Error GoTo 0

                    
                    '**************************************************************
                    '*** PREPARE TO INSERT DOCUMENT AND DETAIL RECORDS
                   
                    '*** Clear variables
                    txtFilterStatement = ""
                    txtFieldsList = ""
                    txtSplitFieldName = ""
                    txtSplitFieldValue = ""
                    txtRouteQueueFieldName = ""
                    txtRouteQueueFieldValue = ""
                    txtValuesList = ""
                    txtOrderByList = ""
                    txtFieldNameHold = ""
                    
                    
                    '****************************************************************************
                    '*** Prepare the BREAK Clause for Split
                    '***  Append the txtValuesList because we don't know which we'll find first.
                    
                    funcWriteToDebugLog Me.name, "    Prepare the BREAK Clause for Split"

                    For intIndex = 0 To lblFieldDescription.Count - 1
                        
                        'Find the Field flagged as "SPLIT Batches On This Field".
                        If txtFieldSplitBatches(intIndex).Text = "1" Then
                            txtSplitFieldName = txtFieldName(intIndex).Text
                            txtSplitFieldValue = rsImaging101BatchPage.Fields("" & txtSplitFieldName & "") & ""
                            txtValuesList = txtValuesList & txtSplitFieldValue & " "
                        End If
                        
                        'Find the Field Flagged as "Route to Batch QUEUE based on this Field"
                        If txtFieldRouteToBatchQueue(intIndex).Text = "1" Then
                            txtRouteQueueFieldName = txtFieldName(intIndex).Text
                            txtRouteQueueFieldValue = rsImaging101BatchPage.Fields("" & txtRouteQueueFieldName & "") & ""
                            txtValuesList = txtValuesList & txtRouteQueueFieldValue & " "
                        End If
                        
                        'Find the Field Flagged as "Route to Batch USER based on this Field"
                        If txtFieldRouteToBatchUser(intIndex).Text = "1" Then
                            txtRouteUserFieldName = txtFieldName(intIndex).Text
                            txtRouteUserFieldValue = rsImaging101BatchPage.Fields("" & txtRouteUserFieldName & "") & ""
                            txtValuesList = txtValuesList & txtRouteQueueFieldValue & " "
                        End If
                        
                        'Find the Field Flagged as "Route to MANAGER based on this Field"
                        If txtFieldRouteToBatchManager(intIndex).Text = "1" Then
                            txtRouteManagerFieldName = txtFieldName(intIndex).Text
                            txtRouteManagerFieldValue = rsImaging101BatchPage.Fields("" & txtRouteManagerFieldName & "") & ""
                            txtValuesList = txtValuesList & txtRouteQueueFieldValue & " "
                        End If


                        
                    Next
                    
                    
                    '****************************************************************************
                    '*** Create a New BATCH record
                    '***    only if the Index Values are Different from the previous record
                    
                    If txtValuesList <> txtValuesListHold Then
                    
                    
                           If txtBatchSplitRECID <> 0 Then
                                'This means a NEW Batch will be created
                                ' 2014-04-25 - Jacob - UN-LOCK the previous Batch, ignore errors
                                
                                funcWriteToDebugLog Me.name, "    UN-LOCK the previous Batch, ignore errors"
                                strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchSplitRECID)
                           End If
                                               
                            'CREATE a New SPLIT Batch
                            '  This should set the "S" lock flag
                            funcWriteToDebugLog Me.name, "    CREATE a New SPLIT Batch"
                            subCreateBatchRecord txtSplitFieldName, txtSplitFieldValue, txtRouteQueueFieldName, txtRouteQueueFieldValue, txtRouteUserFieldName, txtRouteUserFieldValue, txtRouteManagerFieldName, txtRouteManagerFieldValue
                            
                            
                            'Reset the Split Page Counter
                            intBatchPageCount = 0
                    
                    End If
                    
                    
                        
                        '****************************************************************************
                        '*** Increment the New Split Batch Page Counter
                        
                        intBatchPageCount = intBatchPageCount + 1
                        funcWriteToDebugLog Me.name, "    Increment the New Split Batch Page Counter.  intBatchPageCount = " & intBatchPageCount
                        
                        
                        '****************************************************************************
                        '*** COPY the file to the Storage Destination
                        Dim strSourceFile As String
                        Dim strDestinationFile As String
                        
                        strSourceFile = Trim(rsImaging101Batch.Fields("BatchDirectory")) & "\" & txtBatchPageFileName
                        
                        '*** 2023-02-20 - Jacob - Modified to NOT Copy the BatchPage or Annotation files,
                        '                                        but rather leave them in the Original Batch's directory
                        
                        strDestinationFile = strSourceFile
                        
'                        strDestinationFile = txtBatchDirectory & "\" & txtBatchPageFileName
'
'                        txtActionBeforeError = "FileCopy [" & strSourceFile & "] TO [" & strDestinationFile & "]"
'                        funcWriteToDebugLog Me.name, "    FileCopy [" & strSourceFile & "] TO [" & strDestinationFile & "]"
'
'                        FileCopy strSourceFile, strDestinationFile
'
'                        '*** Make SURE the file got copied properly before Deleting the original
'                        If funcFileExists(strDestinationFile) Then
'                            funcWriteToDebugLog Me.name, "    File Exists in Dest [" & strDestinationFile & "]"
'                            funcWriteToDebugLog Me.name, "    KILL original file     [" & strSourceFile & "]"
'
'                            Kill strSourceFile
'                        Else
'
'                            connImaging101Batch.RollbackTrans
'
'                            If rsImaging101Batch.State = adStateOpen Then
'                                rsImaging101Batch.Close
'                            End If
'                            Set rsImaging101Batch = Nothing
'
'                            If rsImaging101BatchPage.State = adStateOpen Then
'                                rsImaging101BatchPage.Close
'                            End If
'                            Set rsImaging101BatchPage = Nothing
'
'                            If rsImaging101BatchSplit.State = adStateOpen Then
'                                rsImaging101BatchSplit.Close
'                            End If
'                            Set rsImaging101BatchSplit = Nothing
'
'                            Set connImaging101Batch = Nothing
'
'                            funcQuickMessage "SHOW", "ERROR: File Not found at Destination after Action (" & txtActionBeforeError & ")... TRANSACTION ROLLED BACK!"
'
'                            Exit Sub
'                        End If
'
                        
                        
                        
'                    '*** Now COPY the Annotation files, if any, to the Storage Destination
'                    Dim strSourceAnnotationFile As String
'                    Dim strDestinationAnnotationFile As String
'                    Dim strPageNumber As String
'
'                    AnnotationFileListBox.Path = Trim(rsImaging101Batch.Fields("BatchDirectory"))
'                    AnnotationFileListBox.Pattern = Left(txtBatchPageFileName, InStrRev(txtBatchPageFileName, ".") - 1) & _
'                                                "_*.ANN"
'                    DoEvents
'
'                    If AnnotationFileListBox.ListCount > 0 Then
'                        For dblAnnotationFileIndex = 0 To AnnotationFileListBox.ListCount - 1
'
'                            AnnotationFileListBox.Selected(dblAnnotationFileIndex) = True
'                            strSourceAnnotationFile = Trim(rsImaging101Batch.Fields("BatchDirectory")) & "\" & AnnotationFileListBox.FileName
'
'                            ' Build the Annotation FilePath
'                            strFullDirectoryPathForAnnotation = txtBatchDirectory
'                            txtActionBeforeError = "Create Directory Structure: " & strFullDirectoryPathForAnnotation
'
'                            'Extract the Page # from the Source Filename
'                            strPageNumber = Mid(AnnotationFileListBox.FileName, InStrRev(AnnotationFileListBox.FileName, "_") + 1, 6)
'
'                            'Create the directory if needed.
'                            funcCreateDirectoryStructure strFullDirectoryPathForAnnotation & ""
'                            strDestinationAnnotationFile = strFullDirectoryPathForAnnotation & "\" & AnnotationFileListBox.FileName
'
'                            txtActionBeforeError = "FileCopy [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"
'                            FileCopy strSourceAnnotationFile, strDestinationAnnotationFile
'
'                        Next
'                    End If
                    
                    
                                            
                        
                        
                        
                        
                        '****************************************************************************
                        '*** FLAG PAGE RECORD as SPLIT, Update and Commit the Transaction
                        rsImaging101BatchPage.Fields!BatchPageStatus = "Split"
                        rsImaging101BatchPage.Fields!BatchRECID = txtBatchSplitRECID
                        rsImaging101BatchPage.Fields!BatchPageOrder = intBatchPageCount
                        rsImaging101BatchPage.Fields!BatchPageFileName = txtBatchPageFileName
                        
    
    '                       connImaging101Batch.CommitTrans
    '                       connImaging101.CommitTrans
            
                        
                        '**************************************************************
                        '*** CONNECT to Batch DB to Create SPLIT Record
                        Set rsImaging101BatchSplit = New ADODB.Recordset
                        txtActionBeforeError = "Connect to Batch DB"
                        Set rsImaging101BatchSplit.ActiveConnection = connImaging101Batch
                        rsImaging101BatchSplit.Source = "SELECT * FROM I101Batches  WHERE BatchRECID = " & txtBatchSplitRECID
                        rsImaging101BatchSplit.CursorLocation = adUseServer
                        rsImaging101BatchSplit.CursorType = adOpenDynamic
                        rsImaging101BatchSplit.LOCKTYPE = adLockOptimistic
                        txtActionBeforeError = "Open Batch DB"
                        rsImaging101BatchSplit.Open
                        
                        '****************************************************************************
                        '*** FLAG PAGE RECORD as INDEXED, Update and Commit the Transaction
                        rsImaging101BatchSplit("BatchPagesNotCommitted") = intBatchPageCount
                        rsImaging101BatchSplit("BatchPagesTotal") = intBatchPageCount
                        
                        
                        
                        '*** 2023-02-20 - Jacob - Modified to NOT Copy the BatchPage or Annotation files,
                        '                                        but rather leave them in the Original Batch's directory
                        
                        rsImaging101BatchSplit("BatchDirectory") = Trim(rsImaging101Batch.Fields("BatchDirectory"))
                        
                    
                    frmCommitStatus.txtPagesCommitted = frmCommitStatus.txtPagesCommitted + 1
                    
                    
                    
                    '**************************************************************
                    '*** UPDATE THE ORIGINAL BATCH
                    '*** FLAG BATCH RECORD as Committed, set counters and Update
                    If intPageIndex = ListView1.ListItems.Count _
                    And frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
                        'If all pages are processed AND no pages requiring action are left
                        rsImaging101Batch.Fields!BatchCommitStatus = "Split-FULL"
                    Else
                        rsImaging101Batch.Fields!BatchCommitStatus = "Split-PARTIAL"
                    End If
                    
                    rsImaging101Batch.Fields!BatchPagesTotal = rsImaging101Batch.Fields!BatchPagesTotal - 1
                    rsImaging101Batch.Fields!BatchSplitDate = Now()
                    rsImaging101Batch.Fields!BatchSplitUser = gsecUserID
                    
                    
                    
                    '****************************************************************************
                    '*** UPDATE TRANSACTIONS
                    
                    txtActionBeforeError = "Update BatchSplit  Table"
                    rsImaging101BatchSplit.Update
                    txtActionBeforeError = "Update BatchPage Table"
                    rsImaging101BatchPage.Update
                    txtActionBeforeError = "Update Batch        Table"
                    rsImaging101Batch.Update
                    
                    
                    '****************************************************************************
                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                    txtActionBeforeError = "Commit Transactions"
                    connImaging101Batch.CommitTrans
                    
                    If rsImaging101Batch.State = adStateOpen Then
                        rsImaging101Batch.Close
                    End If
                    Set rsImaging101Batch = Nothing
                    
                    If rsImaging101BatchPage.State = adStateOpen Then
                        rsImaging101BatchPage.Close
                    End If
                    Set rsImaging101BatchPage = Nothing

                    If rsImaging101BatchSplit.State = adStateOpen Then
                        rsImaging101BatchSplit.Close
                    End If
                    Set rsImaging101BatchSplit = Nothing
                    
                    Set connImaging101Batch = Nothing
                    
                    
                    
                    
                    '*******************************************************
                    '*** Save the Breach Value to Compare on Next Page
                    txtValuesListHold = txtValuesList
                
                    
                    
            Else
            
                frmCommitStatus.txtPagesTotalSkipped = frmCommitStatus.txtPagesTotalSkipped + 1

            End If
            
            DoEvents
        
        Next
        
        '**************************************************************
        '**   FINAL UPDATE OF THE BATCH AFTER ALL PAGES PROCESSED   ***
        '**************************************************************
        
        '**************************************************************
        '*** Establish BATCH DB Connection
        Set connImaging101Batch = New ADODB.Connection
        connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
        connImaging101Batch.ConnectionTimeout = 120
        connImaging101Batch.mode = adModeReadWrite
        connImaging101Batch.Open
        connImaging101Batch.Execute "SET LOCK_TIMEOUT -1"

        
        '*** CONNECT to Batch DB
        Set rsImaging101Batch = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101Batch.ActiveConnection = connImaging101Batch
        rsImaging101Batch.Source = "SELECT * FROM I101Batches  WHERE BatchRECID= " & txtBatchRECID
        rsImaging101Batch.CursorLocation = adUseServer
        rsImaging101Batch.CursorType = adOpenDynamic
        rsImaging101Batch.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101Batch.Open
        
        'Must Convert the PagesCommitted and txtPagesPreviouslyCommitted to INTEGERS
        rsImaging101Batch.Fields!BatchPagesCommitted = CInt(frmCommitStatus.txtPagesCommitted) + CInt(frmCommitStatus.txtPagesPreviouslyCommitted)
        rsImaging101Batch.Fields!BatchPagesNotCommitted = CInt(rsImaging101Batch.Fields!BatchPagesTotal) - CInt(rsImaging101Batch.Fields!BatchPagesCommitted)
        rsImaging101Batch.Fields!BatchcPagesQuestionable = CInt(frmCommitStatus.txtPagesQuestionable)
        rsImaging101Batch.Fields!BatchPagesDoNotFile = CInt(frmCommitStatus.txtPagesDoNotFile)
        rsImaging101Batch.Fields!BatchPagesSeparator = CInt(frmCommitStatus.txtPagesSeparator)
        
        '*** FLAG BATCH RECORD as Committed, set counters and Update
        If frmCommitStatus.txtPagesQuestionable + frmCommitStatus.txtPagesRequiredButEmpty + frmCommitStatus.txtPagesFailedValidation = 0 Then
            'If all pages are processed AND no pages requiring action are left
            rsImaging101Batch.Fields!BatchCommitStatus = "Split-FULL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH SPLIT - FULL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH SPLIT - FULL"
            funcWriteToDebugLog Me.name, "BATCH SPLIT - FULL"
        Else
            rsImaging101Batch.Fields!BatchCommitStatus = "Split-PARTIAL"
            MainMDIForm.ActiveForm.txtChildFormMessage.Text = "BATCH SPLIT - PARTIAL!"
            MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "BATCH SPLIT - PARTIAL"
            funcWriteToDebugLog Me.name, "BATCH PARTIAL - PARTIAL"
            
        End If
        
        rsImaging101Batch.Update
        
        If rsImaging101Batch.State = adStateOpen Then
            rsImaging101Batch.Close
        End If
        Set rsImaging101Batch = Nothing
        
        Set connImaging101Batch = Nothing
        
              
        '*** CREATE BATCH AUDIT RECORD
        funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Split Batch"

        
        frmCommitStatus.SetFocus
        
        funcWriteToDebugLog Me.name, "EXIT:  subSplitBatchToMultipleBatches()"

        
Exit Sub

ERROR_HANDLER:

        funcQuickMessage "SHOW", "subSplitBatchToMultipleBatches: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Data NOT Imported]"
        
        On Error Resume Next
        
        Set rsImaging101Batch = Nothing
        Set rsImaging101BatchPage = Nothing
        
        Set connImaging101 = Nothing
                            
        Screen.MousePointer = vbDefault

End Sub

Private Sub subCreateBatchRecord(strSplitFieldName As String, _
                                                    strSplitFieldValue As String, _
                                                    strRouteQueueFieldName As String, _
                                                    strRouteQueueFieldValue As String, _
                                                    strRouteUserFieldName As String, _
                                                    strRouteUserFieldValue As String, _
                                                    strRouteManagerFieldName As String, _
                                                    strRouteManagerFieldValue As String)

    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
        On Error GoTo CREATE_BATCH_RECORD_ERROR
        
         funcWriteToDebugLog Me.name, "    ENTER: subCreateBatchRecord()"
         funcWriteToDebugLog Me.name, "        Establish DB Connections"

        
        '**************************************************************
        '*** CONNECT to Batch DB to Create SPLIT Record
        Set rsImaging101BatchSplit = New ADODB.Recordset
        txtActionBeforeError = "Connect to Batch DB"
        Set rsImaging101BatchSplit.ActiveConnection = connImaging101Batch
        rsImaging101BatchSplit.Source = "SELECT * FROM I101Batches  WHERE 0=1"
        rsImaging101BatchSplit.CursorLocation = adUseServer
        rsImaging101BatchSplit.CursorType = adOpenDynamic
        rsImaging101BatchSplit.LOCKTYPE = adLockOptimistic
        txtActionBeforeError = "Open Batch DB"
        rsImaging101BatchSplit.Open
                    
        'User Transaction Tracking to prevent partial imports!
'        connImaging101Batch.BeginTrans
        
        txtActionBeforeError = "Open Batches Table"
'        Set rsImaging101BatchSplit = New ADODB.Recordset
'        rsImaging101BatchSplit.Open "I101Batches", connImaging101Batch, adOpenDynamic, adLockPessimistic
        
        txtActionBeforeError = "Add New Record"
        funcWriteToDebugLog Me.name, "        Add New Record"

        rsImaging101BatchSplit.AddNew
        
        txtActionBeforeError = "Assign Variables to Fields"
        funcWriteToDebugLog Me.name, "        Assign Variables to Fields"

        txtBatchSplitRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchRECID")
        funcWriteToDebugLog Me.name, "        txtBatchSplitRECID = " & txtBatchSplitRECID


'       If chkBatchAutoName = vbChecked Then
'            txtBatchName = ""
'            txtBatchDirectory = ""
'            If chkBatchAutoUseBatchID = vbChecked Then
'                txtBatchName = txtBatchName & Format(txtBatchRECID, "0000000000")
'            End If
'            If chkBatchAutoUseDateTime = vbChecked Then
'                If txtBatchName <> "" Then
'                    txtBatchName = txtBatchName & "_"
'                End If
'                txtBatchName = txtBatchName & Format(Now(), "yyyy-mm-dd_hh-mm-ss")
'            End If
'        End If
        DoEvents
        
        
        
        '*** 2023-02-20 - Jacob - Modified to NOT Copy the BatchPage or Annotation files,
        '                                        but rather leave them in the Original Batch's directory
        
'        ' CREATE the Directory Structure for storing this Batch
'        txtBatchDirectory = Left(rsImaging101Batch("BatchDirectory"), InStrRev(rsImaging101Batch("BatchDirectory"), "\", -1)) & Format(txtBatchSplitRECID, "0000000000")
'        funcWriteToDebugLog Me.name, "        CREATE the Directory Structure:  " & txtBatchDirectory
'
'        funcCreateDirectoryStructure txtBatchDirectory


        
        rsImaging101BatchSplit("BatchRECID") = txtBatchSplitRECID
        rsImaging101BatchSplit("ApplicationRECID") = txtApplicationRECID
        rsImaging101BatchSplit("BatchApplication") = ""
        rsImaging101BatchSplit("BatchDesc") = txtBatchDesc

        rsImaging101BatchSplit("BatchName") = "SPLIT-" & strSplitFieldValue & txtBatchSuffix
        funcWriteToDebugLog Me.name, "        SPLIT Batch Name =  " & rsImaging101BatchSplit("BatchName")

        
        '2013-05-31 - Jacob - If no field was set as the RouteTo, then set it to the same queue as the parent Batch
        If Trim(strRouteQueueFieldValue) = "" Then
            rsImaging101BatchSplit("BatchQueue") = rsImaging101Batch("BatchQueue")
            funcWriteToDebugLog Me.name, "        strRouteQueueFieldValue =  EMPTY  -  Set BatchQueue Same as Parent Batch = " & rsImaging101BatchSplit("BatchQueue")

        Else
            rsImaging101BatchSplit("BatchQueue") = strRouteQueueFieldValue
            funcWriteToDebugLog Me.name, "        strRouteQueueFieldValue =  NOT-EMPTY  -  Set BatchQueue Same as Parent Batch = " & rsImaging101BatchSplit("BatchQueue")

        End If

        '2016-09-15 - Jacob - If no field was set as the RouteTo, then set Empty
        '                              THIS WILL OVERRIDE the gSetUserAsBatchOwnerOnSPLIT flag !!!
        If Trim(strRouteUserFieldValue) = "" Then
        
            funcWriteToDebugLog Me.name, "        strRouteUserFieldValue =  EMPTY "

            '2013-05-22 - Jacob - Added option to default BatchOwner to the User doing the SPLIT
            If gSetUserAsBatchOwnerOnSPLIT = "1" Then
                rsImaging101BatchSplit("BatchOwner") = gsecUserName
                funcWriteToDebugLog Me.name, "        gSetUserAsBatchOwnerOnSPLIT =  CHECKED -  Set BatchOwner = " & rsImaging101BatchSplit("BatchOwner")
            
            Else
                rsImaging101BatchSplit("BatchOwner") = ""
                funcWriteToDebugLog Me.name, "        gSetUserAsBatchOwnerOnSPLIT =  UN-CHECKED -  Set BatchOwner = BLANK "
    
            End If

'            rsImaging101BatchSplit("BatchOwner") = ""
'            funcWriteToDebugLog Me.name, "        strRouteUserFieldValue =  EMPTY  -  Set BatchOwner  = EMPTY"

        Else
            rsImaging101BatchSplit("BatchOwner") = strRouteUserFieldValue
            funcWriteToDebugLog Me.name, "        strRouteUserFieldValue =  NOT-EMPTY  -  Set BatchOwner = " & rsImaging101BatchSplit("BatchOwner")

        End If

        '2016-09-15 - Jacob - If no field was set as the RouteTo, then set Empty
        If Trim(strRouteManagerFieldValue) = "" Then
            rsImaging101BatchSplit("BatchManager") = ""
            funcWriteToDebugLog Me.name, "        strRouteManagerFieldValue =  EMPTY  -  Set BatchManager  = EMPTY"

        Else
            rsImaging101BatchSplit("BatchManager") = strRouteManagerFieldValue
            funcWriteToDebugLog Me.name, "        strRouteManagerFieldValue =  NOT-EMPTY  -  Set BatchManager = " & rsImaging101BatchSplit("BatchManager")
        End If

        
        rsImaging101BatchSplit("BatchStatus") = rsImaging101Batch("BatchStatus")
        rsImaging101BatchSplit("BatchPriority") = rsImaging101Batch("BatchPriority")
        rsImaging101BatchSplit("BatchGroup") = "REGULAR"
        
         '2013-05-22 - Jacob - Changed to set SPLIT Batch Date to same as original Batch instead of Now()
        rsImaging101BatchSplit("BatchScanDate") = rsImaging101Batch("BatchScanDate")
        rsImaging101BatchSplit("BatchDirectory") = txtBatchDirectory
        rsImaging101BatchSplit("BatchNotes") = "SPLIT from Batch: " & rsImaging101Batch("BatchName") & vbCrLf & rsImaging101Batch("BatchNotes")
        rsImaging101BatchSplit("BatchPagesCommitted") = 0
        rsImaging101BatchSplit("BatchPagesNotCommitted") = 0
        rsImaging101BatchSplit("BatchPagesTotal") = 0
        rsImaging101BatchSplit("BatchScanUser") = rsImaging101Batch("BatchScanUser")
        rsImaging101BatchSplit("BatchBoxNumber") = rsImaging101Batch("BatchBoxNumber")
        
        '2014-04-25 - Jacob - Added "S" lock to prevent listing in Batch List if Scanning or Splitting
        rsImaging101Batch.Fields("BatchLocked") = "S"
        rsImaging101Batch.Fields("BatchLockedBy") = gsecUserName
        rsImaging101Batch.Fields("BatchLockedDate") = Now()

        
        txtActionBeforeError = "Update Values"
        funcWriteToDebugLog Me.name, "        Update Values"
        rsImaging101BatchSplit.Update
    
        If rsImaging101BatchSplit.State = adStateOpen Then
            rsImaging101BatchSplit.Close
        End If
        
        Set rsImaging101BatchSplit = Nothing
        

    
        Screen.MousePointer = vbDefault

        funcWriteToDebugLog Me.name, "    EXIT: subCreateBatchRecord()"


Exit Sub
    
CREATE_BATCH_RECORD_ERROR:
        MsgBox "CREATE_BATCH_RECORD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch NOT Imported]", vbExclamation
        
        If connImaging101Batch.BeginTrans = True Then
            connImaging101Batch.RollbackTrans
        End If
        If connImaging101Batch.BeginTrans = True Then
            connImaging101Batch.RollbackTrans
        End If
        Screen.MousePointer = vbDefault

        bolCancelPendingXfers = True
    
End Sub


Public Sub subLoadPagesIntoListView()

    Dim intItemCount As Integer
    
    '*** BATCH DB Connection Setup
    
    '*** LOAD PAGES INTO LIST
    '*** Declarations

    Set con = New ADODB.Connection
    con.Open RegImaging101BatchListConnectionString
    
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LOCKTYPE = adLockReadOnly
    End With
    
    '*** txtBatchRECID SET FOR TESTING ONLY
''    txtBatchRECID = 17
    
    rs.Source = "Select BatchPageOrder, BatchPageRECID, BatchPageFileName, BatchDocDesc, BatchPageRotation, CommitViaFTP from " & txtApplicationName & "_BatchPage" & " WHERE BatchRECID = " & txtBatchRECID & " Order by BatchRECID, BatchPageOrder"
    
''    On Error GoTo ERROR_TRAP
    
    'sql statement to select items on the drop down list
''    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID
''    rs.Open ssql, Con
    con.Errors.Clear
    rs.Open
    
    ListView1.ListItems.Clear  ' Reset the List
    intListView1CurrentItem = 0 ' Clear the hold variable
    
    '*** Setup Up PAGE ListView properties - BEGIN
    
'    ListView1.Visible = False
    
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = False
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = False
    ListView1.Sorted = False
    ListView1.SortKey = 0    ' Sort by ListView1.Text
    
    '***  SET COLUMN HEADINGS
    ListView1.ColumnHeaders.Add , , "Page", ListView1.width, lvwColumnLeft
    ListView1.ColumnHeaders.Add , , "BatchPageRECID", 0, lvwColumnLeft
    ListView1.ColumnHeaders.Add , , "BatchPageFileName", 0, lvwColumnLeft
    ListView1.ColumnHeaders.Add , , "DOCDESC", 0, lvwColumnLeft
    ListView1.ColumnHeaders.Add , , "BatchPageRotation", 0, lvwColumnLeft
    ListView1.ColumnHeaders.Add , , "CommitViaFTP", 0, lvwColumnLeft
    
    Do Until rs.EOF
    
           
                Set lstItem = ListView1.ListItems.Add(, , Format(rs!BatchPageOrder, "0000"))
                Set lstSubItem = lstItem.ListSubItems.Add(, , rs!BatchPageRECID)
                Set lstSubItem = lstItem.ListSubItems.Add(, , rs!BatchPageFileName)
                If Not IsNull(rs!BatchDocDesc) Then
                    Set lstSubItem = lstItem.ListSubItems.Add(, , rs!BatchDocDesc)
                Else
                    Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                End If
                If Not IsNull(rs!BatchPageRotation) Then
                    Set lstSubItem = lstItem.ListSubItems.Add(, , rs!BatchPageRotation)
                Else
                    Set lstSubItem = lstItem.ListSubItems.Add(, , "0")
                End If
                If Not IsNull(rs!CommitViaFTP) Then
                    Set lstSubItem = lstItem.ListSubItems.Add(, , rs!CommitViaFTP)
                Else
                    Set lstSubItem = lstItem.ListSubItems.Add(, , "0")
                End If

        '*** Setup Up ListView properties - END
        
        intItemCount = intItemCount + 1
        
        rs.MoveNext

    Loop
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
        
    If txtBatchPagesTotal = "" Then
        txtBatchPagesTotal = "0"
    End If
        
    If intItemCount <> CInt(txtBatchPagesTotal) Then
        result = funcQuickMessage("SHOW", "WARNING!!!" & vbCrLf & vbCrLf & _
                                                                "The number of Pages Loaded is DIFFERENT" & vbCrLf & _
                                                                "      than the Number of Pages this Batch is supposed to have!" & vbCrLf & vbCrLf & _
                                                                "BATCH PAGES = " & Trim(txtBatchPagesTotal) & "    PAGES LOADED = " & CStr(intItemCount) & vbCrLf & vbCrLf & _
                                                                "Please CLOSE and RE-LOAD this Batch!")
    End If
    
'    ListView1.Visible = True
    
End Sub

Private Sub subSetDefaultFieldValues()
    
    'Check BOTH Text and Masked Edit controls... if they have a value... Exit now!
    If Trim((mebIndexValues(intIndex).Text & txtIndexValues(intIndex)) <> "") Then
       Exit Sub
    End If
    
    
    '* If the Field has a Default Value, set it now
    '  Only use the Default value if NOT Flagged for Barcode Only
    If (Not bolProcessingBarcodes _
        And txtFieldDefaultForBarcodeOnly(intIndex).Text <> "1") _
    Or (bolProcessingBarcodes _
        And txtFieldDefaultForBarcodeOnly(intIndex).Text = "1") Then


        '*** Check if we should use the Long Text (txtIndexValues) or regular (mebIndexValues) field...
        '    this will handle saving the value of the the TextBox control
        '    instead of the mebIndexValues Masked Edit control as needed.
        '    This is because the Masked Edit control has a MAX size of 64 Char.
        If txtFieldType(intIndex).Text = "LongText" Then

            'Use the TEXT Control
            Select Case txtFieldDefaultValue(intIndex)
                Case "[Scan Date]"
                    'Format only the Left 10 Characters
                    '  because we re-formatted the dates in the Batch ListView so they
                    '  would sort properly.
                    txtIndexValues(intIndex).Text = Left(frmImaging101BatchList.txtBatchDate, 10)
                Case "[Index Date]"
                    txtIndexValues(intIndex).Text = Left(Now(), 10)
                Case "[Batch Name]"
                    txtIndexValues(intIndex).Text = txtBatchName
                Case "[Batch Queue]"
                    txtIndexValues(intIndex).Text = txtBatchQueue
'                Case "[Batch Prefix]"
'                    txtIndexValues(intIndex).Text = txtBatchPrefix
'                Case "[Batch Suffix]"
'                    txtIndexValues(intIndex).Text = txtBatchSuffix
                Case "[Batch Owner]"
                    txtIndexValues(intIndex).Text = gsecUserID
                Case Else
                    txtIndexValues(intIndex).Text = txtFieldDefaultValue(intIndex)
            End Select
        
        
        Else
            'Use the MASKED EDIT Control
            Select Case txtFieldDefaultValue(intIndex)
                Case "[Scan Date]"
                    'Format only the Left 10 Characters
                    '  because we re-formatted the dates in the Batch ListView so they
                    '  would sort properly.
                    mebIndexValues(intIndex).Text = Format(Left(frmImaging101BatchList.txtBatchDate, 10), mebIndexValues(intIndex).Format)
                Case "[Index Date]"
                    mebIndexValues(intIndex).Text = Format(Now(), mebIndexValues(intIndex).Format)
                Case "[Batch Name]"
                    mebIndexValues(intIndex).Text = txtBatchName
                Case "[Batch Queue]"
                    mebIndexValues(intIndex).Text = txtBatchQueue
'                Case "[Batch Prefix]"
'                    mebIndexValues(intIndex).Text = txtBatchPrefix
'                Case "[Batch Suffix]"
'                    mebIndexValues(intIndex).Text = txtBatchSuffix
                Case "[Batch Owner]"
                    mebIndexValues(intIndex).Text = gsecUserID
                Case Else
                    mebIndexValues(intIndex).Text = txtFieldDefaultValue(intIndex)
            End Select
            
        End If ' txtIndexValues(intIndex).Visible = True
        
    End If
        

End Sub

Private Sub subImportFileToDoc(txtBatchDirectory As String, txtBatchPageFileName As String)

        '****************************************************************************
        '*** IMPORT THE FILE TO THE DOCUMENT
        Dim strSourceFile As String
        
        '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
        MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True

        
        strSourceFile = txtBatchDirectory & "\" & txtBatchPageFileName
        
        SpicerDoc1.ImportPage 0, 0, IN_NEWPAGE_END, strSourceFile, strSourceFile

End Sub
                    
                    
Private Function funcExportTTCImageToJPG() As String


    Dim strSaveFileName As String
    Dim strSaveFileNameExtension As String
    Dim strOutputFilePath As String
    
    Dim lPageCount As Long
    Dim iTemp As Integer
    Dim sTemp As String
    Dim docSave As IDocSave
    
     '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True
    
        strSaveFileName = "I101TempImage"
        
        'Set up full path for export file
        txtAttachmentFileName = Environ("TEMP") & "\" & Trim(strSaveFileName) & ".JPG"

        subRasterizeBatchEX
      
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        
        Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
        
        'Save the Current Page to JPG
        docSave.Save MainMDIForm.ActiveForm.SpicerDoc1.pageID(1), False, API_FF_JFIF, txtAttachmentFileName, txtAttachmentFileName
        DoEvents
        
        
        ' De-initialize the object variable
        
        Set docSave = Nothing

    ' RETURN the FilePath of the Temporary Image File
    funcWriteToDebugLog Me.name, "txtAttachmentFileName= " & txtAttachmentFileName
    funcExportTTCImageToJPG = txtAttachmentFileName
    
    DoEvents
    

End Function

Public Sub subRasterizeBatchEX()

   Dim RasterBatch As IRasterBatch
   Dim lObjectID As Long
   Dim iMergeType As MERGE_TYPE
   Dim iXResolution As Integer
   Dim iYResolution As Integer
   Dim iColor As COLORTYPE
   Dim iBrightness As Integer
   Dim iThreshold As Integer
   Dim iOrientation As ORIENTATION_ANGLE
   Dim lXSize As Long
   Dim lYSize As Long
   Dim iUnit As UNIT_TYPE
   
     '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True
   
   'This flag is to disable the Form Load and Activate subs on the ChildForm
   bolRasterizingDocument = True
   
        ' Set the control to create a NEW Raster Document
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = MainMDIForm.ActiveForm.SpicerConfiguration1.object
        'Set to NOT automatically overwite document on rasterization.
        CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = True
'        'Set to Remove the original
        CFGDocument.RasterOperations(IN_RASTERIZE_REMOVE) = True
        'Deinitialize the object variable
        Set CFGDocument = Nothing


    MainMDIForm.ActiveForm.SpicerView1.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object
    MainMDIForm.ActiveForm.SpicerEdit1.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object
    
   
   ' Set the object variable for the IRasterBatch interface to the Edit Control object
   Set RasterBatch = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   ' Set the rasterize options
   lObjectID = MainMDIForm.ActiveForm.SpicerDoc1.RootID
   
'   iXResolution = 200   ' Keep the original resolution
'   iYResolution = 200
    ' GET the RESOLUTION from the Imaging101Client.INI
    On Error Resume Next
    iXResolution = VBGetPrivateProfileString("Imaging101Client", "frmIndex.subRasterizeBatchEX.Save_Resolution", RegFileName)
    If iXResolution = 0 Then
        iXResolution = 200
        result = WritePrivateProfileString(RegAppname, "frmIndex.subRasterizeBatchEX.Save_Resolution", iXResolution, RegFileName)
    End If
    iYResolution = iXResolution
    On Error GoTo 0
    
   iColor = IN_COLORTYPE_COLOR ' Rasterize to 256 color
'   iColor = IN_COLORTYPE_GRAYSCALE
   iBrightness = -50 ' Maximum darkness for Bilevel Dithered, Enhanced, and CAD
   iThreshold = 255  ' Maximum darkness for bilevel
   iOrientation = IN_ORIENTATION_NONE ' Use original resolution
   iUnit = IN_UNITS_INCH
   lXSize = 8.5 ' 0 = Keep the original size
   lYSize = 11  ' 0 = Keep the original size
   
   ' Rasterize the entire document
   RasterBatch.RasterizeBatchEx lObjectID, iXResolution, iYResolution, iColor, iBrightness, iThreshold, iOrientation, lXSize, lYSize, iUnit
   
   ' De-initialize the object variables
   Set RasterBatch = Nothing
   
   bolRasterizingDocument = False

End Sub

Private Sub txtIndexValues_Change(Index As Integer)

    If Not bolIndexFormLoadComplete Then
        Exit Sub
    End If
    
    If Len(txtIndexValues(Index).Text) > Int(txtFieldSize(Index).Text) Then
        MsgBox "Exceeded field size of " & txtFieldSize(Index).Text & " characters!" & _
                vbCrLf & "Truncating to defined size.", vbOKOnly
        txtIndexValues(Index).Text = Left(txtIndexValues(Index).Text, Int(txtFieldSize(Index).Text))
    End If
    
End Sub

Private Sub txtIndexValues_DblClick(Index As Integer)
    
'    'Re-size the Text field.
'    If txtIndexValues(Index).Height = mebIndexValues(Index).Height Then
'        txtIndexValues(Index).BackColor = vbYellow
'        txtIndexValues(Index).Refresh
'        txtIndexValues(Index).Height = txtIndexValues(Index).Height * 3
'        txtIndexValues(Index).SelStart = 0
'        txtIndexValues(Index).SelLength = 0
'    Else
'        txtIndexValues(Index).BackColor = vbWhite
'        txtIndexValues(Index).Height = mebIndexValues(Index).Height
'        txtIndexValues(Index).SelStart = 0
'        txtIndexValues(Index).SelLength = 0
'    End If
    
End Sub

Private Sub txtIndexValues_LostFocus(Index As Integer)

    'Check to Re-size the field
'    Call txtIndexValues_DblClick(Index)

        txtIndexValues(Index).BackColor = vbWhite
        txtIndexValues(Index).Height = mebIndexValues(Index).Height
        txtIndexValues(Index).SelStart = 0
        txtIndexValues(Index).SelLength = 0
     
End Sub

Private Function funcFtpUploadFile(ByVal strFTPSite As String, _
                                    ByVal intFTPport As Integer, _
                                    ByVal strFTPUserID As String, _
                                    ByVal strFTPPassword As String, _
                                    ByVal strFTPUploadSourceFileName As String) As Boolean

    
    funcWriteToDebugLog Me.name, "ENTERING: funcFtpUploadFile()"
    
    '*** 2020-05-21 - Jacob - Enable BatchMessageMode to allow errors to flow through
    MainMDIForm.ActiveForm.SpicerConfiguration1.BatchMessageMode = True

    '*** CHECK IF FTP UPLOAD
    If txtApplicationCommitBatchOption = "Application & FTP" Then
'    Or txtApplicationCommitBatchOption = "FTP Only" Then
    
        Dim strTempDirectory As String
        Dim strFTPUploadSourceFilePath As String
        Dim strFTPUploadDestinationFileName As String
        Dim bolFTPErrorOccured  As Boolean
        
        bolFTPErrorOccured = False
        
        strTempDirectory = funcGetTempDir()
        
        'EXPORT THE DOCUMENT TO THE TEMP DIRECTORY
        strFTPUploadSourceFilePath = funcExportSaveDocument(strTempDirectory, _
                                                                                  MainMDIForm.ActiveForm.SpicerDoc2, _
                                                                                  MainMDIForm.ActiveForm.SpicerView2, _
                                                                                  MainMDIForm.ActiveForm.SpicerEdit2, _
                                                                                  strFTPUploadSourceFileName, "PDF")
                                                                                  

        'CLOSE SpicerDoc2
        MainMDIForm.ActiveForm.SpicerDoc2.CloseDocument False
        DoEvents
        MainMDIForm.ActiveForm.SpicerDoc2.CloseDocument False
                            
    
    
        'BUILD DESTINATION FILE NAME
        strFTPUploadDestinationFileName = Right(strFTPUploadSourceFilePath, _
                                                                         Len(strFTPUploadSourceFilePath) - InStrRev(strFTPUploadSourceFilePath, "\"))
        
    
        'UPLOAD FILE
        txtActionBeforeError = "FTP Site= " & strFTPSite & "  Command= PUT " & strFTPUserID & ", " & strFTPPassword & ", " & strFTPUploadSourceFilePath & ", " & strFTPUploadDestinationFileName & ", " & False
        funcWriteToDebugLog Me.name, txtActionBeforeError
'        frmFTP.FTPFile strFTPSite, "PUT", strFTPUserID, strFTPPassword, strFTPUploadSourceFilePath, strFTPUploadDestinationFileName, False
        
        '5/17/2011 - Jacob - REPLACED frmFTP with funcFTPPutFile
        bolFTPErrorOccured = funcFTPPutFile(strFTPSite, intFTPport, strFTPUserID, strFTPPassword, strFTPUploadSourceFilePath, strFTPUploadDestinationFileName)
        
        'DELETE THE TEMP SOURCE FILE
        Kill strFTPUploadSourceFilePath

       'Return whether an error occured or not during transfer.
        funcFtpUploadFile = bolFTPErrorOccured

    End If

End Function

Private Function funcBuildDocTypeString() As String

    Dim strDocTypeString As String
    
    strDocTypeString = ""
    
    ' *** NOTE:  The Field Descriptions are case sensitive!
    For intIndex = 0 To lblFieldDescription.Count - 1
        Select Case Trim(lblFieldDescription.item(intIndex).Caption)
            Case Trim(strDOCGROUP)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If txtIndexValues(intIndex).Visible = True Then
                    strDocTypeString = strDocTypeString & txtIndexValues(intIndex).Text
                Else
                    strDocTypeString = strDocTypeString & mebIndexValues(intIndex).Text
                End If
                
            Case Trim(strDOCTYPE)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If txtIndexValues(intIndex).Visible = True Then
                    strDocTypeString = strDocTypeString & txtIndexValues(intIndex).Text
                Else
                    strDocTypeString = strDocTypeString & mebIndexValues(intIndex).Text
                End If
                
            Case Trim(strDOCSUBTYPE)
                '*** Check if the txtIndexValues TextBox control is VISIBLE...
                '    this will handle saving the value of the the TextBox control
                '    instead of the mebIndexValues Masked Edit control as needed.
                '    This is because the Masked Edit control has a MAX size of 64 Char.
                If txtIndexValues(intIndex).Visible = True Then
                    strDocTypeString = strDocTypeString & txtIndexValues(intIndex).Text
                Else
                    strDocTypeString = strDocTypeString & mebIndexValues(intIndex).Text
                End If
        End Select
    
    Next

    funcBuildDocTypeString = strDocTypeString
    
End Function

Private Sub subCommitTransactions()

                    '****************************************************************************
                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                    
                    On Error Resume Next
                    
                    If connImaging101Batch.State = adStateOpen Then
                        connImaging101Batch.CommitTrans
                    End If
                    
                    If connImaging101.State = adStateOpen Then
                        connImaging101.CommitTrans
                    End If
                    
                    If rsImaging101Batch.State = adStateOpen Then
                        rsImaging101Batch.Close
                    End If
                    
                    If rsImaging101BatchPage.State = adStateOpen Then
                        rsImaging101BatchPage.Close
                    End If
                    
                    If rsImaging101Document.State = adStateOpen Then
                        rsImaging101Document.Close
                    End If
                    
                    If rsImaging101DocumentDetail.State = adStateOpen Then
                        rsImaging101DocumentDetail.Close
                    End If
                    
                    Set rsImaging101Batch = Nothing
                    Set rsImaging101BatchPage = Nothing
                    Set rsImaging101Document = Nothing
                    Set rsImaging101DocumentDetail = Nothing

                    Set cmdImaging101 = Nothing
                    Set connImaging101 = Nothing
                    Set connImaging101Batch = Nothing

End Sub


Function CopyFileByChunk(sSource As String, sDestination As String, Optional ByVal ChunkSize As Long) As Long
       
        '*** 2021-02-16 - Jacob - Added CopyFileByChunk() Function
        '*** Copy Large File by Chunk with Progress Notification ***
        'When using FileCopy() to copy large file (particular over network), your project may look like it is "Not Responding" or hang.
        'This single function can be used to copy file by chunk with progress notification between chunk. It can be used anywhere for VB6/VBA projects.
        'If required, a cancel feature of the copy process can be built-in with a Public Boolean flag.
       
        txtActionBeforeError = "***  ENTERING Function CopyFileByChunk()   ***"
        funcWriteToDebugLog Me.name, txtActionBeforeError

       
       Dim FileSize As Long, OddSize As Long, SoFar As Long
       Dim buffer() As Byte, f1 As Integer, f2 As Integer
       Const MaxChunkSize As Long = 2 * 2 ^ 20 '-- 2MB
    
       On Error GoTo CopyFileByChunk_Error
       
        txtActionBeforeError = "'===> OPEN Source File for Binary Access Read "
        funcWriteToDebugLog Me.name, txtActionBeforeError
       f1 = FreeFile: Open sSource For Binary Access Read As #f1
       
       'Create a BLANK File, Existing file will be OVERWRITTEN
        txtActionBeforeError = "'===> CREATE Destination File "
        funcWriteToDebugLog Me.name, txtActionBeforeError
       f2 = FreeFile: Open sDestination For Output As #f2: Close #f2
       
       
        txtActionBeforeError = "'===> GET Source FileSize  "
        funcWriteToDebugLog Me.name, txtActionBeforeError
       FileSize = LOF(f1)
        txtActionBeforeError = "'===>  Source FileSize =  FileSize"
        funcWriteToDebugLog Me.name, txtActionBeforeError

       If FileSize = 0 Then GoTo Exit_CopyFileByChunk ' -- done!
       
        txtActionBeforeError = "'===> OPEN Destination for Binary Access WRITE "
        funcWriteToDebugLog Me.name, txtActionBeforeError
       f2 = FreeFile: Open sDestination For Binary Access Write As #f2
       
       'If NO ChunkSize is specified...use Approximately 1% of File Size
       If ChunkSize <= 0 Then ChunkSize = FileSize \ 100
       
       If ChunkSize = 0 Then
          OddSize = FileSize
       Else
            'If specified ChunkSize is Larger than the MaxChunkSize... then use MaxChunkSize
            If ChunkSize > MaxChunkSize Then
                ChunkSize = MaxChunkSize
            End If
            OddSize = FileSize Mod ChunkSize
       End If
       
       If OddSize Then
            txtActionBeforeError = "'===> OddSize = TRUE : " & OddSize
            funcWriteToDebugLog Me.name, txtActionBeforeError

            ReDim buffer(1 To OddSize)
            
            txtActionBeforeError = "      --->  Get OddSize Destination Chunk "
            funcWriteToDebugLog Me.name, txtActionBeforeError
            Get #f1, , buffer
          
            txtActionBeforeError = "      --->  Put OddSize Destination Chunk "
            funcWriteToDebugLog Me.name, txtActionBeforeError
            Put #f2, , buffer
            
              SoFar = OddSize
              '-- replace with your way of progess notification such as
              'Label1.Caption = SoFar & " bytes out of " & FileSize & " bytes: " & _
                           Format(SoFar / FileSize, "0.0%")
              'Debug.Print SoFar, Format(SoFar / FileSize, "0.0%")
              
                txtActionBeforeError = "      --->  " & SoFar & " bytes out of " & FileSize & " bytes: " & Format(SoFar / FileSize, "0.0%")
                funcWriteToDebugLog Me.name, txtActionBeforeError
              frmCommitStatus.ProgressBarCopyPageToArchive.Value = SoFar / FileSize * 100
    
              DoEvents '-- if required
       End If
       
       If ChunkSize Then
       
            txtActionBeforeError = "'===> ChunkSize = TRUE : " & ChunkSize
            funcWriteToDebugLog Me.name, txtActionBeforeError
            
            ReDim buffer(1 To ChunkSize)
            
            
            '***************************************************
            '*** LOOP top Copy Chunks
            
            Do While SoFar < FileSize
               
                    txtActionBeforeError = "      --->  Get ChunkSize Destination Chunk "
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    Get #f1, , buffer
                    
                    '* WRITE Chunk BEGIN
                    txtActionBeforeError = "      --->  OPEN ChunkSize Destination for Binary Access WRITE "
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    f2 = FreeFile: Open sDestination For Binary Access Write As #f2

                    txtActionBeforeError = "      --->  Find ChunkSize END of Destination File "
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    Seek #f2, LOF(f2) + 1
                    
                    txtActionBeforeError = "      --->  Put ChunkSize Destination Chunk "
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    Put #f2, , buffer
                    
                    'Check LOF (Length of File) = File Size
                    CopyFileByChunk = LOF(f2)
       
                    txtActionBeforeError = "      --->  CLOSE ChunkSize Destination File - To FLUSH Buffers"
                     funcWriteToDebugLog Me.name, txtActionBeforeError
                    
                    Close #f2
                    '* Write Chunk END
                    
                    SoFar = SoFar + ChunkSize
                    '-- replace with your way of progess notification such as
                    'Label1.Caption = SoFar & " bytes out of " & FileSize & " bytes: " & _
                                 Format(SoFar / FileSize, "0.0%")
                    'Debug.Print SoFar, Format(SoFar / FileSize, "0.0%")
                    txtActionBeforeError = "      --->  " & SoFar & " bytes out of " & FileSize & " bytes: " & Format(SoFar / FileSize, "0.0%")
                    funcWriteToDebugLog Me.name, txtActionBeforeError
                    frmCommitStatus.ProgressBarCopyPageToArchive.Value = SoFar / FileSize * 100
                    
                    DoEvents '-- if required
            Loop
            
       End If
       
       
       If CopyFileByChunk <> FileSize Then
          txtActionBeforeError = "<<<   ERROR:   CopyFileByChunk <> FileSize   >>>"
          funcWriteToDebugLog Me.name, txtActionBeforeError

          CopyFileByChunk = -CopyFileByChunk '-- negative denotes error
       End If
       
Exit_CopyFileByChunk:
        txtActionBeforeError = "'===> CLOSE Source File"
        funcWriteToDebugLog Me.name, txtActionBeforeError
        Close #f1
         txtActionBeforeError = "'===> CLOSE Destination File"
        funcWriteToDebugLog Me.name, txtActionBeforeError
       Close #f2
        txtActionBeforeError = "***  EXIT Function CopyFileByChunk()   ***"
        funcWriteToDebugLog Me.name, txtActionBeforeError

        Exit Function
       
CopyFileByChunk_Error:
            txtActionBeforeError = "<<<   ERROR:   Runtime error: " & Err.Number & " : " & Err.Description
            funcWriteToDebugLog Me.name, txtActionBeforeError

            MsgBox "Runtime error " & Err.Number & ":" & vbCrLf & Err.Description, vbCritical, "CopyFileByChunk"
            CopyFileByChunk = -1 '-- negative denotes error
            Resume Exit_CopyFileByChunk
End Function

