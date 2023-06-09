VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportFilesToBatch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Import Files To Batch"
   ClientHeight    =   8520
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11640
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11160
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameButtonBar 
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
      Height          =   612
      Left            =   0
      TabIndex        =   48
      Top             =   7920
      Width           =   11655
      Begin VB.CheckBox chkDeleteFilesAfterImport 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Source Files After Import?"
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
         Left            =   7320
         TabIndex        =   52
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Import Files To Batch"
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
         Height          =   612
         Left            =   9360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmImportToBatch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton cmdDeSelectALL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&De-Select ALL Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmImportToBatch.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton cmdSelectALL 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Select ALL Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   1995
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmImportToBatch.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   0
         Width           =   2000
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   11655
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
         Left            =   10080
         Picture         =   "frmImportToBatch.frx":1196
         ScaleHeight     =   375
         ScaleWidth      =   1455
         TabIndex        =   45
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Import Files to Batch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   9960
         TabIndex        =   46
         Top             =   480
         Width           =   1365
      End
   End
   Begin VB.TextBox txtUserID 
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
      Left            =   8520
      TabIndex        =   42
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdFileSourceDirectoryFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Find"
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
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmImportToBatch.frx":1829
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFileSourceDirectory 
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   840
      TabIndex        =   29
      Top             =   6840
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox txtImportingFile 
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   7608
      Width           =   7935
   End
   Begin VB.TextBox txtBatchDirectory 
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   27
      Top             =   7320
      Width           =   7935
   End
   Begin VB.TextBox txtItemCounter 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9840
      TabIndex        =   26
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtApplicationName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox txtBatchRootDirectory 
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
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2040
      TabIndex        =   16
      Top             =   3996
      Width           =   8532
   End
   Begin VB.CommandButton cmdBatchDirectoryFind 
      Caption         =   "&Find"
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
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmImportToBatch.frx":1DB3
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch Routing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2772
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   10935
      Begin VB.ComboBox cmbBatchOwner 
         Height          =   288
         ItemData        =   "frmImportToBatch.frx":233D
         Left            =   1920
         List            =   "frmImportToBatch.frx":233F
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2340
         Width           =   3135
      End
      Begin VB.ComboBox cmbBatchPriority 
         Height          =   288
         ItemData        =   "frmImportToBatch.frx":2341
         Left            =   6360
         List            =   "frmImportToBatch.frx":2343
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1728
         Width           =   2652
      End
      Begin VB.ComboBox cmbBatchStatus 
         Height          =   288
         ItemData        =   "frmImportToBatch.frx":2345
         Left            =   6360
         List            =   "frmImportToBatch.frx":2347
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2016
         Width           =   2652
      End
      Begin VB.ComboBox cmbBatchQueue 
         Height          =   288
         ItemData        =   "frmImportToBatch.frx":2349
         Left            =   1920
         List            =   "frmImportToBatch.frx":234B
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2052
         Width           =   3135
      End
      Begin VB.ComboBox cmbBatchGroup 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   324
         ItemData        =   "frmImportToBatch.frx":234D
         Left            =   1920
         List            =   "frmImportToBatch.frx":2357
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1728
         Width           =   3135
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
         Left            =   7920
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtBatchName 
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
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtBatchPrefix 
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
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtBatchSuffix 
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
         Height          =   285
         Left            =   6720
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtBatchNotes 
         Height          =   675
         Left            =   1920
         TabIndex        =   7
         Top             =   1056
         Width           =   5295
      End
      Begin VB.TextBox txtBatchDesc 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   768
         Width           =   5892
      End
      Begin VB.Label lblBatchOwner 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Route To User"
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
         Height          =   252
         Left            =   360
         TabIndex        =   41
         Top             =   2340
         Width           =   1572
      End
      Begin VB.Label lblBatchPriority 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch Priority"
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
         Height          =   252
         Left            =   5160
         TabIndex        =   40
         Top             =   1764
         Width           =   1212
      End
      Begin VB.Label lblBatchStatus 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   5160
         TabIndex        =   39
         Top             =   2052
         Width           =   1212
      End
      Begin VB.Label lblBatchQueue 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Route To Queue"
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
         Height          =   252
         Left            =   360
         TabIndex        =   38
         Top             =   2052
         Width           =   1572
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   360
         TabIndex        =   37
         Top             =   1728
         Width           =   1212
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch RECID"
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
         Index           =   0
         Left            =   7920
         TabIndex        =   20
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   6720
         TabIndex        =   19
         Top             =   240
         Width           =   732
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   360
         TabIndex        =   13
         Top             =   768
         Width           =   1452
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   360
         TabIndex        =   12
         Top             =   1056
         Width           =   1212
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
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
         Height          =   252
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.TextBox txtFilePattern 
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Text            =   "*"
      Top             =   4440
      Width           =   4095
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   5880
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   4800
      Width           =   5295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2232
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   5415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File Pattern"
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
      Left            =   5880
      TabIndex        =   54
      Top             =   4440
      Width           =   1092
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "File Location"
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
      Left            =   360
      TabIndex        =   53
      Top             =   4440
      Width           =   1212
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   7800
      TabIndex        =   43
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "File Source Directory"
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
      Height          =   732
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Importing File"
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
      Left            =   120
      TabIndex        =   25
      Top             =   7608
      Width           =   1692
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Import to Directory"
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
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   7320
      Width           =   1692
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
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
      Height          =   252
      Left            =   360
      TabIndex        =   22
      Top             =   840
      Width           =   1092
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
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   17
      Top             =   3996
      Width           =   2052
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Processed"
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
      Left            =   9840
      TabIndex        =   4
      Top             =   7320
      Width           =   1452
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   732
      Left            =   0
      Top             =   7200
      Width           =   11652
   End
End
Attribute VB_Name = "frmImportFilesToBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim txtFullPathName As String
    Dim txtDestinationFilename As String
    Dim txtFileDestination As String
    Dim SourceUNCFilePath As String
    Dim FileName As String
                    
    Dim UNCFilePath  As String
    
    Dim Fileroom As String
    Dim Filecabinet As String
    Dim DocumentDate As String
    Dim DateAdded As String
    Dim DocumentType  As String
    Dim Folder  As String
    Dim FolderDescription  As String
    Dim DocumentSubType   As String
    Dim DocumentExpireDate  As String
    Dim DocumentNote   As String
                    
    Dim BatchID As String
    Dim PageCount As Long
    Dim DocumentRECID As Double
    
    Dim arrFullDirSections() As String
    Dim txtFullDir As String
    Dim txtNewDir As String
    Dim flagMultiPageDocument As Boolean
    Dim flagDocumentInProcess As Boolean
    Dim arrInputLine() As String
    
    ' Set up for Imaging101Batch DB connection
    Dim connImaging101Batch As ADODB.Connection
    Dim cmdImaging101Batch As ADODB.Command
    
    Dim rsImaging101Batch As ADODB.Recordset
    Dim rsImaging101BatchPage As ADODB.Recordset
    
    Dim connLookupList As ADODB.Connection
    Dim rsLookupList As ADODB.Recordset
    
    Dim bolCancelPendingXfers As Boolean




    



Private Sub cmdBatchDirectoryFind_Click()

    txtBatchRootDirectory = funcGetDirectoryLocation("C:\IMAGING101\BATCHES\eCopyIMPORT")
    
End Sub

Private Sub cmdDeSelectALL_Click()

    Dim intIndex As Integer
    
    For intIndex = 0 To File1.ListCount - 1
        File1.Selected(intIndex) = False
    Next

End Sub

Private Sub cmdFileSourceDirectoryFind_Click()

'    txtFileSourceDirectory = funcGetDirectoryLocation(txtFileSourceDirectory)

    File1.Path = txtFileSourceDirectory.Text
    
End Sub

Public Sub cmdImport_Click()



    bolCancelPendingXfers = False

    Dim intIndex As Integer
    Dim txtInputFilePath As String
    
    txtItemCounter = 0
    
    If Len(txtBatchName) = 0 Then
        MsgBox "Please enter a BATCH Name!", vbOKOnly, "No Batch Name..."
        Exit Sub
    End If
    
    'Only Create the Batch and Directory if there are Files to be imported
    If File1.ListCount > 0 Then
        txtActionBeforeError = "FileCopy " & txtInputFilePath & ", " & txtOutputFilePath
        subCreateBatchRecord
        DoEvents
        
        txtActionBeforeError = "funcCreateDirectoryStructure " & txtBatchDirectory
        funcCreateDirectoryStructure txtBatchDirectory
        DoEvents
    Else
        MsgBox "Sorry... Nothing to Import!", vbOKOnly, "No files to Import..."
        Exit Sub
    End If
    
    For intIndex = 0 To File1.ListCount - 1
       
        'Process ONLY Selected Files
        If File1.Selected(intIndex) Then
        
            txtItemCounter = txtItemCounter + 1
            m_ImageCount = txtItemCounter
            txtInputFilePath = File1.Path + "\" + File1.List(intIndex)
            txtImportingFile = txtInputFilePath
            txtOutputFilePath = txtBatchDirectory + "\" + File1.List(intIndex)
            DoEvents
            
            txtActionBeforeError = "FileCopy " & txtInputFilePath & ", " & txtOutputFilePath
            FileCopy txtInputFilePath, txtOutputFilePath
            If bolCancelPendingXfers = True Then
                txtImportingFile = "***  IMPORT ERROR   ***"
                Exit Sub
            End If
            DoEvents
            
            txtActionBeforeError = "subCreateBatchPageRecord " & File1.List(intIndex)
            subCreateBatchPageRecord File1.List(intIndex)
            If bolCancelPendingXfers = True Then
                txtImportingFile = "***  IMPORT ERROR   ***"
                Exit Sub
            End If
            DoEvents
            
            'Delete file if selected
            '   AND the file was NOT Moved due to error or corruption
            If (chkDeleteFilesAfterImport = vbChecked) And (Not bolCancelPendingXfers) Then
                Kill txtInputFilePath
            End If
            
        End If
        
    Next
    
    'If Directory is EMPTY (No Subdirectories or Files), ask if user wants to Delete it
    If (chkDeleteFilesAfterImport = vbChecked) And (Not bolCancelPendingXfers) Then
        File1.Refresh
        If Dir1.ListCount = 0 And File1.ListCount = 0 Then
            result = MsgBox("The Directory is Empty... DELETE it?", vbYesNo)
            If result = vbYes Then
                'Delete the Directory
                RmDir Dir1.Path
                'Move Back to it's Parent directory
                Dir1.Path = Left(Dir1.Path, InStrRev(Dir1.Path, "\"))
                Dir1.Refresh
            End If
        End If
    End If
    
    
    ' 2014-04-25 - Jacob - UN-LOCK the Batch, ignore errors
    strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)


    txtImportingFile = "***  IMPORT COMPLETE   ***"
    MsgBox "***  IMPORT COMPLETE   ***", vbOKOnly, "Import Complete..."
        

'    cmdImport.enabled = False

Exit Sub

ERROR_HANDLER:
        funcQuickMessage "SHOW", "cmdImport_Click ERROR:" & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") "
        
        On Error Resume Next

        bolCancelPendingXfers = True

End Sub
        


Private Sub cmdSelectALL_Click()
    
    Dim intIndex As Integer
    
    For intIndex = 0 To File1.ListCount - 1
        File1.Selected(intIndex) = True
    Next

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Dir1_Change()
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    File1.Pattern = frmImportFilesToBatch.txtFilePattern
    File1.Refresh
    
    'Select ALL Files
    cmdSelectALL_Click
    
    'Set the Batch Name to the SubDirectory name the user is Importing from
    txtBatchName.Text = Right(Dir1.Path, Len(Dir1.Path) - InStrRev(Dir1.Path, "\"))

End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    File1.Pattern = frmImportFilesToBatch.txtFilePattern
    File1.Refresh

End Sub

Private Sub Form_Activate()

    On Error Resume Next
    
    If Trim(txtBatchRootDirectory) = "" Then
        MsgBox "          * * *   WARNING   * * * " & vbCrLf & _
                "No Batch Directory has been Configured for this Application." & vbCrLf & _
                "Please notify the system administrator to correct this." _
                , vbExclamation, "Batch Directory NOT Configured"
        Unload Me
    End If
    
    'Set Focus and Highlight contents
'    txtBatchName.SetFocus
'    txtBatchName.SelStart = 0
'    txtBatchName.SelLength = Len(txtBatchName)
    

End Sub

Private Sub Form_Load()
    
    lblVersion.Caption = " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    ' Get saved settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.Left", RegFileName)
'    Me.width = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.Width", RegFileName)
'    Me.Height = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.Height", RegFileName)
    

    
    txtApplicationRECID = frmImaging101BatchList.txtApplicationRECID
    txtApplicationName = frmImaging101BatchList.cmbApplicationList.Text
    txtBatchRootDirectory = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & txtApplicationRECID, "RootDirectoryPathForBatches")
    
    
    Drive1.Drive = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchImportDrive", RegFileName)
    Dir1.Path = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchImportFilePath", RegFileName)
    frmImportFilesToBatch.txtFilePattern = VBGetPrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchImportFilePattern", RegFileName)
    txtUserID = gsecUserID
    
'    cmdImport.enabled = False
    
    
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

        
    ' Establish the Imaging101 DB Connections
    txtActionBeforeError = "Prepare Imaging101 DB Connections"
    Set connImaging101 = New ADODB.Connection
    Set cmdImaging101 = New ADODB.Command
    Set rsImaging101 = New ADODB.Recordset
    
    
    connImaging101.ConnectionString = RegImaging101ConnectionString
    connImaging101.ConnectionTimeout = 120
    connImaging101.mode = adModeReadWrite
    connImaging101.Open
    Set cmdImaging101.ActiveConnection = connImaging101
    
    ' Establish the Imaging101Batch DB Connections
    txtActionBeforeError = "Prepare Imaging101Batch DB Connections"
    Set connImaging101Batch = New ADODB.Connection
    Set cmdImaging101Batch = New ADODB.Command
    Set rsImaging101Batch = New ADODB.Recordset
    
    
    connImaging101Batch.ConnectionString = RegImaging101BatchListConnectionString
    connImaging101Batch.ConnectionTimeout = 120
    connImaging101Batch.mode = adModeReadWrite
    connImaging101Batch.Open
    Set cmdImaging101Batch.ActiveConnection = connImaging101Batch
    
    
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    File1.Pattern = frmImportFilesToBatch.txtFilePattern
    File1.Refresh
    
    
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
'    '*** LOAD GROUPS LIST DROP-DOWN
'
'    cmbGroupList.AddItem ""
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
'    txtActionBeforeError = "Populate Groups List"
'
'    rs.MoveFirst
'
'    For intIndex = 0 To rs.RecordCount - 1
'        cmbGroupList.AddItem rs.Fields!GroupName
'        cmbGroupList.ItemData(intIndex) = rs.Fields!GroupRECID
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

'    If cmbBatchQueue.ListCount > 0 Then
'        cmbBatchQueue.ListIndex = cmbBatchQueue.TopIndex
'    End If
'
    funcFindItemInComboBox cmbBatchQueue, frmImaging101BatchList.cmbBatchQueue.Text

    '****************************

    '***************************************
    '*** LOAD BATCH STATUS LIST DROP-DOWN
        
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchStatus ORDER BY BatchStatus"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    On Error Resume Next
    
    rs.MoveFirst
    
    On Error GoTo 0
    
    'Add a Blank value to allow clearing the BatchOwner
'    cmbBatchStatus.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchStatus.AddItem rs.Fields!BatchStatus
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
    
    '***************************************
    '*** LOAD BATCH PRIORITY LIST DROP-DOWN
        
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchPriority ORDER BY BatchPriority"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    On Error Resume Next
    
    rs.MoveFirst
    
    On Error GoTo 0
    
    'Add a Blank value to allow clearing the BatchOwner
'    cmbBatchPriority.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchPriority.AddItem rs.Fields!BatchPriority
        rs.MoveNext
    Next
        
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


    
End Sub

Private Sub Text3_Change()

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save Form settings to the registry
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.Left", Me.Left, RegFileName)
'        result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.Width", Me.width, RegFileName)
'        result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.Height", Me.Height, RegFileName)
    End If
    
    result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchRootDirectory", txtBatchRootDirectory, RegFileName)
    
    result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchImportDrive", Drive1.Drive, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchImportFilePath", Dir1.Path, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmImportFilesToBatch.BatchImportFilePattern", frmImportFilesToBatch.txtFilePattern, RegFileName)

End Sub

Private Sub Form_Resize()


'    Me.Height = frameButtonBar.Top + frameButtonBar.Height + 500
'
'    Me.width = frameButtonBar.width + 200
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmImaging101BatchList.Show
    frmImaging101BatchList.subListBatches
    
End Sub

Private Sub txtFilePattern_Change()

    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    File1.Pattern = frmImportFilesToBatch.txtFilePattern
    File1.Refresh
    
End Sub

Private Sub subCreateBatchRecord()
    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
''        On Error GoTo CREATE_BATCH_RECORD_ERROR
        
        'User Transaction Tracking to prevent partial imports!
        connImaging101Batch.BeginTrans
        
        txtActionBeforeError = "Open Batches Table"
        Set rsImaging101Batch = New ADODB.Recordset
        rsImaging101Batch.Open "I101Batches", connImaging101Batch, adOpenDynamic, adLockOptimistic
        
        txtActionBeforeError = "Add New Record"
        rsImaging101Batch.AddNew
        
        txtActionBeforeError = "Assign Variables to Fields"
        
        txtBatchRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchRECID")
    
        txtBatchDirectory = txtBatchRootDirectory + "\" + Format(txtBatchRECID, "0000000000")
    
        
       
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
        
''        txtBatchDirectory = Trim(txtBatchRootDirectory) & "\" & Trim(txtBatchPrefix) & Trim(txtBatchName) & Trim(txtBatchSuffix)
       
'''''''        ' CREATE the Directory Structure for storing this Batch
'''''''        funcCreateDirectoryStructure txtBatchDirectory
        
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
        rsImaging101Batch("BatchScanUser") = txtUserID
        rsImaging101Batch("BatchBoxNumber") = ""
        rsImaging101Batch("BatchInQueueDate") = Now()
        
        '2014-04-25 - Jacob - Added "S" lock to prevent listing in Batch List if Scanning or Splitting
        rsImaging101Batch.Fields("BatchLocked") = "S"
        rsImaging101Batch.Fields("BatchLockedBy") = gsecUserName
        rsImaging101Batch.Fields("BatchLockedDate") = Now()


        txtActionBeforeError = "Update Values"
        rsImaging101Batch.Update
    
    ' Commit the successful transactions
    txtActionBeforeError = "COMMIT BATCH TRANSACTION"
    connImaging101Batch.CommitTrans
    
    
    '*** CREATE BATCH AUDIT RECORD
    funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Import Files - Manual"
    
    
    Screen.MousePointer = vbDefault

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
Private Sub subCreateBatchPageRecord(strFilename As String)
    
        
        On Error GoTo CREATE_BATCH_PAGE_RECORD_ERROR
        
        
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
        rsImaging101BatchPage("BatchPageFileName") = strFilename
        rsImaging101BatchPage("BatchPageOrder") = txtItemCounter
        
        rsImaging101BatchPage("BatchPageIndexed") = ""
        rsImaging101BatchPage("BatchPageIsSeparator") = ""
        rsImaging101BatchPage("BatchPageNote") = ""
        rsImaging101BatchPage("BatchDocDesc") = ""
        rsImaging101BatchPage("BatchPageStatus") = ""
        
        '*** eCopy FILE CONVERSION SECTION
        '***   Auto-Index if file has a ".CPY" extension.
        If UCase(Right(Trim(strFilename), 4)) = ".CPY" Then
            ' Take the Client # from the left part of the Filename
            rsImaging101BatchPage("CLIENTNO") = Left(strFilename, InStr(1, strFilename, ".") - 1)
            rsImaging101BatchPage("DocGroup") = "eCopy AS OF " & Format(Now, "mm-dd-yyyy")
            rsImaging101BatchPage("DocType") = "Client Documentation"
            rsImaging101BatchPage("DocDate") = Format(Now, "mm-dd-yyyy")
            
            '***************************************
            '*** LOAD BATCH PRIORITY LIST DROP-DOWN
                
            Set connLookupList = New ADODB.Connection
            connLookupList.Open RegLookupListConnectionString
            
            Set rsLookupList = New ADODB.Recordset
            Set rsLookupList.ActiveConnection = connLookupList
            
            rsLookupList.Source = "Select * from CLIENT Where CLIENTID = " & rsImaging101BatchPage("CLIENTNO")
            rsLookupList.CursorLocation = adUseClient
            rsLookupList.CursorType = adOpenDynamic
            rsLookupList.LOCKTYPE = adLockReadOnly
            
            connLookupList.Errors.Clear
            
            rsLookupList.Open
            
            If Not rsLookupList.EOF Then
                rsLookupList.MoveFirst
                rsImaging101BatchPage("CLIENTNAME") = rsLookupList("L_NAME") & ", " + rsLookupList("F_NAME")
            Else
                rsImaging101BatchPage("CLIENTNAME") = "*not found*"
            End If
            
            On Error GoTo 0
            
                
            'Close connection and the recordset
            rsLookupList.Close
            Set rsLookupList = Nothing
            connLookupList.Close
            Set connLookupList = Nothing
        
            '****************************
            
        End If
        
        
        txtActionBeforeError = "Update Batch Page Values"
        rsImaging101BatchPage.Update
        
        ' Set BATCHES field values
        txtActionBeforeError = "Assign Variables to Batch Fields"
        rsImaging101Batch("BatchPagesTotal") = rsImaging101Batch("BatchPagesTotal") + 1
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



Private Sub txtFileSourceDirectory_Change()
    File1.Path = txtFileSourceDirectory
    File1.Pattern = frmImportFilesToBatch.txtFilePattern
    File1.Refresh

End Sub
