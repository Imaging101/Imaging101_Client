VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmImaging101Retrieve 
   BackColor       =   &H80000013&
   Caption         =   "Retrieval Document List - Imaging101"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImaging101Retrieve.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmImaging101Retrieve.frx":0442
   ScaleHeight     =   8175
   ScaleWidth      =   12480
   Begin VB.TextBox txtDetailFileType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   43
      Text            =   "txtDetailFileType"
      Top             =   4530
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtFullPathName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   42
      TabStop         =   0   'False
      Text            =   "txtFullPathName"
      Top             =   4815
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtPathSubdirectory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "txtPathSubdirectory"
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "txtFileName"
      Top             =   4245
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtFilterStatement 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "txtFilterStatement"
      Top             =   5100
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtApplicationName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "txtApplicationName"
      Top             =   5385
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton cmdFastFix 
      Caption         =   "Fast Fi&x"
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
      Left            =   10560
      TabIndex        =   36
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstExportedDocuments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   34
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtExportToPDF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   31
      Text            =   "ExportToPDF?"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDocumentGroup 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   30
      Text            =   "txtDocumentGroup"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtDocGroupHead 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "frmImaging101Retrieve.frx":07CC
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDocGroupBody 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "frmImaging101Retrieve.frx":07DE
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDocGroupFoot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9360
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "frmImaging101Retrieve.frx":07F0
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtConcat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "frmImaging101Retrieve.frx":0802
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFoot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmImaging101Retrieve.frx":080E
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "frmImaging101Retrieve.frx":0818
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtHead 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "frmImaging101Retrieve.frx":0822
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDetailRecordCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   22
      Text            =   "txtDetailRecordCount"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtPageCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   21
      Text            =   "txtPageCount"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtDocumentRECID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "txtDocumentRECID"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   13
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdExportToExcel 
         BackColor       =   &H00FFFFFF&
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
         Height          =   735
         Left            =   4410
         Picture         =   "frmImaging101Retrieve.frx":082C
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Export the List of the selected documents to a "".CSV"" / EXCEL file."
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdOpenSelected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Picture         =   "frmImaging101Retrieve.frx":10F6
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Open Selected Documents"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdSelectAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Select All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   630
         Picture         =   "frmImaging101Retrieve.frx":1860
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Select ALL Documents in this List"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdDeSelectAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&DeSel All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1260
         Picture         =   "frmImaging101Retrieve.frx":1BA2
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "De-Select ALL Documents in this List"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
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
         Height          =   735
         Left            =   1890
         Picture         =   "frmImaging101Retrieve.frx":1CEC
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Modify Index Values for Selected Documents"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
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
         Height          =   735
         Left            =   2520
         Picture         =   "frmImaging101Retrieve.frx":2456
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Print Selected Documents"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Send"
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
         Height          =   735
         Left            =   3150
         Picture         =   "frmImaging101Retrieve.frx":2BC0
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Send Selected Documents via eMail"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdExportSelected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Export"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3780
         Picture         =   "frmImaging101Retrieve.frx":2F4A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Export Selected Documents"
         Top             =   15
         Width           =   630
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D&elete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5040
         Picture         =   "frmImaging101Retrieve.frx":3814
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Flag Selected Document as ""Deleted"""
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdRestore 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Restore"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6300
         Picture         =   "frmImaging101Retrieve.frx":3C56
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Restore Deleted documents"
         Top             =   0
         Width           =   630
      End
      Begin VB.CommandButton cmdMove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mo&ve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5670
         Picture         =   "frmImaging101Retrieve.frx":3FE0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Move selected documents to another Application"
         Top             =   0
         Width           =   630
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
         Left            =   9720
         Picture         =   "frmImaging101Retrieve.frx":4422
         ScaleHeight     =   375
         ScaleWidth      =   1575
         TabIndex        =   33
         Top             =   105
         Width           =   1572
      End
      Begin VB.CheckBox chkViewDeletedDocuments 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Deleted"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8400
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkViewDocDetails 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View RecID"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7080
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Items Selected"
         ForeColor       =   &H00C07000&
         Height          =   285
         Left            =   8400
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblItemsSelected 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblItemsSelected"
         ForeColor       =   &H00C07000&
         Height          =   285
         Left            =   7080
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblItemsFound 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblItemsFound"
         ForeColor       =   &H00C07000&
         Height          =   285
         Left            =   7080
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Items Found"
         ForeColor       =   &H00C07000&
         Height          =   285
         Left            =   8400
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   9720
         TabIndex        =   14
         Top             =   480
         Width           =   1245
      End
   End
   Begin MSAdodcLib.Adodc ADOdc1 
      Height          =   375
      Left            =   240
      Top             =   6480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   60
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
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
      Caption         =   "ADOdc1"
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
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   6030
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   6030
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1213
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   59
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4675
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1213
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   59
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
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
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   6030
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6090
      Visible         =   0   'False
      Width           =   6030
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4545
         Picture         =   "frmImaging101Retrieve.frx":4AB5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4200
         Picture         =   "frmImaging101Retrieve.frx":4DF7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   345
         Picture         =   "frmImaging101Retrieve.frx":5139
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         Picture         =   "frmImaging101Retrieve.frx":547B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5318
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin MSAdodcLib.Adodc ADOdcDetail 
      Height          =   375
      Left            =   3720
      Top             =   6480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   60
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
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
      Caption         =   "ADOdcDetail"
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
   Begin MSAdodcLib.Adodc ADOdcDestination 
      Height          =   375
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   3
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   60
      CursorType      =   2
      LockType        =   2
      CommandType     =   8
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
      Caption         =   "ADOdcDestination"
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
   Begin MSAdodcLib.Adodc ADOdcDetailDestination 
      Height          =   375
      Left            =   3720
      Top             =   6960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   3
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   60
      CursorType      =   2
      LockType        =   2
      CommandType     =   8
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
      Caption         =   "ADOdcDetailDestination"
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
   Begin MSAdodcLib.Adodc ADOdcWork 
      Height          =   375
      Left            =   240
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   60
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
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
      Caption         =   "ADOdcWork"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   5205
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImaging101Retrieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim dblNumberOfPagesAfterImport As Double
Dim strCommandSource As String

Dim intNumberPages

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Dim arrTempFilesList() As String
Dim intTempFilesCount




Private Sub chkViewDeletedDocuments_Click()

    'Re-issue the Search
    frmImaging101Search.cmdFind_Click
    subCheckButtonSecurity

End Sub


Public Sub chkViewDocDetails_Click()
    
    If chkViewDocDetails.Value = vbChecked Then
        ListView1.ColumnHeaders(1).width = 1500
'        ListView1.ColumnHeaders(2).width = 1000
'        ListView1.ColumnHeaders(3).width = 1000
'        ListView1.ColumnHeaders(4).width = 1000
'        ListView1.ColumnHeaders(5).width = 1000
'        ListView1.ColumnHeaders(6).width = 1000
'        ListView1.ColumnHeaders(7).width = 1000
    Else
        ListView1.ColumnHeaders(1).width = 0
'        ListView1.ColumnHeaders(2).width = 0
'        ListView1.ColumnHeaders(3).width = 0
'        ListView1.ColumnHeaders(4).width = 0
'        ListView1.ColumnHeaders(5).width = 0
'        ListView1.ColumnHeaders(6).width = 0
'        ListView1.ColumnHeaders(7).width = 0
    
    End If

End Sub






Private Sub cmdConfigOrder_Click()
    frmConfig.Show

End Sub


Private Sub cmdResetFields_Click()
    cmbCustomerNumber = ""
    cmbCustomerName = ""
    mebDateFrom = ""
    mebDateThru = ""
    txtPartNumber = ""
    txtInvoiceNumber = ""
    cmbDocType = ""
    cmbDocGroup = ""
    
End Sub

Private Sub CmdScan_Click()
''    Me.Hide
''    frmScan.Show
    frmImport.Show
End Sub

Private Sub DataCombo1_Click(Area As Integer)
'    AdodcFileRoom.RecordSource = "SELECT DISTINCT Fileroom FROM I101Documents"
'    AdodcFileRoom.Refresh

'    DataCombo1.DataSource = AdodcFileRoom
'    DataCombo1.DataMember = "Documents"
'    DataCombo1.DataField = "Fileroom"
'    DataCombo1.RowSource = AdodcFileRoom
'    DataCombo1.ListField = "Fileroom"
'    DataCombo1.BoundColumn = "Fileroom"
'    DataCombo1.Refresh
End Sub

Private Sub cmdDeSelectALL_Click()

    funcWriteToDebugLog Me.name, "ENTER: cmdDeSelectALL_Click"
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    cmdOpenSelected.Enabled = False
    cmdModify.Enabled = False
    cmdPrint.Enabled = False
    cmdSend.Enabled = False
    
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
            funcWriteToDebugLog Me.name, "i=" & i & " - " & frmImaging101Retrieve.ListView1.ListItems(i).Text
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = False   ' Force item selection
    Next
    
    subShowItemsSelected
    
    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub

Private Sub cmdExportToExcel_Click()


    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    On Error GoTo ERROR_HANDLER
    
    Dim strLineOut As String
    Dim strTempDir As String
    Dim strFilePath As String
    
    strTempDir = funcGetTempDir()

    strFilePath = strTempDir & App.EXEName & "_EXPORT.CSV"   ' & Format(Now(), "yyyy-mm-dd_hhmmss") & ".CSV"
    
    txtActionBeforeError = "Open " & strFilePath & " For output As #99"
    If funcFileExists(strFilePath) Then
        Kill strFilePath
    End If
    Open strFilePath For Append As #99
    

    'Prepare File Header from Column Headers
    strLineOut = ""
    For intColumn = 6 To ListView1.ColumnHeaders.Count
        strLineOut = strLineOut & Chr(34) & ListView1.ColumnHeaders(intColumn).Text & Chr(34) & ","
    Next
    
    Print #99, strLineOut

    For intRow = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
            
        If frmImaging101Retrieve.ListView1.ListItems(intRow).Selected = True Then
            
'            funcWriteToDebugLog Me.name, frmImaging101Retrieve.ListView1.ListItems(intRow).Text
            frmImaging101Retrieve.ListView1.ListItems(intRow).Selected = True   ' Force item selection
            
            strLineOut = ""
            For intColumn = 5 To ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems.Count
            
                strLineOut = strLineOut & Chr(34) & Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(intColumn)) & Chr(34) & ","

            Next
            
            Print #99, strLineOut
          
        End If
          
    Next
    
    Close #99
    
    Call shelldoc(strFilePath)
    
    subCheckButtonSecurity
    
    subShowItemsSelected
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
Exit Sub

ERROR_HANDLER:

    funcQuickMessage "SHOW", "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "The temporary file I am trying to create named" & vbCrLf & _
                                strFilePath & vbCrLf & _
                                "may be Open or In-Use." & vbCrLf & _
                                "Please either close it... or save it with a different name and try again."
                                
    
End Sub

Private Sub cmdFastFix_Click()


    frmImaging101Search.Hide
    frmImaging101Retrieve.Hide
    
    
    frmImaging101Modify.Show
    frmImaging101Modify.subModifyRecords
    
    

End Sub

Private Sub cmdModify_Click()

    'Do the Right-Click action
    ListView1_MouseUp vbRightButton, 0, 0, 0

End Sub



Private Sub cmdMove_Click()

    frmImaging101Search.Hide
    frmImaging101Retrieve.Hide
    
'    bolWaitForMoveCommand = True

    'Allow User to Select Destination Application & Map Fields
    frmImaging101MoveDocumentsBetweenApplications.Show
    
End Sub

Public Sub subMoveDocsBetweenApplications()
    
'    While bolWaitForMoveCommand
''        lblFieldDescription(intIndex).Caption = rs("FieldName")
'        DoEvents
'    Wend

    Me.Enabled = False
    
    'Close the Viewer to make SURE no documents are locked / in-use
    If funcIsFormLoaded2("frmMainMDIForm") Then
        Unload MainMDIForm
    End If
    
    frmImaging101MoveDocumentsBetweenApplications.ProgressBar1.Value = 0
    
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
    
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
            funcWriteToDebugLog Me.name, frmImaging101Retrieve.ListView1.ListItems(i).Text
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
            
           
            ReleaseCapture
            
            
            'Begin Transaction
            
            'Copy Application DETAIL Records
            txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(i).Text
            subCopyDetailRecordAndFiles txtDocumentRECID
            
            'Copy Application Record
            subCopyDocumentRecord txtDocumentRECID
            
            
            'Flag SOURCE Document as MOVED OUT (MO)
            funcSaveFieldToDB RegImaging101ConnectionString, txtApplicationName, "DocumentRECID = " & txtDocumentRECID, "DocumentLocked", "MO"
            
            
            'Create Audit Log entry
            
            
'            'End Transaction
            
            
            
        End If
    Next
    
    
    MsgBox "Move Complete.", vbOKOnly, "Document move Complete"
    
    Me.Enabled = True
    Unload Me
    

End Sub

Private Sub cmdOpenSelected_Click()

    Screen.MousePointer = MousePointerConstants.vbArrowHourglass

    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
            funcWriteToDebugLog Me.name, frmImaging101Retrieve.ListView1.ListItems(i).Text
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
            ListView1_DblClick
            
        End If
    Next
    
    If funcCountListViewItemsSelected(ListView1) > 1 Then
        MainMDIForm.mnuWTileHorizontal_Click
    End If
    
    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub
Private Sub cmdExportSelected_Click()

   
    '*** MDVIP TWEAK - BEGIN
    If frmImaging101Search.txtApplicationName = "MDVIP" Then
        'Show Export Options form
         frmImaging101ExportOptions.Show
         frmImaging101ExportOptions.SetFocus
    Else
        'NOT for MDVIP
        Load frmImaging101ExportOptions
'        frmImaging101ExportOptions.chkExportToPDF.Value = vbUnchecked
        frmImaging101ExportOptions.chkExportToPDF.Value = vbChecked
        subExportSelected_Run
    End If
        
End Sub


Public Sub subExportSelected_Run()

    
    On Error GoTo ERROR_HANDLER
    
    Dim strDirectory As String
    
    
    '*** MDVIP TWEAK - BEGIN
    If frmImaging101Search.txtApplicationName = "MDVIP" Then
        
        Call subMDVIP
        Exit Sub
        
    End If
    '*** MDVIP TWEAK - END
    
    
    
    


    strCommandSource = "cmdExportSelected"
    
    
    'GET Export Options
'    frmImaging101ExportSimple.Show vbModal, Me
    
    '*** GET DIRECTORY to Export to
    strDirectory = funcBrowseForDirectory
'    strDirectory = frmImaging101ExportOptions.txtFullPathForPDFexport.Text
    
    
    If strDirectory = "" Then
        MsgBox "No Directory Selected..." & vbCrLf & "Documents will NOT be Exported!", vbExclamation, "No Directory Selected"
        Screen.MousePointer = MousePointerConstants.vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
   
'    If strDirectory = "D:\" Then
'        strDirectory = "C:\Documents and Settings\" & gsecUserID & "\Local Settings\Application Data\Microsoft\CD Burning\"
'    End If
    
    ' CLEAR the Exported Document Listbox
    lstExportedDocuments.Clear
    
    
    
    '*** Scan for Documents Selected to Export
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
        
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
        
            Load frmClientSpicerControlForm
        
            'Get the DocumentRECID
            txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(i).Text
            funcWriteToDebugLog Me.name, txtDocumentRECID
            
            'Get the # of Images
            txtPageCount = funcGetFieldFromDB(RegImaging101ConnectionString, txtApplicationName, "DocumentRECID=" & txtDocumentRECID, "DocumentImages")
            
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
            
            If txtPageCount = "1" Then
                Dim strDetailFileType As String
                strDetailFileType = funcGetFieldFromDB(RegImaging101ConnectionString, txtApplicationName & "_Detail", "DocumentRECID=" & txtDocumentRECID, "DetailFileType")
            End If
            
            'Import all document pages
            subExportSelectedGetImages
            DoEvents
            
            'Save combined document
            strSaveFilePath = funcExportSelectedSaveDocument(strDirectory, frmClientSpicerControlForm.SpicerDoc1, "", "PDF")
            'Add the saved document to the Listbox
            lstExportedDocuments.AddItem strSaveFilePath
            DoEvents

            On Error GoTo ERROR_HANDLER
            DoEvents
            
            For w = 1 To 10000
                DoEvents
            Next
           
           Unload frmClientSpicerControlForm
           Set frmClientSpicerControlForm = Nothing
           
        End If
           
    Next

    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    MsgBox "DOCUMENT EXPORT COMPLETE!", vbOKOnly, "Document Export Complete"

Exit Sub

ERROR_HANDLER:

    Screen.MousePointer = MousePointerConstants.vbDefault

    funcQuickMessage "SHOW", "subExportSelected_Run - ERROR EXPORTING DOCUMENT(s):  " & Err.Number & " - " & Err.Description
    
    

End Sub
Public Sub subExportSelectedGetImages()

    On Error Resume Next
    
    ReDim arrTempFilesList(0) As String
    intTempFilesCount = 0
    
'''    Dim txtPathSubdirectory As String
'''    Dim txtFileName As String
'''    Dim txtFullPathName As String

    
   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                        " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                        " DetailFileName, DetailFileType, DetailRotation  " & _
                        " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                        " WHERE  DocumentRECID = " & txtDocumentRECID & _
                        " ORDER BY DetailOrder "

    txtActionBeforeError = "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    
    If frmImaging101Retrieve.ADOdcDetail.Recordset.EOF = True Then
        Exit Sub
    Else
        txtActionBeforeError = "ADodcDetail MoveFirst "
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
    End If
   
    
    While Not frmImaging101Retrieve.ADOdcDetail.Recordset.EOF
    
        txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
        txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
        txtDetailFileType = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileType")
        txtFullPathName = txtPathSubdirectory & "\" & txtFileName
        txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")

''        If Not frmImaging101Retrieve.AdodcDetail.Recordset.Fields("DetailRotation") Then
''            txtPageRotation = frmImaging101Retrieve.AdodcDetail.Recordset.Fields("DetailRotation")
''        End If
        
''''''        frmLoadingImages.SetFocus
''''''        frmLoadingImages.txtImageNumber = txtImageNumber + 1
'        funcWriteToDebugLog Me.Name, txtFullPathName & " - " & txtImageNumber & " - " & txtDetailRECID & " - " & txtPageRotation
''        DoEvents
        
'        Dim dblNumberOfPagesBeforeImport As Double
'        Dim dblNumberOfPagesAfterImport As Double
'
'
'        dblNumberOfPagesBeforeImport = frmClientSpicerControlForm.SpicerDoc1.NumberOfPages
        
        '*** If there is only ONE Record, DON'T Import it!
        '***   we will simply Copy it out to make it faster.
        txtDetailRecordCount = frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount
''''        If txtDetailRecordCount = "1" _
''''            And strCommandSource = "cmdExportSelected" Then
''''                Exit Sub
''''        End If
         
        Dim docContents As IDocContents
        Set docContents = frmClientSpicerControlForm.SpicerDoc1.object
         
'        If True Then
'            'Copy the Stored Document to the Local TEMP Directory
'            'FileCopy txtFullPathName, txtLaunchFileName
'            Dim strLocalFilePathName As String
'            strLocalFilePathName = funcGetTempName()
'            funcWriteToDebugLog Me.name, "chkCopyFileToLocalTempDir = vbChecked"
'            funcWriteToDebugLog Me.name, ".CopyFile " & txtFullPathName & ", " & strLocalFilePathName
'            With New FileSystemObject
'                .CopyFile txtFullPathName, strLocalFilePathName, True
'            End With
'            'Re-map the File and Path to the LOCAL file
'            txtFullPathName = strLocalFilePathName
'            txtPageFileName = txtFullPathName
'        End If
'
        
        '*** 2020-05-18 - Jacob - Modified to Save to the LocalTemp\Imaging101 subdirectory
        chkCopyFileToLocalTempDir = 1
        If chkCopyFileToLocalTempDir = vbChecked Then
            'Copy the Stored Document to the Local TEMP Directory
            Dim strLocalFilePathName As String
            strLocalTempDir = funcGetTempDir()
            strLocalTempDir = strLocalTempDir & "Imaging101\"
            
            'Create the directory if needed.
            funcCreateDirectoryStructure strLocalTempDir & ""
                
'            strLocalFilePathName = funcGetTempName()
            strLocalFilePathName = strLocalTempDir & txtDetailRECID & "_" & txtFileName
            funcWriteToDebugLog Me.name, "chkCopyFileToLocalTempDir = vbChecked"
            funcWriteToDebugLog Me.name, ".CopyFile " & txtFullPathName & ", " & strLocalFilePathName
            
            With New FileSystemObject
                If Not .FileExists(strLocalFilePathName) Then
                    .CopyFile txtFullPathName, strLocalFilePathName, True
                End If
            End With
            
            'Re-map the File and Path to the LOCAL file
            txtFullPathName = strLocalFilePathName
            txtPageFileName = txtFullPathName
        End If
        
        
        
        
         
         '*** IMPORT the File Contents
        docContents.ImportPage 0, 0, IN_NEWPAGE_END, txtImageNumber, txtFullPathName
        txtPageCount = docContents.NumberOfPages
        
        Set docContents = Nothing
        
        '*** If Combine Into Single PDF selected then IMPORT the File Contents to the 2nd Doc Control
        
        '*** MDVIP TWEAK - BEGIN
        If frmImaging101Search.txtApplicationName = "MDVIP" Then
            If frmImaging101ExportOptions.optCombineIntoSinglePDF = True Then
                 '*** IMPORT the File Contents into 2nd Doc Control
                Set docContents = frmClientSpicerControlForm.SpicerDoc2.object
         
                docContents.ImportPage 0, 0, IN_NEWPAGE_END, txtImageNumber, txtFullPathName
                txtPageCount = docContents.NumberOfPages
                Set docContents = Nothing
                
            End If
        End If
        
        
        DoEvents
        
'        '*** Update the Arrays to reflect the ACTUAL number of Pages Loaded
'        '     just in case the FIRST document loaded was a MULTI-PAGE file!
'        subUpdatePageArrays dblNumberOfPagesBeforeImport, dblNumberOfPagesAfterImport

        '*** 2020-05-18 - Jacob - Disabled the KILL from Array, because this is now handled by the TEMP File Delete in frmMainMenu / Form_Unload
'        '2017-09-19 - Jacob - Added KILL for the TEMP Files...
'        '                               must create a LIST Array and then delete them all AFTER the RasterrizeBatchEX command
'        '                               because SPICER seems to keep them open till the export is complete.
'        intTempFilesCount = intTempFilesCount + 1
'        ReDim Preserve arrTempFilesList(intTempFilesCount) As String
'        arrTempFilesList(intTempFilesCount) = txtFullPathName
'
'        Debug.Print intTempFilesCount & " " & arrTempFilesList(intTempFilesCount)
        
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveNext
        
    Wend
    

  
'    subSetCurrentPage
    
'    bolRelatedImagesLoaded = True
    
    
End Sub

Private Function funcExportSelectedSaveDocument(strDirectory As String, _
                                                ctlDocControl As Control, _
                                                Optional strSaveFileName As String, _
                                                Optional strSaveFormat As String) As String


    '*********************************************************************
    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - BEGIN
    
''''    Dim txtAttachmentFileName As String
''''    txtAttachmentFileName = InputBox("What would you like to call THIS document?" & vbCrLf & "Press [Cancel] or the [Esc] key to cancel the send.", "Attachment Name", "Imaging101 Document")
''''
''''    'Check if the user Canceled or entered no filename
''''    If Trim(txtAttachmentFileName) = "" Then
''''        Exit Sub
''''    End If
'''''    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".TIF"

    

'    Dim strSaveFileName As String
    Dim strSaveFileNameExtension As String
    Dim txtAttachmentFileName As String
    
    strDirectory = Trim(strDirectory)
    If Right(strDirectory, 1) <> "\" Then
        strDirectory = strDirectory & "\"
    End If
    
    'Define Filename ONLY if NOT Passed
    If Trim(strSaveFileName) = "" Then
    
        strSaveFileName = Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(5)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(6)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(7)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(8)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(9)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(10)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(11)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index).ListSubItems(12)) & _
                            "~" & _
                           Trim(ListView1.ListItems.item(ListView1.SelectedItem.Index))
    End If
    
    'Remove INVALID Characters from the File Name
    strSaveFileName = funcCleanFileName(strSaveFileName)
    strSaveFileName = Left(strSaveFileName, 245)
    strSaveFileNameExtension = Right(txtFullPathName, Len(txtFullPathName) - InStrRev(txtFullPathName, ".") + 1)
    
    '*** If there was only ONE Record, then simply COPY the File out
    If txtDetailRecordCount = "1" _
    And (txtDetailFileType = "MSG" _
            Or txtDetailFileType = "PDF") Then
        'Set up full path for export file
        txtAttachmentFileName = strDirectory & Trim(strSaveFileName) & "." & txtDetailFileType

        FileCopy txtFullPathName, txtAttachmentFileName

    Else
        'Set up full path for export file
'        strDirectory = funcGetTempDir
        txtAttachmentFileName = strDirectory & Trim(strSaveFileName) & "." & strSaveFormat
    
        
        Dim docSave As IDocSave

    
        Debug.Print
        '***  Save the modified pages in the Spicer Document format
        '     The FORMAT is Different for Single-page VS Multi-Page PDF
        If ctlDocControl.NumberOfPages > 1 Then
            '*** Rasterize the Pages before sending
    '         me.subRasterizeBatch
            Me.subRasterizeBatchEX ctlDocControl
            DoEvents
            ' Set the object variable for the IDocSave interface to the Document Control object
            ' that was saved by the Rasterize sub
            Set docSave = ctlDocControl.object
    '        docSave.SaveAsDialog False
    
            If strSaveFormat = "PDF" Or Trim(strSaveFormat) = "" Then
                'PDF - Flate, 619, R/W, Raster only Multipage PDF subset - bilevel, 8 bit, 24 bit
                docSave.Save 0, False, 620, txtAttachmentFileName, "" 'txtAttachmentFileName
'                docSave.Export 0, False, 620, txtAttachmentFileName, txtAttachmentFileName
            Else
                docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, "" 'txtAttachmentFileName
            End If
            
        Else
        
            '*** Rasterize the Pages before sending
             Me.subRasterizeBatchEX ctlDocControl
            DoEvents
            ' Set the object variable for the IDocSave interface to the Document Control object
            ' that was saved by the Rasterize sub
            Set docSave = ctlDocControl.object
    '        docSave.SaveAsDialog False
            
            If strSaveFormat = "PDF" Then
                'PDF - Flate, 101, R/W, Raster only PDF subset - bilevel, 8 bit, 24 bit
                docSave.Save 0, False, 101, txtAttachmentFileName, "" 'txtAttachmentFileName
            Else
                docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, "" 'txtAttachmentFileName
            End If
            
        End If

    End If
    
    
    DoEvents
    ' De-initialize the object variable
    Set docSave = Nothing
    
    
    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - END
    '*********************************************************************
    
    funcExportSelectedSaveDocument = txtAttachmentFileName
    
Exit Function

ERROR_HANDLER:

    MsgBox "funcExportSelectedSaveDocument ERROR: " & Err.Number & " - " & Err.Description
    
End Function

Private Sub funcExportSelectedSaveDocumentPages(strSaveDirectory As String, strHtmlDirectory As String, strPdfExportDirectory As String, lPreviousPageCount As Long)


    Dim strSaveFileName As String
    Dim strSaveFileNameExtension As String
    Dim strOutputFilePath As String
    
    Dim lPageCount As Long
    Dim lTotalPageCount As Long
    
    Dim iTemp As Integer
    Dim sTemp As String
    
    Dim docContents As IDocContents
    Dim docSave As IDocSave
    
    On Error GoTo ERROR_HANDLER
    
    ' Set the object variable for the IDocContents interface to the Document Control object
    Set docContents = frmClientSpicerControlForm.SpicerDoc1.object
    ' Get the number of pages for the document
    lPageCount = docContents.NumberOfPages
    lTotalPageCount = lPreviousPageCount + lPageCount
    
    sTemp = "DOCUMENT RecID: " & txtDocumentRECID & " Previous Pages: " & str(lPreviousPageCount) & " Current doc Pages:" + str(lPageCount) + " Total Pages: " & lTotalPageCount & "." + vbCrLf
    ' Get the pageID for all pages in the document

    For iTemp = lPreviousPageCount + 1 To lTotalPageCount
       
        sTemp = sTemp + "    Page " + str(iTemp) + " ID: " + str(frmClientSpicerControlForm.SpicerDoc1.pageID(iTemp)) + vbCrLf
    
        ' Display the number of pages and the identifiers for each
        funcWriteToDebugLog Me.name, sTemp
        
        ' De-initialize the object variable
        Set docContents = Nothing
    
        strSaveFileName = "iPage_" & iTemp & "_0"

      
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        Set docSave = frmClientSpicerControlForm.SpicerDoc1.object
        
        
        
        '*** Export to JPG
        'Set up full path for export file
         txtAttachmentFileName = strSaveDirectory & "\" & Trim(strSaveFileName) & ".JPG"
         
         'Save the Current Page to JPG
        txtActionBeforeError = "docSave.Save " & txtAttachmentFileName & vbCrLf & "DEBUG TRACE: " & vbCrLf & sTemp
         docSave.Save frmClientSpicerControlForm.SpicerDoc1.pageID(iTemp), False, API_FF_JFIF, txtAttachmentFileName, txtAttachmentFileName
        
'    '*** Export to PDF if the txtExportToPDF flag is set
'        If UCase(txtExportToPDF) = "Y" Then
'
'            txtAttachmentFileName = strPdfExportDirectory & "\" & Trim(strSaveFileName) & ".PDF"
'            'PDF - JPEG  PDF Raster only PDF subset - 8 bit, 24 bit colour   Bilevel or color    82
'            docSave.Save frmClientSpicerControlForm.SpicerDoc1.pageID(iTemp), False, 82, txtAttachmentFileName, txtAttachmentFileName
'
'        End If


        
        DoEvents
        
        '***************************************
        '*** PREPARE iPage_n.htm file
        
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Prepare .\" & rs!DocumentGroup & "\iPage_" & iTemp & ".htm"
        DoEvents
        txtConcat.Text = ""
        LoadText txtConcat, strHtmlDirectory & "\DOCGROUP\DYNAMIC\iPage_n.txt"
        
        'Set the current Page #
        txtConcat.Text = Replace(txtConcat.Text, "{{{PageNumber}}}", iTemp)
        
        'Set the Previous Page #
        If iTemp > 1 Then
            txtConcat.Text = Replace(txtConcat.Text, "{{{PreviousPageNumber}}}", iTemp - 1)
        Else
            txtConcat.Text = Replace(txtConcat.Text, "{{{PreviousPageNumber}}}", iTemp)
            
        End If
        
        'Set the Next Page #
        If iTemp < lTotalPageCount Then
            txtConcat.Text = Replace(txtConcat.Text, "{{{NextPageNumber}}}", iTemp + 1)
        Else
            txtConcat.Text = Replace(txtConcat.Text, "{{{NextPageNumber}}}", iTemp)
        
        End If
    
        
        '*** SAVE the iPage_n.htm
        strOutputFilePath = strSaveDirectory & "\iPage_" & iTemp & ".htm"
        If funcFileExists(strOutputFilePath) Then
            'Delete the File if it already exists.
            Kill strOutputFilePath
        End If
        Open strOutputFilePath For Output As #1
        Print #1, txtConcat.Text
        Close #1
        
        
        '***************************************************
        '*** Add Body Page Detail
        txtDocGroupBody.Text = txtDocGroupBody.Text + vbCrLf & "<a href=""./ipage_" & iTemp & ".htm"" target=""showframe"">Page " & iTemp & "</a><br>"

        
        ' De-initialize the object variable
        
        Set docSave = Nothing

    Next iTemp
    
    
    lPreviousPageCount = lTotalPageCount
    
    DoEvents
    
    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - END
    '*********************************************************************
Exit Sub

ERROR_HANDLER:
        funcQuickMessage "SHOW", "funcExportSelectedSaveDocumentPages: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [DOCUMENT NOT EXPORTED]"
        
End Sub



Public Sub subRasterizeBatch()


   
   
   'This flag is to disable the Form Load and Activate subs on the ChildForm
   bolRasterizingDocument = True
   
        ' Set the control to create a NEW Raster Document
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = frmClientSpicerControlForm.SpicerConfiguration1.object
        'Set to NOT automatically overwite document on rasterization.
        CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = True
'        'Set to Remove the original
        CFGDocument.RasterOperations(IN_RASTERIZE_REMOVE) = True
        CFGDocument.RasterOperations(IN_RASTERIZE_CROP) = False

        'Deinitialize the object variable
        Set CFGDocument = Nothing

   
    BindControls frmClientSpicerControlForm.SpicerDoc1.object
   
   Dim RasterBatch As IRasterBatch
   Dim docContents As IDocContents

   Dim lPageID As Long
   Dim bAllPages As Boolean
   Dim iXResolution As Integer
   Dim iYResolution As Integer
   Dim bColor As Boolean
   Dim bDither As Boolean
   Dim iLighten As Integer
   
   ' Set the object variable for the IRasterBatch interface to the Edit Control object
   Set RasterBatch = frmClientSpicerControlForm.SpicerEdit1.object
   Set docContents = frmClientSpicerControlForm.SpicerDoc1.object
    
   lPageID = docContents.FirstPageID

   ' Set rasterize options
   bAllPages = True   ' Rasterize all of the pages
   iXResolution = 0   ' Keep the same resolution

   iYResolution = 0
   bColor = False   ' Do not rasterize to color
   bDither = False   ' Do not rasterize to dither
   iLighten = 0
   ' Rasterize the entire document
   RasterBatch.RasterizeBatch lPageID, bAllPages, iXResolution, iYResolution, bColor, bDither, iLighten
   
   ' De-initialize the object variable
   Set RasterBatch = Nothing
   Set docContents = Nothing
End Sub

Public Sub subRasterizeBatchEX(ctlDocControl As Control)

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
   
   On Error GoTo ERROR_HANDLER
   
    txtActionBeforeError = "Dim CFGDocument "
   
   'This flag is to disable the Form Load and Activate subs on the ChildForm
   bolRasterizingDocument = True
   
        ' Set the control to create a NEW Raster Document
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = frmClientSpicerControlForm.SpicerConfiguration1.object
        'Set to NOT automatically overwite document on rasterization.
        CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = True
'        'Set to Remove the original
        CFGDocument.RasterOperations(IN_RASTERIZE_REMOVE) = True
        'Deinitialize the object variable
        Set CFGDocument = Nothing


    BindControls ctlDocControl
   
   ' Set the object variable for the IRasterBatch interface to the Edit Control object
   Set RasterBatch = frmClientSpicerControlForm.SpicerEdit1.object
   
    'Remarks
    '
    'The threshold value is ignored unless the MergeColor value is set to bilevel output.
    'The brightness value is ignored unless MergeColor is set to bilevel (enhanced), bilevel (CAD), or bilevel (diffusion dithered).
    'Changing the orientation is a permanent change to the data, not just a display or header change.
    'The xSize and ySize values must either both be set to 0, or both be set to a size value. You cannot use 0 for one and a size value for the other. When both are set to 0, the units value is ignored.
    '
    'The raster image created by this method either replaces the original image, or must be placed into another Document Control, depending on the ReplaceCurrentDocWhenRasterizing property. For details, see Placing new raster images in a Document Control.
   
    'IN_COLORTYPE_COLOR  0   produces a color image containing up to 16 million colors
    'IN_COLORTYPE_BILEVEL    1   produces a black-and-white image with no dithered sections
    'IN_COLORTYPE_BILEVEL_ENHANCED   2   Windows NT/2000/XP only. Produces a monochrome image depending on the format of the original document. For a list of results, click Bilevel Enhanced.
    'IN_COLORTYPE_BILEVEL_CAD    3   Windows NT only. Produces a monochrome image depending on the format of the original document. For a list of results, click Bilevel CAD.
    'IN_COLORTYPE_BILEVEL_DITHER 4   produces a bilevel image in which dithering--generating pixel patterns--is used to simulate gray or color areas of the original image
    'IN_COLORTYPE_24_BIT_COLOR   5   produces a color image containing up to 16 million colors.
    'IN_COLORTYPE_GRAYSCALE  6   produces a grayscale image.
    
   ' Set the rasterize options
   lObjectID = frmClientSpicerControlForm.SpicerDoc1.RootID
   iXResolution = 0   ' Keep the original resolution
   iYResolution = 0
   iColor = IN_COLORTYPE_COLOR    ' Rasterize to COLOR
   iBrightness = 100 ' Defines the lightness or darkness of the
                     '   bilevel (enhanced), bilevel (CAD), or bilevel (diffusion dithered) profile.
                     '   Values can range from -50 (Dark) to 150 (Light)
   iThreshold = 200  ' Defines the lightness or darkness of the
                      '  bilevel profile.
                     '   Values can range from 0 (Light) to 255 (Dark).
   iOrientation = IN_ORIENTATION_NONE ' Use original orientation
   lXSize = 0 ' Keep the original size
   lYSize = 0
   iUnit = IN_UNITS_INCH

   lYSize = 0
   iUnit = IN_UNITS_INCH
   
   ' Rasterize the entire document
   txtActionBeforeError = "RasterBatch.RasterizeBatchEx"
   RasterBatch.RasterizeBatchEx lObjectID, iXResolution, iYResolution, iColor, iBrightness, iThreshold, iOrientation, lXSize, lYSize, iUnit
   
   ' De-initialize the object variables
   Set RasterBatch = Nothing
   
    '*** 2020-05-18 - Jacob - Disabled the KILL from Array, because this is now handled by the TEMP File Delete in frmMainMenu / Form_Unload
'    '2017-09-19 - Jacob - Added KILL for the TEMP File based on the ARRAY created in  subExportSelectedGetImages()
'    Dim intDeleteFilesCounter As Integer
'   For intDeleteFilesCounter = 1 To intTempFilesCount
'        funcWriteToDebugLog Me.name, "DELETE TEMP FILE - Kill " & arrTempFilesList(intDeleteFilesCounter)
'        Debug.Print "KILL " & intDeleteFilesCounter & " " & arrTempFilesList(intDeleteFilesCounter)
'        On Error Resume Next
'        Kill arrTempFilesList(intDeleteFilesCounter)
'   Next
   
   bolRasterizingDocument = False

Exit Sub

ERROR_HANDLER:
        ' De-initialize the object variables
        Set RasterBatch = Nothing
   
        funcQuickMessage "SHOW", "subRasterizeBatchEX: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [DOCUMENT NOT EXPORTED]"
        

End Sub


Private Sub subPrintDocument()

   Dim PrintView As IPrintView
   ' Set the object variable for the IPrintView interface to the View Control object
   Set PrintView = frmClientSpicerControlForm.SpicerView1.object
   
   'Print DIALOG enabled, "PrintDocument" disabled by Jacob 5/15/2008
   PrintView.PrintDialog
   
'   ' Print one copy of all pages of the document in the active window.
'   ' Do not print a banner or a stamp on it.
'   PrintView.PrintDocument IN_PRINT_ALL_PAGES, 0, 0, 1, IN_PMODE_DOCUMENT, _
             False, IN_ZOOM_SCALETOFIT, IN_ORIENT_BEST_FIT, False, False
   ' De-initialize the object variable

   Set PrintView = Nothing


End Sub

Private Sub cmdPrint_Click()

    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    strCommandSource = "cmdPrint"
    
    '*** Scan for Documents Selected to Export
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
            Load frmClientSpicerControlForm
        
            txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(i).Text
            funcWriteToDebugLog Me.name, txtDocumentRECID
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
            'Import all document pages
            subExportSelectedGetImages
        End If
    Next
    
    '*** Moved the Bind, Print and Close out of the Loop to simplify Printing
    '    also enabled the Print Dialog in the subPrintDocument sub
    DoEvents
    'Save combined document
    BindControls frmClientSpicerControlForm.SpicerDoc1
    subPrintDocument
    DoEvents
    'Close the Document to clear
    frmClientSpicerControlForm.SpicerDoc1.CloseDocument False
    DoEvents

    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub

Private Sub cmdRestore_Click()

    result = MsgBox("Are you SURE you wish the RESTORE the Selected Documents?", vbYesNo, "Document Restore Verificaiton")
    
    If result <> vbYes Then
        Screen.MousePointer = MousePointerConstants.vbDefault
        Exit Sub
    End If

    subSetDocumentStatus ""
    
End Sub

Private Sub cmdSelectALL_Click()

    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
            funcWriteToDebugLog Me.name, frmImaging101Retrieve.ListView1.ListItems(i).Text
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
    Next
    
    
    subCheckButtonSecurity
    
    subShowItemsSelected
    
    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub

Private Sub cmdSend_Click()

    Dim strDirectory As String
    Dim strSaveFilePath As String
    
    On Error GoTo ERROR_HANDLER
    
'    strDirectory = Environ("TEMP") & "\I101Send"
    strDirectory = "C:\TEMP\I101Send"
    
    funcCreateDirectoryStructure strDirectory
    
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass

    strCommandSource = "cmdSend"

        Load frmImaging101ExportOptions
        frmImaging101ExportOptions.chkExportToPDF.Value = vbChecked

    ' CLEAR the Exported Document Listbox
    lstExportedDocuments.Clear
    
    '*** Scan for Documents Selected to Export
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
    
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
                
            Load frmClientSpicerControlForm
        
            'Get the DocumentRECID
            txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(i).Text
            funcWriteToDebugLog Me.name, txtDocumentRECID
            
            'Get the # of Images
            txtPageCount = funcGetFieldFromDB(RegImaging101ConnectionString, txtApplicationName, "DocumentRECID=" & txtDocumentRECID, "DocumentImages")
            
            frmImaging101Retrieve.ListView1.ListItems(i).Selected = True   ' Force item selection
            
            If txtPageCount = "1" Then
                Dim strDetailFileType As String
                strDetailFileType = funcGetFieldFromDB(RegImaging101ConnectionString, txtApplicationName & "_Detail", "DocumentRECID=" & txtDocumentRECID, "DetailFileType")
            End If
            
            'Import all document pages
            subExportSelectedGetImages
            DoEvents
            
            'Save combined document
            strSaveFilePath = funcExportSelectedSaveDocument(strDirectory, frmClientSpicerControlForm.SpicerDoc1, "", "PDF")
            'Add the saved document to the Listbox
            lstExportedDocuments.AddItem strSaveFilePath
            DoEvents
           
            On Error GoTo ERROR_HANDLER
            DoEvents
            
            For w = 1 To 10000
                DoEvents
            Next
           
           Unload frmClientSpicerControlForm
           Set frmClientSpicerControlForm = Nothing
           
        End If
        
    Next
    
    Dim bolSendToSMTP As String
    
    On Error Resume Next
    
'    Me.ActiveForm.subSendTo
    bolSendToSMTP = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Control", "ID = 1", "SendEmailViaSMTP")
    
    If bolSendToSMTP = True Then
        subSendToSMTP
    Else
        subSendToOutlook
    End If
    
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
'    MsgBox "DOCUMENT SEND COMPLETE!", vbOKOnly, "Document SEND Complete"

Exit Sub

ERROR_HANDLER:
        funcQuickMessage "SHOW", "cmdSend_Click: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [DOCUMENT NOT EXPORTED]"
        

End Sub

Private Sub Form_Activate()
    
    funcWriteToDebugLog Me.name, "Form_Activate"
    
    subSetRetrieveButtons
    DoEvents
    
    '*** Prevent Error if the List is Empty
    If Me.ListView1.ListItems.Count > 0 Then
'        ListView1.ListItems.item(1).Selected = True
        ListView1.SetFocus
        DoEvents

''''        'Jacob - 4/5/2007 - Disabled subShowItemsSelected... it caused 100% Utilization
''''                            by alternating Focus between the Retrieve and MainMenu forms.
''''        subShowItemsSelected
''''        DoEvents
    Else
        Exit Sub
    End If
    
        
   
End Sub

Private Sub Form_GotFocus()


    funcWriteToDebugLog Me.name, "Form_GotFocus"
    
    subSetRetrieveButtons
    DoEvents
    
    '*** Prevent Error if the List is Empty
    If lblItemsFound <= 0 Then
        ListView1.ListItems.item(1).Selected = True
        DoEvents
        ListView1.SetFocus
    Else
        Exit Sub
    End If

End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    If bolDebug Then
        Me.Caption = Me.Caption & " - DEBUG"
    End If

  
    lblItemsSelected = ""
    lblItemsFound = ""
    
    txtCurrentModule = "frmImaging101Retrieve"

  
'
    '*** Set SQL wildcard string
    RegConnectionWildcard = "%"
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImaging101Retrieve.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImaging101Retrieve.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmImaging101Retrieve.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmImaging101Retrieve.Height", RegFileName)
'    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmImaging101Retrieve.Caption", RegFileName)
'    If Me.Caption = "" Then Me.Caption = "Document List"
    
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
    
    subCheckButtonSecurity
    
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
'  Frame1.Width = Me.ScaleWidth
  txtFullPathName.Top = Me.ScaleHeight - 300
  
  On Error Resume Next
  'This will resize the ListView when the form is resized
  If Me.ScaleHeight > 200 Then
    ListView1.Height = Me.ScaleHeight - ListView1.Top
  End If
  
  If Me.ScaleWidth > 200 Then
      ListView1.width = Me.ScaleWidth - ListView1.Left
  End If
  
    Frame1.width = Me.width
    picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
    lblVersion.Left = picImaging101Logo.Left
  
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
''''''  If mbEditFlag Or mbAddNewFlag Then Exit Sub
''''''
''''''  Select Case KeyCode
''''''    Case vbKeyEscape
''''''      cmdClose_Click
''''''    Case vbKeyEnd
''''''      cmdLast_Click
''''''    Case vbKeyHome
''''''      cmdFirst_Click
''''''    Case vbKeyUp, vbKeyPageUp
''''''      If Shift = vbCtrlMask Then
''''''        cmdFirst_Click
''''''      Else
''''''        cmdPrevious_Click
''''''      End If
''''''    Case vbKeyDown, vbKeyPageDown
''''''      If Shift = vbCtrlMask Then
''''''        cmdLast_Click
''''''      Else
''''''        CmdNext_Click
''''''      End If
''''''  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImaging101Retrieve.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Retrieve.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Retrieve.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Retrieve.Height", Me.Height, RegFileName)
'        result = WritePrivateProfileString(RegAppname, "frmImaging101Retrieve.Caption", Me.Caption, RegFileName)
    End If
  
    Unload frmImaging101ExportOptions
    

End Sub







Private Sub PrimaryCLS_MoveComplete()
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(PrimaryCLS.AbsolutePosition)
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
'  PrimaryCLS.MoveLast
'  PrimaryCLS.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
    
    result = MsgBox("Are you SURE you wish the DELETE the Selected Documents?", vbYesNo, "Document Delete Verificaiton")
    
    If result <> vbYes Then
        Screen.MousePointer = MousePointerConstants.vbDefault
        Exit Sub
    End If

    subSetDocumentStatus "D"
    
End Sub


Private Sub subSetDocumentStatus(strDocumentStatus As String)

    '*** DOCUMENT DELETE LOGIC
    On Error GoTo ErrLock
    
    strCommandSource = "subSetDocumentStatus"
    
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass

    '*** Scan for Documents Selected to Export
    For i = 1 To frmImaging101Retrieve.ListView1.ListItems.Count
    
        If frmImaging101Retrieve.ListView1.ListItems(i).Selected = True Then
        
        
            txtDocumentRECID = frmImaging101Retrieve.ListView1.ListItems(i).Text
            funcWriteToDebugLog Me.name, txtDocumentRECID
            
            
            Dim conn As ADODB.Connection
            Dim rs As ADODB.Recordset
            
            Set conn = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rs = New ADODB.Recordset
            
            On Error GoTo ErrLock
        
            txtActionBeforeError = "Prepare to Open Batch DB Connection"
            '*** Prepare Connection
            With conn
                .ConnectionString = RegImaging101ConnectionString
                .CursorLocation = adUseServer
                .ConnectionTimeout = 120
                .IsolationLevel = adXactReadCommitted
                .mode = adModeWrite
                txtActionBeforeError = "Open Batch DB Connection"
                .Open
                .Execute "SET LOCK_TIMEOUT -1"
            End With
            
            Set cmd.ActiveConnection = conn
    
            '*** Prepare Result Set
            With rs
                .ActiveConnection = conn
                .CursorLocation = adUseServer
                .CursorType = adOpenDynamic
                .LOCKTYPE = adLockOptimistic
            End With
            
            '*** 2020-09-14 - Jacob - Had to add Single Quotes around the "Date" because SQL Sever is now creating
            '                                             the datetime field as datetime2(0) for NEW Applications.
            rs.Source = "UPDATE " & frmImaging101Search.txtApplicationName & _
                        "   SET DocumentLocked = '" & strDocumentStatus & "', " & _
                        "       DocumentLockedBy = '" & gsecUserID & "', " & _
                        "       DocumentLockedDate = '" & Date & "' " & _
                        " WHERE DocumentRECID = " & txtDocumentRECID
            
            conn.Errors.Clear
            rs.Open
    
    
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
    
            DoEvents
        End If
           
    Next
    
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    MsgBox "ACTION COMPLETE!", vbOKOnly, "Action Complete"

    'Re-issue the Find
    frmImaging101Search.cmdFind_Click
    
    
Exit Sub

ErrLock:

    Screen.MousePointer = MousePointerConstants.vbDefault
    
    If conn.Errors.item(0).NativeError = 1222 Then  ' Lock Timeout
        MsgBox "cmdDelete: Unable to set Document as Deleted due to a Record Lock Timeout. - Document Record ID = " & txtDocumentRECID
    Else
        MsgBox "cmdDelete: ERROR: " & Err.Number & " - " & Err.Description & " -  Document Record ID = " & txtDocumentRECID

    End If
    
    conn.Errors.Clear
    
End Sub

Public Sub cmdRefresh_Click()
    
'    ' If DateTHRU is blank, fill the DateFROM
'    If Trim(mebDateThru) = "" Then
'        mebDateThru = mebDateFrom
'    End If
'
'    '***
'    '*** SET UP THE "WHERE" CLAUSE FILTER STATEMENT
'    '***
'
'    txtFilterStatement = ""
'    ' Check for FileRoom
'    If Trim(frmImaging101Retrieve.cmbCustomerNumber) <> "" Then
'         txtFilterStatement = "CUSTOMERNUMBER LIKE '%" + frmImaging101Retrieve.cmbCustomerNumber + "%' "
'    End If
'    ' Check for FileCabinet
'    If Trim(frmImaging101Retrieve.cmbCustomerName) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " " + cmbFileroomCondition + " "
'         End If
'         txtFilterStatement = txtFilterStatement + " NAME LIKE '%" + frmImaging101Retrieve.cmbCustomerName + "%'"
'    End If
'    ' Check for DateFROM
'    If Trim(frmImaging101Retrieve.mebDateFrom) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " " & cmbDocumentTypeCondition + " "
'         End If
'         txtFilterStatement = txtFilterStatement + " TRANSDATE >= '" + Trim(Format(frmImaging101Retrieve.mebDateFrom, "####-##-##")) + "'"
'    End If
'    ' Check for DateTHRU
'    If Trim(frmImaging101Retrieve.mebDateThru) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " And "
'         End If
'         txtFilterStatement = txtFilterStatement + " TRANSDATE <= '" + Trim(Format(frmImaging101Retrieve.mebDateThru, "####-##-##")) + "'"
'    End If
'    ' Check for Part #
'    If Trim(frmImaging101Retrieve.txtPartNumber) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " And "
'         End If
'         txtFilterStatement = txtFilterStatement + " PARTNUMBER LIKE '%" + frmImaging101Retrieve.txtPartNumber + "%'"
'    End If
'
'    ' Check for Invoice #
'    If Trim(frmImaging101Retrieve.txtInvoiceNumber) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " And "
'         End If
'         txtFilterStatement = txtFilterStatement + " INVOICENUMBER LIKE '%" + frmImaging101Retrieve.txtInvoiceNumber + "%'"
'    End If
'
'    ' Check for DocGroup
'    If Trim(frmImaging101Retrieve.cmbDocGroup) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " And "
'         End If
'         txtFilterStatement = txtFilterStatement + " DOCGROUP LIKE '%" + frmImaging101Retrieve.cmbDocGroup + "%'"
'    End If
'    ' Check for DocType
'    If Trim(frmImaging101Retrieve.cmbDocType) <> "" Then
'        If Trim(txtFilterStatement) <> "" Then
'                  txtFilterStatement = txtFilterStatement + " " + cmbFilecabinetCondition + " "
'         End If
'         txtFilterStatement = txtFilterStatement + " DOCTYPE LIKE '%" + frmImaging101Retrieve.cmbDocType + "%'"
'    End If
'    If Trim(txtFilterStatement) <> "" Then
'        txtFilterStatement = " WHERE " & txtFilterStatement
'    End If
'
'    '***
'    '*** SET UP THE SELECT STATEMENT
'    '***     INCLUDING THE "WHERE" AND "ORDER BY" CLAUSES
'
''   Adodc1.RecordSource = "select BatchID, DocumentID , UNCFilePath, Filename, " & _
''                        " Fileroom, Filecabinet, DocumentType, DocumentDate, PageCount, " & _
''                        " Folder, FolderDescription, DocumentSubType, DateAdded, " & _
''                        " DocumentExpireDate, DocumentNote, Field8, Field9, " & _
''                        " Field10, Field11, Field12, Field13, Field14, Field15, " & _
''                        " Field16, Field17, Field18, Field19, Field20 " & _
''                        " FROM I101Documents " & txtfilterstatement & _
''                        " ORDER BY " & _
''                        frmConfig.cmbSort(0) + ", " + frmConfig.cmbSort(1) + ", " + frmConfig.cmbSort(2) + ", " + frmConfig.cmbSort(3)
'
'   Adodc1.RecordSource = "select BatchRecID, DocumentID , UNCFilePath, Filename, " & _
'                        " CUSTOMERNUMBER, NAME, TRANSDATE, PARTNUMBER, INVOICENUMBER, " & _
'                        " DOCGROUP, DOCTYPE " & _
'                        " FROM HOLMANPARTS " & txtFilterStatement & _
'                        " ORDER BY " & _
'                        "CUSTOMERNUMBER, TRANSDATE,DOCGROUP, DOCTYPE"
'   Adodc1.Refresh
'
'   txtItemsFound = Adodc1.Recordset.RecordCount
'
'   chkViewDocDetails_Click
'
'
'   Exit Sub
'
'
'RefreshErr:
'  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
'  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  PrimaryCLS.Cancel
  SetButtons True
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

'  PrimaryCLS.Update
  SetButtons True
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  PrimaryCLS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  PrimaryCLS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub CmdNext_Click()
  On Error GoTo GoNextError

  PrimaryCLS.MoveNext
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  PrimaryCLS.MovePrevious
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub subSetRetrieveButtons()

       '*** Prevent Error if the List is Empty
    If Me.ListView1.ListItems.Count = 0 Then
        cmdModify.Enabled = False
        cmdOpenSelected.Enabled = False
        cmdSelectAll.Enabled = False
        cmdDeSelectAll.Enabled = False
        Exit Sub
    End If
    
    If UCase(gsecRightsModifyIndexes = vbChecked) Then
        cmdModify.Enabled = True
    Else
        cmdModify.Enabled = False
    End If
    
    cmdOpenSelected.Enabled = True
    cmdSelectAll.Enabled = True
    cmdDeSelectAll.Enabled = True
    
'    ListView1.ListItems.item(1).Selected = True
'    DoEvents
    
    If UCase(gsecRightsAdminSystem = vbChecked) Then
        cmdMove.Visible = True
    Else
        cmdMove.Visible = False
    End If
    
    Me.Enabled = True
    ListView1.Enabled = True
    
    ListView1.SetFocus
    DoEvents
        
 
End Sub


Private Sub TabStrip1_Click()

End Sub



Public Sub subPopulateListview()

    On Error GoTo ERROR_TRAP
    
    '*** Setup Up ListView properties - BEGIN
    
    ListView1.Visible = False
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = True
    
        '***  SET COLUMN HEADINGS
        For intListIndex = 0 To Adodc1.Recordset.Fields.Count - 1
            
            Dim strFieldNameForOutput As String
            
            strFieldNameForOutput = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID & " AND FieldName = '" & Adodc1.Recordset.Fields.item(intListIndex).name & "'", "FieldNameForOutput")
            
            If strFieldNameForOutput = "" Then
                ListView1.ColumnHeaders.Add , , Adodc1.Recordset.Fields.item(intListIndex).name, Len(Adodc1.Recordset.Fields.item(intListIndex).name) * 150, lvwColumnLeft
            Else
                ListView1.ColumnHeaders.Add , , strFieldNameForOutput, Len(strFieldNameForOutput) * 150, lvwColumnLeft
            End If
            
        Next
                
'    On Error Resume Next
'
'    Adodc1.Recordset.MoveFirst

    While Not Adodc1.Recordset.EOF
            For intListIndex = 0 To Adodc1.Recordset.Fields.Count - 1
                If intListIndex = 0 Then
                    If Not IsNull(Adodc1.Recordset.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , Adodc1.Recordset.Fields.item(intListIndex).Value)
                    End If
                Else
            
                    '* This null check is to make sure we don't Skip fields caused by an error.
                    If Not IsNull(Adodc1.Recordset.Fields.item(intListIndex).Value) Then
                        ' Not null... show value
                        
                        Select Case Adodc1.Recordset.Fields.item(intListIndex).Type
                            Case adNumeric, adInteger, adDouble, adSingle, adSmallInt
                                '*** FORCE RIGHT ALIGNMENT OF NUMBERS
                                Set lstSubItem = lstItem.ListSubItems.Add(, , Right("             " & Trim(CStr(Adodc1.Recordset.Fields.item(intListIndex).Value)), 12))
                            Case adDBTimeStamp
                                Set lstSubItem = lstItem.ListSubItems.Add(, , Format(Adodc1.Recordset.Fields.item(intListIndex).Value, "yyyy/mm/dd"))
                            Case Else
                                Set lstSubItem = lstItem.ListSubItems.Add(, , Adodc1.Recordset.Fields.item(intListIndex).Value)
                        End Select
                        
                    Else
                        ' Null... show empty string
                        Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                    End If
                    
                End If
            Next
        Adodc1.Recordset.MoveNext
    Wend
    On Error GoTo 0
    
    ' AutoSize ALL Columns
    Dim i As Integer, lparam As Long
    UseHeader = True
    If UseHeader = False Then
        lparam = LVSCW_AUTOSIZE
    Else
        lparam = LVSCW_AUTOSIZE_USEHEADER
    End If
    For i = 0 To ListView1.ColumnHeaders.Count - 1
        SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lparam
    Next
    
    ' Hide the RecordID's
    ListView1.ColumnHeaders(1).width = 0
    ListView1.ColumnHeaders(2).width = 0
    ListView1.ColumnHeaders(3).width = 0
    ListView1.ColumnHeaders(4).width = 0
    ListView1.ColumnHeaders(5).width = 0
    
    ' Size the Key fields to a standard size
''    ListView1.ColumnHeaders(5).Width = 3000
''    ListView1.ColumnHeaders(4).Width = 2000
''    ListView1.ColumnHeaders(5).Width = 1000
''    ListView1.ColumnHeaders(6).Width = 1000
''    ListView1.ColumnHeaders(7).Width = 1000
    
    
    ListView1.Visible = True

    '*** Setup Up ListView properties - END
    
    '*** Show how many items are selected
    subShowItemsSelected

    '*** If only ONE items was selected, Go ahead and OPEN IT!
    If ListView1.ListItems.Count = 1 Then
        ListView1_DblClick
    End If
    
    'Check to see if the RECID / Details should be displayed
    chkViewDocDetails_Click
    
Exit Sub
    
ERROR_TRAP:

    MsgBox "subPopulateListview ERROR: " & Err.Number & " - " & vbCrLf & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ")", vbExclamation
    
End Sub









Private Sub lblItemsSelected_Change()

    If Trim(lblItemsSelected.Caption = "") Then
        cmdMove.Enabled = False
    Else
        cmdMove.Enabled = True
    End If
    
End Sub

Public Sub ListView1_DblClick()
    
    funcWriteToDebugLog Me.name, "ListView1_DblClick"
    
    '*** Prevent Error if the List is Empty
    If Me.ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    'Disable the Retrieve and Search forms so the user can't click Multiple times
    '  to prevent users clicking wildly which can crash the image loading process.
    '*** DO NOT Disable the and MainMDI Viewer... Otherwise the ChildForms can't load properly.
    Me.Enabled = False
    frmImaging101Search.Enabled = False
    
    Dim txtPathSubdirectory As String
    Dim txtFileName As String
    Dim txtFullPathName As String
    Dim txtDocumentRECID As Double
    Dim txtDetailRECID As Double
    Dim txtDetailRecordCount As Double
    Dim txtDetailOrder As Double
    Dim txtDetailRotation As Integer
    
    Dim txtFTStatus As String
    Dim txtFTDirectory As String
    Dim txtFTFileName As String
    

    bolObjectLaunched = False
    
        lstIndex = Me.ListView1.SelectedItem.Index
        
        ' Get Main Item
        txtDocumentRECID = Me.ListView1.ListItems(lstIndex).Text
        
        
        If Not IsArrayEmpty(gFormArrayRetrieve) Then
        
            For i = 0 To UBound(gFormArrayRetrieve)
                
                    Set frmViewForm = gFormArrayRetrieve(i)
                    If frmViewForm.txtDocumentRECID = txtDocumentRECID Then
                            frmViewForm.SetFocus
                            'funcQuickMessage "SHOW", "Document ALREADY LOADED... Exiting Sub  ListView1_DblClick()"
                            '*** Re-enable the Forms
                            Me.Enabled = True
                            frmImaging101Search.Enabled = True
                            '*** GET OUT NOW
                            Exit Sub
                    End If
                 
            Next
            
        End If

        
        
    On Error GoTo DOCUMENT_RETRIEVE_ERRORS
        

        
        
        
''        ' Get Sub-Items
''        txtPathSubdirectory = Me.ListView1.ListItems(lstIndex).ListSubItems(2).Text
''        txtFileName = Me.ListView1.ListItems(lstIndex).ListSubItems(3).Text
        
   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   
    frmImaging101Retrieve.Adodc1.ConnectionTimeout = 300
    frmImaging101Retrieve.Adodc1.CommandTimeout = 600
    
    Dim txtFullTextApp As String
    Dim bolFullTextApp As Boolean
    
    txtFullTextApp = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "IndexFullText") & ""
    
    If txtFullTextApp = "1" Then
        bolFullTextApp = True
    Else
        bolFullTextApp = False
    End If
    
   If bolFullTextApp Then
        frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                             " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                             " DetailFileName, DetailFileType, DetailRotation,  " & _
                             " FTStatus, FTDirectory, FTFileName " & _
                             " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                             " WHERE  DocumentRECID = " & txtDocumentRECID & _
                             " ORDER BY DetailOrder "
    Else
            frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                             " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                             " DetailFileName, DetailFileType, DetailRotation  " & _
                             " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                             " WHERE  DocumentRECID = " & txtDocumentRECID & _
                             " ORDER BY DetailOrder "
    End If
    
    funcWriteToDebugLog Me.name, "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    funcWriteToDebugLog Me.name, "ADodcDetail MoveFirst "
    frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
   
   Dim bolFullTextDoc As Boolean
   bolFullTextDoc = False
   If bolFullTextApp Then
        txtFTStatus = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("FTStatus") & ""
        If Trim(txtFTStatus) <> "" Then
            bolFullTextDoc = True
        End If
    End If
        
    If bolFullTextDoc Then
        'Get the Full-Text path and name and hardcode for a single detail item
        txtFTDirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("FTDirectory") & ""
        txtFTFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("FTFileName") & ""
        
        txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("FTDirectory") & ""
        txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("FTFileName") & ""
        txtDetailRecordCount = 1
    Else
        'Get the original file path and name
        txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
        txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
        txtDetailRecordCount = frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount
    End If
        
        
    txtFullPathName = txtPathSubdirectory & "\" & txtFileName
    txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")
    txtDetailOrder = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailOrder")
     
    
    
                '*** Build the Caption using the first few fields from the ListView
        Dim txtCaption As String
        Dim intListSubItemsIndex
        txtCaption = ""
        For intListSubItemsIndex = 5 To Me.ListView1.ListItems(lstIndex).ListSubItems.Count
            txtCaption = txtCaption + Trim(Me.ListView1.ListItems(lstIndex).ListSubItems(intListSubItemsIndex)) & "|"
            'Don't show fields past Field 9
            If intListSubItemsIndex = 9 Then
                Exit For
            End If
            
        Next intListSubItemsIndex
    
    
    
    '***************************************************************************************
    '***  MOVED TO HERE FROM  MainMDIForm.funcShowImage() -> "Case gI101ModuleRetrieve"
    '***   Also added the "IsArrayEmpty()" function since when the first document is opened
    '***   the array will be Empty and an Error 9 "Subscript Out of Range" would occur
    
'    If Not IsArrayEmpty(arrDisplayedPagesRetrieve) Then
    If IsArrayEmpty(arrDisplayedPagesRetrieve) Then
                  
                i = 0
                
    Else
    
                i = UBound(arrDisplayedPagesRetrieve)
                
                funcWriteToDebugLog Me.name, "UBound(arrDisplayedPagesRetrieve) = " & UBound(arrDisplayedPagesRetrieve)
                
                '*** If the Array is not empty Check if THIS DetailRECID item is already open
                If UBound(arrDisplayedPagesRetrieve) > 0 Then
                    For j = 0 To i
                        If arrDisplayedPagesRetrieve(j) = txtDetailRECID Then
                        
                                'The item IS Already Loaded... Simply Set the Focus to it.
                                gFormArrayRetrieve(j).SetFocus
                                
                           
                                    If MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "Object Launched" Then

                                            txtLaunchFileFullPath = MainMDIForm.ActiveForm.StatusBar1.Panels(2).Text
                                            '**************************************************************************************
                                            '*** 2021-10-20 - Jacob - Check if Windows is already open using Wildcards
                                            
                                            Dim lngWindowExists As Long
                                            lngWindowExists = funcFindWindowLike("*" & txtLaunchFileFullPath)
                                             
                                            If lngWindowExists <> 0 Then
                                                    bolObjectLaunched = True
                                                    funcShowImage = -1
                                                    'Bring the window we found to the top.
                                                    Call FormOnTop(lngWindowExists, True)
                                                    
                                                    'Re-enable the Forms
                                                    Me.Enabled = True
                                                    frmImaging101Search.Enabled = True
                                                    
                                                    funcWriteToDebugLog Me.name, "Document ALREADY LOADED... Exiting Sub  ListView1_DblClick()"
                    
                                                    ' Exit Function changed to Exit Sub
                                                    Exit Sub
                                                    
                                            Else

                                                    'The document window was closed, must unload the MDI Child window to re-load document.
                                                    Unload MainMDIForm.ActiveForm
                                                    'Now exit the Loop to re-load the file
                                                    Exit For
                                                    
                                            End If
                                            
                                    End If
                                    


                        End If
                        
                    Next
                    
                End If
    

    End If

    '***
    '***************************************************************************************
    
    
    
    '***********************************************************
    '*** Handle NULL value in Rotation
    If Not frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRotation") Then
        txtDetailRotation = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRotation")
    End If
    
    '***********************************************************
    '*** Check if file Exists
    If Not funcFileExists(txtFullPathName) Then
        result = MsgBox("SORRY! I can't find file:" + vbNewLine + txtFullPathName + vbNewLine + "PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR!", vbCritical)
'        txtLOGOutputFilePath = Form1.txtOutputFilePath + "\" + txtFullPathName + ".LOG"
'        Open txtLOGOutputFilePath For Append As #4
'        Print #4, "Pass2 - Could Not Open either Original or Fixed file:  " + txtFullPathName
'        Close #4
    'Re-enable the Forms
        Me.Enabled = True
        frmImaging101Search.Enabled = True
        Exit Sub
    End If
    
    
    
        '******************************************************************************************************
        '*** 2022-07-25 - Jacob - Check if we should write a Document Open LOG Record
   
        If bolLogOpenedDocuments = True Then
        
                Dim strAddDocumentOpenLogRecord As String
                Dim strAddDocumentOpenLogRecordResult As Boolean
                Dim strDateFormatted As String
                
                strDateFormatted = Now()


                strAddDocumentOpenLogRecord = "INSERT INTO I101LogDocumentActions ("
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " [LogDate] , "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " [LogAction],  "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " [UserID], "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " [UserName] , "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " [DocumentRECID] , "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " [LogNotes]  ) "

                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & "  VALUES ( "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " '" & Now() & "', "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " 'R', "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " '" & gsecUserID & " ', "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & " '" & gsecUserName & "', "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & txtDocumentRECID & ", "
                strAddDocumentOpenLogRecord = strAddDocumentOpenLogRecord & "NULL ) "

                funcRunSQLCommand RegImaging101ConnectionString, strAddDocumentOpenLogRecord
        
        End If
        
        
        '***********************************************************
        '*** Load the document / object into the Spicer Viewer
        funcWriteToDebugLog Me.name, "MainMDIForm.Show"
        
        MainMDIForm.Show
        
        txtCurrentModule = "frmImaging101Retrieve"
        
        funcWriteToDebugLog Me.name, "frmImaging101Retrieve.Listview1.DblClick | txtCurrentModule = " & txtCurrentModule
        

        
        funcWriteToDebugLog Me.name, "frmImaging101Retrieve.Listview1.DblClick | MainMDIForm.funcShowImage(" & txtPathSubdirectory & ", " & txtFileName & ", " & _
                                            txtDocumentRECID & ", " & txtDetailRECID & ", " & txtCaption & ", " & _
                                            txtDetailRecordCount & ", " & txtDetailOrder & ", " & txtDetailRotation & ", " & _
                                            txtFTStatus & ", " & txtFTDirectory & ", " & txtFTFileName & ", " & _
                                            gI101ModuleRetrieve & ")"
                                            
        Dim intShowImageResult As Integer
        intShowImageResult = MainMDIForm.funcShowImage("" & txtPathSubdirectory, txtFileName, _
                                            txtDocumentRECID, txtDetailRECID, txtCaption, _
                                            txtDetailRecordCount, txtDetailOrder, txtDetailRotation, _
                                            txtFTStatus, txtFTDirectory, txtFTFileName, _
                                            gI101ModuleRetrieve)
        
        '*** 2021-08-10 - Jacob - Added Check for Errors occured during funcShowImage()
   '     If result = 0 And bolErrorOccured = False Then
                'SET the ApplicationRECID field
                
               
                funcWriteToDebugLog Me.name, "Listview1.DblClick | Set frmViewForm = gFormArrayRetrieve(" & i & ") | THIS WILL TRIGGER THE ChildForm1.Form_Load() EVENT."
'                 Set frmViewForm = gFormArrayRetrieve(i)
                 Set frmViewForm = MainMDIForm.ActiveForm
                 
                funcWriteToDebugLog Me.name, "Listview1.DblClick | MainMDIForm.ActiveForm.txtApplicationRECID = " & frmImaging101Search.txtApplicationRECID
                frmViewForm.txtApplicationRECID = frmImaging101Search.txtApplicationRECID
                
                'Initialize the Child Form
                funcWriteToDebugLog Me.name, "Listview1.DblClick | MainMDIForm.ActiveForm.subInitializeChildForm"
                frmViewForm.subInitializeChildForm
'        End If
                Set frmViewForm = Nothing

    
'    End If
    
    'Re-enable the Forms
    Me.Enabled = True
    frmImaging101Search.Enabled = True

    
Exit Sub
    
DOCUMENT_RETRIEVE_ERRORS:

    '*** 2021-10-12 - Jacob - Changed from msgBox to funcQuickMessage and to show Full File Path
    funcQuickMessage "SHOW", "Listview1.DblClick | Retrieve Document ERROR: " & Err.Number & " - " & vbCrLf & Err.Description & vbCrLf & "FilePath= " & txtPathSubdirectory & "\" & txtFileName & vbCrLf & "See DEBUG Log file for details."
    funcWriteToDebugLog Me.name, "Listview1.DblClick | Retrieve Document ERROR: " & Err.Number & " - " & vbCrLf & Err.Description & vbCrLf & "FilePath= " & txtPathSubdirectory & "\" & txtFileName
    On Error Resume Next
    
    'Re-enable the Forms
    Me.Enabled = True
    frmImaging101Search.Enabled = True
    
    
End Sub

Private Sub ListView1_GotFocus()

    funcWriteToDebugLog Me.name, "ListView1_GotFocus"
    
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)

    On Error GoTo ERROR_HANDLER
        
    funcWriteToDebugLog Me.name, "ENTER: ListView1_ItemClick"
    
    ' Force a ListView1_Click upon mouse up/down
    ListView1_Click
    
'    subShowItemsSelected
    
    funcWriteToDebugLog Me.name, "EXIT: ListView1_ItemClick"
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR ListView1_ItemClick: " & Err.Number & " - " & Err.Description

End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER: ListView1_ColumnClick"
    
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ' Set the Sort Column
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Sort It!
    ListView1.Sorted = True
    
    funcWriteToDebugLog Me.name, "EXIT: ListView1_ColumnClick"
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR ListView1_ColumnClick: " & Err.Number & " - " & Err.Description

End Sub


Private Sub ListView1_Click()
    
    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER: ListView1_Click"
    
    '*** Prevent Error if the List is Empty
    If Me.ListView1.ListItems.Count = 0 Then
        funcWriteToDebugLog Me.name, "NO Items in ListView... EXIT:"
        Exit Sub
    End If
    
    funcWriteToDebugLog Me.name, "CALLING subShowItemsSelected"
    
    subShowItemsSelected
    
    funcWriteToDebugLog Me.name, "EXIT: ListView1_Click"
    
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR ListView1_Click: " & Err.Number & " - " & Err.Description

End Sub



Private Sub ListView1_KeyPress(KeyAscii As Integer)
    
    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER: ListView1_KeyPress"
    
    'Catch Enter key
    If KeyAscii = 13 Then
        ListView1_DblClick
    End If

    If KeyAscii = Asc("[") And frmImaging101Retrieve.Visible = True Then
        frmImaging101Search.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("]") And frmImaging101Retrieve.Visible = True Then
        frmImaging101Retrieve.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc("\") And MainMDIForm.Visible = True Then
        MainMDIForm.SetFocus
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    End If

    funcWriteToDebugLog Me.name, "EXIT: ListView1_KeyPress"
    
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR ListView1_KeyPress: " & Err.Number & " - " & Err.Description

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

            'Control + RightClick
            If (Shift = vbCtrlMask) And (KeyCode = vbKeyC) Then
            
                funcQuickMessage "SHOW", "Key Press = Ctrl+C"
                
            End If
    

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER: ListView1_MouseUp"
    
    If Me.ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    subCheckButtonSecurity
    
    cmdOpenSelected.Enabled = True
    
            
    'Right Mouse Click & user  has Modify Index Rights
    If (Button = vbRightButton) And (UCase(gsecRightsModifyIndexes = vbChecked)) Then
    
        funcWriteToDebugLog Me.name, "      ListIndex= " & ListView1.SelectedItem.Index
        'Disable the Retrieve and Search forms so the user can't click Multiple times
        '  to prevent users clicking wildly which can crash the image loading process.
        '*** DO NOT Disable the and MainMDI Viewer... Otherwise the ChildForms can't load properly.
        Me.Enabled = False
        frmImaging101Search.Enabled = False
        
        cmdFastFix_Click
    End If
                    

    funcWriteToDebugLog Me.name, "EXIT: ListView1_MouseUp"
    
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR ListView1_MouseUp: " & Err.Number & " - " & Err.Description

End Sub

Public Sub subCheckButtonSecurity()

    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER: subCheckButtonSecurity"

        If gsecRightsAdminSystem = vbChecked Then
    '        cmdConfigOrder.Enabled = True
            chkViewDocDetails.Visible = True
            chkViewDeletedDocuments.Visible = True
        Else
    '        cmdConfigOrder.Enabled = False
            chkViewDocDetails.Visible = False
            chkViewDeletedDocuments.Visible = False
        End If
        
        ' ALL Other buttons
        If gsecRightsPrint = vbChecked Then
            cmdPrint.Visible = True
        Else
            cmdPrint.Visible = False
        End If
        
        If gsecRightsSendMail = vbChecked Then
            cmdSend.Visible = True
        Else
            cmdSend.Visible = False
        End If
        
        If gsecRightsExport = vbChecked Then
            cmdExportSelected.Visible = True
            cmdExportToExcel.Visible = True
        Else
            cmdExportSelected.Visible = False
            cmdExportToExcel.Visible = False
        End If
        
        If gsecRightsDeleteDocuments = vbChecked Then
            'Align the Restore button
            cmdRestore.Left = cmdDelete.Left
            If chkViewDeletedDocuments.Value = vbChecked Then
                cmdDelete.Visible = False
                cmdRestore.Visible = True
            Else
                cmdDelete.Visible = True
                cmdRestore.Visible = False
            End If
        Else
            cmdDelete.Visible = False
            cmdRestore.Visible = False
        End If
        
        funcWriteToDebugLog Me.name, " BEFORE: If funcCountListViewItemsSelected(ListView1) >= 1 "
        
        ' OPEN - Make sure at least one item is selected
        If funcCountListViewItemsSelected(ListView1) >= 1 Then
            'Items selected - Enable buttons
            cmdOpenSelected.Enabled = True
            cmdPrint.Enabled = True
            cmdSend.Enabled = True
        Else
            cmdOpenSelected.Enabled = False
            cmdPrint.Enabled = False
            cmdSend.Enabled = False
        End If
        
        
        'MODIFY - Check Security
        If gsecRightsModifyIndexes = vbChecked Then
            cmdModify.Visible = True
            ' Make sure Only ONE item is selected
'            If funcCountListViewItemsSelected(ListView1) <> 1 Then
'                'No Items or more than one item selected - Disable buttons
'                cmdModify.enabled = False
'            Else
                cmdModify.Enabled = True
'            End If
        Else
            cmdModify.Visible = False
        End If
        
    funcWriteToDebugLog Me.name, "EXIT: subCheckButtonSecurity"
    
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR subCheckButtonSecurity: " & Err.Number & " - " & Err.Description

End Sub

Private Sub subShowItemsSelected()

    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "ENTER: subShowItemsSelected"

    'Count how many items are "Selected"
    Dim i As Double
    Dim j As Double

    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            j = j + 1
'            funcWriteToDebugLog Me.name, "  IN LOOP: ListView1.ListItems.Count = " & ListView1.ListItems.Count & " i=" & i & " j=" & j
        End If
    Next i
    
'    funcWriteToDebugLog Me.name, "  END LOOP: ListView1.ListItems.Count = " & ListView1.ListItems.Count & " i=" & i & " j=" & j
    
    lblItemsSelected = j
    
Exit Sub

ERROR_HANDLER:
        funcWriteToDebugLog Me.name, "  ERROR subShowItemsSelected: " & Err.Number & " - " & Err.Description
        funcWriteToDebugLog Me.name, "        ListView1.ListItems.Count = " & ListView1.ListItems.Count & " i=" & i & " j=" & j

End Sub

' Perform all necessary binding for activeX controls
Private Sub BindControls(ctlDocControl As Control)

    On Error Resume Next
    
    
    frmClientSpicerControlForm.SpicerView1.BindToDocumentControl ctlDocControl.object
'    SpicerMarkup1.BindToDocumentControl frmClientSpicerControlForm.SpicerDoc1.object
'    SpicerMarkup1.BindToViewControl frmClientSpicerControlForm.SpicerView1.object
    frmClientSpicerControlForm.SpicerEdit1.BindToDocumentControl ctlDocControl.object


'    detailWin.SpicerDetail1.BindToViewControl frmClientSpicerControlForm.SpicerView1.object
'    thumbWin.SpicerThumbnail1.BindToViewControl frmClientSpicerControlForm.SpicerView1.object
'    refWin.SpicerReference1.BindToViewControl frmClientSpicerControlForm.SpicerView1.object
'    layersWin.SpicerLayersWin1.BindToViewControl frmClientSpicerControlForm.SpicerView1.object
End Sub

Private Sub subMDVIP()
    
    Dim strDirectory As String
    Dim strSubDirectory As String
    
    Dim strRootDirectoryPathForHtmlSource As String

    Dim strEntity As String
    Dim strDoctorID As String
    Dim strCustomerID As String
    Dim strCustomerName As String
    Dim strPDFFileName As String
    
    Dim strHtmlDirectory As String
    Dim strPdfExportDirectory As String
    Dim strPdfReaderDirectory As String
    
    Dim fso As FileSystemObject
    Dim strOutputFilePath As String
    Dim strHoldDocumentGroup As String
    Dim intSameDocumentGroupCounter As String
    Dim lPreviousPageCount As Long
    
    On Error Resume Next
    
    Set fso = New FileSystemObject

    

    '**************************************************************************
    '*** Determine if the documents should be exported in PDF format.
    
    
    strHtmlDirectory = frmImaging101ExportOptions.cmbHTMLsourceDir.Text
    If Right(Trim(strHtmlDirectory), 1) <> "\" Then
        strHtmlDirectory = strHtmlDirectory + "\"
    End If
    
    strPdfExportDirectory = frmImaging101ExportOptions.txtFullPathForPDFexport.Text
    If Right(Trim(strPdfExportDirectory), 1) <> "\" Then
        strPdfExportDirectory = strPdfExportDirectory + "\"
    End If
    
    strPdfReaderDirectory = frmImaging101ExportOptions.txtFullPathForPdfReaderFiles.Text
    If Right(Trim(strPdfReaderDirectory), 1) <> "\" Then
        strPdfReaderDirectory = strPdfReaderDirectory + "\"
    End If
    
    strDirectory = frmImaging101ExportOptions.txtFullPathForHTMLexport.Text
    If Right(Trim(strDirectory), 1) <> "\" Then
        strDirectory = strDirectory + "\"
    End If

    '*** Show the Message Form
    frmMessageForm.Show modal, Me
    
    
    If frmImaging101ExportOptions.chkExportToPDF Then
        
        frmMessageForm.txtMessage = "Preparing PDF Export Directory"
        DoEvents
        If Not funcDirectoryExists(strPdfExportDirectory) Then
            funcCreateDirectoryStructure strPdfExportDirectory
        End If
    
    End If
    
   
    
    If frmImaging101ExportOptions.chkExportToHTML Then
    
        '***************************************
        '*** COPY STATIC ROOT FILES
        
        frmMessageForm.txtMessage = "Preparing Root Directory"
        DoEvents
        If Not funcDirectoryExists(strDirectory) Then
            funcCreateDirectoryStructure strDirectory
        End If
        fso.CopyFile strHtmlDirectory & "\ROOT\STATIC\*.*", strDirectory, True
    
    End If
    
    
    
    '***************************************************************************
    '*** Get a List of All Document Groups to create the "Contents.HTM" file
    
    frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Getting List of Document Groups"
    DoEvents
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = ""
    rs.Source = rs.Source & "Select DISTINCT * "
    rs.Source = rs.Source & " FROM " & txtApplicationName
    rs.Source = rs.Source & txtFilterStatement
    rs.Source = rs.Source & " ORDER BY DocumentGroup "
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    '*** Make sure we got at least ONE record... Otherwise Cancel Export
    If rs.RecordCount < 1 Then
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "NO Items Returned... Export will NOT continue!"
        DoEvents
        MsgBox "NO Items Returned... Export will NOT continue!"
        GoTo SKIP_EXPORT
    End If
    
    rs.MoveFirst
    
    
    
    '************************************************************
    ' HOLD the Customer ID and Name for the PDF File Name
    strCustomerID = rs.Fields("CustomerID")
    strCustomerName = rs.Fields("CustomerName")
    
    
    
    If frmImaging101ExportOptions.chkExportToHTML Then
    
        '*************************************************************************
        '*************************************************************************
        '***  BUILD ROOT DIRECTORY  - BEGIN
        '***
        
        '***************************************
        '*** PREPARE ROOT Contents.htm
    
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Prepare .\ROOT\Contents.htm"
        DoEvents
        
        
        
        
    
        'Load Head & Foot... and clear the Body
        LoadText txtHead, strHtmlDirectory & "\ROOT\DYNAMIC\Contents_HEAD.txt"
        LoadText txtFoot, strHtmlDirectory & "\ROOT\DYNAMIC\Contents_FOOT.txt"
        txtBody.Text = ""
    
    End If
    
    
    '****************************************************************************
    '***  BEGIN DOCUMENT GROUP PROCESSING
    
    frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "*** BEGIN DOCUMENT GROUP PROCESSING."
    
    intSameDocumentGroupCounter = 1
    

    For intIndex = 0 To rs.RecordCount - 1
            
        '*********************************************************
        '***  Setup ROOT Detail Body line for Contents.htm
        
        txtDocumentGroup = rs!DocumentGroup
        
        strSubDirectory = txtDocumentGroup
    
        txtDocumentRECID = rs!DocumentRECID
        
        funcWriteToDebugLog Me.name, "txtDocumentRECID = " & txtDocumentRECID
        
        
        '*** If SAME Document Group... Do NOT Reset the Pagecount!
        '    let it increment to combine into a single subdirectory in funcExportSelectedSaveDocumentPages
        
        
        '*************************************************
        '*** EXPORT TO HTML ?
        
        If frmImaging101ExportOptions.chkExportToHTML Then
        
            If txtDocumentGroup <> strHoldDocumentGroup Then
                '*** INITIALIZE the Page Counter
                lPreviousPageCount = 0
        
                txtBody.Text = txtBody.Text & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & _
                                "<a href=""./" & txtDocumentGroup & "_DOCS.htm"" target=""main"">" & txtDocumentGroup & "</a><br><br>"
                        
                        
                 '***************************************
                 '*** Copy STATIC DOCGROUP Files
                 
                 
                 frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "*** Preparing " & strSubDirectory & " Directory"
                 DoEvents
                 funcCreateDirectoryStructure strDirectory & strSubDirectory
                 fso.CopyFile strHtmlDirectory & "\DOCGROUP\STATIC\*.*", strDirectory & strSubDirectory, True
        
             
                '***************************************
                '*** PREPARE DOCGROUP_DOCS.htm file
                
                frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Prepare .\ROOT\" & txtDocumentGroup & "_DOCS.htm"
                DoEvents
                txtConcat.Text = ""
                LoadText txtConcat, strHtmlDirectory & "\ROOT\DYNAMIC\DOCGROUP_DOCS.txt"
                txtConcat.Text = Replace(txtConcat.Text, "{{{DOCGROUP}}}", txtDocumentGroup)
                
                '*** SAVE the ROOT Contents.htm
                strOutputFilePath = strDirectory & txtDocumentGroup & "_DOCS.htm"
                If funcFileExists(strOutputFilePath) Then
                    Kill strOutputFilePath
                End If
                Open strOutputFilePath For Output As #1
                Print #1, txtConcat.Text
                Close #1
                         
             
                '***************************************
                '*** PREPARE DOCGROUP_TITLE.htm file
                
                frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Prepare .\ROOT\" & txtDocumentGroup & "_TITLE.htm"
                DoEvents
                txtConcat.Text = ""
                LoadText txtConcat, strHtmlDirectory & "\ROOT\DYNAMIC\DOCGROUP_TITLE.txt"
                txtConcat.Text = Replace(txtConcat.Text, "{{{DOCGROUP}}}", txtDocumentGroup)
                
                '*** SAVE the ROOT Contents.htm
                strOutputFilePath = strDirectory & txtDocumentGroup & "_TITLE.htm"
                If funcFileExists(strOutputFilePath) Then
                    Kill strOutputFilePath
                End If
                Open strOutputFilePath For Output As #1
                Print #1, txtConcat.Text
                Close #1
                
                
                
                '************************************************************
                '*** Load DOCGROUP Head & Foot... and clear the Body
                LoadText txtDocGroupHead, strHtmlDirectory & "\DOCGROUP\DYNAMIC\Contents_HEAD.txt"
                LoadText txtDocGroupFoot, strHtmlDirectory & "\DOCGROUP\DYNAMIC\Contents_FOOT.txt"
                txtDocGroupBody.Text = ""
            
            End If
                    
        
        End If
        
                
        
        
        '***************************************************
        '*** Import all document pages into the Viewer Control
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Get Related Images - " & txtDocumentRECID

        subExportSelectedGetImages
        DoEvents
        
        
        '**************************************************
        '*** Save combined document
        
        If frmImaging101ExportOptions.chkExportToHTML Then
            '*** Export the Document Pages for HTML
            frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Save Document Pages - " & txtDocumentRECID
            funcExportSelectedSaveDocumentPages strDirectory & txtDocumentGroup, strHtmlDirectory, strPdfExportDirectory, lPreviousPageCount
        End If
        
        '*** Export to PDF if the txtExportToPDF flag is set
        '    and the optBreakPDFbyDocgroup is selected
        If frmImaging101ExportOptions.chkExportToPDF = vbChecked _
        And frmImaging101ExportOptions.optBreakPDFbyDocgroup Then
            frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "SAVE PDF Document for " & txtDocumentGroup
            funcExportSelectedSaveDocument strPdfExportDirectory, frmClientSpicerControlForm.SpicerDoc1, txtDocumentGroup, "PDF"
        End If
        
        
        'Close the Document to clear
        frmClientSpicerControlForm.SpicerDoc1.CloseDocument (False)
        DoEvents
            
        
        '*************************************************
        '*** EXPORT TO HTML ?
        
        If frmImaging101ExportOptions.chkExportToHTML Then
             '**********************************
             '*** SAVE DOCGROUP Contents.htm
        
             txtConcat.Text = txtDocGroupHead.Text & txtDocGroupBody.Text & txtDocGroupFoot.Text
             
             frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Save .\" & txtDocumentGroup & "\Contents.htm"
             DoEvents
             strOutputFilePath = strDirectory & txtDocumentGroup & "\Contents.htm"
             If funcFileExists(strOutputFilePath) Then
                 Kill strOutputFilePath
             End If
             Open strOutputFilePath For Append As #1
             Print #1, txtConcat.Text
             Close #1
   
        End If
        
        '***************************************
        '*** GET the Next Record
        
        '*** Hold the Document Group for the next iteration
        ' Hold the current Document Group to see if there are more with SAME name
        strHoldDocumentGroup = txtDocumentGroup

        rs.MoveNext
    Next
    
    
    '*************************************************
    '*** EXPORT TO HTML ?
    
    If frmImaging101ExportOptions.chkExportToHTML Then
        
        '**********************************
        '*** SAVE ROOT Contents.htm
        
        txtConcat.Text = txtHead.Text & txtBody.Text & txtFoot.Text
        
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Save .\ROOT\Contents.htm"
        DoEvents
        strOutputFilePath = strDirectory & "Contents.htm"
        If funcFileExists(strOutputFilePath) Then
            Kill strOutputFilePath
        End If
        Open strOutputFilePath For Append As #1
        Print #1, txtConcat.Text
        Close #1
    
        '***
        '***  BUILD ROOT DIRECTORY  - END
        '*************************************************************************
        '*************************************************************************
    End If
    
        
    
    
    If frmImaging101ExportOptions.chkExportToPDF = vbChecked _
    And frmImaging101ExportOptions.optCombineIntoSinglePDF Then
        
        
        strPDFFileName = strCustomerName & "_" & strCustomerID
        
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "SAVE PDF File for " & strPDFFileName
        
        funcExportSelectedSaveDocument strPdfExportDirectory, frmClientSpicerControlForm.SpicerDoc2, strPDFFileName, "PDF"
        
        frmClientSpicerControlForm.SpicerDoc2.CloseDocument False
        
    End If
    
    
    '*************************************************
    '*** COPY the PDF Reader Files if requested
    
    If frmImaging101ExportOptions.chkExportToPDF = vbChecked _
    And frmImaging101ExportOptions.chkIncludePdfReader = vbChecked Then
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "*** COPY PDF READER FILES ***"
        DoEvents
        If Not funcDirectoryExists(strPdfExportDirectory) Then
            funcCreateDirectoryStructure strPdfExportDirectory
        End If
        'COPY the Entire contents of the PDF Reader Directory to the PDF Export Directory
        fso.CopyFile strPdfReaderDirectory & "*.*", strPdfExportDirectory, True
        fso.CopyFolder strPdfReaderDirectory & "*", strPdfExportDirectory, True
    End If
    
    
    '*************************************************
    '*** LAUNCH the CD Burn Directory
    
    frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "*** SHOW EXPORT DIRECTORIES ***"
    
    If frmImaging101ExportOptions.chkExportToHTML Then
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Show HTML Directory"
        Call shelldoc(strDirectory)
    End If
    
    If frmImaging101ExportOptions.chkExportToPDF = vbChecked Then
        frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "Show PDF Directory"
        Call shelldoc(strPdfExportDirectory)
    End If
    
    frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & ""
    frmMessageForm.txtMessage = frmMessageForm.txtMessage & vbCrLf & "*** EXPORT COMPLETE ***"
       
    Unload frmImaging101ExportOptions
    
    frmMessageForm.SetFocus
        
SKIP_EXPORT:

    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
    DoEvents
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
 
End Sub
    
Public Sub subSendToOutlook()

'  On Error GoTo ErrorHandle
    On Error GoTo ErrorExit
  
    '***2022-03-22  - Jacob - Changed form "Dim OutApp As Outlook.Application"
  '                               to "Dim OutApp As Object"
  '                               THIS stopped error "ACTIVEX Object NOT Registered."
  Dim OutApp As Object
  Dim OutMail As Object
   
  Set OutApp = CreateObject("Outlook.Application")
  Set OutMail = OutApp.CreateItem(0)
   
  
  With OutMail
'    .Recipients =
    .To = ""
    .CC = ""
    .BCC = ""
    .Subject = "Files from Imaging101 Document Imaging... "
    .Body = ""
    .Display 'or Send
    For i = 0 To lstExportedDocuments.ListCount - 1
        funcWriteToDebugLog Me.name, lstExportedDocuments.List(i)
        .Attachments.Add lstExportedDocuments.List(i)
    Next
  End With
   
       
       
    '********************************
    '*** DELETE the Temporary Files
'    On Error Resume Next
    
    Dim arrFilesToDelete() As String
    For i = 0 To lstExportedDocuments.ListCount - 1
        Kill lstExportedDocuments.List(i)
    Next
    
   
   
ErrorExit:
  Set OutMail = Nothing
  Set OutApp = Nothing
  
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "ERROR subSendToOutlook: " & Err.Number & "  Description: " & Err.Description
    End If

Exit Sub
   
'ErrorHandle:
'  Resume ErrorExit
End Sub


Public Sub subSendToSMTP()


    On Error GoTo ERROR_HANDLER
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    strCommandSource = "subSendToSMTP"
    
    funcWriteToDebugLog Me.name, txtDocumentRECID
    
    
    Dim txtAttachmentFileName As String
    
    'Prepare to send the attachment
'    frmMain.Show
    frmSMTPeMailForm.Show
    
    txtAttachmentFileName = ""
    For i = 0 To lstExportedDocuments.ListCount - 1
        If txtAttachmentFileName = "" Then
            txtAttachmentFileName = lstExportedDocuments.List(i)
        Else
            txtAttachmentFileName = txtAttachmentFileName & ";" & lstExportedDocuments.List(i)
        End If
    Next
    
    frmSMTPeMailForm.subStartup txtAttachmentFileName
    
    Screen.MousePointer = MousePointerConstants.vbDefault


Exit Sub

ERROR_HANDLER:
    

    bolErrorOccured = True
    strErrMsg = "subEmailDocument ERROR: Trace file = [" & strDestinationFile & "]  Error #: " & Err.Number & " - " & Err.Description
'    subWriteToAuditTraceFile txtTraceFilePath, dblDocumentRECID, dblDetailRECID, txtDestinationFilename, strErrMsg
    
    funcWriteToDebugLog Me.name, strErrMsg
'    funcWriteToSystemEventLog frmImaging101AutoExport.NTService1, svcMessageError, strErrMsg
    
    '*** Clean up/free resources used
    DoEvents
    'Close the Document to clear
    frmClientSpicerControlForm.SpicerDoc1.CloseDocument False
    DoEvents

    '*** Close the printer and clear the buffer.
    Printer.EndDoc
    
    '********************************
    '*** DELETE the Temporary File
    On Error Resume Next
    Kill txtAttachmentFileName
                    
    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub


Private Sub subCopyDetailRecordAndFiles(txtDocumentRECID As String)

    
    On Error GoTo ERROR_HANDLER
    
   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   
   frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT * " & _
                        " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                        " WHERE  DocumentRECID = " & txtDocumentRECID & _
                        " ORDER BY DetailOrder "

    txtActionBeforeError = "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    
    If frmImaging101Retrieve.ADOdcDetail.Recordset.EOF = True Then
        Exit Sub
    Else
        txtActionBeforeError = "ADodcDetail MoveFirst "
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
    End If
   
    
    
    frmImaging101Retrieve.ADOdcDetailDestination.ConnectionString = RegImaging101ConnectionString
    frmImaging101Retrieve.ADOdcDetailDestination.RecordSource = "SELECT * " & _
            " FROM " & frmImaging101MoveDocumentsBetweenApplications.txtDestinationApplicationName.Text & "_Detail " & _
            " WHERE 0 = 1"
            
    frmImaging101Retrieve.ADOdcDetailDestination.Refresh
    
    
    '*************************************************
    '*** Get the Root Directory to Store Objects
    funcWriteToDebugLog Me.name, "Get the Root Directory to Store Objects"
    RegRootDirToStoreObjects = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "RootDirectoryPathForImageArchive") & ""



    
    '****************************************************************************
    'Process All Detail Records for the Selected Document
    
    While Not frmImaging101Retrieve.ADOdcDetail.Recordset.EOF
    
        txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
        txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
        txtFullPathName = txtPathSubdirectory & "\" & txtFileName
        txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")

        txtDetailRecordCount = frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount
        
        frmImaging101MoveDocumentsBetweenApplications.ProgressBar1.Max = frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount
        
        
        '****************************************************************************
        '*** GET DETAIL SUBDIRECTORY STRUCTURE AND CREATE IT
        
        txtDocumentDirectoryStructure = RegRootDirToStoreObjects & "\" & _
                                        Format(CStr(frmImaging101MoveDocumentsBetweenApplications.txtDestinationApplicationRECID.Text), "0000") & _
                                        funcGetDetailSubdirectoryString(CDbl(txtDetailRECID))
        
        txtActionBeforeError = "Create Directory Structure: " & txtDocumentDirectoryStructure
        
        funcCreateDirectoryStructure txtDocumentDirectoryStructure & ""
                    
        '****************************************************************************
        '*** COPY THE FILE FROM SOURCE TO DESTINATION
        Dim strDestinationFilePath As String
        
        On Error Resume Next
        
        strDestinationFilePath = txtDocumentDirectoryStructure & "\" & txtFileName
        
        If funcFileExists(txtFullPathName) Then
            
            With New FileSystemObject
               .CopyFile txtFullPathName, strDestinationFilePath, True
            End With
                    
        End If
        
        '****************************************************************************
        '*** COPY THE ANNOTATION FILE(S) FROM SOURCE TO DESTINATION
        strFullDirectoryPathForAnnotationSource = funcGetFullPathForAnnotation(frmImaging101Search.txtApplicationRECID, CDbl(txtDetailRECID))
        strFullDirectoryPathForAnnotationDestination = funcGetFullPathForAnnotation(frmImaging101MoveDocumentsBetweenApplications.txtDestinationApplicationRECID.Text, CDbl(txtDetailRECID))
        
        txtActionBeforeError = "Create strFullDirectoryPathForAnnotationDestination Structure: " & strFullDirectoryPathForAnnotationDestination
         
        'Create the directory if needed.
        funcCreateDirectoryStructure strFullDirectoryPathForAnnotationDestination & ""
        
        Dim strAnnotationFileNameWildCard As String
        strAnnotationFileNameWildCard = strFullDirectoryPathForAnnotationSource & "\" & Left(txtFileName, InStr(txtFileName, ".") - 1) & "*.ANN"
        
        
        If funcFileExists(strAnnotationFileNameWildCard) Then
        
            With New FileSystemObject
               .CopyFile strAnnotationFileNameWildCard, strFullDirectoryPathForAnnotationDestination, True
            End With
        
        End If
        
        On Error GoTo ERROR_HANDLER
        
        '****************************************************************************
        'Create NEW Record in Destination Table
        
        frmImaging101Retrieve.ADOdcDetailDestination.Recordset.AddNew
        
        
        '****************************************************************************
        'Copy Field Values from Source to Destination
        
        For i = 0 To frmImaging101Retrieve.ADOdcDetail.Recordset.Fields.Count - 1
            
            If frmImaging101Retrieve.ADOdcDetailDestination.Recordset.Fields(i).name = "DetailSubdirectory" Then
            
                'Store New Destination DetailSubdirectory
                frmImaging101Retrieve.ADOdcDetailDestination.Recordset.Fields(i).Value _
                        = txtDocumentDirectoryStructure
                        
            Else
            
                frmImaging101Retrieve.ADOdcDetailDestination.Recordset.Fields(i).Value _
                        = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields(i).Value
                        
            End If
            
            DoEvents
        
        Next
        
            
        '****************************************************************************
        '*** SAVE THE NEW DESTINATION RECORD
        frmImaging101Retrieve.ADOdcDetailDestination.Recordset.Update
        
        frmImaging101MoveDocumentsBetweenApplications.ProgressBar1.Value = frmImaging101MoveDocumentsBetweenApplications.ProgressBar1.Value + 1


        
        '****************************************************************************
        '*** DELETE THE SOURCE FILES
        
        If funcFileExists(txtFullPathName) Then
            Kill txtFullPathName
        End If
        
        If funcFileExists(strAnnotationFileNameWildCard) Then
            Kill strAnnotationFileNameWildCard
        End If
        
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveNext
    
    Wend
    

    frmImaging101Retrieve.ADOdcDetailDestination.Recordset.Close

    frmImaging101MoveDocumentsBetweenApplications.ProgressBar1.Value = lblItemsSelected
    
Exit Sub

ERROR_HANDLER:

        funcWriteToDebugLog Me.name, "  ERROR subCopyDetailRecordAndFiles: " & Err.Number & " - " & Err.Description
        MsgBox "  ERROR subCopyDetailRecordAndFiles: " & Err.Number & " - " & Err.Description
        
End Sub


Private Sub subCopyDocumentRecord(txtDocumentRECID As String)

    
   frmImaging101Retrieve.Adodc1.ConnectionString = RegImaging101ConnectionString
   
   frmImaging101Retrieve.Adodc1.RecordSource = "SELECT * " & _
                        " FROM " & frmImaging101Search.txtApplicationName & _
                        " WHERE  DocumentRECID = " & txtDocumentRECID

    txtActionBeforeError = "ADOdc1 Refresh "
    frmImaging101Retrieve.Adodc1.Refresh
    
    If frmImaging101Retrieve.Adodc1.Recordset.EOF = True Then
        Exit Sub
    Else
        txtActionBeforeError = "ADOdc1 MoveFirst "
        frmImaging101Retrieve.Adodc1.Recordset.MoveFirst
    End If
    
    
    frmImaging101Retrieve.ADOdcDestination.ConnectionString = RegImaging101ConnectionString
    frmImaging101Retrieve.ADOdcDestination.RecordSource = "SELECT * " & _
            " FROM " & frmImaging101MoveDocumentsBetweenApplications.txtDestinationApplicationName.Text & _
            " WHERE 0 = 1"
            
    frmImaging101Retrieve.ADOdcDestination.Refresh
    
    
   'Get a List of ALL System (non-User-Defined) Fields
   frmImaging101Retrieve.ADOdcWork.ConnectionString = RegImaging101ConnectionString
   
   frmImaging101Retrieve.ADOdcWork.RecordSource = "SELECT column_name " & _
                                                    "FROM INFORMATION_SCHEMA.Columns " & _
                                                    "WHERE table_name = '" & frmImaging101Search.txtApplicationName & "'" & _
                                                    " AND COLUMN_NAME NOT IN " & _
                                                    " (SELECT FIELDNAME FROM I101Fields WHERE ApplicationRECID = " & frmImaging101Search.txtApplicationRECID & ")" & _
                                                    " ORDER BY column_name "
   
    txtActionBeforeError = "ADOdcWork Refresh "
    frmImaging101Retrieve.ADOdcWork.Refresh
    
    
    
    'Cycle through each Document - Should only be ONE record per Document
    'this is just a Double-check that a record exists.
    While Not frmImaging101Retrieve.Adodc1.Recordset.EOF
    
        
        'Create NEW Record in Destination Table
        frmImaging101Retrieve.ADOdcDestination.Recordset.AddNew
        
        Dim bolUserDefinedField As Boolean
        
        'Cycle through each Field in the Document Record
'        For i = 0 To frmImaging101Retrieve.ADOdc1.Recordset.Fields.count - 1
'
'            bolUserDefinedField = False
            
            For j = 0 To frmImaging101MoveDocumentsBetweenApplications.cmbSourceFieldNameForInput.UBound
                
                'Only handle Defined Fields...
                If Trim(frmImaging101MoveDocumentsBetweenApplications.txtSourceFieldName(j)) <> "" Then
'                And (frmImaging101Retrieve.ADOdc1.Recordset.Fields(i).name _
'                            = frmImaging101MoveDocumentsBetweenApplications.txtSourceFieldName(j)) Then

                    'See if field needs to be Truncated
                    Dim Dest As String
                    Dim Source As String
                    
                    Source = frmImaging101MoveDocumentsBetweenApplications.txtSourceFieldName(j)
                    Dest = frmImaging101MoveDocumentsBetweenApplications.txtDestinationFieldName(j)
                    
                    If UCase(frmImaging101MoveDocumentsBetweenApplications.lblSourceFieldType(j)) = "TEXT" _
                    And CInt(frmImaging101MoveDocumentsBetweenApplications.lblSourceFieldSize(j)) > CInt(frmImaging101MoveDocumentsBetweenApplications.lblDestinationFieldSize(Index)) _
                    Then
                        frmImaging101Retrieve.ADOdcDestination.Recordset.Fields(Dest).Value _
                                = Left(frmImaging101Retrieve.Adodc1.Recordset.Fields(Source).Value, CInt(frmImaging101MoveDocumentsBetweenApplications.lblDestinationFieldSize(j)))
                    Else
                        frmImaging101Retrieve.ADOdcDestination.Recordset.Fields(Dest).Value _
                                = frmImaging101Retrieve.Adodc1.Recordset.Fields(Source).Value
                    End If
                        
                    
                    bolUserDefinedField = False
                
                End If
            Next
                    
        
'        Next
            
        'Handle all System (non-User-Defined) fields
        Dim FieldName As String
        
        While Not ADOdcWork.Recordset.EOF
        
            FieldName = frmImaging101Retrieve.ADOdcWork.Recordset.Fields(0).Value
            
            If FieldName = "DocumentLocked" Then
                'Flag DESTINATION Document as MOVED IN (MI)
                frmImaging101Retrieve.ADOdcDestination.Recordset.Fields(FieldName).Value = "MI"
            Else
                frmImaging101Retrieve.ADOdcDestination.Recordset.Fields(FieldName).Value _
                            = frmImaging101Retrieve.Adodc1.Recordset.Fields(FieldName).Value
            End If
            
            ADOdcWork.Recordset.MoveNext
            
            
        Wend
            
        
        DoEvents
        frmImaging101Retrieve.Adodc1.Recordset.MoveNext
    
    Wend
    
    frmImaging101Retrieve.ADOdcDestination.Recordset.Update

    frmImaging101Retrieve.ADOdcDestination.Recordset.Close

End Sub


