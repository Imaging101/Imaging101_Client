VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmImaging101Search 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Retrieval Search Form - Imaging101"
   ClientHeight    =   4290
   ClientLeft      =   1320
   ClientTop       =   870
   ClientWidth     =   7470
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7470
   Begin VB.CheckBox chkDocumentCountOnly 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Count Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   195
      Left            =   2460
      TabIndex        =   43
      Top             =   855
      Width           =   1110
   End
   Begin VB.TextBox txtSearchTemplateWhereFreehand 
      DragIcon        =   "frmSearch.frx":0442
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
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Text            =   "frmSearch.frx":0884
      Top             =   3840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox cmbSearchTemplateList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   345
      ItemData        =   "frmSearch.frx":08A5
      Left            =   3720
      List            =   "frmSearch.frx":08A7
      TabIndex        =   38
      Top             =   1080
      Width           =   3630
   End
   Begin VB.TextBox txtFieldTableLookupOverridesDefault 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsForOutputOnly 
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
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
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Text            =   "frmSearch.frx":08A9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdFieldList 
      BackColor       =   &H00E0E9EF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      Picture         =   "frmSearch.frx":08B8
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3510
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ComboBox cboFieldSearchConditionLIST 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmSearch.frx":0E42
      Left            =   4920
      List            =   "frmSearch.frx":0E61
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboFieldSearchCondition 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      ItemData        =   "frmSearch.frx":0E95
      Left            =   2280
      List            =   "frmSearch.frx":0E97
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   1560
      Width           =   1335
   End
   Begin VB.FileListBox AnnotationFileListBox 
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
      Left            =   3120
      Pattern         =   "NOFILE.PATTERN"
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFieldSearchCondition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   360
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtMaxItemsToRetrieve 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFilterStatement 
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
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "frmSearch.frx":0E99
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFieldDropDown 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   6360
      Picture         =   "frmSearch.frx":0EAC
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   2640
   End
   Begin VB.TextBox txtFieldIsRequiredForCommit 
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
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldName 
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
      Index           =   0
      Left            =   4560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldSize 
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
      Index           =   0
      Left            =   4200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldType 
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
      Index           =   0
      Left            =   3840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchFieldsRECID 
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
      Index           =   0
      Left            =   1680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldsRECID 
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
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldDefaultValue 
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
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldHighValue 
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
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldLowValue 
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
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtFieldIsSticky 
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
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmbApplicationList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   360
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   3495
   End
   Begin MSMask.MaskEdBox mebIndexValues 
      Height          =   330
      HelpContextID   =   1
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   1575
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtWhereFreehand 
      DragIcon        =   "frmSearch.frx":1436
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
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "frmSearch.frx":1878
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   780
      Left            =   -75
      TabIndex        =   16
      Top             =   0
      Width           =   7365
      Begin VB.CommandButton cmdEditSearchTemplate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Template"
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
         Left            =   3285
         Picture         =   "frmSearch.frx":1889
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAdvanced 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ad&vanced"
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
         Left            =   2460
         Picture         =   "frmSearch.frx":2153
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
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
         Left            =   1635
         Picture         =   "frmSearch.frx":2A1D
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear"
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
         Left            =   810
         Picture         =   "frmSearch.frx":32E7
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find"
         Default         =   -1  'True
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
         Left            =   0
         Picture         =   "frmSearch.frx":3729
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdPackage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Package"
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
         Left            =   3600
         Picture         =   "frmSearch.frx":3FF3
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Help"
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
         Left            =   4080
         Picture         =   "frmSearch.frx":4CBD
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox picImaging101Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5910
         Picture         =   "frmSearch.frx":5587
         ScaleHeight     =   405
         ScaleWidth      =   1440
         TabIndex        =   21
         Top             =   30
         Width           =   1440
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5910
         TabIndex        =   22
         Top             =   420
         Width           =   1245
      End
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4770
      TabIndex        =   14
      Top             =   3630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtApplicationName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3930
      TabIndex        =   15
      Top             =   3630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSearchTemplateRECID 
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
      Left            =   5760
      TabIndex        =   41
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSelectSearchTemplate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Search Template"
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
      Left            =   3720
      TabIndex        =   39
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblWhereFreehand 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type search conditions (Example: DocumentIndexDate like 'feb 25 2013%'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   5655
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
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
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
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmImaging101Search"
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
    Dim bolErrorOccured As Boolean
    Dim intCurrentFieldIndex As Integer
    Dim intFieldSpacing As Integer

    
    
    

Public Sub cmbApplicationList_Click()
    
    'Store the selected application
    funcGetSetUserSettings "SET", "ApplicationSearchForm", cmbApplicationList
    
    
    
    ' Get the Application to Commit Batches to
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    

    
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
        
    '*** 2022-07-21 - Jacob - Added LogOpenedDocuments.  Logic is in Form frmImaging101Retrieve
    rs.Source = "Select ApplicationRECID,ApplicationName, MaxItemsToRetrieve, EnableSearchTemplates, LogOpenedDocuments from I101Applications WHERE ApplicationName= '" & cmbApplicationList.Text & "'"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If Not (rs.EOF Or rs.BOF) Then
        txtApplicationRECID = rs!ApplicationRECID
        txtApplicationName = rs!ApplicationName
        txtMaxItemsToRetrieve = rs!MaxItemsToRetrieve
        bolEnableSearchTemplates = rs!EnableSearchTemplates
        '*** 2022-07-21 - Jacob - Added boolean to Log Opened Documents.  Logic is in Form frmImaging101Retrieve
        bolLogOpenedDocuments = rs!LogOpenedDocuments
    End If
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    
    
    '*************************************************
    '*** See if Search Templates are Enabled
    
    '*** 2021-10-11 - Jacob - DISABLED showing of cmdSearchTemplateList if AdminSystem rights
    'If bolEnableSearchTemplates = True Or gsecRightsAdminSystem = vbChecked Then
    If bolEnableSearchTemplates = True Then

        lblSelectSearchTemplate.Visible = True
        cmbSearchTemplateList.Visible = True
    Else
        lblSelectSearchTemplate.Visible = False
        cmbSearchTemplateList.Visible = False
    End If
    
    '*************************************************
    '*** CLEAR FREEHAND FIELDS
    txtWhereFreehand = ""
    txtSearchTemplateWhereFreehand = ""
    
    '*************************************************
    '*** Get the Root Directory to Store Objects
    funcWriteToDebugLog Me.name, "Get the Root Directory to Store Objects"
    RegRootDirToStoreObjects = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "RootDirectoryPathForImageArchive") & ""

    
    '***************************************************************
    '***  GET SECURITY RIGHTS
    
    funcGetSecurityRights gsecSecurityRECID, txtApplicationRECID
    
    
    If bolAIM_Command_AddFile = True Then
                '*** CHECK IF USER HAS RIGHTS TO FILE VIA I101FILER
            Dim bolRightsFileDocsViaI101FILER As Boolean
            bolRightsFileDocsViaI101FILER = funcGetFieldFromDB(RegImaging101ConnectionString, "I101SecurityRoleApp", "SecurityRECID = " & gsecSecurityRECID & "AND ApplicationRECID = " & txtApplicationRECID, "RightsFileDocsViaI101FILER")
            If bolRightsFileDocsViaI101FILER = False Then
                Me.cmdSave.Enabled = False
                result = funcQuickMessage("SHOW", "SORRY!  You do NOT have Security Rights to File Documents via I101FILER." & vbCrLf & "For Application [" & cmbApplicationList & "]" & vbCrLf & "Please contact a System Administrator.")
                Exit Sub
            End If
    End If

    
    'This timer is simply to bypass a VB Error:  "Unable to unload within this context (Error 365)"
    ' attempting to "Destroy" the fields when switching Applications
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vbenlr98/html/vamsgldcantunloadhere.asp
    'apparently VB won't let you unload certain objects in certain contexts like the "_Click" event
    ' or any event or Sub that this event calls!!!
    ' We simply enable this timer to kick in 1/10th of a second.
    ' the Timer1 sub disables the timer and does what we need it to do.
    Timer1.Interval = 100
    Timer1.Enabled = True
    

    
End Sub




Private Sub cmbApplicationList_KeyPress(KeyAscii As Integer)
 
    '*** Don't want user to key in an invalid value
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
    
End Sub




Private Sub cmbSearchTemplateList_Click()

    'Store the selected Search Template
    funcGetSetUserSettings "SET", "SearchTemplate", cmbSearchTemplateList
    
'    '*** 2020-07-07 - Jacob - Added Document Count option
'    If cmbSearchTemplateList.Text <> "*DOCUMENT COUNT ONLY*" Then
    
        txtSearchTemplateRECID = cmbSearchTemplateList.ItemData(cmbSearchTemplateList.ListIndex)
        
        txtSearchTemplateWhereFreehand = funcGetFieldFromDB(RegImaging101ConnectionString, "I101SearchTemplates", "SearchTemplateRECID = " & txtSearchTemplateRECID, "WhereFreehand")

'    End If
    
End Sub

Private Sub cmbSearchTemplateList_DropDown()

    If Trim(cmbApplicationList) = "" Then
        result = MsgBox("Please select and APPLICATION and try again.", vbInformation, "No Application Selected")
        Exit Sub
    End If


    '******************************************************************
    '*** LOAD SEARCH TEMPLATE LIST DROP-DOWN
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
        
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = ""
    rs.Source = rs.Source & "SELECT DISTINCT  I101SearchTemplates.SearchTemplateName, I101SearchTemplates.SearchTemplateRECID"
    rs.Source = rs.Source & " FROM I101SearchTemplates, I101SearchTemplateUsers "
    rs.Source = rs.Source & " WHERE ApplicationRECID = " & txtApplicationRECID
    rs.Source = rs.Source & " AND      SecurityRECID = " & gsecSecurityRECID
    rs.Source = rs.Source & " AND       I101SearchTemplates.SearchTemplateRECID = I101SearchTemplateUsers.SearchTemplateRECID"
   
    rs.Source = rs.Source & " ORDER BY SearchTemplateName"

    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If rs.RecordCount > 0 Then
    
        rs.MoveFirst
        
        cmbSearchTemplateList.Clear
        
        For intIndex = 0 To rs.RecordCount - 1
            cmbSearchTemplateList.AddItem rs.Fields!SearchTemplateName
            cmbSearchTemplateList.ItemData(intIndex) = rs.Fields!SearchTemplateRECID
            rs.MoveNext
        Next
    
'        '*** 2020-07-07 - Jacob - Added Document Count option
'        cmbSearchTemplateList.AddItem "*DOCUMENT COUNT ONLY*"
'        cmbSearchTemplateList.ItemData(intIndex) = "999999999"

    
    
    End If

    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
End Sub

Private Sub cmbSearchTemplateList_KeyPress(KeyAscii As Integer)

    '*** Don't want user to key in an invalid value
        'Cancel the Keypress by setting the KeyAscii to Zero (0)
        KeyAscii = 0
        
End Sub

Public Sub cmdAdvanced_Click()

        Dim intWhereFreehandHeight As Integer
        intWhereFreehandHeight = 800
        
        If txtWhereFreehand.Visible = False Then
            Me.Height = Me.Height + intWhereFreehandHeight + 150
            txtWhereFreehand.Top = Me.ScaleHeight - intWhereFreehandHeight
            txtWhereFreehand.Left = Me.ScaleLeft
            txtWhereFreehand.width = Me.ScaleWidth
            txtWhereFreehand.Height = intWhereFreehandHeight
            txtWhereFreehand.Text = ""
            txtWhereFreehand.Visible = True
            
            lblWhereFreehand.Top = txtWhereFreehand.Top - lblWhereFreehand.Height
            lblWhereFreehand.Left = txtWhereFreehand.Left
            lblWhereFreehand.Visible = True
            
            cmdFieldList.Top = txtWhereFreehand.Top + (txtWhereFreehand.Height / 2) - (cmdFieldList.Height / 2)
            cmdFieldList.Left = Me.ScaleWidth - cmdFieldList.width
            cmdFieldList.Visible = True
            
        Else
            txtWhereFreehand.Text = ""
            txtWhereFreehand.Visible = False
            Me.Height = Me.Height - intWhereFreehandHeight - 150
            
            lblWhereFreehand.Visible = False
            
            cmdFieldList.Visible = False
            
        End If


End Sub

Public Sub cmdClear_Click()

    Dim strHoldFieldMask As String
    
    For intIndex = 0 To lblFieldDescription.Count - 1
        strHoldFieldMask = mebIndexValues(intIndex).Mask
        mebIndexValues(intIndex).Mask = ""
        mebIndexValues(intIndex).Text = ""
        txtIndexValues(intIndex).Text = ""
        mebIndexValues(intIndex).Mask = strHoldFieldMask
        txtWhereFreehand.Text = ""
        
        '7/4/2017 - Jacob - Added Set Default Values upon Clear Fields
        '8/20/2017 - Jacob - Added Handle AddFile scenario to ONLY set defaults if in Add-File mode.
         If bolAIM_Command_AddFile Then
            subSetDefaultFieldValues (intIndex)
        End If
        
    Next
        
    Call cmdSetFocusOnFirstField
    
End Sub

Public Sub cmdSetFocusOnFirstField()

    '5/11/2015 - Jacob - Set focus on the first input field
    '7/16/2015 - Jacob - Added the On-Error Resume Next
    '                             to avoid "Runtime error 5 - Invalid Procedure call"
    '                             if the field was set to "Prevent Manual Indexing"
    '7/19/2017 - Jacob   Added check for txtFieldIsForOutputOnly ("Prevent Manual Indexing")
    On Error Resume Next
    
    If txtFieldIsForOutputOnly(0) = 1 And bolAIM_Command_AddFile = False Then
    
        If txtFieldType(0) = "LongText" Then
            txtIndexValues(0).SetFocus
            txtIndexValues(0).SelStart = 0
            txtIndexValues(0).SelLength = 4000
        Else
            mebIndexValues(0).SetFocus
            mebIndexValues(0).SelStart = 0
            mebIndexValues(0).SelLength = 4000
        End If
        
    End If
 
    DoEvents

End Sub

Private Sub cmdEditSearchTemplate_Click()

    frmImaging101SearchTemplate.Show

End Sub

Private Sub cmdFieldDropDown_Click(Index As Integer)

    frmDropDownList.Caption = txtFieldName(Index)
    
    If txtFieldType(Index) = "LongText" Then
        frmDropDownList.funcPopulateDropDown cmbApplicationList, txtFieldName(Index), txtIndexValues(Index)
    Else
        frmDropDownList.funcPopulateDropDown cmbApplicationList, txtFieldName(Index), mebIndexValues(Index)
    End If
    
    'Show the list in as Modal AFTER populating it.  Otherwise it stops processing and won't populate.
    If funcIsFormLoaded2("frmDropDownList") Then
         frmDropDownList.Hide
    End If
    frmDropDownList.Show vbModal, Me
    

End Sub



Private Sub cmdFieldList_Click()

    '*** POPULATE the FieldList Drop-Down
    
'    select column_name from INFORMATION_SCHEMA.COLUMNS
'    where table_name = 'CLIENT_FILES'
'    order by column_name

    frmDropDownList.Show
    frmDropDownList.Caption = "Field List"
    
    'CLEAR the List
    frmDropDownList.lstDropDownsList.Clear
    
    'ADD Special Fields - surrounded with Single-Quotes
    frmDropDownList.lstDropDownsList.AddItem "'{LoggedInUserID}'"
    frmDropDownList.lstDropDownsList.AddItem "'{CurrentDate}'"
    
    'FILL the List
    funcFillList frmDropDownList.lstDropDownsList, RegImaging101ConnectionString, "INFORMATION_SCHEMA.COLUMNS", "column_name", "table_name = '" & cmbApplicationList & "'", False, False
    
    
    frmImaging101Search.SetFocus
    frmImaging101Search.txtWhereFreehand.SetFocus
    
End Sub

Public Sub cmdFind_Click()
        
        'CHECK IF SEARCH TEMPLATES ARE ENABLED
        If bolEnableSearchTemplates = True And Trim(cmbSearchTemplateList) = "" Then
            result = MsgBox("Please select a Search Template and try again.", vbInformation, "No Search Template Selected")
            Exit Sub
        End If
            
        
        'Make SURE we execute the mebIndexValues GotFocus event to fill the THRU date it appropriate.
        If intCurrentFieldIndex < mebIndexValues.Count - 2 Then
            mebIndexValues_GotFocus (intCurrentFieldIndex) + 1
        End If
        
        
        'Check if User is Flagged to RESET the Viewer images, etc.
        If gsecViewResetImagesOnFind = vbChecked Then
            '*** Unload ALL open images
            If funcIsFormLoaded2("MainMDIForm") Then
                Dim i As Integer
                i = UBound(arrDisplayedPagesRetrieve)
                
                If i > 0 Then
                    For i = 1 To i
                        Unload MainMDIForm.ActiveForm
                    Next
                End If
            End If
                
            '*** Unload Thumbnail form
            If funcIsFormLoaded2("frmThumb") Then
                Unload frmThumb
            End If
        End If
            
            
        '*** Build Select Statement
        subBuildSelectStatement
        
        If chkDocumentCountOnly = vbChecked Then
                Exit Sub
        End If
        
        If bolErrorOccured Then
            Exit Sub
        End If
        
        '*** Show the Retrieve form
        frmImaging101Retrieve.Show
        frmImaging101Retrieve.WindowState = vbNormal
        
        frmImaging101Retrieve.ListView1.ListItems.Clear
        frmImaging101Retrieve.subPopulateListview
        
        '*** Pass Parameters to the Retrieval Form
        frmImaging101Retrieve.txtApplicationName.Text = cmbApplicationList.Text
        frmImaging101Retrieve.txtFilterStatement.Text = txtFilterStatement
        
        '*** Set the FOCUS to the Retrieval Window
        frmImaging101Retrieve.SetFocus
        
End Sub



Private Sub cmdHelp_Click()

    'Launch the document
    Call shelldoc(".\Imaging101Help.mht")

End Sub

Private Sub cmdPackage_Click()

    frmImaging101Package.Show modal, Me
    

End Sub

Private Sub cmdSave_Click()


    
    On Error GoTo ERROR_HANDLER
'    On Error GoTo 0
        
    'Disable the SAVE button so it can't be clicked again
    Me.cmdSave.Enabled = False
    
    
    
    '****************************************************************************
    '*** Only check for Fields Required or Valid if Not flagged for Skip
    
    Dim bolFieldsRequiredButEmpty As Boolean
    
     For intIndex = 0 To mebIndexValues.Count - 1
         ' If ANY Field is flagged as "Required" but is Empty - Skip this record
         If (txtFieldIsRequiredForCommit(intIndex) = "1") And (mebIndexValues(intIndex).Text = "" And txtIndexValues(intIndex).Text = "" And cboFieldSearchCondition.item(intIndex).Text <> "<=") Then
              bolFieldsRequiredButEmpty = True
              Exit For
         End If
         
          '*** VALIDATE DATE FIELD!
          funcValidateDate intIndex
          If blnDateError = True Then
              bolFieldsRequiredButEmpty = True

              Exit For
          End If
      
     Next
     
     If bolFieldsRequiredButEmpty = True Then
        MsgBox "Please fill in all REQUIRED (red) fields and try the save again.", vbInformation, "Fields Required but Empty"
        Me.cmdSave.Enabled = True
        Exit Sub
     End If
       
    '*** END check for Fields Required or Valid if Not flagged for Skip
    '****************************************************************************

    
    '****************************************************************************
    '*** Validate SAVE Request
            
    result = MsgBox("Are you sure you wish to SAVE this File?", vbYesNo, "Save File?")
    
    If result <> vbYes Then
        Me.cmdSave.Enabled = True
        Exit Sub
    End If

    
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "+++++++++++++++++++++++++++++++++++++++++++"
    funcWriteToDebugLog Me.name, "ENTERING: SAVE File to Imaging101"
        
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
    Dim txtSourceFilename As String
    
    Dim intPositionOfLastPeriod As Integer
    
    Dim intCopyFileRetryCount As Integer
    
    Dim connImaging101 As ADODB.Connection
    Dim rsImaging101Document As ADODB.Recordset
    Dim rsImaging101DocumentDetail As ADODB.Recordset
    
'    Dim bolErrorOccured As Boolean

    '*** Check for Annotations
    MainMDIForm.ActiveForm.subAnnotationLayerSaveCheck
    
    '*** GET SOURCE FILE NAME
    Dim strSourceFile As String
    strSourceFile = MainMDIForm.ActiveForm.txtPageFileName
    
        '*** 2022-07-18 - Jacob - Determine txtSourceFilePath to Kill TEMP Files
        Dim txtSourceFilePath As String
        Dim intPositionOfLastBackslash As Integer
        
        intPositionOfLastBackslash = InStrRev(strSourceFile, "\")
        txtSourceFilePath = Left(strSourceFile, intPositionOfLastBackslash)
        

    '*** SET VIEWER
    MainMDIForm.ActiveForm.txtChildFormMessage.Visible = True
    MainMDIForm.ActiveForm.txtChildFormMessage.Text = "SAVING DOCUMENT!"
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "SAVING DOCUMENT"
    MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
    MainMDIForm.ActiveForm.lstPageList.Visible = False
    MainMDIForm.ActiveForm.SpicerView1.Visible = False
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).AutoSize = sbrContents
    
    

    
    
               
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
                    '* BEGIN TRANSACTIONS
RETRY_TRANSACTION:
                    
                    connImaging101.BeginTrans
                    
                    
                    '****************************************************************************
                    '*** Insert DOCUMENT record into SQL
                    '***    only if the Index Values are Different from the previous record
                    '*** OTHERWISE FIND THE EXISTING DOCUMENT RECORD
                    
'                    If txtValuesList <> txtValuesListHold Then
                        
                        txtActionBeforeError = "Get Next Control DocumentRECID"
                            
                        txtDocumentRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "DocumentRECID")
                        
                        '*** Establish the Imaging101 Document Recordset
                        Set rsImaging101Document = New ADODB.Recordset
                        
                        ' Open the Imaging Application Document Table
                        txtActionBeforeError = "Open " & txtApplicationName
                        rsImaging101Document.Open txtApplicationName.Text, connImaging101, adOpenDynamic, adLockOptimistic, adCmdTable

                        'Add NEW Imaging101 Doocument Record
                        txtActionBeforeError = "Add New Document Record " & txtApplicationName
                        rsImaging101Document.AddNew
                        
                        'Set Field Values
                        txtActionBeforeError = "Assign Document System Field Values " & txtApplicationName
                        
                        rsImaging101Document.Fields("DocumentRECID") = txtDocumentRECID
                        rsImaging101Document.Fields("DocumentScanUserID") = gsecUserID
                        rsImaging101Document.Fields("DocumentScanDate") = Now()
                        rsImaging101Document.Fields("DocumentIndexUserID") = gsecUserID
                        rsImaging101Document.Fields("DocumentIndexDate") = Now()
                        rsImaging101Document.Fields("DocumentCommitUserID") = gsecUserID
                        rsImaging101Document.Fields("DocumentCommitDate") = Now()
                        rsImaging101Document.Fields("DocumentBatchRECID") = 0
                        rsImaging101Document.Fields("DocumentBatchName") = "I101Filer"
'                        rsImaging101Document.Fields("BatchBoxNumber") = rsImaging101Batch.Fields("BatchBoxNumber")
                        
                        rsImaging101Document.Fields("DocumentPages") = 0
                        rsImaging101Document.Fields("DocumentImages") = 0
                        rsImaging101Document.Fields("DocumentNotes") = ""
                        rsImaging101Document.Fields("DocumentLockedBy") = ""
                        rsImaging101Document.Fields("DocumentLockedDate") = Null
                        rsImaging101Document.Fields("DocumentLockExpDate") = Null

'                        '*** Update Application Fields
'                        For intIndex = 0 To lblFieldDescription.count - 1
'                            txtActionBeforeError = "Assign Document User-defined Field Values " & vbCrLf & _
'                                                    "Application = " & txtApplicationName & vbCrLf & _
'                                                    "FieldName = " & txtFieldName(intIndex).Text & vbCrLf & _
'                                                    "Value = " & mebIndexValues(intIndex).Text
'
'                            rsImaging101Document.Fields("" & txtFieldName(intIndex).Text & "") = mebIndexValues(intIndex).Text
'                        Next



                        '**************************************************
                        '*** Cycle through Field Values
                        For intIndex = 0 To mebIndexValues.Count - 1
                                
                                ' Get Field Index by comparing the Form fieldname with the DB result set fieldname
                                For intFieldIndex = 0 To rsImaging101Document.Fields.Count - 1
                                    If rsImaging101Document.Fields(intFieldIndex).name = txtFieldName(intIndex) Then
                                        Exit For
                                    End If
                                Next
                                
                                
                                'If field is LongText, use the txtIndexValues field instead
                                If txtFieldType(intIndex) = "LongText" Then
                                        rsImaging101Document.Fields(intFieldIndex) = Left(Trim(txtIndexValues(intIndex).Text), txtFieldSize(intIndex))
                                Else
                                    If (Not IsNull(mebIndexValues(intIndex))) _
                                            And (mebIndexValues(intIndex) <> "") _
                                            Then
                                            'Save the MASKED EDIT Control Value
                                        
                                            ' If the field is empty, Set to Null value
                                            If Trim(mebIndexValues(intIndex)) = "" Then
                                                rsImaging101Document.Fields(intFieldIndex) = Null
                                            
                                            ' If the field is Date, Format as Date value... but ONLY for the Main field, not the "Thru" clone
                                            '7/13/2017 - Jacob added check for "=" and "Contains" to make sure is doesn't skip these values.
                                            ElseIf txtFieldType(intIndex) = "Date" Then
                                                If cboFieldSearchCondition(intIndex) = ">=" _
                                                Or cboFieldSearchCondition(intIndex) = "=" _
                                                Or cboFieldSearchCondition(intIndex) = "Contains" Then
                                                    rsImaging101Document.Fields(intFieldIndex) = Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format)
                                                    If Err.Number <> 0 Then
                                                        mebIndexValues(intIndex).SetFocus
                                                        mebIndexValues(intIndex).ForeColor = vbRed
                                                    End If
'                                                    '7/14/2017 - Jacob - Disabled this clone date feature since using Filer we don't clone Date fields.
                                                    'Set the "clone date" field equal to the main date field
'                                                    mebIndexValues(intIndex + 1).Text = mebIndexValues(intIndex).Text
                                                End If
                                            
                                            ' If the field is Numeric, convert to Long
                                            ElseIf txtFieldType(intIndex) = "Numeric" Then
                                                rsImaging101Document.Fields(intFieldIndex) = CDbl(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
                                            
                                            ' If the field is Currency, convert to Currency
                                            ElseIf txtFieldType(intIndex) = "Currency" Then
                                                rsImaging101Document.Fields(intFieldIndex) = CCur(Format(mebIndexValues(intIndex).FormattedText, mebIndexValues(intIndex).Format))
                                            
                                            Else
                                                'Use the FieldMask value to format the data
                                                '  also Trim it save only up to the defined length of the field
                                                If Trim(mebIndexValues(intIndex).Format) = "" Then
                                                    'NO Format -- Simply Trim and Clip the Field
                                                    rsImaging101Document.Fields(intFieldIndex) = Left(Trim(mebIndexValues(intIndex).Text), txtFieldSize(intIndex))
                                                Else
                                                    rsImaging101Document.Fields(intFieldIndex) = Left(Trim(Format(mebIndexValues(intIndex).Text, mebIndexValues(intIndex).Format)), txtFieldSize(intIndex))
                                                End If
                                            End If
                                    End If
                                        
                                End If 'LongText
            
                                '* If field flagged as questionable, flag as red
                                If mebIndexValues(intIndex).Text = txtQuestionable Then
                                    mebIndexValues(intIndex).Font.Bold = True
                                    mebIndexValues(intIndex).ForeColor = vbRed
                                Else
                                    mebIndexValues(intIndex).Font.Bold = False
                                    mebIndexValues(intIndex).ForeColor = vbNormal
                                End If
                            
                          
                        Next
                        '**************************************************








                        'Reset Detail Order counter
                        intDetailOrder = 0
                        intPageCount = 0
                        'Hold List of Index values to compare with next record
                        txtValuesListHold = txtValuesList
                    
                        
                    '*************************************************
                    '*** Get the Root Directory to Store Objects
                    funcWriteToDebugLog Me.name, "Get the Root Directory to Store Objects"
                    RegRootDirToStoreObjects = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & frmImaging101Search.txtApplicationRECID, "RootDirectoryPathForImageArchive") & ""
                            
                    
                    '****************************************************************************
                    '****************************************************************************
                    '*** GET DETAIL SUBDIRECTORY STRUCTURE AND CREATE IT
                    
                    txtActionBeforeError = "Get Next Control DetailRECID " & txtApplicationName
                    txtDetailRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "DetailRECID")
                    
                    txtDocumentDirectoryStructure = RegRootDirToStoreObjects & "\" & _
                                                    Format(CStr(txtApplicationRECID), "0000") & _
                                                    funcGetDetailSubdirectoryString(txtDetailRECID)
                    
                    txtActionBeforeError = "Create Directory Structure: " & txtDocumentDirectoryStructure
                    funcCreateDirectoryStructure txtDocumentDirectoryStructure & ""
                    
                    
                    
                    '****************************************************************************
                    '*** DEFINE DESTINATION FILE NAME AND FILE TYPE
                
                    intPositionOfLastPeriod = InStrRev(strSourceFile, ".")
                    
                    If intPositionOfLastPeriod = 0 Then
                        txtDestinationFileType = ""
                    Else
                        txtDestinationFileType = Trim(UCase(Right(strSourceFile, Len(strSourceFile) - intPositionOfLastPeriod)))
                    End If
                    
                    
                    txtDestinationFilename = Format(CStr(txtDetailRECID), "0000000000") & "." & txtDestinationFileType
                    
                    
                    
                    '****************************************************************************
                    '*** Insert DOCUMENT DETAIL record into SQL
                
                    'Increment order counter - the counter gets reset every time a new Batch is created
                    intDetailOrder = intDetailOrder + 1

                    MainMDIForm.ActiveForm.StatusBar1.Panels(2).Text = "Page " & intDetailOrder


                    '*** Establish the Imaging101 Document Recordset
                    Set rsImaging101DocumentDetail = New ADODB.Recordset
                    
                    ' Open the Imaging Application Document Table
                    txtActionBeforeError = "Open " & txtApplicationName & "_Detail"
                    rsImaging101DocumentDetail.Open txtApplicationName.Text & "_Detail", connImaging101, adOpenDynamic, adLockOptimistic, adCmdTable

                    'Add NEW Imaging101 Doocument Record
                    txtActionBeforeError = "Add New Document DETAIL Record " & txtApplicationName & "_Detail ( " & txtDetailRECID & " )"
                    rsImaging101DocumentDetail.AddNew
                        
                    rsImaging101DocumentDetail.Fields("DetailRECID") = txtDetailRECID
                    rsImaging101DocumentDetail.Fields("DocumentRECID") = txtDocumentRECID
                    rsImaging101DocumentDetail.Fields("DetailOrder") = intDetailOrder
                    rsImaging101DocumentDetail.Fields("DetailCreatedDate") = Now()
                    rsImaging101DocumentDetail.Fields("DetailSubdirectory") = txtDocumentDirectoryStructure
                    rsImaging101DocumentDetail.Fields("DetailFileName") = txtDestinationFilename
                    rsImaging101DocumentDetail.Fields("DetailFileType") = txtDestinationFileType
                    rsImaging101DocumentDetail.Fields("DetailRotation") = MainMDIForm.ActiveForm.funcGetPageRotation(intDetailOrder)
                    
                    
                    '****************************************************************************
                    '*** UPDATE the DOCUMENT record with Page & Image Counts
                    
                    rsImaging101Document.Fields("DocumentPages") = intPageCount
                    rsImaging101Document.Fields("DocumentImages") = intDetailOrder
                    
                    
                                    
                    '*********************************************************
                    '*** CLOSE THE IMAGE IN THE VIEWER
                    '*** TO PREVENT "Runtime Error 75: File/Path Access Error"
                    '*** WHEN PROCESSING SINGLE-PAGE PDF's
                    '*** WHICH SEEM TO REMAIN "IN-USE" WHEN OPEN.
                    'Close the document to release it
                    
                    funcWriteToDebugLog Me.name, "MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument"
                    
                    MainMDIForm.ActiveForm.SpicerDoc1.CloseDocument False
                    
                    
                    
                    '****************************************************************************
                    '*** COPY the file to the Storage Destination
                    
                    Dim strDestinationFile As String
                
                    strDestinationFile = txtDocumentDirectoryStructure & "\" & txtDestinationFilename
                    
'                    txtActionBeforeError = "FileCopy [" & strSourceFile & "] TO [" & strDestinationFile & "]"
'                    FileCopy strSourceFile, strDestinationFile

                    intCopyFileRetryCount = 1
                    
                    txtActionBeforeError = "CopyFile [" & strSourceFile & "] TO [" & strDestinationFile & "]"
                    
                    
                    With New FileSystemObject
                       .CopyFile strSourceFile, strDestinationFile, True
                    End With
                    
                    
                    '****************************************************************************
                    '*** Now COPY the Annotation files, if any, to the Storage Destination
                    Dim strSourceAnnotationFile As String
                    Dim strSourceAnnotationDirectory As String
                    Dim strDestinationAnnotationFile As String
                    Dim strPageNumber As String


                    Dim strLocalTempDir As String
                    strLocalTempDir = funcGetTempDir()

                    intPositionOfLastPeriod = InStrRev(strSourceFile, ".")

'                    strSourceAnnotationDirectory = Left(strSourceFile, InStrRev(strSourceFile, "\") - 1) & "\Annotations"
                    
                    strSourceAnnotationDirectory = strLocalTempDir & "Imaging101\Annotations"
                    
                    funcCreateDirectoryStructure strSourceAnnotationDirectory
                   
                    AnnotationFileListBox.Path = strSourceAnnotationDirectory
                    
                    AnnotationFileListBox.Pattern = "Annotation" & "_*.ANN"

                    DoEvents

                    If AnnotationFileListBox.ListCount > 0 Then
                        For dblAnnotationFileIndex = 0 To AnnotationFileListBox.ListCount - 1

                            AnnotationFileListBox.Selected(dblAnnotationFileIndex) = True
                            strSourceAnnotationFile = strSourceAnnotationDirectory & "\" & AnnotationFileListBox.FileName

                            ' Build the Annotation FilePath
                            strFullDirectoryPathForAnnotation = funcGetFullPathForAnnotation(txtApplicationRECID, txtDetailRECID)
                            txtActionBeforeError = "Create Directory Structure: " & txtDocumentDirectoryStructure

                            'Extract the Page # from the Source Filename
                            strPageNumber = Mid(AnnotationFileListBox.FileName, InStrRev(AnnotationFileListBox.FileName, "_") + 1, 6)

                            'Create the directory if needed.
                            funcCreateDirectoryStructure strFullDirectoryPathForAnnotation & ""
                            strDestinationAnnotationFile = strFullDirectoryPathForAnnotation & "\" & _
                                                    Format(CStr(txtDetailRECID), "0000000000") & _
                                                    "_" & _
                                                    Format(CStr(strPageNumber), "000000") & _
                                                    ".ANN"

'                            txtActionBeforeError = "FileCopy [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"
'                            FileCopy strSourceAnnotationFile, strDestinationAnnotationFile

                            intCopyFileRetryCount = 1

                            txtActionBeforeError = "FileCopy [" & strSourceAnnotationFile & "] TO [" & strDestinationAnnotationFile & "]"

                            With New FileSystemObject
                                'Move the Annotation file
                                txtActionBeforeError = "MOVE Annotation File from " & strSourceAnnotationFile & " to " & strDestinationAnnotationFile
                               .MoveFile strSourceAnnotationFile, strDestinationAnnotationFile
                                
                                '*** Check to Make SURE the Files were copied properly!
                                If Not funcFileExists(strDestinationAnnotationFile) Then
                                    funcQuickMessage "SHOW", "ERROR: File Not found at Destination after Action (" & txtActionBeforeError & ")"
                                End If
                                
                            End With

                        Next
                    End If
                    
                    
                    
                    
                    '*** Check to Make SURE the Files were copied properly!
                    If Not funcFileExists(strDestinationFile) Then
                    
                        connImaging101.RollbackTrans
                        
                        funcQuickMessage "SHOW", "ERROR: File Not found at Destination after Action (" & txtActionBeforeError & ")... TRANSACTION ROLLED BACK!"
                    
                        If rsImaging101Document.State = adStateOpen Then
                            rsImaging101Document.Close
                        End If
                        Set rsImaging101Document = Nothing
                        
                        If rsImaging101DocumentDetail.State = adStateOpen Then
                            rsImaging101DocumentDetail.Close
                        End If
                        Set rsImaging101DocumentDetail = Nothing
                        
                        Set rsImaging101 = Nothing
                        
                        Set connImaging101 = Nothing
                        Set cmdImaging101 = Nothing
                        
                        
                        Exit Sub
                    End If
                        
                    
                    
        
                    
                    '****************************************************************************
                    '*** UPDATE TRANSACTIONS AND CLOSE RECORD SETS
                    
                    txtActionBeforeError = "Update DOCUMENT " & txtDocumentRECID
                    rsImaging101DocumentDetail.Update
                    
                    txtActionBeforeError = "Update DOCUMENT DETAIL" & txtDetailRECID
                    rsImaging101Document.Update
                    
                    
                    
                    
                    '****************************************************************************
                    '*** COMMIT TRANSACTIONS AND CLOSE RECORD SETS
                    
                    connImaging101.CommitTrans
                    
                    If rsImaging101Document.State = adStateOpen Then
                        rsImaging101Document.Close
                    End If
                    Set rsImaging101Document = Nothing
                    
                    If rsImaging101DocumentDetail.State = adStateOpen Then
                        rsImaging101DocumentDetail.Close
                    End If
                    Set rsImaging101DocumentDetail = Nothing
                    

                    Set cmdImaging101 = Nothing
                    
                    Set connImaging101 = Nothing
                    
                    
                    
                    '****************************************************************************
                    '****************************************************************************
                
        
    MainMDIForm.ActiveForm.txtChildFormMessage.Text = "DOCUMENT SAVED!"
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "DOCUMENT SAVED"
       
    On Error Resume Next
    
    funcWriteToDebugLog Me.name, "Kill " & strSourceFile
    Kill strSourceFile
    
    funcWriteToDebugLog Me.name, "Kill " & strSourceFile & ".pdf  (TEMP Files - for MSG or EML)"
    Kill strSourceFile & ".pdf"
    

    funcWriteToDebugLog Me.name, "Kill ~*.txt (TEMP Files - Text)"
    Kill txtSourceFilePath & "~*.txt"

    '*** UNLOAD the Table Lookup and Doc Types forms
    funcWriteToDebugLog Me.name, "UNLOAD the Table Lookup and Doc Types forms"
        If funcIsFormLoaded2("frmLookupList") Then
            Unload frmLookupList
            Set frmLookupList = Nothing
        End If

        If funcIsFormLoaded2("frmDocTypeList") Then
            Unload frmDocTypeList
            Set frmDocTypeList = Nothing
        End If



    'Make the Save button invisible,but re-enable it for the next time
    Me.cmdSave.Visible = False
    Me.cmdSave.Enabled = True


    Call cmbApplicationList_Click

    'Now re-enable the Clone date fields
    Call subEnableDisableCloneDateFieldsAndButtons
    

    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "EXITING: cmdSave"
    
    funcWriteToDebugLog Me.name, "+++++++++++++++++++++++++++++++++++++++++++"
    

    'RESET the GLOBAL Add File boolean variable to show we are done Adding the file
    bolAIM_Command_AddFile = False
    'Rename the Form so I101Filer can process the next file.
    Me.Caption = "Retrieval Search - Imaging101"
    DoEvents


Exit Sub


ERROR_HANDLER:
        
        On Error Resume Next
        
        Dim dbErr As ADODB.Error
        Dim strErrMsg As String
        
        
        If (connImaging101.Errors.Count > 0) Then
            If (connImaging101.Errors(0).SQLState = "40001") Or (connImaging101.Errors(0).SQLState = "40001") Then
                'Handle transaction commit failure - Serialization Error
                funcQuickMessage "SHOW", "SAVE Failure DURING ACTION: (" & txtActionBeforeError & ") - RETRYING TRANSATCION"
                connImaging101.RollbackTrans
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
                        
        funcQuickMessage "SHOW", "SaveToImaging101 ERROR: " & strErrMsg & vbNewLine & vbNewLine & _
                        "  DURING ACTION: (" & txtActionBeforeError & ") " & vbNewLine & vbNewLine & _
                        "  [Transaction Rolled Back - Page NOT Committed]" & vbNewLine & vbNewLine & _
                        "  File Name   = " & strSourceFile
                        
        On Error Resume Next
        
        Set rsImaging101Document = Nothing
        Set rsImaging101DocumentDetail = Nothing
        
        Set connImaging101 = Nothing
                    
        Set cmdImaging101 = Nothing
        
        Screen.MousePointer = vbDefault

End Sub





Private Sub Form_Activate()
       
    
    'Set focus on the first input field
    '  if NO Errors occured during Input / subBuildSelectStatement
    '  FormActivate happens each time this form gets the focus
'    If Not bolErrorOccured Then
'        mebIndexValues(0).SetFocus
'    End If

    
    
End Sub


Private Sub Form_Load()
    
    bolSearchFormLoadComplete = False
    


    txtWhereFreehand = ""
    txtSearchTemplateWhereFreehand = ""
    

    
    '*** Disable buttons to prevent users from Clicking on them
    '    prior to the form being ready
    cmdFind.Enabled = False
    cmdClear.Enabled = False
'    If gsecRightsDocPackage = vbChecked Then
'        cmdPackage.Visible = True
'    Else
'        cmdPackage.Visible = False
'    End If
'    cmdPackage.enabled = False
'    cmdHelp.enabled = False
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImaging101Search.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImaging101Search.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "frmImaging101Search.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmImaging101Search.Height", RegFileName)
    On Error GoTo 0
    


'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
'''''''    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    On Error GoTo FORM_LOAD_ERROR
    
    '***************************************
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
        
    con.Open RegImaging101ConnectionString
    
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
    
    con.Errors.Clear
    
    rs.Open
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
    
    ' GET The Application this User used last
    Dim i As Integer
    Dim txtApplication As String
    
    txtApplication = funcGetSetUserSettings("GET", "ApplicationSearchForm", "")
    
    
        ' Walk down the Application list... there was no easier way to set the
        '   ListIndex to the right value to Trigger the "cmbApplicationList_Click" event
        For i = 0 To cmbApplicationList.ListCount - 1
            If txtApplication = cmbApplicationList.List(i) Then
                ' This will Trigger the "cmbApplicationList_Click" event
                '   to Load the list of Batches
                cmbApplicationList.ListIndex = i
                Exit For
            End If
        Next i

    '*** Load Last Used Search Template
    If bolEnableSearchTemplates = True Then
        cmbSearchTemplateList = funcGetSetUserSettings("GET", "SearchTemplate", "")
        txtSearchTemplateRECID = funcGetFieldFromDB(RegImaging101ConnectionString, "I101SearchTemplates", "SearchTemplateName = '" & cmbSearchTemplateList & "'", "SearchTemplateRECID")
        txtSearchTemplateWhereFreehand = funcGetFieldFromDB(RegImaging101ConnectionString, "I101SearchTemplates", "SearchTemplateRECID = " & txtSearchTemplateRECID, "WhereFreehand")
    End If

Exit Sub



FORM_LOAD_ERROR:
    If Err.Number = 3021 Then
        MsgBox "You do not have SECURITY RIGHTS to ANY Applications!" & vbCrLf & "Please contact your System Administrator.", vbCritical, "Missing Security Rights"
        Exit Sub
    End If
    result = MsgBox("FORM_LOAD_ERROR: " & Err.Number & " - " & Err.Description, vbOKCancel)
    Err.Clear
    If result = vbOK Then
        'Try again
        Resume Next
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImaging101Search.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Search.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Search.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101Search.Height", Me.Height, RegFileName)
'''        Result = WritePrivateProfileString(RegAppname, "frmImaging101BatchList.Caption", Me.Caption, RegFileName)
    End If


End Sub


Sub subLoadFieldDefinitions()

    '*** THIS SUBROUTINE LOADS ALL THE APPLICATION FIELD DEFINITION INFORMATION
    '***  INCLUDING FIELD FORMAT VALUES INTO AN ARRAY.
    
    bolSearchFormLoadComplete = False
    Me.Enabled = False
    
    funcWriteToDebugLog Me.name, "ENTER: subLoadFieldDefinitions"
    
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    '2020-09-18 - Jacob - Added HideForSearchIndex <> '1' or NULL -- To ignore fields flagged as HideForSearchIndex
    rs.Source = "Select * from I101Fields WHERE ApplicationRECID = " & txtApplicationRECID & " AND (HideForSearchIndex <> '1'  OR HideForSearchIndex is NULL)  ORDER BY FieldOrderBatch"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
   On Error GoTo ERROR_TRAP
    
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
    
    
    '*** DESTROY FIELDS ARRAYS
   On Error Resume Next
    
    Dim intUnloadIndex As Integer
    For intUnloadIndex = 1 To lblFieldDescription.Count - 1
        Unload lblFieldDescription(intUnloadIndex)
        Unload mebIndexValues(intUnloadIndex)
        Unload txtIndexValues(intUnloadIndex)
        Unload txtBatchFieldsRECID(intUnloadIndex)
        Unload txtFieldsRECID(intUnloadIndex)
        Unload txtFieldDefaultValue(intUnloadIndex)
        Unload txtFieldLowValue(intUnloadIndex)
        Unload txtFieldHighValue(intUnloadIndex)
        Unload txtFieldIsSticky(intUnloadIndex)
        Unload txtFieldType(intUnloadIndex)
        Unload txtFieldSize(intUnloadIndex)
        Unload txtFieldName(intUnloadIndex)
        Unload txtFieldIsRequiredForCommit(intUnloadIndex)
        Unload txtFieldSearchCondition(intUnloadIndex)
        Unload cboFieldSearchCondition(intUnloadIndex)
'        Unload txtFieldIsRequiredForSplit(intUnloadIndex)
'        Unload txtFieldSplitBatches(intUnloadIndex)
        Unload cmdFieldDropDown(intUnloadIndex)
        Unload txtFieldIsForOutputOnly(intUnloadIndex)
        Unload txtFieldTableLookupOverridesDefault(intUnloadIndex)

    Next
   
   On Error GoTo ERROR_TRAP
        
   intFieldSpacing = 40

    'RE-Size Form Based on How many fields we Expect if more than 10
'    If rs.RecordCount > 10 Then
'        'Increase the size of the form by the number of Fields we expect
        Dim intNewHeight As Integer
        intNewHeight = lblFieldDescription(0).Top + ((lblFieldDescription(0).Height + intFieldSpacing) * rs.RecordCount) + 700
'        If intNewHeight > 6000 Then
            Me.Height = intNewHeight
'        Else
'            Me.Height = 6000
'        End If
'    End If
    
    '*** intFieldIndex allows us to Add the Second Date field to search through
    '     we set it to (-1) to make sure we start at Zero (0) in the Loop
    Dim intFieldIndex As Integer
    Dim bolFirstPass As Boolean
    
    intFieldIndex = -1
    
    '*** intTabCounter allows us to number the TAB ORDER of the fields and buttons properly
    '     we set it to (1) to make sure we start after the last FIXED/Pre-defined field
    intTabCounter = 1
    
    Dim intDatabaseIndex As Integer
    
    For intDatabaseIndex = 0 To rs.RecordCount - 1
    
    
        
        'Initialize the bolFirstPass flag to track fields we want to create duplicates of...
        bolFirstPass = True

       
        
CREATE_FIELD_OBJECTS:

            intFieldIndex = intFieldIndex + 1
            
            Dim intFieldTop As Integer
            
        '* Create Field Objects - BEGIN
            If intFieldIndex > 0 Then
                Load lblFieldDescription(intFieldIndex)
                Load mebIndexValues(intFieldIndex)
                Load txtIndexValues(intFieldIndex)
                Load txtBatchFieldsRECID(intFieldIndex)
                Load txtFieldsRECID(intFieldIndex)
                Load txtFieldDefaultValue(intFieldIndex)
                Load txtFieldLowValue(intFieldIndex)
                Load txtFieldHighValue(intFieldIndex)
                Load txtFieldIsSticky(intFieldIndex)
                Load txtFieldType(intFieldIndex)
                Load txtFieldSize(intFieldIndex)
                Load txtFieldName(intFieldIndex)
                Load txtFieldIsRequiredForCommit(intFieldIndex)
                Load txtFieldSearchCondition(intFieldIndex)
                Load cboFieldSearchCondition(intFieldIndex)
                Load cmdFieldDropDown(intFieldIndex)
                Load txtFieldIsForOutputOnly(intFieldIndex)
                Load txtFieldTableLookupOverridesDefault(intFieldIndex)
                
                'Set the top to slightly below the previous field
                intFieldTop = lblFieldDescription(intFieldIndex - 1).Top + lblFieldDescription(intFieldIndex - 1).Height + intFieldSpacing
            Else
                'Set top to where the first field is
                intFieldTop = lblFieldDescription(intFieldIndex).Top
            End If
        '* Create Field Objects - END
                
                
'                Set lblFieldDescription(intFieldIndex).Container = Frame2
                lblFieldDescription(intFieldIndex).Top = intFieldTop
                lblFieldDescription(intFieldIndex).Enabled = True
                lblFieldDescription(intFieldIndex).Visible = True
                lblFieldDescription(intFieldIndex).Caption = ""
                lblFieldDescription(intFieldIndex).ForeColor = vbNormal

'                Set mebIndexValues(intDatabaseIndex).Container = Frame2
                mebIndexValues(intFieldIndex).Top = intFieldTop
                mebIndexValues(intFieldIndex).Enabled = True
                mebIndexValues(intFieldIndex).Visible = False
                '2023-05-03 - Jacob - Disabled setting TabIndex here.  It is handled below.
                'mebIndexValues(intFieldIndex).TabIndex = intDatabaseIndex + 1
                mebIndexValues(intFieldIndex).Text = ""
            
'                Set txtIndexValues(intDatabaseIndex).Container = Frame2
                txtIndexValues(intFieldIndex).Top = intFieldTop
                txtIndexValues(intFieldIndex).Enabled = True
                txtIndexValues(intFieldIndex).Visible = False
                '2023-05-03 - Jacob - Disabled setting TabIndex here.  It is handled below.
                'txtIndexValues(intFieldIndex).TabIndex = intDatabaseIndex + 1
                txtIndexValues(intFieldIndex).Text = ""

'                Set txtBatchFieldsRECID(intFieldIndex).Container = Frame2
                txtBatchFieldsRECID(intFieldIndex).Enabled = True
                txtBatchFieldsRECID(intFieldIndex).Visible = False

'                Set txtFieldsRECID(intFieldIndex).Container = Frame2
                txtFieldsRECID(intFieldIndex).Enabled = True
                txtFieldsRECID(intFieldIndex).Visible = False

'                Set txtFieldDefaultValue(intFieldIndex).Container = Frame2
                txtFieldDefaultValue(intFieldIndex).Enabled = True
                txtFieldDefaultValue(intFieldIndex).Visible = False
                txtFieldDefaultValue(intFieldIndex).Text = ""
                
'                Set txtFieldLowValue(intFieldIndex).Container = Frame2
                txtFieldLowValue(intFieldIndex).Enabled = True
                txtFieldLowValue(intFieldIndex).Visible = False
                txtFieldLowValue(intFieldIndex).Text = ""
            
'                Set txtFieldHighValue(intFieldIndex).Container = Frame2
                txtFieldHighValue(intFieldIndex).Enabled = True
                txtFieldHighValue(intFieldIndex).Visible = False
                txtFieldHighValue(intFieldIndex).Text = ""
                
'                Set txtFieldIsSticky(intFieldIndex).Container = Frame2
                txtFieldIsSticky(intFieldIndex).Enabled = True
                txtFieldIsSticky(intFieldIndex).Visible = False
                txtFieldIsSticky(intFieldIndex).Text = ""
            
'                Set txtFieldType(intFieldIndex).Container = Frame2
                txtFieldType(intFieldIndex).Enabled = True
                txtFieldType(intFieldIndex).Visible = False
                txtFieldType(intFieldIndex).Text = ""
            
'                Set txtFieldSize(intFieldIndex).Container = Frame2
                txtFieldSize(intFieldIndex).Enabled = True
                txtFieldSize(intFieldIndex).Visible = False
                txtFieldSize(intFieldIndex).Text = ""
                
'                Set txtFieldName(intFieldIndex).Container = Frame2
                txtFieldName(intFieldIndex).Enabled = True
                txtFieldName(intFieldIndex).Visible = False
                txtFieldName(intFieldIndex).Text = ""
            
'                Set txtFieldIsRequiredForCommit(intFieldIndex).Container = Frame2
                txtFieldIsRequiredForCommit(intFieldIndex).Enabled = True
                txtFieldIsRequiredForCommit(intFieldIndex).Visible = False
                txtFieldIsRequiredForCommit(intFieldIndex).Text = ""
            
            
'                Set txtFieldSearchCondition(intFieldIndex).Container = Frame2
                txtFieldSearchCondition(intFieldIndex).Top = tintFieldTop
                txtFieldSearchCondition(intFieldIndex).Enabled = False
                txtFieldSearchCondition(intFieldIndex).Visible = False
                txtFieldSearchCondition(intFieldIndex).Text = ""
            
            
'                Set cboFieldSearchCondition(intFieldIndex).Container = Frame2
                cboFieldSearchCondition(intFieldIndex).Top = intFieldTop
                cboFieldSearchCondition(intFieldIndex).Enabled = True
                cboFieldSearchCondition(intFieldIndex).Visible = True

                txtFieldIsForOutputOnly(intFieldIndex).Enabled = True
                txtFieldIsForOutputOnly(intFieldIndex).Visible = False
                txtFieldIsForOutputOnly(intFieldIndex).Text = ""
            
                '*** Create the DROP-DOWN Button
'                Set FieldDropDownList(intFieldIndex).Container = Frame2
                cmdFieldDropDown(intFieldIndex).Top = intFieldTop
                cmdFieldDropDown(intFieldIndex).Enabled = True
                intTabCounter = intTabCounter + 1
                cmdFieldDropDown(intFieldIndex).TabStop = False
                'Make the DropDownList button VISIBLE only if Checked for the current field
                '2017-04-16 Jacob - Added and NOT in Add-File Mode
                If rs.Fields!FieldDropDownList = vbChecked And (bolAIM_Command_AddFile = False Or rs.Fields!FieldDropDownListAlsoOnFiler = vbChecked) Then
                    cmdFieldDropDown(intFieldIndex).Visible = True
                Else
                    cmdFieldDropDown(intFieldIndex).Visible = False
                End If
                
                
        
        'Clear any Values carried over from the first (Master) field
        lblFieldDescription(intFieldIndex) = ""
        mebIndexValues(intFieldIndex).Mask = ""
        mebIndexValues(intFieldIndex).Format = ""
        mebIndexValues(intFieldIndex).Text = ""
    
        '* Assign Field Values
        txtFieldsRECID(intFieldIndex) = rs.Fields!FieldsRECID
        If (IsNull(rs.Fields!FieldNameForInput)) Or (rs.Fields!FieldNameForInput <> "") Then
            lblFieldDescription(intFieldIndex) = rs.Fields!FieldNameForInput
        Else
            lblFieldDescription(intFieldIndex) = rs.Fields!FieldName
        End If
        
        
        If Not IsNull(rs.Fields!FieldMask) Then mebIndexValues(intFieldIndex).Mask = rs.Fields!FieldMask
        If Not IsNull(rs.Fields!FieldFormat) Then mebIndexValues(intFieldIndex).Format = rs.Fields!FieldFormat
        
        If Not IsNull(rs.Fields!FieldDefaultValue) Then txtFieldDefaultValue(intFieldIndex) = rs.Fields!FieldDefaultValue
        If Not IsNull(rs.Fields!FieldLowValue) Then txtFieldLowValue(intFieldIndex) = rs.Fields!FieldLowValue
        If Not IsNull(rs.Fields!FieldHighValue) Then txtFieldHighValue(intFieldIndex) = rs.Fields!FieldHighValue
        If Not IsNull(rs.Fields!FieldIsSticky) Then txtFieldIsSticky(intFieldIndex) = rs.Fields!FieldIsSticky
        If Not IsNull(rs.Fields!FieldType) Then txtFieldType(intFieldIndex) = rs.Fields!FieldType
        If Not IsNull(rs.Fields!FieldSize) Then txtFieldSize(intFieldIndex) = rs.Fields!FieldSize
        If Not IsNull(rs.Fields!FieldName) Then txtFieldName(intFieldIndex) = rs.Fields!FieldName
        If Not IsNull(rs.Fields!FieldSearchCondition) Then txtFieldSearchCondition(intFieldIndex) = rs.Fields!FieldSearchCondition
        
         '1/19/2011 - Jacob - FieldIsForOutputOnly handles the "Prevent Manual Indexing" application field option
        If Not IsNull(rs.Fields!FieldIsForOutputOnly) Then txtFieldIsForOutputOnly(intFieldIndex) = rs.Fields!FieldIsForOutputOnly
        If Not IsNull(rs.Fields!FieldTableLookupOverridesDefault) Then txtFieldTableLookupOverridesDefault(intFieldIndex) = rs.Fields!FieldTableLookupOverridesDefault

        If bolAIM_Command_AddFile Then
            If txtFieldIsForOutputOnly(intFieldIndex) = vbChecked Then
                mebIndexValues(intFieldIndex).Enabled = False
                txtIndexValues(intFieldIndex).Enabled = False
            End If
        End If
        
        For i = 0 To cboFieldSearchConditionLIST.ListCount - 1
            cboFieldSearchCondition(intFieldIndex).AddItem cboFieldSearchConditionLIST.List(i)
        Next
        funcFindItemInComboBox cboFieldSearchCondition(intFieldIndex), txtFieldSearchCondition(intFieldIndex)
        
        
        If Not IsNull(rs.Fields!FieldIsRequiredForCommit) Then txtFieldIsRequiredForCommit(intFieldIndex) = rs.Fields!FieldIsRequiredForCommit
        
        
        '10/7/2012 - Jacob - Moved this down so txtFieldType test will work
        '*** Determine whether to use the Text or Masked Edit Control
        If txtFieldType(intFieldIndex) = "LongText" Then
            mebIndexValues(intFieldIndex).TabStop = False
            txtIndexValues(intFieldIndex).TabStop = True
            txtIndexValues(intFieldIndex).Visible = True
            '2023-05-03 - Jacob - Corrected TabIndex ordering to use "intFieldIndex", instead of "intDatabaseIndex + 1"
            'txtIndexValues(intFieldIndex).TabIndex = intDatabaseIndex + 1
            txtIndexValues(intFieldIndex).TabIndex = intFieldIndex
        Else
            txtIndexValues(intFieldIndex).TabStop = False
            mebIndexValues(intFieldIndex).TabStop = True
            mebIndexValues(intFieldIndex).Visible = True
            '2023-05-03 - Jacob - Corrected TabIndex ordering to use "intFieldIndex", instead of "intDatabaseIndex + 1"
            'mebIndexValues(intFieldIndex).TabIndex = intDatabaseIndex + 1
            mebIndexValues(intFieldIndex).TabIndex = intFieldIndex
        End If
        

        'Handle AddFile scenario BEFORE the Date Conditional check... otherwise it will not execute
         If bolAIM_Command_AddFile Then
        
            If txtFieldIsRequiredForCommit(intFieldIndex) = vbChecked Then
                'Show that field is required!
                'lblFieldDescription(intFieldIndex) = lblFieldDescription(intFieldIndex)
                lblFieldDescription(intFieldIndex).ForeColor = vbRed
            Else
                lblFieldDescription(intFieldIndex).ForeColor = vbNormal
            End If
            
            'Set the Default Values for this field
            Call subSetDefaultFieldValues(intFieldIndex)

        End If
   
        
        
        If txtFieldType(intFieldIndex).Text = "Date" And bolFirstPass = True And bolAIM_Command_AddFile = False Then
            'Increase Form Height to account for this extra field
            Me.Height = Me.Height + lblFieldDescription(0).Height
            cboFieldSearchCondition(intFieldIndex) = ">="
            cboFieldSearchCondition(intFieldIndex).Enabled = False
            'Make sure we don't create it more than once.
            bolFirstPass = False
            'Create another Date field identical to this one
            GoTo CREATE_FIELD_OBJECTS
        End If
        
        '* If setting up a Range Field - append the text "(Thru)"
        If bolFirstPass = False And bolAIM_Command_AddFile = False Then
            lblFieldDescription(intFieldIndex) = lblFieldDescription(intFieldIndex - 1) '& " (Thru)"
            cboFieldSearchCondition(intFieldIndex) = "<="
            cboFieldSearchCondition(intFieldIndex).Enabled = False
            'RESET the bolFirstPass FLAG to get the next record
            bolFirstPass = True
        End If
        
            
        'Get NEXT FIELD only if bolFirstPass = true -- NOT creating second Range Date Search Field
        If bolFirstPass = True Then
            rs.MoveNext
        End If
        
      
        
    Next
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    

    
    funcWriteToDebugLog Me.name, "EXIT: subLoadFieldDefinitions"

    
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    
    funcWriteToDebugLog Me.name, "LoadFieldFormats - Error: " & Err.Number & " - " & Err.Description
    result = MsgBox("LoadFieldFormats - Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
End Sub


Private Sub Form_Resize()

    Dim dblMinWidth As Double
    
    If Me.WindowState <> vbMinimized Then
        dblMinWidth = cmbSearchTemplateList.Left + cmbSearchTemplateList.width + 220
        If Me.width < dblMinWidth Then
            Me.width = dblMinWidth
        End If
        
        Frame1.width = Me.width
        picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
        lblVersion.Left = picImaging101Logo.Left
        
        '*** 2020-04-24 - Jacob - Added RESIZE for All Fields
        '*** 2020-07-30 - Jacob -Moved inside WindowState minimized to avoid "Run-time error '380': Invalid property value"
        Dim intResizeFieldsIndex As Integer
        For intResizeFieldsIndex = 0 To lblFieldDescription.Count - 1
    
                    cmdFieldDropDown(intResizeFieldsIndex).Left = Me.ScaleWidth - cmdFieldDropDown(intResizeFieldsIndex).width - 10
    
                    If txtFieldType(intResizeFieldsIndex).Text = "LongText" Then
                       txtIndexValues(intResizeFieldsIndex).width = Me.ScaleWidth - txtIndexValues(intResizeFieldsIndex).Left - cmdFieldDropDown(intResizeFieldsIndex).width - 10
                    Else
                        mebIndexValues(intResizeFieldsIndex).width = Me.ScaleWidth - mebIndexValues(intResizeFieldsIndex).Left - cmdFieldDropDown(intResizeFieldsIndex).width - 10
                    End If
    
        Next
        
    End If
    

    
End Sub

Private Sub Form_Unload(Cancel As Integer)


    Dim bolKillI101AimFile As Boolean
    
    
    bolKillI101AimFile = False

    
    

    'If Adding File using AIM, ask if user wants to delete the file.
    If bolAIM_Command_AddFile = True Then
    
         '*** GET SOURCE FILE NAME to Use if we need to Delete the I101AIM file, since we will have already Unloaded the Viewer.
        Dim strSourceFile As String
        strSourceFile = MainMDIForm.ActiveForm.txtPageFileName
    
        result = MsgBox("You have NOT Saved this file." & vbCrLf & "Would you like to DELETE it?", vbYesNoCancel + vbQuestion, "Delete File?")
        If result = vbYes Then
            result = MsgBox("Are you SURE you want to DELETE this File?", vbYesNoCancel + vbCritical, "Delete File Confirmation")
            
            If result = vbYes Then
                     bolKillI101AimFile = True
                     'Prevent the Child form from asking to save Annotations
                     bolAnnotationAdded = False
            ElseIf result = vbCancel Then
                    Cancel = True
                    Exit Sub
            End If
        
        ElseIf result = vbCancel Then
                    Cancel = True
                    Exit Sub
        End If
        
        If result = vbNo Then
        '*** 2023-03-16 - Jacob - Added vbCrLF's to make message more readable
            result = MsgBox("Please note that I101 FILER will keep trying to open this File until you save it." & vbCrLf & "To prevent this" & vbCrLf & "CLICK the [STOP] Button" & "in the I101 Filer window.", vbOKCancel + vbInformation, "Delete File Confirmation")
            bolKillI101AimFile = False
            If result = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        
    End If
    
    

    
    
    If funcIsFormLoaded2("frmIndex") And UnloadMode = vbFormControlMenu Then
        
        'INDEXING FORM IS OPEN
        'DO NOT UNLOAD the VIEWER
        
    Else
        
        'GO AHEAD AND UNLOAD THE VIEWER
        'AND SHOW THE MAIN MENU
        
        If funcIsFormLoaded2("frmDropDownList") Then
            Unload frmDropDownList
            Set frmDropDownList = Nothing
        End If
        
        If funcIsFormLoaded2("frmLookupList") Then
            Unload frmLookupList
            Set frmLookupList = Nothing
        End If
        
        If funcIsFormLoaded2("frmDocTypeList") Then
            Unload frmDocTypeList
            Set frmDocTypeList = Nothing
        End If
        
        If funcIsFormLoaded2("frmThumb") Then
            Unload frmThumb
            Set frmThumb = Nothing
        End If
        
        If funcIsFormLoaded2("frmAnnotate") Then
            Unload frmAnnotate
            Set frmAnnotate = Nothing
        End If
        
        If funcIsFormLoaded2("frmImaging101Retrieve") Then
            Unload frmImaging101Retrieve
            Set frmImaging101Retrieve = Nothing
        End If
        
        If funcIsFormLoaded2("MainMDIForm") Then
            Unload MainMDIForm
            Set MainMDIForm = Nothing
        End If
    
        'Only Show Menu when unloading if NOT in SYSTRAY mode.
        If Not bolSysTrayActive Then
            frmMainMenu.Show
            frmMainMenu.WindowState = vbNormal
            frmMainMenu.SetFocus
        End If
        
    End If
        
    '*** 2023-03-15 - Jacob - Added delete for PDF Temp file created for MSG and EML's
    
    On Error Resume Next
    
    'DELETE the I101AIM File?
    If bolKillI101AimFile = True Then
    
                With New FileSystemObject
                
                        'Move the Annotation file
                        txtActionBeforeError = "funcFileExists(" & strSourceFile & ")"
                        funcWriteToDebugLog Me.name, txtActionBeforeError
                        If funcFileExists(strSourceFile) Then
                            txtActionBeforeError = "DELETE File: " & strSourceFile
                            funcWriteToDebugLog Me.name, "DELETE File: " & strSourceFile
                            .DeleteFile strSourceFile
                        End If
                        
                        If funcFileExists(strSourceFile & ".pdf") Then
                            xtActionBeforeError = "DELETE File: " & strSourceFile & ".pdf (TEMP file for MSG & EML)"
                            funcWriteToDebugLog Me.name, ".DeleteFile" & strSourceFile & ".pdf (TEMP file for MSG & EML)"
                            .DeleteFile strSourceFile & ".pdf"
                        End If
                        
                        Dim strLocalTempDir As String
                        strLocalTempDir = funcGetTempDir()
                        
                        If Dir(strLocalTempDir & "\Imaging101\Annotations\*.ANN") <> "" Then
                            txtActionBeforeError = "DELETE Temp Annotation Files."
                            funcWriteToDebugLog Me.name, txtActionBeforeError
                            .DeleteFile strLocalTempDir & "\Imaging101\Annotations\*.ANN"
                        End If

                        
                End With
                            
    End If
    
    On Error Resume Next

    funcWriteToDebugLog Me.name, "Kill ~*.txt (TEMP Text Files)"
    Kill txtSourceFilePath & "~*.txt"

    
    

        
    'RESET the GLOBAL Add File boolean variable to show we are done Adding the file
    bolAIM_Command_AddFile = False
    
    
    Set frmImaging101Search = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mebIndexValues_Change(Index As Integer)

'    'Check to see if something was entered into the FIRST Index Field
'    If Trim(mebIndexValues(Index).Text) <> "" Then
'        cmdPackage.enabled = True
'    Else
'        cmdPackage.enabled = False
'    End If

End Sub

Private Sub mebIndexValues_GotFocus(Index As Integer)

'    '*** Re-enable buttons
'    cmdFind.Enabled = True
'    cmdClear.Enabled = True
'    cmdHelp.Enabled = True
'    bolSearchFormLoadComplete = True
    
    
    '* If Date is Blank, then Errors - Copy the Date to the Thru field
    If (txtFieldType(Index) = "Date") And (Trim(mebIndexValues(Index).Text) = "") Then
        If InStr(1, cboFieldSearchCondition(Index), "<=") > 0 Then
            '* By design, the Thru date field is always immediatelly after the From
            mebIndexValues(Index).Text = mebIndexValues(Index - 1).Text
        End If
    End If
    
    '*** Highlight the Field
    mebIndexValues.item(Index).SelStart = 0
'    mebIndexValues.item(Index).SelLength = Len(mebIndexValues.item(Index).Text)
    mebIndexValues.item(Index).SelLength = 99
    
    intCurrentFieldIndex = Index

    
End Sub

Private Sub mebIndexValues_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If bolSearchFormLoadComplete <> True Then
        Exit Sub
    End If
    
    
    'Catch Enter key
    If KeyAscii = 13 Then
        cmdFind_Click
    End If

'    If KeyAscii = Asc("[") And frmImaging101Retrieve.Visible = True Then
'        frmImaging101Search.SetFocus
'        'Cancel the Keypress by setting the KeyAscii to Zero (0)
'        KeyAscii = 0
'    End If
'
'    If KeyAscii = Asc("]") And frmImaging101Retrieve.Visible = True Then
'        frmImaging101Retrieve.SetFocus
'        'Cancel the Keypress by setting the KeyAscii to Zero (0)
'        KeyAscii = 0
'    End If
'
'    If KeyAscii = Asc("\") And MainMDIForm.Visible = True Then
'        MainMDIForm.SetFocus
'        'Cancel the Keypress by setting the KeyAscii to Zero (0)
'        KeyAscii = 0
'    End If
    


End Sub

Private Sub mebIndexValues_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

'    'Check to see if something was entered into the FIRST Index Field
'    If Trim(mebIndexValues(Index).Text) <> "" Then
'        cmdPackage.Enabled = True
'    Else
'        cmdPackage.Enabled = False
'    End If

End Sub

Private Sub mebIndexValues_LostFocus(Index As Integer)

        'Set FOCUS to the Thru Date field
'        If cboFieldSearchCondition(Index).Text = ">=" Then
'             If Index < mebIndexValues.UBound Then
'                mebIndexValues(Index + 1).SetFocus
'            End If
             
'        End If
        
End Sub

Private Sub mebIndexValues_Validate(Index As Integer, Cancel As Boolean)
    
    On Error GoTo ERROR_TRAP
    
    If (txtFieldType(Index) = "Date") And (Trim(mebIndexValues(Index).Text) <> "") Then
        '* Remove the "Prompt" characters
        strDateFormatted = Replace(Trim(mebIndexValues(Index).FormattedText), "_", "")
        strDateFormatted = CDate(strDateFormatted)
       
    End If
            

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

Public Sub subSetFieldValue(Index As Integer, strFieldValue As String)

    While Not bolSearchFormLoadComplete
        DoEvents
    Wend
    
    mebIndexValues(Index).Text = strFieldValue
    
    cmdFind_Click
    
    
End Sub

Private Sub picImaging101Logo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Constant (Button) Value Description
    'vbLeftButton      1     Left button is pressed
    'vbRightButton     2     Right button is pressed
    'vbMiddleButton    4     Middle button is pressed
    '
    'Constant (Shift) Value Description
    'vbShiftMask      1     SHIFT key is pressed.
    'vbCtrlMask       2     CTRL key is pressed.
    'vbAltMask        4     ALT key is pressed.

    If gsecRightsAdminSystem And Button = vbRightButton And Shift = vbCtrlMask Then
        
        cmdAdvanced_Click
        
    End If

End Sub

Private Sub Timer1_Timer()
    'This timer is simply to bypass a VB Error:  "Unable to unload within this context (Error 365)"
    ' attempting to "Destroy" the fields when switching Applications
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vbenlr98/html/vamsgldcantunloadhere.asp
    'apparently VB won't let you unload certain objects in certain contexts like the "_Click" event
    ' or any event or Sub that this event calls!!!
    
    On Error Resume Next
    
    'Disable this timer now
    Timer1.Enabled = False
    
    
    'Set the Form CAPTION depending on the function being performed
    If bolAIM_Command_AddFile = True Then
            
        Me.Caption = "I101Filer AddFile - Imaging101"
               
    Else
    
        Me.Caption = "Retrieval Search - Imaging101"
        
    End If
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    If bolDebug Then
        Me.Caption = Me.Caption & " - DEBUG"
    End If

    
    If cmdFieldList.Visible = True Then
        cmdAdvanced_Click
    End If
    
    If funcIsFormLoaded2("frmDropDownList") Then
        Unload frmDropDownList
    End If
    
            If funcIsFormLoaded2("frmDocTypeList") Then
                Unload frmDocTypeList
                DoEvents
                Set frmdoctypeslist = Nothing
            End If
            
            If funcIsFormLoaded2("frmLookupList") Then
                Unload frmLookupList
                DoEvents
                Set frmLookupList = Nothing
            End If
            
    frmImaging101Retrieve.ListView1.ListItems.Clear
    frmImaging101Retrieve.subCheckButtonSecurity
    
    bolSearchFormLoadComplete = False
    
    'Now load the new field definitions
    Call subLoadFieldDefinitions
    
  
    
    'Now DISABLE the Clone date fields
    Call subEnableDisableCloneDateFieldsAndButtons
    
    Me.Show
    
  
    DoEvents
    Me.Enabled = True
    
    
    If bolAIM_Command_AddFile Then
    
            Call cmdSetFocusOnFirstField
    
            funcWriteToDebugLog Me.name, "frmDocTypeList.Show"
            frmDocTypeList.Show
'            frmDocTypeList.WindowState = vbNormal
            DoEvents
            
            funcWriteToDebugLog Me.name, "frmLookupList.Show"
            frmLookupList.Show
'            frmLookupList.WindowState = vbNormal
            
            If funcIsFormLoaded2("frmLookupList") Then
                frmLookupList.Visible = True
'                funcWriteToDebugLog Me.name, "funcMakeTopMost frmLookupList, True"
'                funcMakeTopMost frmLookupList, True
'                funcMakeTopMost frmLookupList, False
'                frmLookupList.SetFocus
    '            DoEvents
            End If
            

    Else
    
            '7/4/2017 - Jacob - Moved the Clear Fields inside the AddFile check so it won't Clear Default values
            '5/11/2015 - Jacob - Clear Fields to Set Focus on the FIRST Field
             Call cmdClear_Click

    End If
    

    
    bolSearchFormLoadComplete = True

    
    
End Sub

Private Sub subEnableDisableCloneDateFieldsAndButtons()

    '**************************************************
    '*** 2/28/2013 - New Loop to disable the "clone" date field if in AddFile mode - Cycle through Field Values
        For intIndex = 0 To mebIndexValues.Count - 1
        
            'Handle the Clone Date Fields
            If txtFieldType(intIndex) = "Date" And cboFieldSearchCondition(intIndex) = "<=" Then
                    
                If bolAIM_Command_AddFile = True Then
                
                    lblFieldDescription(intIndex).Visible = False
                    cboFieldSearchCondition(intIndex).Visible = False
                    mebIndexValues(intIndex).Visible = False
                    cmdSave.Default = True
                                    
                Else
                
                    lblFieldDescription(intIndex).Visible = True
                    cboFieldSearchCondition(intIndex).Visible = True
                    mebIndexValues(intIndex).Visible = True
                    cmdFind.Default = True
                    
                End If
                
            End If

             If bolAIM_Command_AddFile = False Then
                'Re-Enable ALL Fields for Search Input
                mebIndexValues(intIndex).Enabled = True
                txtIndexValues(intIndex).Enabled = True
            End If
            
        Next
        
        
        
        If bolAIM_Command_AddFile = True Then
                            
                    cmdFind.Visible = False
                    cmdFind.Enabled = False
                    
                    cmdAdvanced.Visible = False
                    cmdAdvanced.Enabled = False
                    
                    cmdSave.Left = cmdFind.Left
        
        Else
                    cmdFind.Visible = True
                    cmdFind.Enabled = True
            
                    If gsecRightsAdminSystem Or gsecAdvancedSearch Then
                        cmdAdvanced.Visible = True
                        cmdAdvanced.Enabled = True
                    Else
                        cmdAdvanced.Visible = False
                        cmdAdvanced.Enabled = False
                    End If
                    
                    '*** 2020-05-15 - Jacob - Added chec for gsecRightsEditSerchTemplates
                    If gsecRightsAdminSystem Or gsecRightsEditSearchTemplates Then
                        cmdEditSearchTemplate.Visible = True
                        cmdEditSearchTemplate.Enabled = True
                    Else
                        cmdEditSearchTemplate.Visible = False
                        cmdEditSearchTemplate.Enabled = False
                    End If
        
        End If
        
        cmdClear.Visible = True
        cmdClear.Enabled = True
        
        cmdHelp.Visible = False
        cmdHelp.Enabled = False
        
        cmdPackage.Visible = False
        cmdPackage.Enabled = False
        
End Sub

Private Sub subBuildSelectStatement()
'    Dim txtFilterStatement As String   'Jacob - 12/22/1006: Changed to invisible Text field on Form
    Dim txtFieldsList As String
    Dim txtOrderByList As String
    Dim strDateFormatted As String
    Dim intViewDeletedDocuments As Integer
    
    On Error GoTo ERROR_TRAP
    
    '*** Clear variables
    txtFilterStatement = ""
    txtFieldsList = ""
    txtOrderByList = ""
    
    'RESET the Error occured flag
    bolErrorOccured = False
    
    
    '*** Prepare the WHERE Clause and List of Fields
    For intIndex = 0 To lblFieldDescription.Count - 1
        ' Use field ONLY if NOT EMPTY
        If Trim(mebIndexValues(intIndex) & txtIndexValues(intIndex)) <> "" Then
            If Trim(txtFilterStatement) <> "" Then
                txtFilterStatement = txtFilterStatement & " AND "
            End If
            
            If txtFieldType(intIndex) = "Date" Then
                '* Remove the "Prompt" characters
                strDateFormatted = Replace(Trim(mebIndexValues(intIndex).FormattedText), "_", "")
                strDateFormatted = CDate(strDateFormatted)
                
                '* If FieldType is "Date" Handle Date From and Thru scenario
                If InStr(1, cboFieldSearchCondition(intIndex), "<=") = 0 Then
                    txtFilterStatement = txtFilterStatement & txtFieldName(intIndex) & " >= '" & strDateFormatted & "' "
                Else
                    txtFilterStatement = txtFilterStatement & txtFieldName(intIndex) & " <= '" & strDateFormatted & "' "
                End If
            Else
                '* If NOT Date, do a  search based on the FieldSearchCondition
                
                '*** 2008-11-13 - Jacob: Added to allow User-defined Search Conditions
                Dim strFieldSearchCondition As String
                Dim strFieldWildCardBegin As String
                Dim strFieldWildCardEnd As String
                Dim strFieldValue As String
                
                Select Case cboFieldSearchCondition(intIndex)  'txtFieldSearchCondition(intIndex)
                    Case "Contains"
                        strFieldSearchCondition = "LIKE"
                        strFieldWildCardBegin = "%"
                        strFieldWildCardEnd = "%"
                    Case "Begins With"
                        strFieldSearchCondition = "LIKE"
                        strFieldWildCardBegin = ""
                        strFieldWildCardEnd = "%"
                    Case "IN"
                        strFieldSearchCondition = "IN"
                        strFieldWildCardBegin = ""
                        strFieldWildCardEnd = ""

                    Case Else
                        strFieldSearchCondition = cboFieldSearchCondition(intIndex)   'txtFieldSearchCondition(intIndex)
                        strFieldWildCardBegin = ""
                        strFieldWildCardEnd = ""
                
                End Select
                
                
                '*** 2004-04-08 - Jacob: Added the REPLACE to prevent errors if Single Apostrophe was allowed in field.
'                txtFilterStatement = txtFilterStatement & txtFieldName(intIndex) & " LIKE '%" & Replace(Trim(mebIndexValues(intIndex).FormattedText), "'", "''") + "%' "
                '*** 2008-11-13 - Jacob: Modified to allow User-defined Search Conditions
                
                '*** 2012-10-23 - Jacob - Added "Long Text" Check & strValue field.
                If txtFieldType(intIndex) = "LongText" Then
                    strFieldValue = Replace(Trim(txtIndexValues(intIndex).Text), "'", "''")
                    txtFilterStatement = txtFilterStatement & txtFieldName(intIndex) & " " & strFieldSearchCondition & " '" & strFieldWildCardBegin & strFieldValue & strFieldWildCardEnd & "' "
                ElseIf cboFieldSearchCondition(intIndex) = "IN" Then
                    strFieldValue = Trim(mebIndexValues(intIndex).FormattedText)
                    txtFilterStatement = txtFilterStatement & txtFieldName(intIndex) & " " & strFieldSearchCondition & " " & strFieldWildCardBegin & strFieldValue & strFieldWildCardEnd
                Else
                    strFieldValue = Replace(Trim(mebIndexValues(intIndex).FormattedText), "'", "''")
                    txtFilterStatement = txtFilterStatement & txtFieldName(intIndex) & " " & strFieldSearchCondition & " '" & strFieldWildCardBegin & strFieldValue & strFieldWildCardEnd & "' "
                End If
                
                
            
            End If
        End If
        
        'Only Add the Fields to the FieldList & Order-by if NOT the "Thru" field
        If (txtFieldType(intIndex) = "Date") And InStr(1, cboFieldSearchCondition(intIndex), "<=") > 0 Then
            'SKIP THIS FIELD
        Else
            ' Append Field Names
            txtFieldsList = txtFieldsList & ", " & txtFieldName(intIndex)
            ' Create OrderBy list
            If intIndex < 5 Then
                If txtOrderByList <> "" Then
                    txtOrderByList = txtOrderByList & ", "
                End If
                txtOrderByList = txtOrderByList & txtFieldName(intIndex)
            End If
        End If

    Next
    
    'Build the WHERE Clause
    If Trim(txtFilterStatement) <> "" Then
        txtFilterStatement = " WHERE " & txtFilterStatement
        
        'Add the FREEHAND conditions
        If Trim(txtWhereFreehand) <> "" Then
            txtFilterStatement = txtFilterStatement & " AND ( " & txtWhereFreehand & " ) "
        End If
        
        'Add the SEARCH TEMPLATE FREEHAND conditions
        If Trim(txtSearchTemplateWhereFreehand) <> "" Then
            'Replace Special Fields, if used, with Actual Values
            '2019-12-13 - Jacob - Removed the single quotes around gsecUserID to allow using "Wildcards" in Search Template.  Added single-quotes in Dropdown, instead.
            txtSearchTemplateWhereFreehand = Replace(txtSearchTemplateWhereFreehand, "{LoggedInUserID}", gsecUserID)
            txtSearchTemplateWhereFreehand = Replace(txtSearchTemplateWhereFreehand, "{CurrentDate}", Format(Now(), "mm-dd-yyyy"))
           
            txtFilterStatement = txtFilterStatement & " AND ( " & txtSearchTemplateWhereFreehand & " ) "
        End If
        
        'Check if user selected to View Documents flagged as Deleted
        intViewDeletedDocuments = frmImaging101Retrieve.chkViewDeletedDocuments.Value
        If intViewDeletedDocuments = vbChecked Then
            '2014-12-20 - Jacob - ONLY Show Documents Deleted by Users... NOT via AutoImport
            txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked = 'D' ) "
        Else
            'Don't Display Deleted Documents
            '2014-12-20 - Jacob - Changed to Handle "D" deleted from Retrieval and "DA" deleted via AutoExport.
           'txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked <> 'D' OR DocumentLocked = '' OR DocumentLocked is NULL ) "
            txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked NOT LIKE 'D%' OR DocumentLocked = '' OR DocumentLocked is NULL ) "
        End If
    
        'Don't display documents that were MOVED OUT (MO) except in FreeHand search
        If InStr(txtWhereFreehand, "DocumentLocked") = 0 Then
            txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked <> 'MO' OR DocumentLocked = '' OR DocumentLocked is NULL  ) "
        End If
        
    Else    'If Trim(txtFilterStatement) <> ""
        
        
        If Trim(txtWhereFreehand) = "" And Trim(txtSearchTemplateWhereFreehand) = "" Then
            'If NO Search items submitted then Get out of here now
            'set the bolErrorOccured flag so the retrieve window is not loaded or populated
            result = MsgBox("Please enter at least one (1) item to search for and try again.", vbOKOnly, "NO Search Parameters Entered.")
            bolErrorOccured = True
            Exit Sub
        End If
        
        'Add the FREEHAND conditions
        If Trim(txtWhereFreehand) <> "" Then
            txtWhereFreehand = Replace(txtWhereFreehand, "{LoggedInUserID}", gsecUserID)
            txtWhereFreehand = Replace(txtWhereFreehand, "{CurrentDate}", Format(Now(), "mm-dd-yyyy"))

            txtFilterStatement = " WHERE ( " & txtWhereFreehand & " ) "
        End If
        
        'Add the SEARCH TEMPLATE FREEHAND conditions
        If Trim(txtSearchTemplateWhereFreehand) <> "" Then
            'Replace Special Fields, if used, with Actual Values
            txtSearchTemplateWhereFreehand = Replace(txtSearchTemplateWhereFreehand, "{LoggedInUserID}", gsecUserID)
            txtSearchTemplateWhereFreehand = Replace(txtSearchTemplateWhereFreehand, "{CurrentDate}", Format(Now(), "mm-dd-yyyy"))
           
            If txtWhereFreehand = "" Then
                txtFilterStatement = " WHERE ( " & txtSearchTemplateWhereFreehand & " ) "
            Else
                txtFilterStatement = txtFilterStatement & " AND ( " & txtSearchTemplateWhereFreehand & " ) "
            End If
        End If
            
        'Now see if we should display Deleted documents
        intViewDeletedDocuments = frmImaging101Retrieve.chkViewDeletedDocuments.Value
        If intViewDeletedDocuments = vbChecked Then
            '2014-12-20 - Jacob - ONLY Show Documents Deleted by Users... NOT via AutoImport
            txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked = 'D' ) "
        Else
            'Don't Display Deleted Documents
            '2014-12-20 - Jacob - Changed to Handle "D" deleted from Retrieval and "DA" deleted via AutoExport.
'                txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked <> 'D' OR DocumentLocked = '' OR DocumentLocked is NULL ) "
            txtFilterStatement = txtFilterStatement & " AND ( DocumentLocked NOT LIKE 'D%' OR DocumentLocked = '' OR DocumentLocked is NULL ) "
        End If

    End If
    
    
    '***********************************************************************************************************
    '*** 2023-01-31 - Jacob - BEGIN:  PREVENT SQL INJECTION
    
    Dim bolSQLInjectionDetected As Boolean
    Dim intSQLInjectionLoop As Integer
    Dim intPositionOfFieldBeforeEqualSign As Integer
    Dim bolFoundFieldName As Boolean
    Dim strTextBeforeEqualSign As String
    
    'Loop through entire txtFilterStatement
    For intSQLInjectionLoop = 1 To Len(txtFilterStatement)
    
        'Check for ASCII Characters > 127
        If Asc(Mid(txtFilterStatement, intSQLInjectionLoop, 1)) > 127 Then
                bolSQLInjectionDetected = True
        End If
        
        'Check for other attack vectors, there must ALWAYS be a FieldName before an Equal (=) sign.
'        If Mid(txtFilterStatement, intSQLInjectionLoop, 1) = "=" Then
'
'                bolFoundFieldName = False
'                'Loop through all Fields
'                For intIndex = 0 To lblFieldDescription.Count - 1
'                        Debug.Print UCase(txtFieldName(intIndex))
'                        If InStrRev(UCase(txtFilterStatement), UCase(txtFieldName(intIndex)), intSQLInjectionLoop - 1) > 0 Then
'                            bolFoundFieldName = True
'                            'Since we found a Field Name, exit the loop
'                            Exit For
'                        Else
'                            'bolFoundFieldName = False
'                        End If
'                Next
'
'                If bolFoundFieldName = False Then
'                    bolSQLInjectionDetected = True
'                End If
'
'        End If
        
    Next
    
    'Check for Keyword Attack vectors
    Dim strSelectKeywords As String
    Dim strModifyKeywords As String
    Dim strCommentKeywords As String
    Dim strAllKeywords As String
    Dim strAllKeywordsArray() As String
    
    'Do NOT include the keyword "union" because it would prevent searching for words like "Credit Union", etc.
    strSelectKeywords = "information_schema|insert|update|delete|truncate|drop|reconfigure|sysobjects|waitfor|xp_cmdshell|;"
    strModifyKeywords = "information_schema|truncate|drop|reconfigure|sysobjects|waitfor|xp_cmdshell"
    strCommentKeywords = "--|/*|*/"
    
    'Concatenate keywords
    strAllKeywords = strSelectKeywords + "|" + strModifyKeywords + "|" + strCommentKeywords
    
    'Split keywords into an Array to iterate through
    strAllKeywordsArray = Split(strAllKeywords, "|")
    
    'Check txtFilterStatement for all keywords
    For intSQLInjectionLoop = 0 To UBound(strAllKeywordsArray)
            If InStr(txtFilterStatement, strAllKeywordsArray(intSQLInjectionLoop)) > 0 Then
                    bolSQLInjectionDetected = True
            End If
    Next
    
    If bolSQLInjectionDetected = True Then
            Dim strSQLInjectionMessage As String
            strSQLInjectionMessage = "***   WARNING:  " + vbCrLf + "***         SQL INJECTION ATTEMPT DETECTED !!!" + vbCrLf + "***         SEARCH CANCELLED."
            funcWriteToDebugLog Me.name, strSQLInjectionMessage
            funcQuickMessage "SHOW", strSQLInjectionMessage
            Beep
            TimePause 2
            Beep
            TimePause 2
            Beep
            'Set WHERE to return NOTHING.  Can NOT just Exit Sub without the RecordSource being set.
            txtFilterStatement = " WHERE NULL IS NOT NULL "
    End If

    '*** 2023-01-31 - Jacob - END:    PREVENT SQL INJECTION
    '***********************************************************************************************************
 
    
    
    '*** WE ARE SETTING THE frmImaging101Retrieve (Search Results List) FORM CONTROLS
    Dim strSelectRange As String
    
    frmImaging101Retrieve.Adodc1.ConnectionString = RegImaging101ConnectionString
    frmImaging101Retrieve.Adodc1.ConnectionTimeout = 300
   
   
    '*** 2020-04-23 - Jacob - Added  DocumentBatchRECID to the retrieval list
    If chkDocumentCountOnly = vbChecked Then
            frmImaging101Retrieve.Adodc1.RecordSource = "SELECT COUNT(*) " & _
            " FROM " & txtApplicationName & _
            txtFilterStatement
            
            frmImaging101Retrieve.Adodc1.Refresh
            
            funcQuickMessage "SHOW", "Total Records Found = " & frmImaging101Retrieve.Adodc1.Recordset.Fields.item(0).Value
            Exit Sub
        
    ElseIf txtMaxItemsToRetrieve = "0" Then
        strSelectRange = "SELECT "
    Else
        strSelectRange = "SELECT TOP " & txtMaxItemsToRetrieve
    End If
    
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, txtFilterStatement
    funcWriteToDebugLog Me.name, ""
   

   frmImaging101Retrieve.Adodc1.RecordSource = strSelectRange & " " & _
                        " DocumentRECID , " & _
                        " DocumentLocked, DocumentLockedBy, DocumentLockedDate, DocumentLockExpDate " & _
                        txtFieldsList & ", " & _
                        " DocumentImages, DocumentPages, DocumentBatchName, DocumentBatchRECID, " & _
                        " BatchBoxNumber, " & _
                        " DocumentScanDate, DocumentScanUserID, " & _
                        " DocumentIndexDate , DocumentIndexUserID, " & _
                        " DocumentCommitDate , DocumentCommitUserID " & _
                        " FROM " & txtApplicationName & _
                        txtFilterStatement & _
                        " ORDER BY " & txtOrderByList

   
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, frmImaging101Retrieve.Adodc1.RecordSource
    funcWriteToDebugLog Me.name, ""
   
   frmImaging101Retrieve.Adodc1.Refresh
   
   frmImaging101Retrieve.lblItemsFound.Caption = frmImaging101Retrieve.Adodc1.Recordset.RecordCount
   
''   frmImaging101Retrieve.chkViewDocDetails_Click
''    frmImaging101Retrieve.SetFocus
    
   
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    
    If Err.Number = 13 Then
        result = MsgBox("Field Format Error: " & Err.Number & " - " & Err.Description & vbCrLf & "PLEASE CHECK YOUR INPUT.", vbOK)
        Me.SetFocus
        mebIndexValues(intIndex).SetFocus
        bolErrorOccured = True
        Err.Clear
        Exit Sub
    End If

    result = MsgBox("LoadFieldFormats - Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
    
End Sub


Private Sub txtIndexValues_Change(Index As Integer)


    If Not bolSearchFormLoadComplete Then
        Exit Sub
    End If
    
    If Len(txtIndexValues(Index).Text) > Int(txtFieldSize(Index).Text) Then
        MsgBox "Exceeded field size of " & txtFieldSize(Index).Text & " characters!" & _
                vbCrLf & "Truncating to defined size.", vbOKOnly
        txtIndexValues(Index).Text = Left(txtIndexValues(Index).Text, Int(txtFieldSize(Index).Text))
    End If
    
End Sub

Private Sub txtIndexValues_LostFocus(Index As Integer)

    'Check to Re-size the field
'    Call txtIndexValues_DblClick(Index)

        txtIndexValues(Index).BackColor = vbWhite
        txtIndexValues(Index).Height = mebIndexValues(Index).Height
        txtIndexValues(Index).SelStart = 0
        txtIndexValues(Index).SelLength = 0
     
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


Private Sub subSetDefaultFieldValues(intFieldIndex As Integer)
    
    'Check BOTH Text and Masked Edit controls... if they have a value... Exit now!
    If Trim((mebIndexValues(intFieldIndex).Text & txtIndexValues(intFieldIndex)) <> "") Then
       Exit Sub
    End If
    
    
    '* If the Field has a Default Value, set it now

        Dim txtTodaysDate As String
        txtTodaysDate = Format(Now(), "mm-dd-yyyy")
        
        '*** Check if the txtIndexValues TextBox control is VISIBLE...
        '    this will handle saving the value of the the TextBox control
        '    instead of the mebIndexValues Masked Edit control as needed.
        '    This is because the Masked Edit control has a MAX size of 64 Char.
        If txtIndexValues(intFieldIndex).Visible = True Then
            'Use the TEXT Control
            Select Case txtFieldDefaultValue(intFieldIndex)
                Case "[Scan Date]"
                    txtIndexValues(intFieldIndex).Text = txtTodaysDate
                Case "[Index Date]"
                    txtIndexValues(intFieldIndex).Text = txtTodaysDate
                Case Else
                    txtIndexValues(intFieldIndex).Text = txtFieldDefaultValue(intFieldIndex)
            End Select
        
        
        Else
            'Use the MASKED EDIT Control
            Select Case txtFieldDefaultValue(intFieldIndex)
                Case "[Scan Date]"
                    mebIndexValues(intFieldIndex).Text = Format(txtTodaysDate, mebIndexValues(intFieldIndex).Format)
                Case "[Index Date]"
                    mebIndexValues(intFieldIndex).Text = Format(txtTodaysDate, mebIndexValues(intFieldIndex).Format)
                Case Else
                    mebIndexValues(intFieldIndex).Text = txtFieldDefaultValue(intFieldIndex)
            End Select
            
        End If ' txtIndexValues(intFieldIndex).Visible = True

        

End Sub



