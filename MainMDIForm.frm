VERSION 5.00
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.6#0"; "PIXEZOCX.OCX"
Begin VB.MDIForm MainMDIForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Viewer - Imaging101"
   ClientHeight    =   7905
   ClientLeft      =   1320
   ClientTop       =   750
   ClientWidth     =   9105
   Icon            =   "MainMDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox PictureButtonBar 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9105
      TabIndex        =   1
      Top             =   0
      Width           =   9105
      Begin VB.PictureBox picImaging101Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   7320
         Picture         =   "MainMDIForm.frx":0CCA
         ScaleHeight     =   405
         ScaleWidth      =   1440
         TabIndex        =   26
         Top             =   0
         Width           =   1440
      End
      Begin VB.TextBox txtFindText 
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Thumbnails"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdZoomOut 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zoom Out"
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
         Picture         =   "MainMDIForm.frx":135D
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Zoom Out"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Launch to Associated Application"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
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
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdSendTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "SendTo"
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
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Send To (eMail)"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdThumbNails 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Thumbnails"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Thumbnails"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdZoomFit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fit to Window"
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
         Left            =   600
         Picture         =   "MainMDIForm.frx":18E7
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fit Image to Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CheckBox chkStayOnTop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "StayOnTop"
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
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnnotate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Annotate"
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
         Left            =   3720
         Picture         =   "MainMDIForm.frx":1E71
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Annotate image"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdLaunch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Launch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Launch to Associated Application"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdImageRotateLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rotate"
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
         Picture         =   "MainMDIForm.frx":23FB
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Rotate Left"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdImageRotateRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rotate"
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
         Left            =   3120
         Picture         =   "MainMDIForm.frx":2985
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Rotate Right"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdZoomIn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zoom In"
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
         Left            =   1320
         Picture         =   "MainMDIForm.frx":2F0F
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Zoom In"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdGotoPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Go to Page"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Go to a Page by Number"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Page"
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
         Left            =   4560
         Picture         =   "MainMDIForm.frx":3499
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Previous Page"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdNextPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Page"
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
         Left            =   5640
         Picture         =   "MainMDIForm.frx":3823
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Next Page"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdSaveZoom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save Zoom"
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
         Left            =   1920
         Picture         =   "MainMDIForm.frx":3BAD
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Zoom"
         Top             =   0
         Width           =   615
      End
      Begin VB.CheckBox chkViewAnnotations 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Annotations"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1932
      End
      Begin VB.CommandButton cmdPrevImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<- Image"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdNextImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Image ->"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdEnhance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enhance OFF"
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
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Enhance OFF"
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdGotoImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Goto Image"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin PixezocxLib.PixEzImage PixEzImage1 
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _Version        =   65542
         _ExtentX        =   661
         _ExtentY        =   873
         _StockProps     =   100
         TAG_ENV_TIFFJPEGCOMPRESSION=   6
         TAG_ENV_ANNOTATIONVERSION=   1
         TAG_ENV_PDFRENDERING=   0
         TAG_OPEN_DIR    =   ""
         TAG_OPEN_SCHEMA =   ""
         TAG_OPEN_EXT    =   ""
         TAG_OPEN_ROOT   =   ""
         TAG_OPEN_DETECTSCHEMA=   1
         TAG_OPEN_FILENAMES=   ""
         TAG_WINDOW_CURPAGE=   0
         PIXEZ_SELECT    =   ""
         TAG_BORDER_COLOR_BG=   6579300
         TAG_BORDER_COLOR_ENVFOCUS=   254
         TAG_BORDER_COLOR_WINFOCUS=   16711422
         TAG_BRIGHTNESS  =   128
         TAG_CONTRAST    =   128
         TAG_BLUEBRIGHTNESS=   128
         TAG_BLUECONTRAST=   128
         TAG_GREENBRIGHTNESS=   128
         TAG_GREENCONTRAST=   128
         TAG_REDBRIGHTNESS=   128
         TAG_REDCONTRAST =   128
         TAG_DOC_OPENATTRIBUTE=   66
         TAG_FILLORDER   =   1
         TAG_HFLIP       =   0
         TAG_PAN_HEIGHT  =   0
         TAG_PAN_WIDTH   =   0
         TAG_PAN_XPOS    =   0
         TAG_PAN_YPOS    =   0
         TAG_PAN_SCALING =   4
         TAG_PAN_TITLE   =   "Pan Window"
         TAG_PAN_SHOW    =   0
         TAG_ONE_ACCELMODE=   0
         TAG_ONE_ACTION_CLOCKWISE=   35
         TAG_ONE_ACTION_CTRCLOCKWISE=   3
         TAG_ONE_ACTION_DEFINEREG=   64
         TAG_ONE_ACTION_DEFINEREGASPECT=   64
         TAG_ONE_ACTION_PAN=   1
         TAG_ONE_ACTION_SWITCHTOTREE=   64
         TAG_ONE_ACTION_ZOOMINREG=   64
         TAG_ONE_ACTION_ZOOMINREGASPECT=   64
         TAG_ONE_ACTION_ZOOMOUTCORNER=   44
         TAG_ONE_ACTION_CONTEXTMENU=   32
         TAG_ONE_ACTION_ANNOTATIONITEMTRIGGER=   12
         TAG_ONE_ACTION_ANNOTATIONMODEACTION=   0
         TAG_ONE_ACTION_ANNOTATIONMODEACTIONASP=   64
         TAG_ONE_SCROLLBARS=   2
         TAG_ONE_SETTINGS_RANGE=   0
         TAG_ONE_MOUSEOPTION=   0
         TAG_ORIENTATION =   1
         TAG_OVERSCAN    =   0
         TAG_PHOTOMETRICINTERPRETATION=   0
         TAG_PRINT_COLLATE=   1
         TAG_PRINT_COPIES=   1
         TAG_PRINT_DEVICENO=   0
         TAG_PRINT_DEVNAME1=   ""
         TAG_PRINT_DEVNAME2=   ""
         TAG_PRINT_RANGEMODE=   0
         TAG_PRINT_REGION=   0
         TAG_PRINT_SCALE =   0
         TAG_PRINT_SHOWDLG=   0
         TAG_REGION_COUNT=   0
         TAG_REGION_MODE =   0
         TAG_ROTATION    =   1
         TAG_SCALING     =   4
         TAG_DITHER      =   1
         TAG_VIEWASGRAY  =   0
         TAG_SCALE_X     =   1
         TAG_SCALE_Y     =   1
         TAG_SCAN_ALLOW_TURNOVER=   0
         TAG_SCAN_COLORFORMAT=   8388608
         TAG_SCAN_COMPRESSION=   4
         TAG_SCAN_CURPAGE=   0
         TAG_SCAN_DISPLAYPAGE=   0
         TAG_SCAN_DIR    =   ""
         TAG_SCAN_DUPLEX =   0
         TAG_SCAN_EXT    =   ""
         TAG_SCAN_FILENAME=   ""
         TAG_SCAN_INSERTMODE=   1
         TAG_SCAN_SCHEMA =   ""
         TAG_SCAN_WARNOVERWRITE=   0
         TAG_SCAN_MULTIPAGE=   0
         TAG_SCAN_USESCHEMA=   0
         TAG_SCAN_MAXPAGES=   -1
         TAG_SCAN_ORIENTATION=   1
         TAG_SCAN_PACK   =   196608
         TAG_SCAN_PRECEDENCE=   1
         TAG_SCAN_ROOT   =   ""
         TAG_SCAN_USELONGNAMES=   0
         TAG_SAVE_MERGEANNOTATIONS=   0
         TAG_SAVE_COLORFORMAT=   0
         TAG_SAVE_COMPRESSION=   4
         TAG_SAVE_DIR    =   ""
         TAG_SAVE_EXT    =   ""
         TAG_SAVE_FILENAME=   ""
         TAG_SAVE_ORIENTATION=   1
         TAG_SAVE_PACK   =   196608
         TAG_SAVE_PRECEDENCE=   1
         TAG_SAVE_RANGESTR=   ""
         TAG_SAVE_ROOT   =   ""
         TAG_SAVE_WARNOVERWRITE=   0
         TAG_SAVE_MULTIPAGE=   1
         TAG_SAVE_USESCHEMA=   0
         TAG_SAVE_USELONGNAMES=   0
         TAG_THRESH_X    =   0
         TAG_THRESH_Y    =   0
         TAG_TREE_COLOR_BG=   8421504
         TAG_TREE_COLOR_NODETEXT=   0
         TAG_TREE_COLOR_NODESELTEXT=   16777215
         TAG_TREE_COLOR_THUMBTEXT=   0
         TAG_TREE_COLOR_THUMBSELTEXT=   16777215
         TAG_TREE_COLOR_LINE=   0
         TAG_TREE_THUMBSTYLE=   528
         TAG_TREE_UIFLAGS=   1280
         TAG_WINDOW_STYLE=   0
         TAG_XPOSITION   =   0
         TAG_YPOSITION   =   0
         TAG_INVERT      =   0
         TAG_SCAN_SCANROT=   0
         TAG_SCAN_AUTOCOLORFORMAT=   0
         TAG_TREE_cxTHUMBNAIL=   34
         TAG_TREE_cxTHUMBCELL=   52
         TAG_TREE_cyTHUMBNAIL=   44
         TAG_TREE_cyTHUMBLINE=   72
         TAG_SAVE_PROGRESSMODE=   0
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
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   7320
         TabIndex        =   27
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   975
      Width           =   9105
   End
   Begin VB.Menu mnuBatch 
      Caption         =   "&Batch"
      Begin VB.Menu mnuBatchReplaceCurrentPage 
         Caption         =   "Replace Current Page"
      End
      Begin VB.Menu mnuBatchInsertPagesBefore 
         Caption         =   "Insert Pages BEFORE Current Page"
      End
      Begin VB.Menu mnuBatchInsertPagesAfter 
         Caption         =   "Insert Pages AFTER Current Page"
      End
      Begin VB.Menu mnuBatchAppendPages 
         Caption         =   "Append Pages to Batch"
      End
      Begin VB.Menu mnuBatchDeletePage 
         Caption         =   "Delete Page from Batch"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuScaleToGray 
         Caption         =   "Enhance On/Off"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLockRotation 
         Caption         =   "Lock Rotation"
      End
   End
   Begin VB.Menu Window 
      Caption         =   "&Window"
      Begin VB.Menu mnuWTileHorizontal 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuWTileVertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuWCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange"
      End
   End
   Begin VB.Menu mnuSwitchTo 
      Caption         =   "&SwitchTo"
      Index           =   0
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnu_MarkupControl 
      Caption         =   "&Markup Control"
      Visible         =   0   'False
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu AboutDoc 
         Caption         =   "About Doc"
      End
      Begin VB.Menu AboutView 
         Caption         =   "About View"
      End
      Begin VB.Menu AboutMarkup 
         Caption         =   "About Markup"
      End
      Begin VB.Menu AboutImaging101 
         Caption         =   "About Imaging101"
      End
   End
End
Attribute VB_Name = "MainMDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim docContents As IDocContents
    Dim ActivePage As IActivePage
    ' frmViewForm is the equivalent of "Me.ActiveForm"
    Dim frmViewForm As ChildForm1
    Dim txtFullPathName As String
    Dim boolPageRotated As Boolean
    Dim strScanMode As String
    
    Dim m_ImageCount As Long
    Dim m_PageCount As Long
    Dim m_ImageSkipCount As Long
    
    Dim dblHoldBatchPageOrder As Long
    
    Dim intLoop As Integer
    Dim intSwitchToHold As Integer
    

    
    
    
Private Sub AboutDoc_Click()
    
    On Error Resume Next
    
    Me.ActiveForm.SpicerDoc1.AboutBox
    
    
    
End Sub

Private Sub AboutImaging101_Click()

    On Error Resume Next
    
    frmAbout.Show
    
End Sub

Private Sub AboutMarkup_Click()
    
    On Error Resume Next
    
    Me.ActiveForm.SpicerMarkup1.AboutBox
    
End Sub

Private Sub AboutView_Click()
    
    On Error Resume Next
    
    Me.ActiveForm.SpicerView1.AboutBox
    
End Sub

Private Sub chkStayOnTop_Click()

    If chkStayOnTop = vbChecked Then
        funcMakeTopMost MainMDIForm, True
    Else
        funcMakeTopMost MainMDIForm, False
    End If
    
End Sub

Private Sub chkViewAnnotations_Click()

    On Error Resume Next
    
    Call Me.ActiveForm.subAnnotationLayerShowHide

End Sub

Public Sub cmdAnnotate_Click()

    On Error GoTo ERROR_HANDLER
    
    If funcIsFormLoaded2("frmAnnotate") Then
        Me.ActiveForm.subAnnotationLayerSaveCheck
        Unload frmAnnotate
        Exit Sub
    End If
    
    '*** 2022-11-06 - Jacob - Added check for intI101Module = gI101ModuleIndex to allow Annotations while Indexing Batches,
    '                                         becasue UBound is always Zero
    If intI101Module = gI101ModuleIndex _
    Or (UBound(arrDisplayedPagesRetrieve) + UBound(arrDisplayedPagesIndex) > 0) Then
    '    frmAnnotate.Top = Me.Top
    '    frmAnnotate.Left = Me.Left - frmAnnotate.width
        frmAnnotate.Show
        Me.ActiveForm.subAnnotationLayerCreate
    End If
    
Exit Sub

ERROR_HANDLER:
    MsgBox "cmdAnnotate_Click:  Error with Annotations... Document not available!"
    Unload frmAnnotate
    
End Sub

Public Sub mnu_Markup_Click()
   On Error GoTo ErrorOccurred
   
''   Call IUserTools_ActiveTool(19)
   Exit Sub
   
ErrorOccurred:
   Exit Sub
End Sub





Private Sub cmdEdit_Click()
    
    On Error Resume Next
    
    'Launch ORIGINAL page/document
    Me.ActiveForm.subLaunch "Edit"

    Exit Sub
    
cmdEdit_ERROR:
    MsgBox "ERROR: " & Err.Number & " - " & Err.Description & _
    vbCrLf & "Unable to Launch document for Edit."
    
End Sub

Public Sub cmdEnhance_Click()


    'Toggle ScaleToGray enhancement On/Off
    mnuScaleToGray_Click
    
    'Execute the cmdScaleToGray Sub in the Active MDI Child form
    Me.ActiveForm.cmdScaleToGray
    
End Sub

Private Sub cmdFind_Click()

    On Error GoTo ErrorOccurred
    
   Dim ActivePage As IActivePage
   Dim sText As String
   ' Set the object variable for the IActivePage interface to the View Control object
   Set ActivePage = Me.ActiveForm.SpicerView1.object
   ' Enter the exact text to find
'   sText = InputBox("Enter the exact text to find.", "Find Text")
   sText = txtFindText
   ' Find the text, matching the case and searching for the next match
   ActivePage.FindTextMatch sText, IN_TXTSRCH_CASEINSENSITIVE, IN_DIR_NEXT

   ' De-initialize the object variable
   Set ActivePage = Nothing
   
Exit Sub

ErrorOccurred:
    
'    funcQuickMessage "SHOW", "You have reached the end of the document."
'   ActivePage.FindTextMatch sText, IN_TXTSRCH_CASEINSENSITIVE, IN_DIR_FIRST

    Resume Next

End Sub

Private Sub cmdGotoImage_Click()

    Me.ActiveForm.subGotoImage
    

End Sub

Public Sub cmdLaunch_Click()

    On Error Resume Next
    
    'Launch page/document
    Me.ActiveForm.subLaunch
    
End Sub

Private Sub Command1_Click()
    '***************************************************
    '*** SAVE ALL PAGES AS A MULTIPAGE TIFF
    Dim txtAttachmentFileName As String
    txtAttachmentFileName = InputBox("What would you like to call THIS document?" & vbCrLf & "Press [Cancel] or the [Esc] key to cancel the send.", "Attachment Name", "Imaging101 Document")
    
    'Check if the user Canceled or entered no filename
    If Trim(txtAttachmentFileName) = "" Then
'        GoTo SESSION_LOGOFF
    End If
'    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".TIF"
    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".PDF"
    
    '*** Rasterize the Pages before sending
'     Me.ActiveForm.subRasterizeBatchEX
'     Me.ActiveForm.subRasterizeBatch
    
    Dim docSave As IDocSave
    ' Set the object variable for the IDocSave interface to the Document Control object
    ' that was saved by the Rasterize sub
    Set docSave = Me.ActiveForm.SpicerDoc1.object
    
'    docSave.SaveAsDialog False
    
    ' Save the modified pages in the Spicer Document format
    If Me.ActiveForm.txtPageCount > 1 Then
'        docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, txtAttachmentFileName
        docSave.Save 0, False, 619, txtAttachmentFileName, txtAttachmentFileName
    Else
'        docSave.Save 0, False, API_FF_TIFFM, txtAttachmentFileName, txtAttachmentFileName
        docSave.Save 0, False, 101, txtAttachmentFileName, txtAttachmentFileName
    End If


    ' De-initialize the object variable
    Set docSave = Nothing
    '***************************************************

End Sub

Private Sub cmdNextImage_Click()

    Me.ActiveForm.subGetNextImage
    
End Sub

Private Sub cmdPrevImage_Click()

    Me.ActiveForm.subGetPrevImage

End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrorOccurred

   Dim PrintView As IPrintView
   Dim msgPrint As String
   Dim printChoice
   ' Set the object variable for the IPrintView interface to the View Control object
   Set PrintView = Me.ActiveForm.SpicerView1.object
   ' Ask what print command to use
''   msgPrint = "Do you want to use the Print dialog box? " + Chr(13) + _
              "Click Yes to use the Print dialog box. " + Chr(13) + _
              "Click No to use the PrintDocument method."

''   printChoice = MsgBox(msgPrint, vbYesNo + vbQuestion, "Print Choices")
''   If printChoice = vbYes Then
      PrintView.PrintDialog 'Display the Print dialog box
''   Else
      ' Print one copy of all pages of the document in the active window.
      ' Do not print a banner or a stamp on it.
''      PrintView.PrintDocument IN_PRINT_ALL_PAGES, 0, 0, 1, IN_PMODE_DOCUMENT, _
                           False, IN_ZOOM_SCALETOFIT, IN_ORIENT_BEST_FIT, _
                           False, False

''   End If
   'De-initialize the object variable
   Set PrintView = Nothing
   Exit Sub
   
ErrorOccurred:
''''   ErrorHandler
   MsgBox "Sorry an error has occured... Error #" & Err.Number & " - " & Err.Description
   Exit Sub

End Sub

Private Sub cmdSaveZoom_Click()
    
    On Error Resume Next
    
    Dim ScaleScrollRotation As IScaleScrollRotation
    Dim dScale As Double
    Dim X As Long
    Dim Y As Long
    Dim ZoomLevel As Integer
    
    ' Set the object variable for the IScaleScrollRotation interface to the View Control object
    Set ScaleScrollRotation = Me.ActiveForm.SpicerView1.object
    ' Find the scale factor and x,y coordinates for the active page
''''    ScaleScrollRotation.GetZoomFactor 0, dScale, x, y
    dScale = Me.ActiveForm.SpicerView1.ZoomFactorScale(0)
    X = Me.ActiveForm.SpicerView1.ZoomFactorX(0)
    Y = Me.ActiveForm.SpicerView1.ZoomFactorY(0)
    ZoomLevel = Me.ActiveForm.SpicerView1.ZoomLevel(0)
    
    result = WritePrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomFactorScale", dScale, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomFactorX", X, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomFactorY", Y, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomLevel", ZoomLevel, RegFileName)
    
    ' De-initialize the IMiscellaneous object variable
    Set ScaleScrollRotation = Nothing
    

    
End Sub

Private Sub cmdSendTo_Click()

    Dim bolSendToSMTP As String
    
    On Error Resume Next
    
'    Me.ActiveForm.subSendTo
    bolSendToSMTP = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Control", "ID = 1", "SendEmailViaSMTP")
    
    If bolSendToSMTP = True Then
        Me.ActiveForm.subSendToSMTP
    Else
        Me.ActiveForm.subSendToOutlook
    End If


End Sub

Private Sub cmdThumbNails_Click()
    
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    
    If funcIsFormLoaded2("frmThumb") Then
        'If the form was already loaded,
        '  and the user clicked the button again... close it
        Unload frmThumb
        Set frmThumb = Nothing
    Else
        DoEvents
        frmThumb.Show
        DoEvents
        frmThumb.SpicerThumbnail1.BindToViewControl Me.ActiveForm.SpicerView1.object
        frmThumb.SpicerThumbnail1.Visible = True

        Me.SetFocus
'        'Walk the images Backwards to end up on Page 1
'        For intIndex = 1 To CInt(Me.ActiveForm.StatusBar1.Panels(4).Text)
'            Me.cmdNextPage_Click
'            DoEvents
'        Next
'        Me.Show
    End If
    
    Screen.MousePointer = vbNormal

End Sub



Private Sub cmdZoomFit_Click()
    
    On Error Resume Next
    
    Me.ActiveForm.cmdZoomFit_Click

End Sub



Private Sub Command2_Click()


End Sub

Private Sub mnuLockRotation_Click()

    On Error Resume Next
    
   If Me.mnuLockRotation.Checked Then
        Me.mnuLockRotation.Checked = vbUnchecked
        Me.mnuLockRotation.Caption = "Lock Rotation"
   Else
        ' Toggle the menu item between checked and not checked
        Me.mnuLockRotation.Checked = vbChecked
        Me.mnuLockRotation.Caption = "Unlock Rotation"
        
        Me.ActiveForm.funcLockRotation
        
    End If

        

End Sub



Private Sub MDIForm_Load()
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    If bolDebug Then
        Me.Caption = Me.Caption & " - DEBUG"
    End If

    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{"
    funcWriteToDebugLog Me.name, "ENTERING MainMDIForm.Form_Load"
    
    ' Get saved settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "MainMDIForm.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "MainMDIForm.Left", RegFileName)
    Me.width = VBGetPrivateProfileString(RegAppname, "MainMDIForm.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "MainMDIForm.Height", RegFileName)
    On Error GoTo 0
    
    
    '***************************************************************
    '***  CHECK BUTTON SECURITY RIGHTS
    
    subCheckButtonSecurity
    
    
    '*** Initialize the Displayed Pages Array
    ReDim arrDisplayedPagesRetrieve(0)
    ReDim arrDisplayedPagesIndex(0)
'    ReDim gFormArrayRetrieve(0)
'    ReDim gFormArrayIndex(0)
    
''''    PictureButtonBar.Height = cmdPrevPage.Height + cmdPrevImage.Height + cmdZoomIn.Height + 10
    
    

End Sub

Public Sub subCheckButtonSecurity()

    '***************************************************************
    '***  CHECK BUTTON SECURITY RIGHTS
    
    '*** 2020-05-27 - Jacob - MOVED BUTTON SETTINGS FROM funcShowImage()
'
'        cmdEdit.enabled = True
        cmdEnhance.Enabled = True
        cmdZoomFit.Enabled = True
        cmdZoomIn.Enabled = True
        cmdZoomOut.Enabled = True
        cmdSaveZoom.Enabled = True
        cmdGotoImage.Enabled = True
        cmdImageRotateLeft.Enabled = True
        cmdImageRotateRight.Enabled = True

    If gsecRightsSendMail = vbChecked Then
        cmdSendTo.Visible = True
        cmdSendTo.Enabled = True
    Else
        cmdSendTo.Visible = False
    End If
    
    If gsecRightsLaunchDoc = vbChecked Then
        cmdLaunch.Visible = True
        cmdLaunch.Enabled = True
    Else
        cmdLaunch.Visible = False
    End If
    
    If gsecRightsPrint = vbChecked Then
        cmdPrint.Visible = True
        cmdPrint.Enabled = True
    Else
        cmdPrint.Visible = False
    End If
    
    If gsecRightsAnnotate = vbChecked Then
        cmdAnnotate.Visible = True
        cmdAnnotate.Enabled = True
    Else
        cmdAnnotate.Visible = False
    End If
    
    If gsecRightsThumbnails = vbChecked Then
        cmdThumbNails.Visible = True
        cmdThumbNails.Enabled = True
    Else
        cmdThumbNails.Visible = False
    End If
    
    

End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Do NOT allow this form to be unloaded while it's still loading.
'    If funcIsFormLoaded2("frmIndex") And (Not bolIndexFormLoadComplete) Then
    '2009-12-01 - Jacob - Changed to prevent unloading by User during indexing
    If (funcIsFormLoaded2("frmIndex") Or bolAIM_Command_AddFile = True) And UnloadMode = vbFormControlMenu Then
        MsgBox "Sorry, the Viewer cannot be unloaded " & vbCrLf & "while the Indexing window is open!", vbInformation, "Indexing window open"
        Cancel = True
        Exit Sub
    End If
    
    

End Sub

Private Sub MDIForm_Resize()


        picImaging101Logo.Left = Me.ScaleWidth - picImaging101Logo.width - 10
        lblVersion.Left = picImaging101Logo.Left

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    On Error Resume Next
    
    funcWriteToDebugLog Me.name, ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
    funcWriteToDebugLog Me.name, "ENTERING:  MDIFORM.Form_Unload()"
    
    funcWriteToDebugLog Me.name, "UBound(arrDisplayedPagesRetrieve) = " & UBound(arrDisplayedPagesRetrieve)
    funcWriteToDebugLog Me.name, UBound(arrDisplayedPagesRetrieve)

    'Save Form settings to the registry
    funcWriteToDebugLog Me.name, "Save Form settings to the registry"
    
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "MainMDIForm.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "MainMDIForm.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "MainMDIForm.Width", Me.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "MainMDIForm.Height", Me.Height, RegFileName)
    End If

    On Error GoTo ERROR_HANDLER
    
    If funcIsFormLoaded2("frmThumb") Then
        funcWriteToDebugLog Me.name, "Unload frmThumb"
        Unload frmThumb
        funcWriteToDebugLog Me.name, "Set frmThumb = Nothing"
        Set frmThumb = Nothing
    End If
    
    If funcIsFormLoaded2("frmAnnotate") Then
        funcWriteToDebugLog Me.name, "Unload frmAnnotate"
        Unload frmAnnotate
        funcWriteToDebugLog Me.name, "Set frmAnnotate = Nothing"
        Set frmAnnotate = Nothing
    End If
    
    
    'SAFE way of saying: Set Me = Nothing
    funcWriteToDebugLog Me.name, "BEGIN SAFE UNLOAD of MainMDIForm"
    Dim Form As Form
    For Each Form In Forms
            If Form Is Me Then
                    funcWriteToDebugLog Me.name, "MainMDIForm = Nothing"
                    Set Form = Nothing
                    funcWriteToDebugLog Me.name, "Exit For"
                    Exit For
            End If
    Next Form
    
    funcWriteToDebugLog Me.name, "EXIT SUB - MainMDIForm"
    funcWriteToDebugLog Me.name, "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, "}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}"
    funcWriteToDebugLog Me.name, ""
    
Exit Sub

ERROR_HANDLER:
    
    funcWriteToDebugLog Me.name, "############################################"
    funcWriteToDebugLog Me.name, "Entering MainMDIForm.Unload ERROR_HANDLER"
    funcWriteToDebugLog Me.name, "MainMDIForm.Unload ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "frmIndex.Unload ERROR: " & Err.Number & " - " & Err.Description, vbInformation
    funcWriteToDebugLog Me.name, "Resume Next"
    funcWriteToDebugLog Me.name, "############################################"
    Resume Next
    



End Sub

Private Sub mnuBatchAppendPages_Click()


    On Error GoTo ERROR_HANDLER
    
    m_ImageCount = 0
    m_PageCount = 0
    m_ImageSkipCount = 0
    
    strScanMode = "Append Pages"

    subScanBatch "A"

    
    
Exit Sub

ERROR_HANDLER:
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
    End If


End Sub
Private Sub subScanBatch(strScanModePrefix As String)

    On Error GoTo ERROR_HANDLER

    Dim rc As Long

    PixEzImage1.ScanStateFlush

    PixEzImage1.ScanAllowLongNames = 1     ' Allow long names
    PixEzImage1.ScanInsertMode = 2         ' overwrite
    PixEzImage1.ImageFileSchemaDetect = 0  ' We are specifying a schema, so do not try to detect
    
    PixEzImage1.ScanUseSchema = 1
    PixEzImage1.ScanFileDir = Trim(frmIndex.txtBatchDirectory)  ' Set directory for images
    PixEzImage1.SavePrecedence = 0   ' 0 = Let Color Format setting determine file types
    PixEzImage1.ScanFileRoot = CStr(Format(frmIndex.txtBatchRECID, "0000000000")) & "-" & Format(Now(), "yymmddhhmmss") & "-A"   ' Set schema root. All files will begin with this
     
    ' Make sure the ScanFileSchema does not have more than nine (9) pound signs (#'s) otherwise
    '  we get spaces instead of zero's
    PixEzImage1.ScanFileSchema = "$####;" ' Set schema: root name plus two digits.
    
    rc = PixEzImage1.ScanSelect
    
    If rc = -4526 Then
        Exit Sub
    End If
    
'    PixEzImage1.ShowScannerSettingsDialog
    rc = frmScannerSettings.StartForm(PixEzImage1)

    If rc = 0 Then
        Exit Sub
    End If

'    '*** See if user wants to continue the scan
'    If Not frmContinueBatchScanSimple.subScanContinue Then
'        Exit Sub
'    End If

    PixEzImage1.ScanFileExt = funcPixGetFileExt(PixEzImage1)
    

    PixEzImage1.ScanBatch



ERROR_HANDLER:
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
    End If

End Sub
Public Sub mnuBatchDeletePage_Click()

    On Error GoTo ERROR_HANDLER
    
    '03-21-2017 - DO NOT Allow Deleting if only ONE page left
    If frmIndex.ListView1.ListItems.Count = 1 Then
        result = MsgBox("SORRY!  You CANNOT delete the ONLY page left." & vbCrLf & "Please EXIT Indexing and Delete the BATCH instead.", vbInformation, "Cannot Delete Last Page")
        Exit Sub
    End If
    
    frmIndex.cmdDeleteSelectedPage.BackColor = vbYellow
    
    'Delete the currently selected page
    result = MsgBox("Are you Sure you wish to Delete the Selected Page? ", vbYesNo, "Delete Batch Page")
    If result <> vbYes Then
        frmIndex.cmdDeleteSelectedPage.BackColor = vbWhite
        Exit Sub
    End If
    
    frmIndex.cmdDeleteSelectedPage.BackColor = vbRed
    
    'Hold to Current Position in the List (Page number)
    dblHoldBatchPageOrder = frmIndex.ListView1.ListItems.item(frmIndex.ListView1.SelectedItem.Index).Text
    
    '*** DELETE the Page
    subDeleteBatchPageRecord
    
    'Re-populate the List of Pages
    Call frmIndex.subLoadPagesIntoListView
    
    If dblHoldBatchPageOrder > frmIndex.ListView1.ListItems.Count Then
        'Set the list index to the last image on the list
        frmIndex.ListView1.ListItems.item(frmIndex.ListView1.ListItems.Count).Selected = True
    Else
        'Set the list index to the same position on the list
        frmIndex.ListView1.ListItems.item(dblHoldBatchPageOrder).Selected = True
    End If

    'Re-display the image
    Call frmIndex.ListView1_Click
    
    frmIndex.cmdDeleteSelectedPage.BackColor = vbWhite

Exit Sub

ERROR_HANDLER:
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
    End If


End Sub

Private Sub mnuBatchInsertPagesAfter_Click()

    On Error GoTo ERROR_HANDLER
    
    m_ImageCount = 0
    m_PageCount = 0
    m_ImageSkipCount = 0
    
    strScanMode = "Insert Pages AFTER"
    
    subScanBatch "IA"

    
    
Exit Sub

ERROR_HANDLER:
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
    End If

End Sub

Private Sub mnuBatchInsertPagesBefore_Click()

    On Error GoTo ERROR_HANDLER
    
    m_ImageCount = 0
    m_PageCount = 0
    m_ImageSkipCount = 0
    
    strScanMode = "Insert Pages BEFORE"
    
    subScanBatch "IB"

    
Exit Sub

ERROR_HANDLER:
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
    End If

End Sub

Private Sub mnuBatchReplaceCurrentPage_Click()

    On Error GoTo ERROR_HANDLER
    
    Dim strOriginalFileName As String
    Dim strTempFileName As String
    Dim rc As Long
    
    funcWriteToDebugLog Me.name, "******** ENTERING mnuBatchReplaceCurrentPage_Click() *************"
    strScanMode = "Replace Page"
    
    rc = PixEzImage1.ScanSelect
    
    '*** See if user wants to continue the scan
'    If Not frmContinueBatchScanSimple.subScanContinue Then
    If rc = -4526 Then
        Exit Sub
    End If
    
'    PixEzImage1.ShowScannerSettingsDialog
'    funcWriteToDebugLog  Me.Name, Err.Number & " " & Err.Description
    
    rc = frmScannerSettings.StartForm(PixEzImage1)

    If rc = 0 Then
        Exit Sub
    End If
    
    'Close the displayed document to allow renaming it
    Me.ActiveForm.SpicerDoc1.CloseDocument False

    
    txtFullPathName = frmIndex.txtBatchDirectory & "\" & frmIndex.txtBatchPageFileName
    strTempFileName = frmIndex.txtBatchDirectory & "\" & "REPLACE.TMP"
    strOriginalFileName = txtFullPathName
    
    'Rename the original file in case the scan fails
    On Error Resume Next
    funcWriteToDebugLog Me.name, "Name " & strOriginalFileName & " As " & strTempFileName
    Name strOriginalFileName As strTempFileName
    
    If Err.Number <> 0 Then
        funcWriteToDebugLog Me.name, "mnuBatchReplaceCurrentPage ERROR: " & Err.Number & "  Description: " & Err.Description & _
                vbCrLf & "UNABLE to RENAME file:" & vbCrLf & strOriginalFileName & _
                vbCrLf & " TO " & vbCrLf & strTempFileName

        funcQuickMessage "SHOW", "mnuBatchReplaceCurrentPage ERROR: " & Err.Number & "  Description: " & Err.Description & _
                vbCrLf & "UNABLE to RENAME file:" & vbCrLf & strOriginalFileName & _
                vbCrLf & " TO " & vbCrLf & strTempFileName
        Exit Sub
    End If
    
    On Error GoTo ERROR_HANDLER

    funcWriteToDebugLog Me.name, "PixEzImage1.ScanFileName = " & strOriginalFileName

    PixEzImage1.ScanFileName = strOriginalFileName
    PixEzImage1.ScanUseSchema = False
    PixEzImage1.ScanSingle
    PixEzImage1.FlushPages
    PixEzImage1.Close
    
    'Scan succeeded... zap Temp file
    On Error Resume Next
    funcWriteToDebugLog Me.name, "Kill " & strTempFileName
    Kill strTempFileName
    
    If Err.Number <> 0 Then
        funcWriteToDebugLog Me.name, "mnuBatchReplaceCurrentPage ERROR: " & Err.Number & "  Description: " & Err.Description & _
                vbCrLf & "UNABLE to DELETE file: " & vbCrLf & strTempFileName
        funcQuickMessage "SHOW", "mnuBatchReplaceCurrentPage ERROR: " & Err.Number & "  Description: " & Err.Description & _
                vbCrLf & "UNABLE to DELETE file: " & vbCrLf & strTempFileName
        Exit Sub
    End If
    
    'Re-display the image
    Call frmIndex.ListView1_Click
    
Exit Sub

ERROR_HANDLER:
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        funcQuickMessage "SHOW", "mnuBatchReplaceCurrentPage ERROR: " & Err.Number & "  Description: " & Err.Description & _
                "Will attempt to Re-name  the TEMPORARY file :" & vbCrLf & strTempFileName & _
                vbCrLf & " BACK TO it's Original File Name: " & _
                vbCrLf & strOriginalFileName
                
        'Rename the Temp file back to the original name so Index doesn't freak out!
        On Error Resume Next
        Name strTempFileName As strOriginalFileName
        
        If Err.Number <> 0 Then
            funcQuickMessage "SHOW", "mnuBatchReplaceCurrentPage ERROR: " & Err.Number & "  Description: " & Err.Description & _
                    vbCrLf & "UNABLE to RENAME file:" & vbCrLf & strTempFileNames & _
                    vbCrLf & "BACK TO: " & vbCrLf & trOriginalFileName
        End If
        
    End If
    
    Call frmIndex.ListView1_Click
    
End Sub

Public Sub mnuScaleToGray_Click()

    On Error Resume Next
    
   If Me.mnuScaleToGray.Checked Then
        Me.mnuScaleToGray.Checked = vbUnchecked
        Me.cmdEnhance.Caption = "Enhance ON"
   Else
        ' Toggle the menu item between checked and not checked
        Me.mnuScaleToGray.Checked = vbChecked
        Me.cmdEnhance.Caption = "Enhance OFF"
    End If

    'Execute the cmdScaleToGray Sub in the Active MDI Child form
    Me.ActiveForm.cmdScaleToGray
    
End Sub





Public Sub mnuWCascade_Click()
   ' Cascade child forms.
   Me.Arrange vbCascade
End Sub

Public Sub mnuWTileHorizontal_Click()
   ' Tile child forms (horizontal).
   Me.Arrange vbTileHorizontal
End Sub

Public Sub mnuWTileVertical_Click()
   ' Tile child forms (horizontal).
   Me.Arrange vbTileVertical
End Sub
Public Sub mnuWArrange_Click()
   ' Arrange all child form icons.
   Me.Arrange vbArrangeIcons
End Sub



Public Sub cmdGotoPage_Click()
   
    Dim txtHoldPageNumber As String
   
    On Error Resume Next
    
    txtHoldPageNumber = Me.ActiveForm.txtPageNumber
    
    'Check if Annotations were added and users wishes to Save them
    Me.ActiveForm.subAnnotationLayerSaveCheck
    
    Me.ActiveForm.SpicerView1.PageGotoDialog
   
    Me.ActiveForm.subSetCurrentPage
    
    '*** HAD TO TRICK THE LIST BOX TO CLEAR THE PREVIOUSLY SELECTED ITEM
    '*** commented the above code because  "lstPageList.Selected" triggers a "Click" event
    Me.ActiveForm.lstPageList.Selected(Me.ActiveForm.txtPageNumber - 1) = True
    Me.ActiveForm.lstPageList.Selected(txtHoldPageNumber - 1) = False
    
    
    result = MainMDIForm.funcZoomToSavedFactor

    If bolErrorOccured Then
            MsgBox "funcZoomToSavedFactor() ERROR:" & vbCrLf & " While attempting to set Saved ZOOM Factor." & vbCrLf & vbCrLf & result, vbCritical
    End If
      

End Sub

Public Sub cmdNextPage_Click()

    On Error Resume Next
    
    'Check if Annotations were added and users wishes to Save them
    Me.ActiveForm.subAnnotationLayerSaveCheck
    
'''    Me.ActiveForm.SpicerView1.GotoPageRelative 1
'''
'''    Me.ActiveForm.subSetCurrentPage
    
    '*** HAD TO TRICK THE LIST BOX TO CLEAR THE PREVIOUSLY SELECTED ITEM
    '*** commented the above code because  "lstPageList.Selected" triggers a "Click" event
    Me.ActiveForm.lstPageList.Selected(Me.ActiveForm.txtPageNumber - 1) = False
    Me.ActiveForm.lstPageList.Selected(Me.ActiveForm.txtPageNumber) = True
   
    strZoomResult = MainMDIForm.funcZoomToSavedFactor

    If bolErrorOccured Then
            MsgBox "funcZoomToSavedFactor() ERROR:" & vbCrLf & " While attempting to set Saved ZOOM Factor." & vbCrLf & vbCrLf & result, vbCritical
    End If
   
End Sub

Public Sub cmdPrevPage_Click()

    On Error Resume Next
    
    'Check if Annotations were added and users wishes to Save them
    Me.ActiveForm.subAnnotationLayerSaveCheck
    
'''    Me.ActiveForm.SpicerView1.GotoPageRelative -1
'''
'''    Me.ActiveForm.subSetCurrentPage
    
    '*** HAD TO TRICK THE LIST BOX TO CLEAR THE PREVIOUSLY SELECTED ITEM
    '*** commented the above code because  "lstPageList.Selected" triggers a "Click" event
    Me.ActiveForm.lstPageList.Selected(Me.ActiveForm.txtPageNumber - 1) = False
    Me.ActiveForm.lstPageList.Selected(Me.ActiveForm.txtPageNumber - 2) = True
    
    
    strZoomResult = MainMDIForm.funcZoomToSavedFactor

    If bolErrorOccured Then
            MsgBox "funcZoomToSavedFactor() ERROR:" & vbCrLf & " While attempting to set Saved ZOOM Factor." & vbCrLf & vbCrLf & result, vbCritical
    End If

End Sub

Public Sub cmdZoomIn_Click()

    On Error Resume Next
    
    'Execute the cmdZoomIn Sub in the Active MDI Child Form
    Me.ActiveForm.cmdZoomIn

End Sub

Public Sub cmdZoomOut_Click()

    On Error Resume Next
    
    'Execute the cmdZoomOut Sub in the Active MDI Child Form
    Me.ActiveForm.cmdZoomOut
    
End Sub



Function funcShowImage(strBatchDirectory As String, strBatchPageFileName As String, _
                        ByVal txtDocumentRECID As String, ByVal txtDetailRECID As String, _
                        ByVal txtCaption As String, ByVal txtTotalImages As Double, _
                        ByVal txtImageNumber As Double, ByVal txtPageRotation As Double, _
                        ByVal txtFTStatus As String, ByVal txtFTDirectory As String, ByVal txtFTFileName As String, _
                        intI101Module As Integer)
                
                
        '*** NOTE:  Make sure to Check if File Exists PRIOR to calling this function
        
        'On Error Resume Next
        
        On Error GoTo ERROR_HANDLER:
    
        Dim i As Integer
        
        '*** 2021-07-23 - Jacob - Added  and Moved Dim's to top of Func.
        Dim strLocalTempDir As String
        Dim txtLaunchFilePath As String
        Dim txtFullPathName As String
        Dim txtLaunchFileName As String
        Dim txtLaunchFileFullPath As String
        Dim sErrMessage As String
        
        Select Case intI101Module
        
            Case gI101ModuleRetrieve
            
                funcWriteToDebugLog Me.name, "**** funcShowImage | Case gI101ModuleRetrieve "

                'Allow loading MULTIPLE documents in the viewer
                
                i = UBound(arrDisplayedPagesRetrieve)
                funcWriteToDebugLog Me.name, "UBound(arrDisplayedPagesRetrieve) = " & UBound(arrDisplayedPagesRetrieve)
                
'                '*** If the Array is not empty Check if THIS DetailRECID item is already open
'                If UBound(arrDisplayedPagesRetrieve) > 0 Then
'                    For j = 0 To i
'                        If arrDisplayedPagesRetrieve(j) = txtDetailRECID Then
'                            'The item IS Already Loaded... Simply Set the Focus to it.
'                            gFormArrayRetrieve(j).SetFocus
'                            Exit Function
'                        End If
'                    Next
'                End If
                
                txtFullPathName = strBatchDirectory & "\" & strBatchPageFileName
                funcWriteToDebugLog Me.name, "txtFullPathName = " & txtFullPathName
                
                '2023-03-15 - Jacob - Added IF to correct Error with Uninitialized gFormArrayRetrieve
                'Create a NEW Child Form and increment the Array
                If IsArrayEmpty(gFormArrayRetrieve) Then
                        ReDim gFormArrayRetrieve(0)
                Else
                        ReDim Preserve gFormArrayRetrieve(LBound(gFormArrayRetrieve) To UBound(gFormArrayRetrieve) + 1)
                End If
                    
                Set frmViewForm = gFormArrayRetrieve(UBound(gFormArrayRetrieve))
                
                '*** 2022-07-28 - Jacob - Added txtFormArrayRetrieveIndex textbox to ChildForm to set which Form it is, for Unloading
                frmViewForm.txtFormArrayRetrieveIndex = UBound(gFormArrayRetrieve)
    
                 funcWriteToDebugLog Me.name, "UBound(gFormArrayRetrieve) = " & UBound(gFormArrayRetrieve)
   
                'Add THIS DetailRECID to the list of Displayed RECID's
                ReDim Preserve arrDisplayedPagesRetrieve(LBound(arrDisplayedPagesRetrieve) To UBound(arrDisplayedPagesRetrieve) + 1)
                arrDisplayedPagesRetrieve(UBound(arrDisplayedPagesRetrieve)) = txtDetailRECID
                
                funcWriteToDebugLog Me.name, "UBound(arrDisplayedPagesRetrieve) = " & UBound(arrDisplayedPagesRetrieve)
            
            Case gI101ModuleIndex
                
                
                funcWriteToDebugLog Me.name, "**** funcShowImage | Case gI101ModuleIndex"
                
                txtFullPathName = strBatchDirectory & "\" & strBatchPageFileName
                
                'Set up the Child Form - Set i=1 to Allow SINGLE PAGES ONLY
                i = 0
                ReDim Preserve gFormArrayIndex(i)
                Set frmViewForm = gFormArrayIndex(i)
                Set gFormArrayIndex(i) = frmViewForm
                
'                frmViewForm.Show
                
                DoEvents

                If UBound(gFormArrayIndex) > 0 Then
                    ' If there is an open document in the SpicerDoc1 Control
                    ' Close the document in the SpicerDoc1 control and
                    ' To check if the document has been changed, set CloseDocument to "True"
                    ' Allow loading only a SINGLE document
                    funcWriteToDebugLog Me.name, "**** funcShowImage | FormArrayIndex(" & i & ").SpicerDoc1.CloseDocument"

                    gFormArrayIndex(i).SpicerDoc1.CloseDocument False
                End If
            
                '*** 2023-01-03 - Jacob - Edited Comment to clarify how BatchPageRECID is passed.
                'Add the BatchPageRECID (sent in the txtDetailRECID field )  to the list of Displayed RECID's
                
                funcWriteToDebugLog Me.name, "**** funcShowImage | Add the BatchPageRECID to arrDisplayedPagesIndex() "
                ReDim Preserve arrDisplayedPagesIndex(0 To i)
                arrDisplayedPagesIndex(i) = txtDetailRECID
                
                
           Case Else
                funcWriteToDebugLog Me.name, "**** funcShowImage | Case Else "
                Exit Function
                
        End Select
        
        
        '*** Store the Array and DetailRECID to the New Viewform
        '    to allow removing from the arrDisplayedPages() Array when unloaded
        
        funcWriteToDebugLog Me.name, "**** funcShowImage | *** Set frmViewForm.Caption - THIS FORCES A Form.Load() event on the Child Form "
        
        '*** THIS FORCES A Form.Load() event on the Child Form
        

        If txtCaption = "" Then
            funcWriteToDebugLog Me.name, "**** funcShowImage | frmViewForm.Caption = " & txtFullPathName
            frmViewForm.Caption = txtFullPathName
        Else
            funcWriteToDebugLog Me.name, "**** funcShowImage | frmViewForm.Caption = " & txtCaption
            frmViewForm.Caption = txtCaption
        End If
        
        funcWriteToDebugLog Me.name, "**** funcShowImage | Store the Array and DetailRECID to the New Viewform"

        
        frmViewForm.txtModuleIndex = intI101Module
        frmViewForm.txtArrayIndex = i
        frmViewForm.txtDetailRECID = txtDetailRECID
        frmViewForm.txtDocumentRECID = txtDocumentRECID
        frmViewForm.txtImageCount = txtTotalImages
        frmViewForm.txtImageNumber = txtImageNumber
        frmViewForm.txtPageNumber = 1
        frmViewForm.txtPageRotation = txtPageRotation
        frmViewForm.txtPageFileName = txtFullPathName
        frmViewForm.txtBatchPageFileName = strBatchPageFileName
        frmViewForm.txtFileDirectory = strBatchDirectory
        
        frmViewForm.txtFTStatus = txtFTStatus
        

        
        'Jacob - 1/2/1014 - TEMPORARY BYPASS to Avoid ERROR:
        '                           "Microsoft Visual C++ Runtime Library - Buffer overrrun detected"
        
'        If UCase(Right(txtFullPathName, 4)) = "XLSX" Then
'            'Force error
'            Err.Number = 1
'            Err.Description = "ERROR OPENING FILE"
'        Else
            
 
                
                
        Set frmViewForm = Nothing
       

        funcShowImage = 0

        funcWriteToDebugLog Me.name, "**** funcShowImage | End Function | funcShowImage = " & funcShowImage

'2023-03-08 - Jacob - Added this Error Handler to try to catch crashes.
Exit Function

ERROR_HANDLER:

    'Set Error Flag to FORCE the Launch
    bolErrorOccured = True
    funcWriteToDebugLog Me.name, "**** funcShowImage |  Err #: " & Err.Number & " | " & Err.Description

        
End Function


Public Sub cmdImageRotateLeft_Click()
    
    On Error Resume Next
    
    'ActiveForm is the "Active" instance of the ChildForm1
    Me.ActiveForm.cmdImageRotateLeft
    
End Sub

Public Sub cmdImageRotateRight_Click()

    On Error Resume Next
    
    'ActiveForm is the "Active" instance of the ChildForm1
    Me.ActiveForm.cmdImageRotateRight
    
End Sub

Public Sub subShowActiveLayer()

    Dim UserTools As IUserTools
    
    ' Set the object variable for the IUserTools interface to the Markup Control object
    Set UserTools = frmViewForm.SpicerMarkup1.object
    
    ' Open the Change Active Layer dialog box
    UserTools.ActiveLayerDialog
    
    ' De-initialize the object variable
    Set UserTools = Nothing
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

End Sub

Public Function funcZoomToSavedFactor() As String

    Dim ScaleScrollRotation As IScaleScrollRotation
    Dim dScale As Double
    Dim X As Long
    Dim Y As Long
    Dim ZoomLevel As Integer
    
    ' Set the object variable for the IScaleScrollRotation interface to the View Control object
    Set ScaleScrollRotation = Me.ActiveForm.SpicerView1.object
    ' Find the scale factor and x,y coordinates for the active page
''''    ScaleScrollRotation.GetZoomFactor 0, dScale, x, y

    On Error Resume Next
    dScale = VBGetPrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomFactorScale", RegFileName)
    X = VBGetPrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomFactorX", RegFileName)
    Y = VBGetPrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomFactorY", RegFileName)
    ZoomLevel = VBGetPrivateProfileString(RegAppname, "frmViewForm.SpicerView1.ZoomLevel", RegFileName)
    
    '2021-07-14 - Jacob - Added this Error Handler to prevent crashing due to bad PDF documents... and initiallize bolErrorOccured
    On Error GoTo ERROR_HANDLER
    bolErrorOccured = False
    
    ' ***  ZOOM FACTORS are NOT BEING SET -- SEEMS TO BE "CONSTANTS" ???
'    frmViewForm.SpicerView1.ZoomFactorX(0) = X
'    frmViewForm.SpicerView1.ZoomFactorY(0) = Y
'    frmViewForm.SpicerView1.ZoomFactorScale(0) = dScale
   
'''   Dim intResolutionX As Integer
'''   Dim intResolutionY As Integer
'''
'''   ResolutionX = frmViewForm.SpicerDoc1.GetResolutionX(0)
'''   ResolutionY = frmViewForm.SpicerDoc1.GetResolutionY(0)
   
    Select Case ZoomLevel
    
        Case IN_ZOOM_IN  '(1)
            'magnify
            '*Ignore
            
        Case IN_ZOOM_OUT  '(2)
            'reduce
            '*Ignore
            
        Case IN_ZOOM_SCALETOFIT  '(3)
            'fit to page
            ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_SCALETOFIT
            
        Case IN_ZOOM_HORIZFIT  '(4)
            'fit horizontally in window
            ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_HORIZFIT
            
        Case IN_ZOOM_VERTFIT  '(5)
            'fit vertically in window
            ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_HORIZFIT
 
        Case IN_ZOOM_ACTUALSIZE  '(6)
            'actual size
            ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_ACTUALSIZE
 
        Case IN_ZOOM_SCALEFACTOR  '(7)
            'specific scale factor
            ScaleScrollRotation.SetZoomFactor 0, IN_ZOOM_CUSTOM_CENTER, dScale, X, Y
            ' Automatically sets: ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_SCALEFACTOR
            
        Case IN_ZOOM_1TO1  '(8)
            '1-to-1 scale factor
            ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_1TO1
 
        Case Else
            '*Ignore
   End Select

    
    ' De-initialize the IMiscellaneous object variable
    Set ScaleScrollRotation = Nothing
    funcZoomToSavedFactor = ""
    
'2021-07-14 - Jacob - Added this Error Handler to prevent crashing due to bad PDF documents.
Exit Function

ERROR_HANDLER:

    ' De-initialize the IMiscellaneous object variable
    Set ScaleScrollRotation = Nothing

    'Set Error Flag to FORCE the Launch
    bolErrorOccured = True
    funcZoomToSavedFactor = "funcZoomToSavedFactor() | Err #: " & Err.Number & " | " & Err.Description

End Function

Public Sub subRemoveDisplayedPageFromArray(intI101Module As Integer, txtArrayIndex As String)

    '*** REMOVE the Array Element for an Unloaded Document
    Dim i As Integer
'    For j = 0 To UBound(arrDisplayedPages, intI101Module): funcWriteToDebugLog  Me.Name, j & " " & intI101Module & " " & arrDisplayedPages(j): Next
        
        Select Case intI101Module
        
            Case gI101ModuleRetrieve
            
                For i = 1 To UBound(arrDisplayedPagesRetrieve)
                    If i >= txtArrayIndex And i < UBound(arrDisplayedPagesRetrieve) Then
                        'Replace the value of the current array item with the value of the next one
                        arrDisplayedPagesRetrieve(i) = arrDisplayedPagesRetrieve(i + 1)
                    End If
                Next
                
                ReDim Preserve arrDisplayedPagesRetrieve(UBound(arrDisplayedPagesRetrieve) - 1)
                funcWriteToDebugLog Me.name, "UBound(arrDisplayedPagesRetrieve) = " & UBound(arrDisplayedPagesRetrieve)
            
                If UBound(arrDisplayedPagesRetrieve) < 1 Then
                    cmdAnnotate.Enabled = False
                    cmdEdit.Enabled = False
                    cmdEnhance.Enabled = False
                    cmdThumbNails.Enabled = False
                    cmdZoomFit.Enabled = False
                    cmdZoomIn.Enabled = False
                    cmdZoomOut.Enabled = False
                    cmdSaveZoom.Enabled = False
                    cmdImageRotateLeft.Enabled = False
                    cmdImageRotateRight.Enabled = False
                    cmdPrint.Enabled = False
                    cmdLaunch.Enabled = False
                    cmdSendTo.Enabled = False
                    cmdPrevImage.Enabled = False
                    cmdNextImage.Enabled = False
                    cmdPrevPage.Enabled = False
                    cmdNextPage.Enabled = False
                    cmdGotoImage.Enabled = False
                    cmdGotoPage.Enabled = False
                    
                End If
                
            Case gI101ModuleIndex
                ReDim Preserve arrDisplayedPagesIndex(UBound(arrDisplayedPagesIndex) - 1)

            Case Else
             
        End Select
        
             If UBound(arrDisplayedPagesRetrieve) < 1 And UBound(arrDisplayedPagesIndex) < 1 Then
                If funcIsFormLoaded2("frmAnnotate") Then
                    Unload frmAnnotate
                    Set frmAnnotate = Nothing
                End If
            End If
             


'    For j = 0 To UBound(gI101ModuleArrayUbound, arrDisplayedPages): funcWriteToDebugLog  Me.Name, j & " " & arrDisplayedPages(j): Next


End Sub

Public Sub subRemoveChildFormFromArray(intI101Module As Integer, txtFormArrayIndex As String)


    On Error Resume Next
    
    '*** REMOVE the Array Element for an Unloaded Document
    Dim intFormIndex As Integer
    Dim intWorkIndex As Integer
    Dim intIndexToRemove As Integer
    
    intIndexToRemove = CInt(txtFormArrayIndex)
    
    Dim gFormArrayRetrieveWork() As New ChildForm1
    

        
        
        Select Case intI101Module
        
            Case gI101ModuleRetrieve
            
                For j = 0 To UBound(gFormArrayRetrieve)
                        Debug.Print j & " " & intI101Module & " " & gFormArrayRetrieve(j).Caption
                Next
                
                If UBound(gFormArrayRetrieve) <= 0 Then
                        Erase gFormArrayRetrieve
                Else
                        'Reorganize array indexes
                        For intFormIndex = LBound(gFormArrayRetrieve) To UBound(gFormArrayRetrieve)
                            
                            If intFormIndex > intIndexToRemove Then
                                Set gFormArrayRetrieve(intFormIndex - 1) = gFormArrayRetrieve(intFormIndex)
                                gFormArrayRetrieve(intFormIndex - 1).txtFormArrayRetrieveIndex = intFormIndex - 1
                            End If
                            
'                            If intFormIndex <> intIndexToRemove Then
'                                'Replace the value of the current array item with the value of the next one
'                                ReDim Preserve gFormArrayRetrieveWork(LBound(gFormArrayRetrieveWork) To intWorkIndex)
'                                Set frmViewForm = gFormArrayRetrieve(intFormIndex)
'                                Set gFormArrayRetrieveWork(intWorkIndex) = frmViewForm
'                                intWorkIndex = intWorkIndex + 1
'                            End If
                            
                        Next
                        
                
                        'Re-dimension the Form Array to one less
                        'ReDim Preserve gFormArrayRetrieve(UBound(gFormArrayRetrieve) - 1)
                        ReDim Preserve gFormArrayRetrieve(LBound(gFormArrayRetrieve) To UBound(gFormArrayRetrieve) - 1)
        
                        funcWriteToDebugLog Me.name, "UBound(gFormArrayRetrieve) = " & UBound(gFormArrayRetrieve)
                        
                End If
            

                
            Case gI101ModuleIndex
            
                        'For BATCHES, gFormArrayIndex can ONLY be one (1)
                        Debug.Print 1 & " " & intI101Module & " " & gFormArrayIndex(0).Caption

                        'ReDim Preserve gFormArrayIndex(UBound(gFormArrayIndex) - 1)
                        Erase gFormArrayIndex

            Case Else
             
        End Select
        
        
        If IsArrayEmpty(gFormArrayRetrieve) Then
                    cmdAnnotate.Enabled = False
                    cmdEdit.Enabled = False
                    cmdEnhance.Enabled = False
                    cmdThumbNails.Enabled = False
                    cmdZoomFit.Enabled = False
                    cmdZoomIn.Enabled = False
                    cmdZoomOut.Enabled = False
                    cmdSaveZoom.Enabled = False
                    cmdImageRotateLeft.Enabled = False
                    cmdImageRotateRight.Enabled = False
                    cmdPrint.Enabled = False
                    cmdLaunch.Enabled = False
                    cmdSendTo.Enabled = False
                    cmdPrevImage.Enabled = False
                    cmdNextImage.Enabled = False
                    cmdPrevPage.Enabled = False
                    cmdNextPage.Enabled = False
                    cmdGotoImage.Enabled = False
                    cmdGotoPage.Enabled = False
        End If
    
    '***2022-04-10 - Jacob - Commented out the following if
'             If UBound(arrDisplayedPagesRetrieve) < 1 And UBound(arrDisplayedPagesIndex) < 1 Then
'                If funcIsFormLoaded2("frmAnnotate") Then
'                    Unload frmAnnotate
'                    Set frmAnnotate = Nothing
'                End If
'            End If
             



End Sub



Private Sub PixEzImage1_AfterScanSave()

    Dim txtFullPathName As String
    
    
    
    
    Select Case strScanMode
    
        Case "Replace Page"
            'Logic handled in the mnuBatchReplaceCurrentPage_Click() sub
            '  because AfterScanSave is NOT fired on single page scan.
            
        Case "Insert Pages BEFORE"
            txtFullPathName = frmIndex.txtBatchDirectory & "\" & frmIndex.txtBatchPageFileName
            dblHoldBatchPageOrder = frmIndex.ListView1.ListItems.item(frmIndex.ListView1.SelectedItem.Index).Text
            subSavePage strScanMode, frmIndex.txtApplicationName, frmIndex.txtBatchDirectory, frmIndex.txtBatchRECID, dblHoldBatchPageOrder
            
            'Re-populate the List of Pages
            Call frmIndex.subLoadPagesIntoListView
            'Set the list index to the same image we were on which should be the first one inserted
            frmIndex.ListView1.ListItems.item(dblHoldBatchPageOrder + 1).Selected = True
        
            
        Case "Insert Pages AFTER"
            txtFullPathName = frmIndex.txtBatchDirectory & "\" & frmIndex.txtBatchPageFileName
            dblHoldBatchPageOrder = frmIndex.ListView1.ListItems.item(frmIndex.ListView1.SelectedItem.Index).Text
            subSavePage strScanMode, frmIndex.txtApplicationName, frmIndex.txtBatchDirectory, frmIndex.txtBatchRECID, dblHoldBatchPageOrder
    
            'Re-populate the List of Pages
            Call frmIndex.subLoadPagesIntoListView
            'Set the list index to the same image we were on which should be the first one inserted
            frmIndex.ListView1.ListItems.item(dblHoldBatchPageOrder + 1).Selected = True
        
        Case "Append Pages"
            txtFullPathName = frmIndex.txtBatchDirectory & "\" & frmIndex.txtBatchPageFileName
            dblHoldBatchPageOrder = frmIndex.ListView1.ListItems.Count
            subSavePage strScanMode, frmIndex.txtApplicationName, frmIndex.txtBatchDirectory, frmIndex.txtBatchRECID, dblHoldBatchPageOrder
            
            'Re-populate the List of Pages
            Call frmIndex.subLoadPagesIntoListView
            'Set the list index to the same image we were on which should be the first one inserted
            frmIndex.ListView1.ListItems.item(dblHoldBatchPageOrder + 1).Selected = True
        
            
    End Select
    
    'Re-display the image
    Call frmIndex.ListView1_Click
            


End Sub






Private Sub subSavePage(ScanMode As String, strApplicationName As String, strBatchDirectory As String, dblBatchRECID As Double, dblBatchPageOrder As Long)

    On Error GoTo ERROR_HANDLER
    
    funcWriteToDebugLog Me.name, "***** ENTERING subSavePage"
    funcWriteToDebugLog Me.name, "  ScanMode = " & ScanMode
    funcWriteToDebugLog Me.name, "  strApplicationName = " & strApplicationName
    funcWriteToDebugLog Me.name, "  strBatchDirectory = " & strBatchDirectory
    funcWriteToDebugLog Me.name, "  dblBatchRECID " & dblBatchRECID
    funcWriteToDebugLog Me.name, "  dblBatchPageOrder = " & dblBatchPageOrder
    
    ' Send the image to the Viewer form and save to file if requested

    Dim strFileName As String
    Dim Temp As Long
    
    m_ImageCount = m_ImageCount + 1
    m_PageCount = m_PageCount + 1
    m_ImageSkipCount = m_ImageSkipCount + 1
    
    
    ' Save Image if requested
''''    strFileName = Format(dblBatchRECID, "000000000") & "-" & _
''''                    Format(Now(), "yymmddhhmmss") & "-" & _
''''                    Format(m_ImageCount, "0000") & ".TIF"
                    
    strFileName = Right(PixEzImage1.PageFileName, Len(PixEzImage1.PageFileName) - InStrRev(PixEzImage1.PageFileName, "\"))

    funcWriteToDebugLog Me.name, "strFilename = " & strFileName
        
    DoEvents

    'Send Image to the Viewer
    If chkScanDisplayImages = vbChecked Then
        If m_PageCount = 1 Or m_ImageSkipCount = CDbl(txtScanImageSkipCount) Then
            funcWriteToDebugLog Me.name, "Imaging101ScanViewer.AddImage " & strFullBatchDirectory & "\" & strFileName
            Imaging101ScanViewer.AddImage strFullBatchDirectory & "\" & strFileName
            m_ImageSkipCount = 0
        End If
    End If
    
    DoEvents
    
    'Create the Batch Page Record - Pass the filename as a parameter
    subCreateBatchPageRecord strScanMode, strApplicationName, strFileName, dblBatchRECID, dblBatchPageOrder

Exit Sub

ERROR_HANDLER:

'    MsgBox "chkBatchAutoName ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
''''''''''    GetError
    Resume Next
    
End Sub


Private Sub subCreateBatchPageRecord(strScanMode As String, strApplicationName As String, strFileName As String, dblBatchRECID As Double, dblBatchPageOrder As Long)
    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
        On Error GoTo CREATE_BATCH_PAGE_RECORD_ERROR
        
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
                    " WHERE BatchRECID = " & dblBatchRECID
        
        connImaging101Batch.Errors.Clear
        rsImaging101Batch.Open
        rsImaging101Batch.MoveFirst
        
        txtActionBeforeError = "Open Batch Page Table"
        Set rsImaging101BatchPage = New ADODB.Recordset
        rsImaging101BatchPage.Open strApplicationName & "_BatchPage", connImaging101Batch, adOpenDynamic, adLockOptimistic
        
        'User Transaction Tracking to make sure the Batch and BatchPage tables are updated together!
        connImaging101Batch.BeginTrans
        
        '*** Adjust the QC values
        Select Case strScanMode
        
            Case "Insert Pages BEFORE"
            
                Set cmd = New ADODB.Command
                Set cmd.ActiveConnection = connImaging101Batch
                'Increment the page numbers of all Pages after the newly scanned one
                cmd.CommandText = "Update " & strApplicationName & "_BatchPage " & _
                                    " SET BatchPageOrder = BatchPageOrder + 1 " & _
                                    " WHERE BatchPageOrder >= " & dblBatchPageOrder & _
                                    "   AND BatchRECID = " & dblBatchRECID

                cmd.Execute , , adCmdText

                
            Case "Insert Pages AFTER"
            
                Set cmd = New ADODB.Command
                Set cmd.ActiveConnection = connImaging101Batch
                cmd.CommandText = "Update " & strApplicationName & "_BatchPage " & _
                                    " SET BatchPageOrder = BatchPageOrder + 1 " & _
                                    " WHERE BatchPageOrder >= " & dblBatchPageOrder + 1 & _
                                    "   AND BatchRECID = " & dblBatchRECID
                cmd.Execute , , adCmdText

            
            Case "Append Pages"
                'No Page REnumbering required
                
        End Select
        
        
        txtActionBeforeError = "Add New Record"
        rsImaging101BatchPage.AddNew
        
        txtActionBeforeError = "Assign Variables to Batch Page Fields"
        rsImaging101BatchPage("BatchPageRECID") = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchPageRECID")
        rsImaging101BatchPage("BatchRECID") = dblBatchRECID
        rsImaging101BatchPage("BatchPageFileName") = strFileName
        
        rsImaging101BatchPage("BatchPageIndexed") = ""
        rsImaging101BatchPage("BatchPageIsSeparator") = ""
        rsImaging101BatchPage("BatchPageNote") = ""
        rsImaging101BatchPage("BatchDocDesc") = ""
        rsImaging101BatchPage("BatchPageStatus") = ""
'        rsImaging101BatchPage("BatchPageCommitDate") = ""
        rsImaging101BatchPage("BatchPageCommitUser") = ""
        
        ' Set BATCHES field values
        txtActionBeforeError = "Assign Variables to Batch Fields"
        
        Dim intBatchPagesTotal As Integer
        intBatchPagesTotal = rsImaging101Batch("BatchPagesTotal")
        intBatchPagesTotal = intBatchPagesTotal + 1
        
        rsImaging101Batch("BatchPagesTotal") = intBatchPagesTotal
        
        '*** Adjust the QC values
        Select Case strScanMode
        
            Case "Insert Pages BEFORE"
                rsImaging101BatchPage("BatchPageOrder") = dblBatchPageOrder
                rsImaging101Batch("BatchPagesQCInserted") = rsImaging101Batch("BatchPagesQCInserted") + 1
                
            Case "Insert Pages AFTER"
                rsImaging101BatchPage("BatchPageOrder") = dblBatchPageOrder + 1
                rsImaging101Batch("BatchPagesQCInserted") = rsImaging101Batch("BatchPagesQCInserted") + 1
            
            Case "Append Pages"
                rsImaging101BatchPage("BatchPageOrder") = intBatchPagesTotal
                rsImaging101Batch("BatchPagesQCAppended") = rsImaging101Batch("BatchPagesQCAppended") + 1
                
        End Select
        
        
        rsImaging101Batch("BatchPagesNotCommitted") = rsImaging101Batch("BatchPagesNotCommitted") + 1
        
        '*** UPDATE TABLES
        txtActionBeforeError = "Update Batch Page Values"
        rsImaging101BatchPage.Update
        
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
    lngFileSize = FileLen(FullFilePath)
    
    funcKillFileIfSmallerThan = False
    
    If lngFileSize < FileMinimumSize Then
'        MsgBox FullFilePath & " Size: " & lngFileSize & " Min: " & FileMinimumSize
        Kill FullFilePath
    
        'Set to true to notify the calling routing that the image WAS deleted
        funcKillFileIfSmallerThan = True

    End If
    
End Function


Private Sub subDeleteBatchPageRecord()
    
        
'''        Screen.MousePointer = vbHourglass
'''        txtPagesImported = 0
                
        On Error GoTo DELETE_BATCH_PAGE_RECORD_ERROR
        
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
        
        txtActionBeforeError = "Open Batch Page Table"
        Set rsImaging101BatchPage = New ADODB.Recordset
        rsImaging101BatchPage.Open frmIndex.txtApplicationName & "_BatchPage", connImaging101Batch, adOpenDynamic, adLockOptimistic
        
        'User Transaction Tracking to make sure the Batch and BatchPage tables are updated together!
        connImaging101Batch.BeginTrans
        
        '*** Adjust the QC values
                Set cmd = New ADODB.Command
                Set cmd.ActiveConnection = connImaging101Batch
                
                
                'DELETE THE SELECTED PAGE
                cmd.CommandText = "DELETE FROM " & frmIndex.txtApplicationName & "_BatchPage " & _
                                    " WHERE BatchPageRECID = " & frmIndex.txtBatchPageRECID
                cmd.Execute , , adCmdText

                
                'DECREMENT the page numbers of all Pages after the one to be deleted
                cmd.CommandText = "Update " & frmIndex.txtApplicationName & "_BatchPage " & _
                                    " SET BatchPageOrder = BatchPageOrder - 1 " & _
                                    " WHERE BatchPageOrder >= " & dblHoldBatchPageOrder & _
                                    "   AND BatchRECID = " & frmIndex.txtBatchRECID
                                    
                cmd.Execute , , adCmdText

                
        '*** 2023-02-23 - Jacob - Corrected Page Counts
        txtActionBeforeError = "Assign Variables to Batch Fields"
        
        Dim intBatchPagesTotal As Integer
        intBatchPagesTotal = rsImaging101Batch("BatchPagesTotal")
        intBatchPagesTotal = intBatchPagesTotal - 1
        rsImaging101Batch("BatchPagesTotal") = intBatchPagesTotal
        
        frmIndex.txtBatchPagesTotal = intBatchPagesTotal
        
        rsImaging101Batch("BatchPagesQCDeleted") = rsImaging101Batch("BatchPagesQCDeleted") + 1
        
        rsImaging101Batch("BatchPagesNotCommitted") = rsImaging101Batch("BatchPagesNotCommitted") - 1
        
        txtActionBeforeError = "Update Batch Values"
        rsImaging101Batch.Update
        
    
        'Close the displayed document to allow Deleting it
        Me.ActiveForm.SpicerDoc1.CloseDocument False
            
        '*** 2023-02-22 - Jacob - Modified to PREVENT Deleting the File if this page was a Copy of another
        rsImaging101BatchPage.Close
            
        rsImaging101BatchPage.Source = "SELECT count(*) " & _
                    " FROM " & frmIndex.txtApplicationName & "_BatchPage" & _
                    " WHERE BatchRECID = " & frmIndex.txtBatchRECID & _
                    "     AND  BatchPageFileName = '" & frmIndex.txtBatchPageFileName & "'"
        
        rsImaging101BatchPage.Open
        
        Dim intRecordCount As Integer
        intRecordCount = rsImaging101BatchPage.GetString

        ' Do NOT delete the File if there are still any records pointing to the same FileName.
        ' For example, if this was a "Copy" of another record.
        If intRecordCount = 0 Then
                '*** KILL the IMAGE FILE
                Kill frmIndex.txtBatchDirectory & "\" & frmIndex.txtBatchPageFileName & "'"
        End If


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
    
DELETE_BATCH_PAGE_RECORD_ERROR:
        
'        & ") - [Transaction Rolled Back - Batch Page NOT Created]"
        
        gYesNo = vbNo
        frmYesNo.lblYesNoMessage = "subDeleteBatchPageRecord ERROR: " & Err.Number & " - " & Err.Description & _
        vbCrLf & "DURING ACTION: (" & txtActionBeforeError & ")" & _
        vbCrLf & "Can NOT Delete File: " & frmIndex.txtBatchDirectory & "\" & frmIndex.txtBatchPageFileName & _
        vbCrLf & vbCrLf & "Do you want to continue to Delete this page?"
        
        frmYesNo.Top = Me.Top
        frmYesNo.Left = Me.Left
        frmYesNo.Show vbModal, Me
        
        On Error Resume Next
        
        If gYesNo = vbYes Then
            ' Commit the successful transactions to DELETE the bad file
            txtActionBeforeError = "COMMIT BATCH PAGE TRANSACTION"
            connImaging101Batch.CommitTrans
        Else
            txtActionBeforeError = "ROLL-BACK BATCH PAGE TRANSACTION"
            connImaging101Batch.RollbackTrans
        End If
        
        
        
        rsImaging101Batch.Close
        Set rsImaging101Batch = Nothing
        rsImaging101BatchPage.Close
        Set rsImaging101BatchPage = Nothing
'        Set connImaging101Batch = Nothing
        
        Screen.MousePointer = vbDefault

        bolCancelPendingXfers = True

    
End Sub

Private Sub txtFindText_GotFocus()

    cmdFind.Default = True

End Sub
