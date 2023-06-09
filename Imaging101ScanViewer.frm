VERSION 5.00
Object = "{71C182E1-878D-11D1-8108-020701190C00}#8.0#0"; "view.ocx"
Object = "{22B7B2BB-4EFA-11D2-81FC-0000D1108734}#8.0#0"; "Edit.ocx"
Object = "{895CDC7A-8837-11D1-8109-020701190C00}#8.0#0"; "docctrl.ocx"
Object = "{C8B15BE2-E8D8-11D1-818A-0000D1108734}#8.0#0"; "SpiConfg.ocx"
Object = "{CA5948F6-E0F8-11D1-9A59-0000929B58F0}#8.0#0"; "mark.ocx"
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.6#0"; "PIXEZOCX.OCX"
Begin VB.Form Imaging101ScanViewer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imaging101 Scan Viewer"
   ClientHeight    =   7170
   ClientLeft      =   6540
   ClientTop       =   315
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   6615
   Visible         =   0   'False
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "Imaging101ScanViewer.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   11
      Top             =   0
      Width           =   1572
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
      Height          =   250
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   855
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
      Height          =   250
      Left            =   855
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   250
      Left            =   2064
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin SPICERMARKUPLib.SpicerMarkup SpicerMarkup1 
      Left            =   2880
      Top             =   4440
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERCONFIGURATIONLib.SpicerConfiguration SpicerConfiguration1 
      Left            =   2160
      Top             =   4440
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERDOCUMENTLib.SpicerDoc SpicerDoc1 
      Left            =   1440
      Top             =   4440
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICEREDITLib.SpicerEdit SpicerEdit1 
      Left            =   720
      Top             =   4440
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERVIEWLib.SpicerView SpicerView1 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
      _Version        =   524288
      _ExtentX        =   661
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.TextBox txtFileName 
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
      Left            =   3000
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox EdtTotal 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   780
      Picture         =   "Imaging101ScanViewer.frx":0693
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox EdtImag 
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
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton CmdPrev 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   0
      Picture         =   "Imaging101ScanViewer.frx":0A1D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin PixezocxLib.PixEzImage ezPageView 
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   3135
      _Version        =   65542
      _ExtentX        =   5530
      _ExtentY        =   5106
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
      TAG_BORDER_COLOR_BG=   16777215
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
      TAG_TREE_COLOR_BG=   16777215
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
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   4920
      TabIndex        =   12
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total:"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "Imaging101ScanViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Poor man's viewer :)
'
' This form just uses an array of Picture controls to store
' all the images returned by TwainPRO
Option Explicit
Dim bolDiscardPage As Boolean


Public Sub AddImage(fn As String)
    Dim Index As Long
    
    ' If Preview Only is checked ON keep only ONE image at a time
'    If Imaging101ScanMain.chkScanPreviewOnly = vbChecked Then
'        Index = CDbl(EdtTotal.Text) + 1
'        EdtTotal.Text = Index
'    Else
        Index = 1
        EdtTotal.Text = Index
'    End If
    
    ' Using simple VB Image Control
'    On Error Resume Next
'    If Image1.Count > 1 Then
'        Unload Image1(Index)
'    End If
'    On Error GoTo 0
    
'''    'Resize
'''    If Me.Visible = False Then
'''       Me.Show
'''    End If
    
'    Load Image1(Index)
'    Image1(Index).Picture = Imaging101ScanMain.TwainPRO.Picture
'    Image1(Index).Visible = True
'    Image1(Index).ZOrder

    '*** Using Spicer Image aX
'''    SpicerDoc1.CloseDocument False
'''    SpicerDoc1.OpenFile (fn)
'''    SpicerView1.BindToDocumentControl SpicerDoc1.object
'''    SpicerEdit1.BindToViewControl SpicerView1.object
'''    SpicerEdit1.ReplaceCurrentDocWhenRasterizing = 1
'''    mnuCropAuto_Click

'    SpicerEdit1.Save Me.SpicerEdit1.object, False, 0, fn, ""
    
    '*** Using Pegasus ImageXpress
'    ImagXpress1.hDib = Imaging101ScanMain.TwainPRO.hDib
'    ImagXpress1.ZoomToFit ZOOMFIT_BEST

''    Load Pic(Index)
''    Pic(Index).Picture = Stage.TwainPRO.Picture
''    Pic(Index).Visible = True
''    Pic(Index).ZOrder
    
    DoEvents
    
    If Index > 1 Then
       CmdPrev.enabled = True
    End If
    
    cmdNext.enabled = False
    EdtImag.Text = Index
    
End Sub

Private Sub CmdNext_Click()
'    Dim Index As Long
'
'    Index = CDbl(EdtImag.Text) + 1
'    EdtImag.Text = Index
'    Image1(Index).ZOrder
'    'Pic(Index).ZOrder
'    If Index >= CDbl(EdtTotal.Text) Then
'       CmdNext.Enabled = False
'    End If
'    If CmdPrev.Enabled = False Then
'       CmdPrev.Enabled = True
'    End If
'    Resize
End Sub

Private Sub CmdPrev_Click()
'    Dim Index As Long
'
'    Index = CDbl(EdtImag.Text) - 1
'    EdtImag.Text = Index
'    Image1(Index).ZOrder
'    'Pic(Index).ZOrder
'    If Index <= 1 Then
'       CmdPrev.Enabled = False
'    End If
'    If CmdNext.Enabled = False Then
'       CmdNext.Enabled = True
'    End If
'    Resize
End Sub

Private Sub Resize()

End Sub




Private Sub ezPageView_AfterScanSave()
'Occurs after each page in a batch scanning operation has been saved.
'Use the AfterScanSave event, for example, to display a progress monitor based on
'  the number pages scanned and saved.

        '*************************************************************************
        '*** 8/10/2009 - Jacob - Commented the FlushPages because it was causing
        '***                     slowdowns on Networks, where the Scanner buffer
        '***                     would overflow... and slow down scans by half or more.

        '''   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_AfterScanSave (before FlushPages)- " & ezPageView.PageFileName
        '''   DoEvents
        '''
        '''   ezPageView.FlushPages
        '''   DoEvents
        '''
        '''   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_AfterScanSave (after FlushPages) - " & ezPageView.PageFileName
        '''   DoEvents
   

        '*** If in DUPLEX Mode, Check if this is the REAR Image'
        '***  to see if it is Blank
        'Does the scanner support Duplex?
       
        Imaging101ScanMainPix.lblStatus = "Saving Page #: " & Abs(ezPageView.ScanCurrentPage) & "    Image #: " & ezPageView.ScanAbsolutePage
        DoEvents
        
        If Imaging101ScanViewer.ezPageView.ScanDuplex = 1 Then
                'Is Duplex enabled?
                    'Check for a REMAINDER using the MOD math operator
                    'If there is NO remainder then the page is EVEN,
                    'meaning that it is the REAR image
                    If (ezPageView.ScanAbsolutePage Mod 2) = 0 Then
                        Dim lngFileSize As Long
                        
                        'FileLen function returns the file size in BYTES.
                        lngFileSize = FileLen(ezPageView.PageFileName)
                        
                        If lngFileSize >= Imaging101ScanMainPix.txtMinimumImageSize Then
                                'Only save rear page if larger than Minimum Size requested
                                Imaging101ScanMainPix.subPostScan ezPageView.PageFileName
                                 funcWriteToDebugLog Me.name, ezPageView.ImageFileExt
                                 funcWriteToDebugLog Me.name, ezPageView.ImageFileRoot
                                 funcWriteToDebugLog Me.name, ezPageView.ImageFileDir
                                 funcWriteToDebugLog Me.name, ezPageView.ImageFileSchema
                       Else
'                                '*** THIS MAY NOT WORK IF FILE IS STILL OPEN AT THIS POINT! ***
'                                'Attempt to Close the current document/file
'                                Dim strFileNameToKill As String
'                                strFileNameToKill = ezPageView.PageFileName
'                                ezPageView.Close
'                                funcKillFileIfSmallerThan strFileNameToKill, Imaging101ScanMainPix.txtMinimumImageSize
                        End If
                    Else
                        'Duplex - Always save FRONT of Page
                        Imaging101ScanMainPix.subPostScan ezPageView.PageFileName

                    End If
        Else
            'Simplex - Always save FRONT of Page
            Imaging101ScanMainPix.subPostScan ezPageView.PageFileName

        End If
        
End Sub

Private Sub ezPageView_BeforeScan()
'Occurs just BEFORE each page is scanned,
'  AFTER checking whether or not there are more pages to scan.

   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_BeforeScan - " & ezPageView.PageFileName
   
   DoEvents
    ' Stop Scanning If user clicked the STOP SCANNING button
    '  or if the bolCancelPendingXfers boolean variable was set to TRUE
    If bolCancelPendingXfers Then
''''        'Save the Advanced Capabilities and set Cancel=True to prevent scanning a page.
''''        cmdScannerAdvancedCapabilitiesSave_Click
        ezPageView.ScanCancel = 1
        'Reset the CancelPendingXfers flag
        bolCancelPendingXfers = False
    End If

End Sub

Private Sub ezPageView_DoneScanning()

'The DoneScanning event occurs only after exiting from the scanning loop; that is,
'  if using the Prepare Scanner dialog (the frmContinueScanBatch form) or equivalent,
'  it occurs only after the user clicks Stop Scanning.
'
'Use the DoneScanning event, for example, to control the enabled/disabled stated
'  of various menu commands that you may want to dim while scanning.

   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_DoneScanning - " & " - " & ezPageView.PageFileName
   
   bolCancelPendingXfers = True
   
   DoEvents
   


End Sub



Private Sub ezPageView_ErrorFilter(ErrorNumber As Long, ErrorDescription As String)

'   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_ErrorFilter " & " - ERROR: " & ErrorNumber & " - " & ErrorDescription
'   frmMessageForm.Show
'   frmMessageForm.txtMessage = frmMessageForm.txtMessage & "I101ScanViewer - ezPageView_ErrorFilter " & " - ERROR: " & ErrorNumber & " - " & ErrorDescription & vbCrLf
'   frmMessageForm.OKButton.enabled = True
'   frmMessageForm.cmdCopyToClipboard.enabled = True

End Sub

Private Sub ezPageView_FeederEmpty(NumPages As Long, StopScan As Long, UseFeeder As Long, ScanBacks As Long)
'Occurs during a batch scanning operation when the scanner’s
'  document feeder becomes empty.

'When the FeederEmpty event occurs, the feeder is empty and scanning is suspended.
'  The toolkit is still in its scanning loop. You then can call the built-in
'  Prepare Scanner dialog (frmContinueScanBatch) or its equivalent,
'  and then fill in the values of the FeederEmpty event’s parameters with
'  information gathered from this dialog.

'Note that if you use the included Prepare Scanner dialog (frmContinueBatchScan)
'  it has a ShowPrepareScannerDialog function with all four parameters.
'  Your application simply needs to pass these values back to the FeederEmpty event.

'To continue scanning, do not call ScanBatch again.
'  Instead, set the StopScan parameter to zero and return from FeederEmpty.
'  Scanning will then resume with either the next stack or the backs of the
'  previous stack, depending on the value of the ScanBacks parameter.

   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_FeederEmpty - " & ezPageView.PageFileName
   DoEvents
    
    
    ezPageView.Refresh
    
    'PixSLForm.ShowPixSLDialog ezPageView, NumPages, StopScan, UseFeeder, ScanBacks
    frmContinueBatchScan.ShowPrepareScannerDialog ezPageView, NumPages, StopScan, UseFeeder, ScanBacks
    
    If StopScan Then
        If frmScanBatchCancel.Visible = True Then
            frmScanBatchCancel.Hide
        End If
    End If
    

End Sub

Private Sub ezPageView_ScanBatchControl(Reserved As Integer)

' Occurs once per page when it is safe to change the filename.

' Within the scanning loop (the section of code that causes a page to be fed
'  into the scanner, scanned, the image transferred to the computer, and then saved),
'  there is only one time when it is safe to change the name of the resulting
'  image file. The ScanBatchControl event fires at this time, allowing your application
'  to change the name as necessary, perhaps based upon some characteristic of the image,
'  such as barcode data that has been detected, or a separator page.

'  You can also call the ScanDiscardPage method from within the ScanBatchControl event
'  to throw away the page. This may be useful when you are performing patchcode
'  or blank page detection and wish to discard the patchcode page or blank page when found.

' Note   Each time the application changes the filename, the current document
'  is closed and a new document is opened. This means that any existing images
'  in the tree view will disappear and a new tree view will begin to appear as
'  more pages are scanned into the new document. If your documents consist of
'  single page files, only one image will ever appear in the tree view window while scanning.

   funcWriteToDebugLog Me.name, "I101ScanViewer - ezPageView_ScanBatchControl - " & ezPageView.PageFileName
   DoEvents

'    If bolDiscardPage = True Then
'        ' Get rid of the page
'        ezPageView.ScanDiscardPage
'        bolDiscardPage = False
'    End If


End Sub

Private Sub Form_Load()
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    
    ' Get saved settings from the registry
    On Error Resume Next
    Imaging101ScanViewer.Top = VBGetPrivateProfileString(RegAppname, "Imaging101ScanViewer.Top", RegFileName)
    Imaging101ScanViewer.Left = VBGetPrivateProfileString(RegAppname, "Imaging101ScanViewer.Left", RegFileName)
    Imaging101ScanViewer.width = VBGetPrivateProfileString(RegAppname, "Imaging101ScanViewer.Width", RegFileName)
    Imaging101ScanViewer.Height = VBGetPrivateProfileString(RegAppname, "Imaging101ScanViewer.Height", RegFileName)
    On Error GoTo 0

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
'    SpicerView1.Height = Me.ScaleHeight - SpicerView1.Top
'    SpicerView1.width = Me.ScaleWidth
'
    ezPageView.Top = Me.Top + lblVersion.Top + lblVersion.Height + 10
    ezPageView.Height = Me.ScaleHeight - Me.Top
    ezPageView.width = Me.ScaleWidth
    
'    ImagXpress1.Height = Me.ScaleHeight - SpicerView1.Top
'    ImagXpress1.Width = Me.ScaleWidth
    
    
'      If frmAnnotate.Visible = True Then frmAnnotate.Unload

    On Error GoTo 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    'Save Form settings to the registry
    If Me.Top > 0 And Me.Left > 0 Then
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanViewer.Top", Imaging101ScanViewer.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanViewer.Left", Imaging101ScanViewer.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanViewer.Width", Imaging101ScanViewer.width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "Imaging101ScanViewer.Height", Imaging101ScanViewer.Height, RegFileName)
    End If

    On Error GoTo 0

End Sub



Private Sub SpicerView1_DblClick()

    cmdZoomFit_Click
    
End Sub




Private Sub SpicerView1_UserScroll()
''    MsgBox "userscroll"
    
End Sub

Private Sub SpicerView1_VectorObjectPlaced(ByVal LayerID As Long, ByVal vectObjectID As Long)
    
    Me.SetFocus
    
End Sub

Private Sub cmdZoomFit_Click()

''    SpicerView1.ZoomLevel(0) = IN_ZOOM_SCALETOFIT

End Sub

Private Sub cmdZoomIn_Click()
   Dim ScaleScrollRotation As IScaleScrollRotation
   ' Set the object variable for the IScaleScrollRotation interface to the View Control object
   Set ScaleScrollRotation = Me.SpicerView1.object
   ' Scale the current page to fit to the window
   ScaleScrollRotation.ZoomStepSize(0) = 1.5
   ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_IN
   ' De-initialize the IScaleScrollRotation object variable
   Set ScaleScrollRotation = Nothing

End Sub

Private Sub cmdZoomOut_Click()

   Dim ScaleScrollRotation As IScaleScrollRotation
   ' Set the object variable for the IScaleScrollRotation interface to the View Control object
   Set ScaleScrollRotation = Me.SpicerView1.object
   ' Scale the current page to fit to the window
   ScaleScrollRotation.ZoomStepSize(0) = 1.5
   ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_OUT
   ' De-initialize the IScaleScrollRotation object variable
   Set ScaleScrollRotation = Nothing


End Sub

Private Sub mnuCropAuto_Click()
   On Error GoTo ErrorOccurred
   Dim RasterTools As IRasterTools
   ' Set the object variable for the IRasterTools interface to the Edit Control object
   Set RasterTools = Me.SpicerEdit1.object
   ' Crop the active bilevel raster
   RasterTools.CropToMinimalSize
   ' De-initialize the object variable
   Set RasterTools = Nothing
   Exit Sub
   ' Process any errors. For sample code, click ErrorHandler.

ErrorOccurred:
'   ErrorHandler
   Exit Sub
End Sub



Private Sub EZ_GetBarcodeLoop()
    'Title:
    'KB 5726- How can I separate batches based on barcodes, if the barcode is on the second page?
    '
    'Question:
    'How can I separate batches based on barcodes, if the barcode is on the second page?
    '
    'Answer:
    'You will have to process the batch after scanning. Here is one way to do it:
    'Call this sub from the ezPageView_DoneScanning() sub
    
'    Dim strfname As String
'    Dim cnt As Integer
'    strfname = Empty 'Set the page index to the first page of the file
'    For cnt = 1 To ezPageView.PageCount
'        ezPageView.PageIndex = cnt
'        strfname = EZ_GetBarcode(ezPageView)
'        If strfname <> Empty Then
'            ezPageView.SaveFileName = strfname
'            ezPageView.SaveRangeMode = 2
'            ezPageView.SaveRange = CStr(ezPageView.PageIndex - 1) & "-" & CStr(ezPageView.PageIndex)
'            ezPageView.SavePages
'            strfname = Empty
'        End If
'    Next cnt
End Sub


Private Function EZ_GetBarcode(PixImage As PixEzImage) As String
'    Dim barfilter As New BinaryBarcodeDetection
'    Dim Image1 As IImage
'    Dim cnt As Integer
'    Dim bcVal As BarcodeResult
'    Set Image1 = barfilter.Run(PixImage.image)
'    If Image1.Results.Count > 0 Then
'        For cnt = 1 To Image1.Results.Count
'            Select Case Image1.Results.item(i).Key
'                Case "Barcode" ' Get the barcode text
'                    Set bcVal = Image1.Results.item(i).Value
'                    GetBarcode = "M:\tmp\ben\" & bcVal.Text & ".tif"
'                    Exit Function
'            End Select
'        Next cnt
'    End If
End Function

