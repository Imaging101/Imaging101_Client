VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{71C182E1-878D-11D1-8108-020701190C00}#8.0#0"; "view.ocx"
Object = "{22B7B2BB-4EFA-11D2-81FC-0000D1108734}#8.0#0"; "Edit.ocx"
Object = "{895CDC7A-8837-11D1-8109-020701190C00}#8.0#0"; "docctrl.ocx"
Object = "{C8B15BE2-E8D8-11D1-818A-0000D1108734}#8.0#0"; "SpiConfg.ocx"
Object = "{CA5948F6-E0F8-11D1-9A59-0000929B58F0}#8.0#0"; "mark.ocx"
Object = "{B5893B58-701E-4110-9871-1DA14CF9C1DC}#14.2#0"; "GdPicture.NET.14.tlb"
Begin VB.Form ChildForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SpicerView1"
   ClientHeight    =   10770
   ClientLeft      =   5325
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   7515
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin GdPicture_NET_14Ctl.GdViewer GdViewer1 
      Height          =   795
      Left            =   6210
      TabIndex        =   27
      Top             =   1095
      Visible         =   0   'False
      Width           =   690
      BackColor       =   "0"
      EnableDeferredPainting=   "True"
      PageDisplayMode =   "MultiplePagesView"
      EnableTextSelection=   "True"
      PreserveViewRotation=   "True"
      AnnotationResizeRotateHandlesScale=   "1"
      AnnotationDropShadow=   "True"
      AnnotationEnableMultiSelect=   "True"
      AllowDropFile   =   "False"
      HQAnnotationRendering=   "True"
      RenderGdPictureAnnots=   "True"
      EnableICM       =   "False"
      ZoomCenterAtMousePosition=   "False"
      EnabledProgressBar=   "True"
      DrawPageBorders =   "True"
      PageBordersPenSize=   "1"
      PageBordersColor=   "-16777216"
      PdfShowDialogForPassword=   "True"
      PdfShowOpenFileDialogForDecryption=   "True"
      PdfEnableFileLinks=   "True"
      PdfIncreaseTextContrast=   "False"
      PdfVerifyDigitalCertificates=   "False"
      ScrollBars      =   "True"
      ForceScrollBars =   "False"
      EnableMenu      =   "True"
      EnableFuzzySearch=   "False"
      Zoom            =   "1"
      ViewRotation    =   "RotateNoneFlipNone"
      MouseMode       =   "MouseModePan"
      MagnifierWidth  =   "160"
      MagnifierHeight =   "90"
      MagnifierZoomX  =   "2"
      MagnifierZoomY  =   "2"
      ZoomStep        =   "25"
      RectBorderSize  =   "1"
      ScrollSmallChange=   "1"
      ScrollLargeChange=   "50"
      SilentMode      =   "True"
      ForceTemporaryMode=   "False"
      IgnoreDocumentResolution=   "False"
      LockViewer      =   "False"
      ZoomMode        =   "ZoomMode100"
      EnableMouseWheel=   "True"
      DocumentAlignment=   "DocumentAlignmentMiddleCenter"
      DocumentPosition=   "DocumentPositionMiddleCenter"
      AnimateGIF      =   "True"
      DisplayQuality  =   "DisplayQualityAutomatic"
      DisplayQualityAuto=   "True"
      PdfDisplayFormField=   "True"
      PdfEnableLinks  =   "True"
      KeepDocumentPosition=   "False"
      MouseWheelMode  =   "MouseWheelModeZoom"
      Gamma           =   "1"
      RectIsEditable  =   "True"
      RegionsAreEditable=   "True"
      ClipRegionsToPageBounds=   "True"
      ClipAnnotsToPageBounds=   "True"
      ContinuousViewMode=   "True"
      MouseButtonForMouseMode=   "MouseButtonLeft"
      ForeColor       =   "Black"
      Location        =   "414, 73"
      Name            =   "GdViewer"
      Size            =   "46, 53"
      Object.TabIndex        =   "0"
   End
   Begin VB.TextBox txtFormArrayRetrieveIndex 
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
      Left            =   105
      TabIndex        =   26
      Text            =   "txtFormArrayRetrieveIndex"
      Top             =   10020
      Visible         =   0   'False
      Width           =   2085
   End
   Begin MSComctlLib.ProgressBar ProgressBarLoading 
      Height          =   330
      Left            =   780
      TabIndex        =   25
      Top             =   75
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtDocHeight 
      Height          =   285
      Left            =   5535
      TabIndex        =   24
      Text            =   "txtDocHeight"
      Top             =   9150
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox txtDocWidth 
      Height          =   285
      Left            =   5535
      TabIndex        =   23
      Text            =   "txtDocWidth"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtBatchPageFileName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3915
      TabIndex        =   22
      Text            =   "txtBatchPageFileName"
      Top             =   10005
      Visible         =   0   'False
      Width           =   2745
   End
   Begin SPICERCONFIGURATIONLib.SpicerConfiguration SpicerConfiguration1 
      Left            =   4515
      Top             =   1050
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERMARKUPLib.SpicerMarkup SpicerMarkup1 
      Left            =   3915
      Top             =   1050
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICEREDITLib.SpicerEdit SpicerEdit2 
      Left            =   6075
      Top             =   480
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICEREDITLib.SpicerEdit SpicerEdit1 
      Left            =   5355
      Top             =   450
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERDOCUMENTLib.SpicerDoc SpicerDoc2 
      Left            =   4515
      Top             =   450
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERDOCUMENTLib.SpicerDoc SpicerDoc1 
      Left            =   3915
      Top             =   450
      _Version        =   524288
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin SPICERVIEWLib.SpicerView SpicerView2 
      Height          =   1095
      Left            =   1995
      TabIndex        =   21
      Top             =   450
      Width           =   855
      _Version        =   524288
      _ExtentX        =   1508
      _ExtentY        =   1931
      _StockProps     =   0
   End
   Begin SPICERVIEWLib.SpicerView SpicerView1 
      Height          =   1095
      Left            =   1035
      TabIndex        =   20
      Top             =   465
      Width           =   735
      _Version        =   524288
      _ExtentX        =   1296
      _ExtentY        =   1931
      _StockProps     =   0
   End
   Begin VB.TextBox txtFTStatus 
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
      Left            =   3705
      TabIndex        =   19
      Text            =   "txtFTStatus"
      Top             =   8265
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtFileDirectory 
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
      Left            =   5505
      TabIndex        =   18
      Text            =   "txtFileDirectory"
      Top             =   9705
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox flbAnnotFiles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5535
      Pattern         =   "X.X"
      TabIndex        =   17
      Top             =   8625
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CheckBox chkCopyFileToLocalTempDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "chkCopyFileToLocalTempDir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   16
      Top             =   8250
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.TextBox txtChildFormMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5565
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2085
      Visible         =   0   'False
      Width           =   7050
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5475
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstPageList 
      Height          =   1425
      ItemData        =   "frmChildForm1.frx":0000
      Left            =   0
      List            =   "frmChildForm1.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPageFileName 
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
      Left            =   3705
      TabIndex        =   12
      Text            =   "txtPageFileName"
      Top             =   9705
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtAnnotationLayerID 
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
      Left            =   3705
      TabIndex        =   11
      Text            =   "txtAnnotationLayerID"
      Top             =   9345
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtApplicationRECID 
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
      Left            =   3705
      TabIndex        =   10
      Text            =   "txtApplicationRECID"
      Top             =   8985
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPageNumber 
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
      Left            =   1905
      TabIndex        =   9
      Text            =   "txtPageNumber"
      Top             =   9345
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPageRotation 
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
      Left            =   1905
      TabIndex        =   8
      Text            =   "txtPageRotation"
      Top             =   9705
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtModuleIndex 
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
      Left            =   105
      TabIndex        =   7
      Text            =   "txtModuleIndex"
      Top             =   8625
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPageCount 
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
      Left            =   1905
      TabIndex        =   6
      Text            =   "txtPageCount"
      Top             =   8985
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtImageNumber 
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
      Left            =   3705
      TabIndex        =   5
      Text            =   "txtImageNumber"
      Top             =   8625
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtImageCount 
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
      Left            =   1905
      TabIndex        =   4
      Text            =   "txtImageCount"
      Top             =   8625
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtDocumentRECID 
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
      Left            =   105
      TabIndex        =   3
      Text            =   "txtDocumentRECID"
      Top             =   9705
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtArrayIndex 
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
      Left            =   105
      TabIndex        =   2
      Text            =   "txtArrayIndex"
      Top             =   8985
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtDetailRECID 
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
      Left            =   105
      TabIndex        =   1
      Text            =   "txtDetailRECID"
      Top             =   9345
      Visible         =   0   'False
      Width           =   1695
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10515
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Object.ToolTipText     =   "File Type"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLoadingPages 
      Caption         =   "*********** Loading Pages... ***********"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   3075
      TabIndex        =   15
      Top             =   330
      Visible         =   0   'False
      Width           =   735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ChildForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bolChildFormLoadComplete As Boolean
Private bolRelatedImagesLoaded As Boolean
'Private bolAnnotationAdded As Boolean   'Moved to VariableDeclarations module as bolAnnotationsAdded

    '*** 2021-07-14 - Jacob - Moved  Dim bolLaunchThisFileType from MainMDIForm.funcShowImage()
    Private bolLaunchThisFileType As Boolean

Private arrPageRotation() As Integer
Private arrDetailRECID() As Double
Private arrPageLayerID() As Double
Private arrPageFileName() As String

'*** 2022-07-19 - Jacob - Added Arrays for Page Width and Page Height
Private arrPageWidth() As Double
Private arrPageHeight() As Double

Private intLockedRotation As Integer

Private docContents As IDocContents
Private ActivePage As IActivePage

Private m_NativeImageID As Long
Private m_GdPictureImaging As New GdPictureImaging
Private m_GdPictureDocumentConverter As New GdPictureDocumentConverter
Private m_GdPicturePDFReducer As New GdPicturePDFReducer
Private m_GdPicturePDF As New GdPicturePDF




Private Sub cmdPrint_Click()
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    Set PrintView = Me.SpicerView1.object
    PrintView.PrintDialog
    Set PrintView = Nothing

End Sub

Private Sub cmdAnnotate_Click()
   
   funcWriteToDebugLog Me.name, "*** ENTERING cmdAnnotate_Click()"
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error GoTo ErrorOccurred
   
   '
   ' Enable Layer
   '
   Dim docContents As IDocContents
   Dim lNewLayerID As Long
   Dim UserTools As IUserTools

   'Set FLAG to show that an Annotation or Modification may have occured.
   bolAnnotationAdded = True

   
   ' Set the object variable for the IDocContents interface to the Document Control object
   Set docContents = Me.SpicerDoc1.object
   Set UserTools = Me.SpicerMarkup1.object
   
    If (ActiveLayerID) Then
        Me.SpicerMarkup1.ActiveTool = IN_TOOL_TEXT
    Else
' To ensure that you can use all of the commands in
'  the Markup Control, bind one Markup Control to a
'  View Control and to a Document Control at the same
'  time, as long as the Document Control is also
'  bound to the same View Control.
        UserTools.BindToViewControl Me.SpicerView1
        docContents.NewLayer Me.SpicerView1.ActivePageId, IN_LAYER_FULLEDIT
        UserTools.ActiveTool = IN_TOOL_TEXT
'        Me.SpicerMarkup1.BindToDocumentControl Me.SpicerDoc1
'        Me.SpicerView1.BindToDocumentControl Me.SpicerDoc1
    End If
'   Me.SpicerDoc1.Save Me.SpicerDoc1.LayerID, True, API_FF_EDT, "c:\test.edt", "Layer " & Me.SpicerDoc1.LayerID
   
   ' De-initialize the object variables
   Set docContents = Nothing
   Set UserTools = Nothing
   
Exit Sub
   '
   
ErrorOccurred:
   ErrorHandler
   
End Sub





Private Sub Form_Activate()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    
    funcWriteToDebugLog Me.name, "Form_Activate() | *** ENTERING Form_Activate()"
    
    'If Rasterizing a Document or Loading Related Images,
    '   DO NOT process the Load and Activate subs
    If bolRasterizingDocument = True Or bolRelatedImagesLoaded = False Or bolObjectLaunched = True Then
            funcWriteToDebugLog Me.name, "Form_Activate() | *** EXIT Form_Activate() | bolRasterizingDocument=" & bolRasterizingDocument & "  bolRelatedImagesLoaded=" & bolRelatedImagesLoaded & "  bolObjectLaunched=" & bolObjectLaunched
           Exit Sub
    End If
    
    
    arrPageFileName(txtPageNumber) = txtPageFileName
        
    '*** Show the Batch Menu options ONLY if viewing a Batch Page and frmIndex is visible
    '     HAVE to USE the funcIsFormLoaded function to make sure we don't LOAD the Index form
    If InStr(1, Me.Caption, "BATCH:") And funcIsFormLoaded2("frmIndex") = True Then
        
        txtApplicationRECID = frmIndex.txtApplicationRECID
        
        MainMDIForm.cmdEdit.Visible = False
        MainMDIForm.mnuBatch.Visible = False
        
        If frmIndex.txtBatchCommitStatus <> "Committed-FULL" Then
        
            MainMDIForm.mnuBatch.Visible = True
        
            If gsecAllowModificationOfOrigDocs = vbChecked Then
                MainMDIForm.cmdEdit.Visible = True
            Else
                MainMDIForm.cmdEdit.Visible = False
            End If
    
        End If
        
        subSetCurrentPage
        
    ElseIf Trim(txtApplicationRECID) = "txtApplicationRECID" Then
        txtApplicationRECID = frmImaging101Search.txtApplicationRECID
        MainMDIForm.mnuBatch.Visible = False
        
        MainMDIForm.cmdEdit.Visible = False
    
    End If
    
    
    '*** 2020-09-29 - Jacob - Moved Get Security Rights UP from after Rasterizing, etc. check
    '***************************************************************
    '***  GET SECURITY RIGHTS
    
    funcGetSecurityRights gsecSecurityRECID, txtApplicationRECID
    MainMDIForm.subCheckButtonSecurity
    
    
   

    

    
    
''''    '*** Initialize the Child Form if NOT a Batch Page
''''    '    The frmIndex.ListView1_Click() will call the subInitializeChildForm() for each image selected
''''    If Not InStr(1, Me.Caption, "BATCH:") Then
''''        subDebugChildForm "Form-Activate: Before subInitializeChildForm"
''''        subInitializeChildForm
''''        subDebugChildForm "Form-Activate: After  subInitializeChildForm"
''''    End If
    

'XXX
'    subSetCurrentPage

    
    
    If funcIsFormLoaded2("frmThumb") Then
        frmThumb.SpicerThumbnail1.BindToViewControl Me.SpicerView1.object
        frmThumb.SpicerThumbnail1.Visible = True
        frmThumb.SetFocus
        DoEvents
    End If
    
    
End Sub

Public Sub subInitializeChildForm()

        funcWriteToDebugLog Me.name, "*** ENTERING subInitializeChildForm() | txtModuleIndex = " & txtModuleIndex
    
        
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    

    On Error GoTo ERROR_TRAP
    
    MainMDIForm.PictureButtonBar.Enabled = False
    
'   ' Get the pageID for all pages in the document
'   For iTemp = 1 To lPageCount
'      sTemp = sTemp + "Page" + Str(iTemp) + " ID: " + Str(docContents.pageID(iTemp)) + Chr(13)
'   Next iTemp
'   ' Display the number of pages and the identifiers for each
'   MsgBox sTemp, vbInformation
    
''''    'Add the Rotation for the FIRST Page to the List
''''    subDebugChildForm "subInitializeChildForm: Before  ReDim Preserve arrPageRotation(0 To 2)"
''''    ReDim Preserve arrPageRotation(0 To 2)
''''    ReDim Preserve arrDetailRECID(0 To 2)
''''    ReDim Preserve arrPageLayerID(0 To 2)
''''    subDebugChildForm "subInitializeChildForm: After   ReDim Preserve arrPageRotation(0 To 2)"
''''
''''    arrPageRotation(1) = txtPageRotation
''''
       
       
    If bolObjectLaunched = False Then
            
            funcWriteToDebugLog Me.name, "bolObjectLaunched = False"

    ''''    If Trim(txtDetailRECID) <> "" Then
            arrDetailRECID(1) = txtDetailRECID
            arrPageFileName(1) = txtPageFileName
            arrPageRotation(1) = txtPageRotation
    ''''    End If
    ''''
    
            funcWriteToDebugLog Me.name, "arrDetailRECID(1) = " & txtDetailRECID
            funcWriteToDebugLog Me.name, "arrPageFileName(1) = " & txtPageFileName
            funcWriteToDebugLog Me.name, "arrPageRotation(1) = " & txtPageRotation

        '*** Set the Active Page as the Current Page
'        'XXX
'        funcWriteToDebugLog Me.name, "subInitializeChildForm: subSetCurrentPage"
'        subSetCurrentPage
    
        '3/4/2013 - Jacob - Disabled subAnnotationLayerLoad... because it happens in subSetCurrentPage
'            subAnnotationLayerLoad
    
        '*** THE SPICER DOCUMENT PAGECOUNT INFO
        '***     IS SET IN  subSetCurrentPage
    
        If (txtCurrentModule = "frmImaging101Retrieve") _
            And (bolRelatedImagesLoaded = False) _
            And (bolAIM_Command_AddFile = False) Then
            
            funcWriteToDebugLog Me.name, "(txtCurrentModule = ""frmImaging101Retrieve"") And (bolRelatedImagesLoaded = False)"
            
            '*** RETRIEVING DOCUMENT -- LOAD RELATED PAGES
            funcWriteToDebugLog Me.name, "*** RETRIEVING DOCUMENT -- LOAD RELATED PAGES"
            
            frmImaging101Search.Enabled = False
            frmImaging101Retrieve.Enabled = False
    
            funcWriteToDebugLog Me.name, "*** subGetRelatedImages"
            subGetRelatedImages
            
            Me.lstPageList.Visible = True
            Me.lstPageList.SetFocus
            Me.lstPageList.Selected(0) = True

            frmImaging101Search.Enabled = True
            frmImaging101Retrieve.Enabled = True
            
        Else
        
                On Error Resume Next
                
                '*** RESET
                bolObjectLaunched = False
                lstPageList.Clear
                
                
                funcWriteToDebugLog Me.name, "(txtCurrentModule <> ""frmImaging101Retrieve"") OR (bolRelatedImagesLoaded <> False)"
        
                Dim txtFullPathName As String
                
                If bolAIM_Command_AddFile = True Then
                        'if Adding a file via AIM / I101Filer, don't copy it to the Local Temp dir
                        txtFullPathName = txtPageFileName
                Else
                        txtFullPathName = funcCopyFileToLocalTemp(txtBatchPageFileName, txtPageFileName, txtDetailRECID)
                End If

                '***     This is the FIRST Page - OPEN the FILE for the FIRST Detail Record
'                funcWriteToDebugLog Me.name, "*** OPEN the FILE"
'                 funcWriteToDebugLog Me.name, "txtPathSubdirectory = " & txtPathSubdirectory
'                 funcWriteToDebugLog Me.name, "txtFileName = " & txtFileName
'                 funcWriteToDebugLog Me.name, "txtFullPathName = " & txtFullPathName
'                 funcWriteToDebugLog Me.name, "txtPageFileName = " & txtPageFileName
'                 funcWriteToDebugLog Me.name, "txtDetailRECID = " & txtDetailRECID
'                 funcWriteToDebugLog Me.name, "txtPageRotation = " & txtPageRotation
'                 funcWriteToDebugLog Me.name, "dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport


                '**********************************************************************************************************
                '*** Check if File should be Launched
        
                bolObjectLaunched = funcCheckIfFileShouldBeLaunched(txtFullPathName)
                        
                        
                '***********************************************************************************************************
                '*** Load or Import ONLY if NOT flagged to AutoLaunch and No file open Errors detected
       
                If bolLaunchThisFileType = True Or Err.Number <> 0 Then
                
                    '*** DON'T LOAD OR IMPORT
                    Unload Me
                    
                Else

                    Set docContents = Me.SpicerDoc1.object
                    
                    '*** 2022-07-28 - Jacob - Added CloseDocument
                    docContents.CloseDocument False
                    
                    funcWriteToDebugLog Me.name, "subInitializeChildForm() | docContents.OpenFile " & txtFullPathName
                    docContents.OpenFile txtFullPathName
                    
                    Set docContents = Nothing
                    
    '                 Debug.Print docContents.NumberOfLayers(docContents.NewestObjectID)
                    
                    funcWriteToDebugLog Me.name, "subInitializeChildForm() | *** AFTER docContents.OpenFile | Err.Number = " & Err.Number
                    

                End If
                
                    
                    
                '***************************************************************************************************************
                '*** Check for Import Errors
                
                Call funcCheckForOpenOrImportErrors(Err.Number, Err.Description, txtFullPathName)
                
                If bolObjectLaunched = False Then
                
                        Me.txtPageCount = SpicerDoc1.NumberOfPages
        
                        DoEvents
                    
                    
                    
            Else
            
                        Me.txtPageCount = 1
        
                        DoEvents


            End If
            
            '***************************************************************************************************************
            '*** UPDATE PAGE ARRAYS
            
            funcWriteToDebugLog Me.name, "subInitializeChildForm() | subUpdatePageArrays "
            txtPageCount = SpicerDoc1.NumberOfPages
            subUpdatePageArrays 1, txtPageCount

            '***************************************************************************************************************
            '*** BIND THE SPICER CONTROLS
            
            Call BindControls

            

            '***************************************************************
            '*** Re-enable Forms
            
            If bolAIM_Command_AddFile = True Then
                    frmImaging101Search.Enabled = True
            Else
                    frmIndex.Enabled = True
            End If
            
            MainMDIForm.Enabled = True
            Me.Enabled = True
            
            '***************************************************************
              '*** Make Viewer objects Visible again
        
              Me.SpicerView1.Visible = True
              Me.lstPageList.Visible = True
              Me.lblLoadingPages.Visible = False
              Me.txtChildFormMessage.Visible = False
            
            DoEvents
            
        End If
        


    
        '*** BIND the Thumbnail control to the View control IF it's loaded.
        If funcIsFormLoaded2("frmThumb") = True Then
            funcWriteToDebugLog Me.name, "frmThumb.SpicerThumbnail1.BindToViewControl Me.SpicerView1.object "
            frmThumb.SpicerThumbnail1.BindToViewControl Me.SpicerView1.object
        End If
    
    
        '***************************************************
        '*** SET VIEWER CONFIGURATION OPTIONS
        funcWriteToDebugLog Me.name, "SET VIEWER CONFIGURATION OPTIONS"
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = Me.SpicerConfiguration1.object
        'Allow Zooming in the Viewer
        CFGDocument.AllowMouseZoomInWindow = True
        'Allow Showing the Context Menu on Right-Click
        CFGDocument.AllowRightContextMenuInWindow = True
    
        'Deinitialize the object variable
        Set CFGDocument = Nothing
        '****************************************************
    
    
    
    End If
    

    
    
    'Initialize the Flag for Annotations/Vector-Object Added
    bolAnnotationAdded = False
    
''''''    Unload frmLoadingImages

    funcWriteToDebugLog Me.name, "funcEnableFormsAfterLoadingImages"
    funcEnableFormsAfterLoadingImages
    
'    funcWriteToDebugLog Me.name, "Form_Activate"
'   Call Form_Activate
    
    MainMDIForm.PictureButtonBar.Enabled = True
    
    'Make the StatusBar visible
    StatusBar1.Visible = True
    
    
    

    
    
    
Exit Sub
   
ERROR_TRAP:

   ErrorHandler
   
   MainMDIForm.PictureButtonBar.Enabled = True

End Sub

Private Sub Form_Initialize()


    subDebugChildForm "*** ENTERING ChildForm1 Form_Initialize()"
    ReDim Preserve arrPageRotation(0 To 2)
    ReDim Preserve arrDetailRECID(0 To 2)
    ReDim Preserve arrPageLayerID(0 To 2)
    ReDim Preserve arrPageFileName(0 To 2)



End Sub

Private Sub Form_Load()
    
    On Error Resume Next

    funcWriteToDebugLog Me.name, "*** ENTERING ChildForm1 Form_Load()"

'    FileSystem.ChDir "C:\Open Text\Imagenation"
    
    'If Rasterizing a Document, DO NOT process the Load and Activate subs
    
    bolRelatedImagesLoaded = False
    
    lblLoadingPages.Top = Me.ScaleTop
    lblLoadingPages.Left = Me.ScaleLeft
    lblLoadingPages.Visible = True
    
    '*** SET UP GdPicture License
    funcWriteToDebugLog Me.name, "***Dim oLic As New LicenseManager"
    Dim oLicenseManager  As New LicenseManager
    Dim strLicense As String
    Dim bolRegistrationSuccessfull As Boolean
    
    strLicense = "I5p9Je_FCWBI3qNyfWRQXRG8oXwZW7dr6TqhmOIvDuGv4ZkI0BYI5mimOpNfOtBwMsaIq-Trm8yl55nGpWfQT2BSMtpiriOaEJIEA0ZbFArmyP24AsRnYFWXOP7"
    strLicense = strLicense + "lkpjqb10PQy7lm1mDEWQA5I08KxxowKsBMByCtjm4F6XA4vxdD6TSIDmglxnQJp_jH-9HtWMcitd7u-CSdA-s_SHqvcri2N_plJ_a4xJ7PZ7ipRMx5MrXpbNf_"
    strLicense = strLicense + "7qvO1Mo0iGKHfPaXGZKOi_PZ3_4DtGvriywnpffxlafWRH_EVCBjhk99J6CaRWHiH7JY-h4LNftASJLLa_w3_0woaJMHaVmtEyth3J618Yicc3PAUy3Lr5kd9gn"
    strLicense = strLicense + "0RVVWCYglQTnn-70gZSZbAwQ-MsWM63vWu0Sw_Xd6M_I-KVuaG423-72ouvBJRK-AhvYfVn4rSbBx1KqC7RJSent6DhGlb-V7vHyvFTBvJfj8mbTQcdaiI9EMrz"
    strLicense = strLicense + "_iUUVqIZnjm7UdzXEyCZXWukH8nsfeywEXGdebpM6p-Ig3uA5ZG9sIYVRJzIY8obRI4gGC5Nk_vd7KSL639iSSiRvTyQBY4IHXg6Cb-KkE5VkxgfeGIwv_SZncI"
    strLicense = strLicense + "MmI1gT9y1t8BbBF2jj3pnePVxj2YE29irA06JcICpXy-8IpQGPJ509p22BdGaFxCxcyg1VjSngxtmgBtnH32vAgJWkh47uHwUoufywC8iI5EvJ6Ml-jpvl-dPS96l"
    strLicense = strLicense + "uIXM55yY8EET9gzToDNCs783bKSnna64KaP6oQEg-GSeFSqY5_7HJ5UHR_2bhRQpJtTzD2vRNn34nIEq-XM6Nmd6U31MBQ2hGTDOiUIaZUGPhEtf2Erd1GSvCnFTVq"
    strLicense = strLicense + "_3gixuM5HHnueHwXCPQhHJDP3QHSjpVlvgVT4tb0OtaN5W1qH6H6zWurpBF3vdPH6HKq8h0dwkn6GoXxiyvwnxk_RJOE5cZi9v25t8kRmNH4LrsYwkoTFY2K2HTapTwJ9AOJ1PJ76Q="
    
    funcWriteToDebugLog Me.name, "*** Call RegisterKEY(...)"
    bolRegistrationSuccessfull = oLicenseManager.RegisterKEY(strLicense)
   
    If bolRegistrationSuccessfull = False Then
            funcWriteToDebugLog Me.name, "*** RegisterKEY FAILED"
            MsgBox "Viewer License Registration Failed."
            Exit Sub
    End If

    '*** 2020-04-23 - Jacob - HARD-CODED the chkCopyFileToLocalTempDir Option to ALWAYS copy files to the Local Temp directory.  This prevents WAN / Slow Router issues.
'    chkCopyFileToLocalTempDir = VBGetPrivateProfileString(RegAppname, "ChildForm1.chkCopyFilesToLocalTempDir", RegFileName)
    chkCopyFileToLocalTempDir = 1
    funcWriteToDebugLog Me.name, "chkCopyFileToLocalTempDir = " & chkCopyFileToLocalTempDir

''''    '*** Initialize the Displayed Pages Array
''''    ReDim arrPageRotation(0)

    'Can't load the frmLoadingImages form as MODAL because it stops
    ' all further program action!  SO does setting forms as Enabled=False !!!
    
'    frmLoadingImages.Show vbModal
'    funcDisableFormsWhileLoadingImages
    
    
''''''    FormOnTop frmLoadingImages.hwnd, True
''''''    frmLoadingImages.Show
''''''
''''''    frmLoadingImages.txtImageNumber = 1
''''''    DoEvents
    
    'Add the Rotation for the FIRST Page to the List

'    subDebugChildForm "Form Load: After   ReDim Preserve arrPageRotation(0 To 2)"
    

''''    arrPageRotation(1) = txtPageRotation
''''
''''    If Trim(txtDetailRECID) <> "" Then
''''        arrDetailRECID(1) = txtDetailRECID
''''    End If
        
    
End Sub


Private Sub Form_LostFocus()

    funcWriteToDebugLog Me.name, "LostFocus"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Do NOT allow this form to be unloaded while it's still loading.
    If (funcIsFormLoaded2("frmIndex") And (Not bolIndexFormLoadComplete)) _
    Or (bolAIM_Command_AddFile = True And UnloadMode = vbUser) Then
        Cancel = True
        Exit Sub
    End If

    
''''''''''    If SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId) > 1 Then
''''''''''
''''''''''        result = MsgBox("You have placed Annotations on this page..." & vbCrLf & "Would you like to save them?", vbYesNo, "Save Annotations?")
''''''''''        If result = vbYes Then
''''''''''            subAnnotationLayerSave
''''''''''        End If
''''''''''
''''''''''    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    'Check if Annotations were added and users wishes to Save them
    subAnnotationLayerSaveCheck
    
'*** 2/23/2004 -JACOB - Disabled the CloseDocument Code below because it was generating
'                       a dialog box asking “Save changes to document 00000000991 (TIF)?”

'    Dim docContents As IDocContents
'        ' Set the object variable for the IDocContents interface to the Document Control object
'    Set docContents = Me.SpicerDoc1.object
'    ' Close the document in the SpicerDoc1 control and
'    ' Set CloseDocument to "True" to check if the document has been changed.
'     docContents.CloseDocument False
'    ' De-initialize the object variable
'    Set docContents = Nothing

'*** 4/27/2012 - JACOB - Added CloseDocument again.  Apparently SPICER/OpenText fixed the problem
'                                    Was giving a "Permission Denied" error when attempting to Delete a PDF file copied to LOCAL HD
    Me.SpicerDoc1.CloseDocument False

    'DELETE ALL FILES COPIED TO LOCAL TEMP DIR
    For i = 2 To UBound(arrPageFileName)
        If Right(arrPageFileName(i), 4) = ".tmp" Then
            If funcFileExists(arrPageFileName(i)) Then
                funcWriteToDebugLog Me.name, "KILL " & arrPageFileName(i)
                Kill arrPageFileName(i)
            End If
        End If
    Next
    
    
    
    ' Save Annotation Layer(s)
    
    funcWriteToDebugLog Me.name, "Form.Unload | UNLOAD frmAnnotate & frmThumb"


    If funcIsFormLoaded2("frmAnnotate") Then
        Unload frmAnnotate
        Set frmAnnotate = Nothing
    End If
    
    If funcIsFormLoaded2("frmThumb") Then
        Unload frmThumb
        Set frmThumb = Nothing
    End If
    
    
        'Remove the entry for THIS document from the Displayed Pages Array
    '  Should only work for Retrieved Pages -- NOT Batch Pages
'        MainMDIForm.subRemoveDisplayedPageFromArray CInt(txtModuleIndex), txtArrayIndex
     '*** 2022-07-28 - Jacob - Replaced txtArrayIndex with txtFormArrayRetrieveIndex
     
    Dim intI101Module As Integer
     intI101Module = CInt(txtModuleIndex)
     
    Select Case intI101Module
        
            Case gI101ModuleRetrieve
     
                    MainMDIForm.subRemoveChildFormFromArray intI101Module, txtFormArrayRetrieveIndex

            Case gI101ModuleIndex
            
                    MainMDIForm.subRemoveChildFormFromArray intI101Module, txtArrayIndex
                    
            Case Else
            
    End Select

    'SAFE way of saying: Set Me = Nothing
    funcWriteToDebugLog Me.name, "Form.Unload | BEGIN SAFE UNLOAD of ChildForm1"
    
    Dim Form As Form
    For Each Form In Forms
            If Form Is Me Then
                    funcWriteToDebugLog Me.name, "ChildForm1 = Nothing"
                    Set Form = Nothing
                    funcWriteToDebugLog Me.name, "Exit For"
                    Exit For
            End If
    Next Form
    

    
End Sub

Public Sub ErrorHandler()
   Dim sErrMessage As String
   ' Display a message box for system errors.
   ' Everything above 0 catches Visual Basic errors.
   ' Below 0 catches Spicer ActiveX Control errors.
   If Err.Number >= 0 Then
      funcQuickMessage "SHOW", "Error returned:" + Chr(13) + "Error Number:" + str(Err.Number) + Chr(13) + "Description:" + Err.Description + Chr(13) + "Last DLL Error:" + str(Err.LastDllError) + Chr(13) + "Source:" + Err.Source + Chr(13)
   Else
      ' Display a message box for errors in the Document Control

      sErrMessage = Me.SpicerDoc1.ErrorMessage(Err.Number)
      funcQuickMessage "SHOW", "Error returned:" + Chr(13) + "Error Number:" + str(Err.Number) + _
      Chr(13) + "Description:" + sErrMessage + Chr(13) + "Last DLL Error:" + _
      str(Err.LastDllError) + Chr(13) + "Source:" + Err.Source + Chr(13)
   End If
   Err.Clear
End Sub

Private Sub Form_Resize()
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    'Me.Top = Me.SpicerView1.PictureButtonBar.bottom
    
    'Size Page List
    Me.lstPageList.Left = Me.ScaleLeft
    Me.lstPageList.Top = Me.ScaleTop
    
    If Me.ScaleHeight > Me.StatusBar1.Height Then
        Me.lstPageList.Height = Me.ScaleHeight - Me.StatusBar1.Height
    End If
    
    'Size Spicer Viewer
    Me.SpicerView1.Top = Me.ScaleTop
    Me.SpicerView1.Left = Me.ScaleLeft + lstPageList.width
    If Me.ScaleHeight > Me.StatusBar1.Height Then
        Me.SpicerView1.Height = Me.ScaleHeight - Me.StatusBar1.Height - 120
    End If
    
    Me.SpicerView1.width = Me.ScaleWidth - Me.SpicerView1.Left
    
    'Center the Message Text Box
    txtChildFormMessage.Left = (Me.ScaleWidth - txtChildFormMessage.width) / 2
    
    'Re-size the StatusBar Panels
    For i = 1 To StatusBar1.Panels.Count
        StatusBar1.Panels(i).width = Me.ScaleWidth / StatusBar1.Panels.Count
    Next
    
'    On Error GoTo 0
    
End Sub

Private Sub lstPageList_Click()
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error Resume Next
    
    '*** 2023-03-07 - Jacob - NO IDEA WHY I DID THIS... WITHOUT CHECKING IF WE ARE IN THE INDEXING MODULE ???
    'gFormArrayIndex(0).Show
    
    'Check if Annotations were added and users wishes to Save them
    Me.subAnnotationLayerSaveCheck
    
    Me.SpicerView1.GotoPage lstPageList.Text

    Me.subSetCurrentPage
    
    Me.SpicerView1.SetFocus
    Me.lstPageList.SetFocus
    
    '***2022-03-23 - Jacob - Changed from "result=" to "strZoomResult" to avoid Type Mismatch error.
    Dim strZoomResult As String
    strZoomResult = MainMDIForm.funcZoomToSavedFactor

    If bolErrorOccured Then
            MsgBox "funcZoomToSavedFactor() ERROR:" & vbCrLf & " While attempting to set Saved ZOOM Factor." & vbCrLf & vbCrLf & strZoomResult, vbCritical
    End If

End Sub

Private Sub lstPageList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyDown, vbKeyUp
            lstPageList_Click
    End Select
    
End Sub

Private Sub SpicerEdit1_NewestPageID(ByVal pageID As Long)

         'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    ' Declare the interface variable
    Dim docContents As IDocContents
    ' Assign the interface to the Document Control object
    Set docContents = Me.SpicerDoc2.object
    ' Use commands from the interface, such as OpenFile
    docContents.CloseDocument (False)
    docContents.InsertDocument (pageID)
    ' To clean up memory, set the object variable to Nothing when complete
    
    Set docContents = Nothing
    
    
''    Me.SpicerDoc2.CloseDocument (False)
''    Me.SpicerDoc2.InsertDocument (pageID)
''    DoEvents
''    Me.SpicerView2.BindToDocumentControl Me.SpicerDoc2.object
''    Me.SpicerView2.Refresh

End Sub

Private Sub SpicerView1_ChangePage(ByVal DocwinID As Long, ByVal pageID As Long, ByVal pageNum As Integer)
    
''    funcWriteToDebugLog Me.name, "SpicerDoc1.NumberOfLayers: " & SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId)
'    txtPageNumber = pageNum
'
'    txtDetailRECID = arrDetailRECID(txtPageNumber)
'
'    BindControls
'
'    If SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId) < 1 Then
'        subAnnotationLayerLoad
'    End If

    subAnnotationLayerSaveCheck

End Sub

Private Sub SpicerView1_Click()

    Select Case txtCurrentModule
        Case "frmImaging101Search", "frmImaging101Retrieve"
            'Set the focus to the Active Form containing the SpicerView1 control
            Me.SetFocus
            subDebugChildForm "SpicerView1_Click(): Case Search/Retrieve"

    End Select
    
End Sub

Private Sub SpicerView1_ClickVectorObject(ByVal LayerID As Long, ByVal vectObjectID As Long)
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    SpicerView1.MinimizeAnnotations
    

End Sub

Private Sub SpicerView1_DblClick()

    Me.cmdZoomFit_Click

End Sub




Private Sub SpicerView1_LostFocus()

'''    subAnnotationLayerSaveCheck

End Sub

Private Sub SpicerView1_UserScroll()
''    MsgBox "userscroll"
    
End Sub

Private Sub SpicerView1_UserZoom()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
''    MsgBox "Zoom"
    Select Case txtCurrentModule
    
        Case "frmIndex"
            If funcIsFormLoaded2("frmLookupList") _
            And gOpenBatchInReadOnlyMode = False Then
                If frmLookupList.Visible = True Then
                    frmLookupList.txtTableLookupField.SetFocus
                End If
            End If
            
        Case "frmOCRClip"
        
        
               Dim Miscellaneous As IMiscellaneous
               Dim sFilename As String
               ' Set the object variable for the IMiscellaneous interface to the View Control object
               Set Miscellaneous = Me.SpicerView1.object
               ' Get the name and location for the OCR region
'               sFilename = InputBox("Please enter the file name and location to save the image to.", _
                           "Saving OCR Region")
               ' Save the entire current page as a TIFF LSB - CCITT G4 file
               ''Miscellaneous.OCRRegion sFilename, API_FF_TIFFL, IN_UNITS_PROPORTIONAL, 0, 0, 0, 0
            
                sFilename = "C:\PegasusSoftware\SmartscanICRv30\ActiveX\VB\DIB1.BMP"
                
                'Me.SpicerView1.ZoomFactorX = 85189
                'Me.SpicerView1.ZoomFactorY = 4944
                'Me.SpicerView1.ZoomFactorX = 89784
                'Me.SpicerView1.ZoomFactorY = 6801

                
               Miscellaneous.OCRRegion sFilename, API_FF_BMP, IN_UNITS_PROPORTIONAL, 80000, 2500, 89784, 6801
               ' De-initialize the IMiscellaneous object variable
               Set Miscellaneous = Nothing
                    
        
        
'            Dim docContents As IDocContents
'            Dim lPageID As Long
'            Dim lLayerID As Long
'            Set docContents = Me.SpicerDoc1.object
'            lPageID = Me.SpicerView1.ActivePageId
'            lLayerID = docContents.layerID(lPageID, 1)
'            Set docContents = Nothing
'
'            Dim RasterBatch As IRasterBatch
'            ' Set the object variable for the IRasterBatch interface to the Edit Control object
'            Set RasterBatch = Me.SpicerEdit1.object
'            ' Crop the image on layer 1028 at the specified coordinates (in pixels)
'            RasterBatch.CropToDefinedSize lLayerID, 200, 200, 300, 300
'            ' De-initialize the object variable
'            Set RasterBatch = Nothing
'
'
'            Dim Miscellaneous As IMiscellaneous
'            ' Set the object variable for the IMiscellaneous interface to the View Control object
'            Set Miscellaneous = Me.SpicerView1.object
'            ' Copy the active document to the Clipboard
'            Miscellaneous.CopyDocumentToOleObject
'            ' De-initialize the IMiscellaneous object variable
'            Set Miscellaneous = Nothing
'
            
            
            ''''mod_ISpicerView.ISpicerView_hWnd
            frmOCRClip.cmdAnalyze_Click
    End Select
        
    DoEvents
End Sub

Private Sub SpicerView1_VectorObjectPlaced(ByVal LayerID As Long, ByVal vectObjectID As Long)
    
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error GoTo ERROR_TRAP
    
    funcWriteToDebugLog Me.name, "ENTER Sub SpicerView1_VectorObjectPlaced()"
    
    Me.SetFocus
    
    funcWriteToDebugLog Me.name, "subResetAnnotationButtons"
'    subResetAnnotationButtons
    
'    funcWriteToDebugLog  Me.Name, "SelectTool IN_TOOL_NOTOOL"
'    SelectTool IN_TOOL_NOTOOL
    
    bolAnnotationAdded = True
    
    funcWriteToDebugLog Me.name, "EXIT Sub SpicerView1_VectorObjectPlaced()"
   
Exit Sub
   
ERROR_TRAP:

   funcWriteToDebugLog Me.name, "ERROR_TRAP:  SpicerView1_VectorObjectPlaced: Error=" & Err.Number & " - " & Err.Description
   ErrorHandler

   
End Sub

Public Sub cmdZoomFit_Click()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
''''   Dim ScaleScrollRotation As IScaleScrollRotation
''''   ' Set the object variable for the IScaleScrollRotation interface to the View Control object
''''   Set ScaleScrollRotation = Me.SpicerView1.object
''''   Set ActivePage = Me.SpicerDoc1.object
''''   ' Scale the current page to fit to the window
''''   ScaleScrollRotation.ZoomLevel(Me.SpicerDoc1.object) = IN_ZOOM_SCALETOFIT
''''   ' De-initialize the IScaleScrollRotation object variable
''''   Set ScaleScrollRotation = Nothing
    SpicerView1.ZoomLevel(0) = IN_ZOOM_SCALETOFIT

End Sub

Public Sub cmdZoomIn()
   
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
   Dim ScaleScrollRotation As IScaleScrollRotation
   ' Set the object variable for the IScaleScrollRotation interface to the View Control object
   Set ScaleScrollRotation = Me.SpicerView1.object
   ' Scale the current page to fit to the window
   ScaleScrollRotation.ZoomStepSize(0) = 1.5
   ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_IN
   ' De-initialize the IScaleScrollRotation object variable
   Set ScaleScrollRotation = Nothing

End Sub

Public Sub cmdZoomOut()
   
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
   Dim ScaleScrollRotation As IScaleScrollRotation
   ' Set the object variable for the IScaleScrollRotation interface to the View Control object
   Set ScaleScrollRotation = Me.SpicerView1.object
   ' Scale the current page to fit to the window
   ScaleScrollRotation.ZoomStepSize(0) = 1.5
   ScaleScrollRotation.ZoomLevel(0) = IN_ZOOM_OUT
   ' De-initialize the IScaleScrollRotation object variable
   Set ScaleScrollRotation = Nothing


End Sub

Public Sub cmdScaleToGray()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
'    Load frmViewForm
'    Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHIMAGES) = IN_OFF
'    Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHLINEART) = IN_OFF
'    Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHTEXT) = IN_OFF
'    Me.SpicerView1.Refresh
'    DoEvents
'    Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHIMAGES) = IN_ON
'    Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHLINEART) = IN_ON
'    Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHTEXT) = IN_ON
'    Me.SpicerView1.Refresh
'    DoEvents
'    Unload frmViewForm
    
    funcWriteToDebugLog Me.name, "cmdScaleToGray()"
    
    Me.SpicerView1.RefreshMode = False
 
   If MainMDIForm.mnuScaleToGray.Checked Then
        ' Switch scale to gray ON for the active page (IN_TOGGLE would toggle On/Off)
        Me.SpicerView1.ScaleToGray(0) = IN_ON
        Me.SpicerView1.Sample(0) = IN_OFF
        Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHIMAGES) = IN_ON
        Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHLINEART) = IN_ON
        Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHTEXT) = IN_ON
       
   Else
        ' Switch scale to gray OFF for the active page
        Me.SpicerView1.ScaleToGray(0) = IN_OFF
        Me.SpicerView1.Sample(0) = IN_OFF
        Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHIMAGES) = IN_OFF
        Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHLINEART) = IN_OFF
        Me.SpicerView1.ViewDisplay(0, IN_VIEWDISPLAY_SMOOTHTEXT) = IN_OFF
        
    End If
    
    Me.SpicerView1.RefreshMode = True
        
'    Me.SpicerView1.Refresh
    DoEvents

End Sub

Public Sub cmdImageRotateLeft()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error GoTo RotateLeft_ERROR
    
'    ' Flag image to be overwritten
'    subConfigRasterOperationsOverwrite
'
'   Dim RasterTools As IRasterTools
'   ' Set the object variable for the IRasterTools interface to the Edit Control object
'   Set RasterTools = Me.SpicerEdit1.object
'   ' Turn the raster image 180 degrees
'   RasterTools.DataRotate IN_ROTATION_90_CCW
'   ' De-initialize the object variable
'   Set RasterTools = Nothing
'
'
'   Dim docSave As IDocSave
'   ' Set the object variable for the IDocSave interface to the Document Control object
'   Set docSave = Me.SpicerDoc1.object
'   ' Export the first page in the document, retaining its current format
''   docSave.Export Me.SpicerDoc1.FirstPageID, False, 0, txtFullPathName, "JPEG"
'   docSave.Save 0, False, 0, txtFullPathName, "1"
'   ' De-initialize the object variable
'   Set docSave = Nothing

   
    subRotateImage IN_ROTATION_90_CCW, True
   
Exit Sub

RotateLeft_ERROR:
    MsgBox "Sorry... this document type cannot be saved as rotated!", vbInformation, "Rotate Left"
    
End Sub


Public Sub cmdImageRotateRight()

         'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error GoTo RotateRight_ERROR
    
'    ' Flag image to be overwritten
'    subConfigRasterOperationsOverwrite
'
'   Dim RasterTools As IRasterTools
'   ' Set the object variable for the IRasterTools interface to the Edit Control object
'   Set RasterTools = Me.SpicerEdit1.object
'   ' Turn the raster image 180 degrees
'   RasterTools.DataRotate IN_ROTATION_90
'   ' De-initialize the object variable
'   Set RasterTools = Nothing
'
'
'   Dim docSave As IDocSave
'   ' Set the object variable for the IDocSave interface to the Document Control object
'   Set docSave = Me.SpicerDoc1.object
'   ' Export the first page in the document, retaining its current format
''   docSave.Export Me.SpicerDoc1.FirstPageID, False, 0, txtFullPathName, "JPEG"
'   docSave.Save 0, False, 0, txtFullPathName, "1"
'   ' De-initialize the object variable
'   Set docSave = Nothing
    
 
    subRotateImage IN_ROTATION_90_CW, True
    
Exit Sub
    
RotateRight_ERROR:
    MsgBox "Sorry... this document type cannot be saved as rotated!", vbInformation, "Rotate Right"
    
End Sub

Public Sub subRotateImage(ByVal PageRotation As String, Optional SaveRotation As Boolean)
   
         'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
   On Error GoTo Rotate_ERROR
   
   Dim ScaleScrollRotation As IScaleScrollRotation
   ' Set the object variable for the IScaleScrollRotation interface to the View Control object
   Set ScaleScrollRotation = Me.SpicerView1.object
   
   
   ' Set the rotation of the active page an Save ONLY if it Changed
   If ScaleScrollRotation.Rotation(0) <> CInt(PageRotation) Then
        '*** CORRECT SPICER BUG IN REPORTING ScaleScrollRotation ***
        '***  Incorrectly returns the Rotation Value by reversing the 90 and 270 rotations.
        Select Case CInt(PageRotation)
             Case IN_ROTATION_90
                 PageRotation = IN_ROTATION_270
             Case IN_ROTATION_270
                 PageRotation = IN_ROTATION_90
        End Select
        
        '*** Now Rotate the Image
        ScaleScrollRotation.Rotation(0) = CInt(PageRotation)
        
        If SaveRotation = True Then
             'Save ROTATION to the appropriate Page record
             Select Case txtModuleIndex
                 
                 Case gI101ModuleRetrieve
                 
                    If bolAIM_Command_AddFile = True Then
                        'Save Rotation to the Array but NOT the DB...  it will be saved when the file is "Saved"
                        arrPageRotation(txtPageNumber) = ScaleScrollRotation.Rotation(0)
                    Else
                        arrPageRotation(txtPageNumber) = ScaleScrollRotation.Rotation(0)
                        funcSaveFieldToDB RegImaging101ConnectionString, frmImaging101Search.txtApplicationName & "_Detail", "DetailRECID = " & arrDetailRECID(txtPageNumber), "DetailRotation", arrPageRotation(txtPageNumber)
                    End If
                    
                 Case gI101ModuleIndex
                    arrPageRotation(txtPageNumber) = ScaleScrollRotation.Rotation(0)
                    frmIndex.ListView1.ListItems.item(frmIndex.ListView1.SelectedItem.Index).ListSubItems(4).Text = ScaleScrollRotation.Rotation(0)
                    funcSaveFieldToDB RegImaging101ConnectionString, frmIndex.txtApplicationName & "_BatchPage", "BatchPageRECID = " & arrDetailRECID(txtPageNumber), "BatchPageRotation", arrPageRotation(txtPageNumber)
             
             End Select
        End If
   
   End If
   
   
   ' De-initialize the IScaleScrollRotation object variable
   Set ScaleScrollRotation = Nothing
   
Exit Sub
    
Rotate_ERROR:
    MsgBox "Sorry... this document type cannot be saved as rotated!", vbInformation, "subRotateImage"

End Sub



Public Sub subConfigRasterOperationsOverwrite()
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    Dim CFGDocument As ICFGDocument
    
    'Set the object variable for the ICFGDocument interface to the configuration control object
    Set CFGDocument = MainMDIForm.ActiveForm.SpicerConfiguration1.object
    
    'Set to automatically overwite document on rasterization.
    CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = True
''' ***  Commented because Shawn @ Spicer tech support said this could be causing the Rotation error
'''    'Set to remove overwritten raster layers from the source document window
'''    CFGDocument.RasterOperations(IN_RASTERIZE_REMOVE) = True
    
    'Deinitialize the object variable
    Set CFGDocument = Nothing
    
End Sub

Public Sub subSetCurrentPage()

      funcWriteToDebugLog Me.name, "*** ENTERING subSetCurrentPage() "

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    

    
      Dim docContents As IDocContents
      Dim ActivePage As IActivePage
      Dim iPageNum As Integer
      
      On Error GoTo SETCURRENTPAGE_ERROR
      
    '5/12/2014 - Jacob - Disabled the followng If block that was causing problems on Commit because "txtPageCount" was not being set
'      If Me.SpicerDoc1.NewestObjectID = 0 Then
'        Exit Sub
'      End If
            
      ' Set the object variable for the IDocContents and IActivePage interfaces
      Set docContents = Me.SpicerDoc1.object
      Set ActivePage = Me.SpicerView1.object
      
      
      ' Get the pageID for the current page of active document
      lPageID = ActivePage.ActivePageId
      ' Get the page number for the active page
      iPageNum = docContents.GetPageNumber(lPageID)
      ' Get the number of pages for the document
      lPageCount = docContents.NumberOfPages
    
      txtPageCount = lPageCount
      txtPageNumber = iPageNum
      txtDetailRECID = arrDetailRECID(txtPageNumber)
      txtPageFileName = arrPageFileName(txtPageNumber)
      
'    StatusBar1.Panels(1).Text = "Image " & txtImageNumber & " of " & txtImageCount

    '*** 2022-07-19 - Jacob - Added Arrays for Page Width and Page Height
    If txtPageCount > 1 Then
        StatusBar1.Panels(1).Text = "Multi-Page"
    Else
         StatusBar1.Panels(1).Text = "Single-Page"
    End If
    
    StatusBar1.Panels(2).Text = "Page " & CStr(iPageNum) & " of " & CStr(lPageCount)
    
    Dim intExtPos As Integer
    intExtPos = InStrRev(txtPageFileName, ".")
    If intExtPos > 0 Then
        StatusBar1.Panels(3).Text = Right(txtPageFileName, Len(txtPageFileName) - intExtPos) & "  (" & arrPageWidth(txtPageNumber) & " x " & arrPageHeight(txtPageNumber) & ")"  ' Display File Type
    Else
        StatusBar1.Panels(3).Text = "No Ext  (" & arrPageWidth(txtPageNumber) & " x " & arrPageHeight(txtPageNumber) & ")"
    End If

    If txtFTStatus <> "" Then
        StatusBar1.Panels(4).Text = "Full Text"
    Else
        StatusBar1.Panels(4).Text = ""
    End If

    StatusBar1.Panels(5).Text = txtDetailRECID    ' Display File Type Name


    '* Set Buttons for Multi-Image Documents
'    If CInt(txtImageCount) > 1 Then
'        MainMDIForm.cmdGotoImage.Visible = True
'    Else
'        MainMDIForm.cmdGotoImage.Visible = False
'    End If
'
'    If (CInt(txtImageCount) > 1) And (CInt(txtImageNumber) < CInt(txtImageCount)) Then
'        MainMDIForm.cmdNextImage.Visible = True
'    Else
'        MainMDIForm.cmdNextImage.Visible = False
'    End If
'
'    If CInt(txtImageNumber) > 1 Then
'        MainMDIForm.cmdPrevImage.Visible = True
'    Else
'        MainMDIForm.cmdPrevImage.Visible = False
'    End If
    
    
    
    '* Set Buttons for Multi-Page Documents
    If CInt(lPageCount) > 1 Then
        MainMDIForm.cmdGotoPage.Enabled = True
    Else
        MainMDIForm.cmdGotoPage.Enabled = False
    End If
    
    If (CInt(lPageCount) > 1) And (CInt(iPageNum) < CInt(lPageCount)) Then
        MainMDIForm.cmdNextPage.Enabled = True
    Else
        MainMDIForm.cmdNextPage.Enabled = False
    End If
    
    If CInt(iPageNum) > 1 Then
        MainMDIForm.cmdPrevPage.Enabled = True
    Else
        MainMDIForm.cmdPrevPage.Enabled = False
    End If
    
    '*** Rotate Page as Needed
'    ' First Reset to the ORIGINAL rotation
'    subRotateImage 0
    ' Now Rotate to the saved rotation if needed
    
    subDebugChildForm "subSetCurrentPage():  Before  subRotateImage arrPageRotation(iPageNum)"
    
    If MainMDIForm.mnuLockRotation.Checked = True Then
        subRotateImage intLockedRotation, True
    Else
        subRotateImage arrPageRotation(iPageNum)
    End If
    
    
'    BindControls
        
    '*** Check If Scale To Gray is enabled
    funcWriteToDebugLog Me.name, "cmdScaleToGray"
    cmdScaleToGray
        
    '*** LOAD ANNOTATION LAYER
    funcWriteToDebugLog Me.name, "subSetCurrentPage():  BEFORE subAnnotationLayerLoad"
    
    subAnnotationLayerLoad

    subDebugChildForm "subSetCurrentPage():  AFTER subAnnotationLayerLoad"

'    End If
    
    
    
      ' De-initialize the object variables
      Set docContents = Nothing
      
Exit Sub

SETCURRENTPAGE_ERROR:
    Dim sErrMessage As String
    sErrMessage = Me.SpicerDoc1.ErrorMessage(Err.Number)
    funcWriteToDebugLog Me.name, "SetCurrentPage ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage

    MsgBox "SetCurrentPage ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    Set docContents = Nothing
    Set ActivePage = Nothing
 
End Sub


Public Sub subGetNextImage()

    funcWriteToDebugLog Me.name, "*** ENTERING subGetNextImage()"

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    Dim txtPathSubdirectory As String
    Dim txtFileName As String
    Dim txtFullPathName As String

   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                        " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                        " DetailFileName, DetailFileType, DetailRotation  " & _
                        " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                        " WHERE  DocumentRECID = " & txtDocumentRECID & _
                        " AND    DetailOrder > " & txtImageNumber & _
                        " ORDER BY DetailOrder "

    txtActionBeforeError = "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    txtActionBeforeError = "ADodcDetail MoveFirst "
    frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
   
    txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
    txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
    txtFullPathName = txtPathSubdirectory & "\" & txtFileName

    txtImageNumber = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailOrder")
    txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")
    txtPageRotation = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRotation")
    
    SpicerDoc1.CloseDocument False
    SpicerDoc1.OpenFile txtFullPathName
    
    subSetCurrentPage
    
End Sub


Public Sub subGetPrevImage()

    funcWriteToDebugLog Me.name, "*** subGetPrevImage()"
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    Dim txtPathSubdirectory As String
    Dim txtFileName As String
    Dim txtFullPathName As String

   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                        " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                        " DetailFileName, DetailFileType, DetailRotation  " & _
                        " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                        " WHERE  DocumentRECID = " & txtDocumentRECID & _
                        " AND    DetailOrder = " & CInt(txtImageNumber) - 1 & _
                        " ORDER BY DetailOrder "

    txtActionBeforeError = "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    txtActionBeforeError = "ADodcDetail MoveFirst "
    frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
   
    txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
    txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
    txtFullPathName = txtPathSubdirectory & "\" & txtFileName

    txtImageNumber = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailOrder")
    txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")
    txtPageRotation = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRotation")
    
    SpicerDoc1.CloseDocument False
    SpicerDoc1.OpenFile txtFullPathName
    
    subSetCurrentPage
    
End Sub


Public Sub subGotoImage()
    
    funcWriteToDebugLog Me.name, "*** subGetPrevImage()"
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error GoTo ERROR_HANDLER
    
    Dim txtPathSubdirectory As String
    Dim txtFileName As String
    Dim txtFullPathName As String

    Dim dblPageNumber As Double
    
    dblPageNumber = InputBox("Enter Image # to Go to ( 1 - " & txtImageCount & ") ", " Goto Image", 1)
'    If dblPageNumber = vbCancel Then
'        Exit Sub
'    End If
    
   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   frmImaging101Retrieve.ADOdcDetail.RecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                        " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                        " DetailFileName, DetailFileType, DetailRotation  " & _
                        " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                        " WHERE  DocumentRECID = " & txtDocumentRECID & _
                        " AND    DetailOrder = " & dblPageNumber & _
                        " ORDER BY DetailOrder "

    txtActionBeforeError = "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    If frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount > 0 Then
        txtActionBeforeError = "ADodcDetail MoveFirst "
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
    Else
        txtActionBeforeError = "Page NOT Found "
        MsgBox "Image NOT Found!", vbCritical, "Image NOT Found!"
        Exit Sub
    End If
    
    txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
    txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
    txtFullPathName = txtPathSubdirectory & "\" & txtFileName

    txtImageNumber = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailOrder")
    txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")
'    txtPageRotation = frmImaging101Retrieve.AdodcDetail.Recordset.Fields("DetailRotation")
    
    SpicerDoc1.CloseDocument False
    SpicerDoc1.OpenFile txtFullPathName
    
    subSetCurrentPage
    
Exit Sub
    
    
ERROR_HANDLER:

    If Err.Number = 13 Then
        Exit Sub
    End If
    
    MsgBox "Goto Page ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    
End Sub


Public Sub subGetRelatedImages()

        funcWriteToDebugLog Me.name, "*** ENTERING subGetRelatedImages()"
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    'Prepare Progress Bar
    ProgressBarLoading.Value = 0
    ProgressBarLoading.Visible = True

    'On Error Resume Next
    On Error GoTo ERROR_HANDLER
    
    '*** DISABLE ALL THE VISIBLE FORMS to try to prevent the Crash on roll of the Mouse Roller
    frmImaging101Search.Enabled = False
    frmImaging101Retrieve.Enabled = False
    MainMDIForm.Enabled = False
    Me.Enabled = False
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass

    Dim txtPathSubdirectory As String
    Dim txtFileName As String
    Dim txtFullPathName As String
    '*** These fields are to update the Arrays to reflect the ACTUAL number of Pages Loaded
    '     just in case the FIRST document loaded was a MULTI-PAGE file!
    Dim dblNumberOfPagesBeforeImport As Double
    Dim dblNumberOfPagesAfterImport As Double
    

  
    '**************************************************
    '*** Initialize for the First Loaded document
    dblNumberOfPagesBeforeImport = 0
    dblNumberOfPagesAfterImport = 0
    
'''    '*** Update the Arrays to reflect the ACTUAL number of Pages Loaded
'''    '     just in case the FIRST document loaded was a MULTI-PAGE file!
'''    funcWriteToDebugLog Me.name, "subUpdatePageArrays dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport & _
'''                                    ", dblNumberOfPagesAfterImport = " & dblNumberOfPagesAfterImport & ")"
'''    subUpdatePageArrays dblNumberOfPagesBeforeImport, dblNumberOfPagesAfterImport
    
    'SKIP the LOAD if loaded document is a Full-Text document
    If txtFTStatus <> "" Then
        funcWriteToDebugLog Me.name, "Full Text Document -> SKIP_LOAD"
        Screen.MousePointer = MousePointerConstants.vbDefault
        GoTo SKIP_LOAD
    End If
    
   '**********************************************************
   '*** FIND Records for Additional detail objects/files
   
    funcWriteToDebugLog Me.name, "FIND Records for Additional detail objects/files"
    
   frmImaging101Retrieve.ADOdcDetail.ConnectionString = RegImaging101ConnectionString
   
    '*** 2022-07-07 - Jacob - ADDED Coded to Handle AMA Insurance "STAPLE" pages
    Dim strRecordSource As String
    Dim strStapleID As String
    

    Select Case gsecSiteInformationClientLong
    
        Case "AMA Insurance Agency, Inc."
    
                        strStapleID = funcGetFieldFromDB(RegImaging101ConnectionString, frmImaging101Search.txtApplicationName, "DocumentRECID = " & txtDocumentRECID, "StapleID") & ""
        
                        'Select from Master and Detail to include the STAPLE pages
                        'Order by Master in Descending order to add to end, because STAPLE pages do NOT have a Master #
                        strRecordSource = "SELECT DetailRECID , " & frmImaging101Search.txtApplicationName & "_Detail.DocumentRECID, " & _
                                                                " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                                                                " DetailFileName, DetailFileType, DetailRotation  " & _
                                                                " FROM " & frmImaging101Search.txtApplicationName & ", " & frmImaging101Search.txtApplicationName & "_Detail  " & _
                                                                " WHERE  (" & frmImaging101Search.txtApplicationName & ".DocumentRECID = " & frmImaging101Search.txtApplicationName & "_Detail.DocumentRECID  )" & _
                                                                "                   AND  (" & frmImaging101Search.txtApplicationName & ".DocumentRECID = " & txtDocumentRECID
                                                                
                        If Trim(strStapleID) <> "" Then
                                         strRecordSource = strRecordSource & " OR " & frmImaging101Search.txtApplicationName & ".DocID = '" & strStapleID & "' "
                        End If
                        
                        'Add Close Parenthesis for the AND
                         strRecordSource = strRecordSource & ") "
                        'Add Order By
                        strRecordSource = strRecordSource & " ORDER BY Master DESC, DetailOrder ASC "

        Case Else
        
                        strRecordSource = "SELECT DetailRECID , DocumentRECID, " & _
                                                                " DetailOrder, DetailCreatedDate, DetailSubdirectory, " & _
                                                                " DetailFileName, DetailFileType, DetailRotation  " & _
                                                                " FROM " & frmImaging101Search.txtApplicationName & "_Detail " & _
                                                                " WHERE  DocumentRECID = " & txtDocumentRECID & _
                                                                " ORDER BY DetailOrder "
                                                                
                '*** 2022-03-29 - Jacob Removed " AND    DetailOrder > " & txtImageNumber to include the First Detail Record
    
    End Select
    
   funcWriteToDebugLog Me.name, "strRecordSource = " & strRecordSource
    frmImaging101Retrieve.ADOdcDetail.RecordSource = strRecordSource

    '*** 2022-07-07 - Jacob - END Code to Handle AMA Insurance "STAPLE" pages


    txtActionBeforeError = "ADodcDetail Refresh "
    frmImaging101Retrieve.ADOdcDetail.Refresh
    
    If frmImaging101Retrieve.ADOdcDetail.Recordset.EOF = True Then
        funcWriteToDebugLog Me.name, "frmImaging101Retrieve.ADOdcDetail.Recordset.EOF = True -> SKIP_LOAD"
        Screen.MousePointer = MousePointerConstants.vbDefault
        GoTo SKIP_LOAD
    Else
        txtActionBeforeError = "ADodcDetail MoveFirst   -  RecordCount = " & frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount
        funcWriteToDebugLog Me.name, "frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst"
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveFirst
    End If
   

    
    
    While Not frmImaging101Retrieve.ADOdcDetail.Recordset.EOF
    
        txtPathSubdirectory = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailSubdirectory")
        txtFileName = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailFileName")
        txtFullPathName = txtPathSubdirectory & "\" & txtFileName
        txtPageFileName = txtFullPathName
        
'        txtImageNumber = frmImaging101Retrieve.AdodcDetail.Recordset.Fields("DetailOrder")
        txtDetailRECID = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRECID")
        
        If Not frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRotation") Then
            txtPageRotation = frmImaging101Retrieve.ADOdcDetail.Recordset.Fields("DetailRotation")
        End If
        
        If lstPageList.ListCount > 0 Then
                 dblNumberOfPagesBeforeImport = SpicerDoc1.NumberOfPages
        End If
        

        '**********************************************************************************************************
        '*** RE-MAP the File Path to the Local FilePath
        
                txtFullPathName = funcCopyFileToLocalTemp(txtFileName, txtFullPathName, txtDetailRECID)
                txtPageFileName = txtFullPathName
                
                '*** 2022-07-21 - Jacob - Get File Extension to Check for EML & MSG below
                Dim intPositionOfPeriod As Integer
                Dim strFileExtension As String
                intPositionOfPeriod = InStrRev(txtFullPathName, ".")
                strFileExtension = UCase(Right(txtFullPathName, Len(txtFullPathName) - intPositionOfPeriod))


        '**********************************************************************************************************
        '*** Check if File should be Launched

        bolObjectLaunched = funcCheckIfFileShouldBeLaunched(txtFullPathName)
        
        
                 
       '***********************************************************************************************************
       '*** Load or Import ONLY if NOT flagged to AutoLaunch and No file open Errors detected
       
        If bolLaunchThisFileType = True Or Err.Number <> 0 Then
        
            '*** DON'T LOAD OR IMPORT
            
        ElseIf dblNumberOfPagesAfterImport = 0 Then
        
                '****************************************************************************************************
                '***     This is the FIRST Page - OPEN the FILE for the FIRST Detail Record
                funcWriteToDebugLog Me.name, "*** OPEN the FILE"
                 funcWriteToDebugLog Me.name, "txtPathSubdirectory = " & txtPathSubdirectory
                 funcWriteToDebugLog Me.name, "txtFileName = " & txtFileName
                 funcWriteToDebugLog Me.name, "txtFullPathName = " & txtFullPathName
                 funcWriteToDebugLog Me.name, "txtPageFileName = " & txtPageFileName
                 funcWriteToDebugLog Me.name, "txtDetailRECID = " & txtDetailRECID
                 funcWriteToDebugLog Me.name, "txtPageRotation = " & txtPageRotation
                 funcWriteToDebugLog Me.name, "dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport


                Set docContents = Me.SpicerDoc1.object
                
                '*** 2022-07-21 - Jacob - Added Check for EML and MSG
                If strFileExtension = "EML" Or strFileExtension = "MSG" Then
                        funcWriteToDebugLog Me.name, "subGetRelatedImages() | docContents.OpenFile " & txtFullPathName & ".pdf"
                        docContents.OpenFile txtFullPathName & ".pdf"
                Else
                        funcWriteToDebugLog Me.name, "subGetRelatedImages() | docContents.OpenFile " & txtFullPathName
                        docContents.OpenFile txtFullPathName
                End If
                
                Set docContents = Nothing



                funcWriteToDebugLog Me.name, "*** AFTER docContents.OpenFile | Err.Number = " & Err.Number

        Else
        
                '****************************************************************************************************
                '***     NOT the FIRST Page - IMPORT the FILE into the CURRENT SpicerDoc1
                 
                 
                 funcWriteToDebugLog Me.name, "*** IMPORT the FILE into the CURRENT SpicerDoc1"
                 funcWriteToDebugLog Me.name, "txtPathSubdirectory = " & txtPathSubdirectory
                 funcWriteToDebugLog Me.name, "txtFileName = " & txtFileName
                 funcWriteToDebugLog Me.name, "txtFullPathName = " & txtFullPathName
                 funcWriteToDebugLog Me.name, "txtPageFileName = " & txtPageFileName
                 funcWriteToDebugLog Me.name, "txtDetailRECID = " & txtDetailRECID
                 funcWriteToDebugLog Me.name, "txtPageRotation = " & txtPageRotation
                 funcWriteToDebugLog Me.name, "dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport
                
                 '*** IMPORT the File Contents
                 
                Set docContents = Me.SpicerDoc1.object
        
                '*** 2022-07-21 - Jacob - Added Check for EML and MSG
                If strFileExtension = "EML" Or strFileExtension = "MSG" Then
                        funcWriteToDebugLog Me.name, "subGetRelatedImages() | ImportPage 0, 0, IN_NEWPAGE_END, " & txtFullPathName & ", " & txtFullPathName & ".pdf"
                         docContents.ImportPage 0, 0, IN_NEWPAGE_END, txtFullPathName, txtFullPathName
                Else
                         funcWriteToDebugLog Me.name, "subGetRelatedImages() | ImportPage 0, 0, IN_NEWPAGE_END, " & txtFullPathName & ", " & txtFullPathName
                         docContents.ImportPage 0, 0, IN_NEWPAGE_END, txtFullPathName, txtFullPathName
                End If
                
'                         Debug.Print docContents.NumberOfLayers(docContents.NewestObjectID)
                         
                         Set docContents = Nothing
                         
                funcWriteToDebugLog Me.name, "*** AFTER ImportPage | Err.Number = " & Err.Number
        
        End If
        
        '*** Calculate Progress Bar Percentage
        ProgressBarLoading.Value = (frmImaging101Retrieve.ADOdcDetail.Recordset.AbsolutePosition / frmImaging101Retrieve.ADOdcDetail.Recordset.RecordCount) * 100
        DoEvents
        
        
        '***************************************************************************************************************
        '*** Check for Import Errors
        
        Call funcCheckForOpenOrImportErrors(Err.Number, Err.Description, txtFullPathName)
        
        
        '***************************************************************************************************************
        '*** BIND THE SPICER CONTROLS
        Call BindControls

        
        dblNumberOfPagesAfterImport = SpicerDoc1.NumberOfPages
                 
        funcWriteToDebugLog Me.name, "dblNumberOfPagesAfterImport = " & dblNumberOfPagesAfterImport
                 
        '*** Update the Arrays to reflect the ACTUAL number of Pages Loaded
        '     just in case the FIRST document loaded was a MULTI-PAGE file!
        funcWriteToDebugLog Me.name, "subUpdatePageArrays dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport & ", dblNumberOfPagesAfterImport = " & dblNumberOfPagesAfterImport
        subUpdatePageArrays dblNumberOfPagesBeforeImport, dblNumberOfPagesAfterImport
        
        Me.txtPageCount = dblNumberOfPagesAfterImport
  
        '*** GET NEXT RECORD
        funcWriteToDebugLog Me.name, "frmImaging101Retrieve.ADOdcDetail.Recordset.MoveNext"
        frmImaging101Retrieve.ADOdcDetail.Recordset.MoveNext
        
        
    Wend
    
    
SKIP_LOAD:

'    funcWriteToDebugLog Me.name, "subGetRelatedImages: Call subSetCurrentPage"
''XXX
'    subSetCurrentPage
    
    bolRelatedImagesLoaded = True
    
      
    '***************************************************************
      '*** Re-enable Seach & rForm
    frmImaging101Search.Enabled = True
    frmImaging101Retrieve.Enabled = True
    MainMDIForm.Enabled = True
    Me.Enabled = True
    
    '***************************************************************
      '*** Make Viewer objects Visible again

      Me.SpicerView1.Visible = True
      Me.lstPageList.Visible = True
      Me.lblLoadingPages.Visible = False
    
      'Me.SpicerView1.SetFocus
      
      
'        '*** 2023-02-21 - Jacob - Find "Attachments:"
'       Dim ActivePage As IActivePage
'       Dim sText As String
'       ' Set the object variable for the IActivePage interface to the View Control object
'       Set ActivePage = Me.SpicerView1.object
'
'       ' Text to find
'       sText = "Attachments:"
'       ' Find the text, matching the case and searching for the next match
'       ActivePage.FindTextMatch sText, IN_TXTSRCH_CASEINSENSITIVE, IN_DIR_NEXT
'
'       ' De-initialize the object variable
'       Set ActivePage = Nothing
       
      
    '*** Hide Progress Bar
    ProgressBarLoading.Visible = False

    
    Screen.MousePointer = MousePointerConstants.vbDefault

Exit Sub

ERROR_HANDLER:

    Dim sErrMessage As String
    sErrMessage = Me.SpicerDoc1.ErrorMessage(Err.Number)
    funcWriteToDebugLog Me.name, "subGetRelatedImages() ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    
        
      
    '***************************************************************
      '*** Re-enable Seach & rForm
    frmImaging101Search.Enabled = True
    frmImaging101Retrieve.Enabled = True
    MainMDIForm.Enabled = True
    Me.Enabled = True
    
    '***************************************************************
      '*** Make Viewer objects Visible again

      Me.SpicerView1.Visible = True
      Me.lstPageList.Visible = True
      Me.lblLoadingPages.Visible = False

    
        '*** Hide Progress Bar
    ProgressBarLoading.Visible = False

    Screen.MousePointer = MousePointerConstants.vbDefault

    'MsgBox "subGetRelatedImages() ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage

    'Resume
    
End Sub
Public Sub subUpdatePageArrays(dblNumberOfPagesBeforeImport As Double, dblNumberOfPagesAfterImport As Double)


        funcWriteToDebugLog Me.name, "*** ENTERING subUpdatePageArrays(dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport & _
        ", dblNumberOfPagesAfterImport = " & dblNumberOfPagesAfterImport & ")"
        
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    Dim dblArrayPageLoop As Double
        
        '*** 2022-07-18 - Jacob - Added IF for before and after being the same.
        If dblNumberOfPagesAfterImport > dblNumberOfPagesBeforeImport Then
        
            dblNumberOfPagesForImportedDocument = dblNumberOfPagesAfterImport - dblNumberOfPagesBeforeImport
            funcWriteToDebugLog Me.name, "dblNumberOfPagesForImportedDocument = " & dblNumberOfPagesForImportedDocument
       
       Else
       
            dblNumberOfPagesForImportedDocument = 0
            funcWriteToDebugLog Me.name, "dblNumberOfPagesForImportedDocument = " & dblNumberOfPagesForImportedDocument
            
        End If
        
       
        '*** HANDLE PAGE NUMBERING DIFFERENTLY FOR BATCHES THAN FOR RETRIEVAL
        If txtCurrentModule = "frmImaging101Retrieve" Then
            '*** 2022-07-18 - Jacob - Added IF for before and after being the same.
            If dblNumberOfPagesAfterImport > dblNumberOfPagesBeforeImport Then
                dblStartingPageNumber = dblNumberOfPagesBeforeImport + 1
                funcWriteToDebugLog Me.name, "dblStartingPageNumber = " & dblStartingPageNumber
            Else
                dblStartingPageNumber = dblNumberOfPagesBeforeImport
                funcWriteToDebugLog Me.name, "dblStartingPageNumber = " & dblStartingPageNumber
            End If
            
        Else
            dblStartingPageNumber = dblNumberOfPagesBeforeImport
            funcWriteToDebugLog Me.name, "dblStartingPageNumber = " & dblStartingPageNumber
        End If
        
        subDebugChildForm "subUpdatePageArrays():  Before ReDim Preserve arrPageRotation(0 To UBound(arrPageRotation) + 1)"
        ReDim Preserve arrPageRotation(0 To dblNumberOfPagesAfterImport)
        ReDim Preserve arrDetailRECID(0 To dblNumberOfPagesAfterImport)
        ReDim Preserve arrPageLayerID(0 To dblNumberOfPagesAfterImport)
        ReDim Preserve arrPageFileName(0 To dblNumberOfPagesAfterImport)
        
        '*** 2022-07-19 - Jacob - Added Arrays for Page Width and Page Height
        ReDim Preserve arrPageWidth(0 To dblNumberOfPagesAfterImport)
        ReDim Preserve arrPageHeight(0 To dblNumberOfPagesAfterImport)
        
        subDebugChildForm "subUpdatePageArrays():  After  ReDim Preserve arrPageRotation(0 To UBound(arrPageRotation) + 1)"

'        For dblArrayPageLoop = dblNumberOfPagesBeforeImport + 1 To dblNumberOfPagesAfterImport
     
        funcWriteToDebugLog Me.name, "For dblArrayPageLoop = " & dblStartingPageNumber & " To " & dblNumberOfPagesAfterImport

        For dblArrayPageLoop = dblStartingPageNumber To dblNumberOfPagesAfterImport '- 1
    
            'Add page to the Page ListView
            funcWriteToDebugLog Me.name, "lstPageList.AddItem " & Format(dblArrayPageLoop, "0000")
            lstPageList.AddItem Format(dblArrayPageLoop, "0000")
            DoEvents
            
            arrPageRotation(dblArrayPageLoop) = txtPageRotation
            arrDetailRECID(dblArrayPageLoop) = txtDetailRECID
            arrPageFileName(dblArrayPageLoop) = txtPageFileName
            
             '*** 2022-07-19 - Jacob - Added Arrays for Page Width and Page Height
            
            arrPageWidth(dblArrayPageLoop) = txtDocWidth
            arrPageHeight(dblArrayPageLoop) = txtDocHeight

        Next

End Sub

Public Sub subAnnotationLayerSave()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
   Dim docSave As IDocSave
   Dim strFileName As String
   Dim strFullDirectoryPathForAnnotation As String
   Dim strDestinationFileName As String
   
    On Error GoTo ERROR_HANDLER
    
   ' Set the object variable for the IDocSave interface to the Document Control object
   Set docSave = MainMDIForm.ActiveForm.SpicerDoc1.object
   
                    
    '*** Prepare filename depending on if the page being viewed is a BATCH, ADDFILE or RETRIEVAL page.
    If InStr(1, Me.Caption, "BATCH:") Then
        'Batch Page
        '  the txtDetailRECID is actually the BatchPageRECID
        'Get the current Filename without the Extension
        
        '*** 6/25/2009 - Jacob - Modified to handle Files with NO Extension
        '*** 11/9/2009 - Jacob - Corrected to prefix with "frmIndex." to prevent naming ANN files
        '                        including the extension.
        intPositionOfLastPeriod = InStrRev(frmIndex.txtBatchPageFileName, ".")
        
        If intPositionOfLastPeriod = 0 Then
            strFileName = frmIndex.txtBatchPageFileName
        Else
            strFileName = Left(frmIndex.txtBatchPageFileName, InStrRev(frmIndex.txtBatchPageFileName, ".") - 1)
        End If
        
        strDestinationFileName = frmIndex.txtBatchDirectory & "\" & _
                                    strFileName & _
                                    "_" & _
                                    Format(CStr(Me.txtPageNumber), "000000") & _
                                    ".ANN"
                                    
                                    
        funcWriteToDebugLog Me.name, "subAnnotationLayerSave() ANNOTATION BATCH FileName = " & strDestinationFileName
                                    
    ElseIf bolAIM_Command_AddFile Then
    
        '*** Filer Add File
        
        ' Build the Annotation FilePath
        Dim strLocalTempDir As String
        strLocalTempDir = funcGetTempDir()

        intPositionOfLastPeriod = InStrRev(strSourceFile, ".")

        strFullDirectoryPathForAnnotation = strLocalTempDir & "Imaging101\Annotations"
        txtActionBeforeError = "Create Directory Structure: " & strFullDirectoryPathForAnnotation
         
        'Create the directory if needed.
        funcCreateDirectoryStructure strFullDirectoryPathForAnnotation & ""
        
        strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & _
                                "Annotation" & _
                                "_" & _
                                Format(CStr(Me.txtPageNumber), "000000") & _
                                ".ANN"
                                
        funcWriteToDebugLog Me.name, "subAnnotationLayerSave() ANNOTATION AddFile FileName = " & strDestinationFileName
       
    
    Else
    
        '*** Retrieval Page
        
        ' Build the Annotation FilePath
        strFullDirectoryPathForAnnotation = funcGetFullPathForAnnotation(txtApplicationRECID, txtDetailRECID)
        txtActionBeforeError = "Create Directory Structure: " & strFullDirectoryPathForAnnotation
         
        'Create the directory if needed.
        funcCreateDirectoryStructure strFullDirectoryPathForAnnotation & ""
        strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & _
                                Format(CStr(Me.txtDetailRECID), "0000000000") & _
                                "_" & _
                                Format(CStr(Me.txtPageNumber), "000000") & _
                                ".ANN"
                                
        funcWriteToDebugLog Me.name, "subAnnotationLayerSave() ANNOTATION RETRIEVE FileName = " & strDestinationFileName
        
    End If
   
   ' Check if Annotation File Exists
   If Not funcFileExists(strDestinationFileName) Then
   
            ' Export the current LAYER, retaining its current format
             '  will NOT check if annotation file exists, simply overwrite it.
            docSave.Export arrPageLayerID(txtPageNumber), False, 0, strDestinationFileName, "I101 Annotation " & MainMDIForm.ActiveForm.txtAnnotationLayerID
    
   Else
   
        ' RENAME the Annotation File
        Dim strDestinationFileName_TEMP
        Dim fso As New FileSystemObject
        Set fso = New Scripting.FileSystemObject
        strDestinationFileName_TEMP = strDestinationFileName & "_TMP"
        
        On Error Resume Next
        
        Err.Clear
        fso.MoveFile strDestinationFileName, strDestinationFileName_TEMP
        lngErrNum = Err.Number      ' Save the error number that occurred.
        strErrDescr = Err.Description
        If lngErrNum = 0 Then
            ' No error occurred - File was Renamed properly
            
            ' Export the current LAYER, retaining its current format
             '  will NOT check if annotation file exists, simply overwrite it.
            docSave.Export arrPageLayerID(txtPageNumber), False, 0, strDestinationFileName, "I101 Annotation " & MainMDIForm.ActiveForm.txtAnnotationLayerID
    
            Kill strDestinationFileName_TEMP
        
        Else
            On Error GoTo ERROR_HANDLER
            
            'Raise an Error
            Err.Raise lngErrNum, "subAnnotationLayerSave", strErrDescr & " - Unable to Rename the old Annotation file - " & strDestinationFileName
                
        End If

    End If  ' funcFileExists(strDestinationFileName)
    
   ' De-initialize the object variable
   Set docSave = Nothing
   
Exit Sub

ERROR_HANDLER:

   ' De-initialize the object variable
   Set docSave = Nothing
   
    funcQuickMessage "SHOW", "subAnnotationLayerSave ERROR: " & Err.Number & " - " & Err.Description & " SOURCE: " & Err.Source
    funcWriteToDebugLog Me.name, "subAnnotationLayerSave ERROR: " & Err.Number & " - " & Err.Description & " SOURCE: " & Err.Source
    
End Sub
Public Sub subAnnotationLayerLoad()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    
    On Error GoTo ERROR_HANDLER

    If SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId) > 1 Then
        Exit Sub
    End If
    
    Dim lPageID As Long
    Dim docContents As IDocContents
    Dim strFileName As String
    Dim strFullDirectoryPathForAnnotation As String
    Dim strDestinationFileName As String
    Dim intLocationOfPeriod As Integer
    
    
    ' Build the Annotation FilePath
'     strFullDirectoryPathForAnnotation = funcGetFullPathForAnnotation(txtApplicationRECID, txtDetailRECID)
'     txtActionBeforeError = "Create Directory Structure: " & strFullDirectoryPathForAnnotation
     
     'Create the directory if needed.
''     funcCreateDirectoryStructure strFullDirectoryPathForAnnotation & ""
                     
     '*** Prepare filename
'     strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & Format(CStr(txtDetailRECID), "0000000000") & ".ANN"
    
    '*** Prepare filename depending on if the page being viewed is a BATCH or RETRIEVAL page.
    If InStr(1, Me.Caption, "BATCH:") Then
            'Batch Page
            '  the txtDetailRECID is actually the BatchPageRECID
            'Get the current Filename without the Extension
            
            
            intLocationOfPeriod = InStrRev(frmIndex.txtBatchPageFileName, ".")
            If intLocationOfPeriod > 0 Then
                strFileName = Left(frmIndex.txtBatchPageFileName, InStrRev(frmIndex.txtBatchPageFileName, ".") - 1)
            Else
                strFileName = frmIndex.txtBatchPageFileName
            End If
            
            strDestinationFileName = frmIndex.txtBatchDirectory & "\" & _
                                        strFileName & _
                                        "_" & _
                                        Format(CStr(Me.txtPageNumber), "000000") & _
                                        ".ANN"
            
            funcWriteToDebugLog Me.name, "subAnnotationLayerLoad() ANNOTATION BATCH FileName = " & strDestinationFileName
            
    ElseIf bolAIM_Command_AddFile Then
    
        Dim strLocalTempDir As String
        strLocalTempDir = funcGetTempDir()
        
        strFullDirectoryPathForAnnotation = strLocalTempDir & "Imaging101\Annotations"

        funcCreateDirectoryStructure strFullDirectoryPathForAnnotation
                
        strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & _
                                            "Annotation" & _
                                            "_" & _
                                            Format(CStr(Me.txtPageNumber), "000000") & _
                                            ".ANN"
    
    Else
    
            'Retrieval Page
            
            ' Build the Annotation FilePath
            strFullDirectoryPathForAnnotation = funcGetFullPathForAnnotation(txtApplicationRECID, txtDetailRECID)
            funcWriteToDebugLog Me.name, "subAnnotationLayerLoad() | Else 'Retrieval Page | Set strFullDirectoryPathForAnnotation: " & strFullDirectoryPathForAnnotation
            
            '***************************************************************
            'See if there are Multiple .ANN files for this DetailRECID
            ' If there are more than one... it was a Multi-Page document
            '  otherwise it was a single page.
            
            Dim intAnnotFileCount As Integer
            Dim strAnnotFileCountMask As String
            
            strAnnotFileCountMask = Format(CStr(Me.txtDetailRECID), "0000000000")
            
             
            If funcDirectoryExists(strFullDirectoryPathForAnnotation) Then
            
'                flbAnnotFiles.Path = strFullDirectoryPathForAnnotation
'                flbAnnotFiles.Pattern = Format(CStr(Me.txtDetailRECID), "0000000000") & _
'                                         "_*.ANN"
'
'                intAnnotFileCount = flbAnnotFiles.ListCount
'
'                funcWriteToDebugLog Me.name, "ANNOTATION intAnnotFileCount = " & intAnnotFileCount
                
'                 If intAnnotFileCount > 1 Then
'                    'More than ONE (1) ANN file... Load based on the Page #
                    strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & _
                                            Format(CStr(Me.txtDetailRECID), "0000000000") & _
                                            "_" & _
                                            Format(CStr(Me.txtPageNumber), "000000") & _
                                            ".ANN"
                                                            
                     '*** Now Check if the Annotation File Exists... if NOT then look for one with Page = 000001
                     '     because Batch Commit would save multiple pages of the same document with page 1
                     '     since, in theory, there was only ONE annotation per DocumentPageRECID.
                     '     But make sure it only loads if there is ONLY ONE document file for THIS RECID.
                     
                    If Not funcFileExists(strDestinationFileName) Then
                        
                        'Get the number of DOCUMENT Files for THIS txtDetailRECID...
                        'use the flbAnnotFiles list box since we already have it available
                        '*** 2021-02-22 - Jacob - Moved ".Pattern" ABOVE ".Path" to try to improve performance.
                        '                                              Though, technically, THIS should NEVER happen
                        flbAnnotFiles.Pattern = Format(CStr(Me.txtDetailRECID), "0000000000") & ".*"
                        flbAnnotFiles.Path = txtFileDirectory
                                               
                        If flbAnnotFiles.ListCount > 1 And txtPageCount > 1 Then
                            strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & _
                                                    Format(CStr(Me.txtDetailRECID), "0000000000") & _
                                                    "_" & _
                                                    Format(1, "000000") & _
                                                    ".ANN"
                        End If
                    End If
                    
'                ElseIf intAnnotFileCount = 1 Then
'                    'Single ANN file... Load only Annotation #1
''                    strDestinationFileName = strFullDirectoryPathForAnnotation & "\" & _
'                                             Format(CStr(Me.txtDetailRECID), "0000000000") & _
'                                             "_" & _
'                                             "000001" & _
'                                             ".ANN"
'
'                    flbAnnotFiles.ListIndex = 0
'                    strDestinationFileName = flbAnnotFiles.Path & "\" & flbAnnotFiles.FileName
'
'                Else
'
'                    Exit Sub
'
'                End If
                
            funcWriteToDebugLog Me.name, "subAnnotationLayerLoad() ANNOTATION RETRIEVE FileName = " & strDestinationFileName
    
                 

                
            End If   ' funcDirectoryExists(strFullDirectoryPathForAnnotation)
        
    End If  ' InStr(1, Me.Caption, "BATCH:")
     

     '*** Now Check if the Annotation File Exists... if NOT then DON'T try to Load/Import the Layer.
    If funcFileExists(strDestinationFileName) Then
         
        funcWriteToDebugLog Me.name, "subAnnotationLayerLoad() FILE EXISTS - Import Layer = " & strDestinationFileName
         
         ' Get the page ID for current page of active document
         lPageID = SpicerView1.ActivePageId
         ' Set the object variable for the IDocContents interface to the Document Control object
         Set docContents = SpicerDoc1.object
         ' Import the selected file
         docContents.ImportLayer lPageID, 0, strDestinationFileName
         
         'Save the Layer ID
         arrPageLayerID(txtPageNumber) = SpicerDoc1.NewestObjectID
         
         SpicerMarkup1.ActiveLayer = SpicerDoc1.NewestObjectID
    
         ' De-initialize object variable
         Set docContents = Nothing
         
        subAnnotationLayerShowHide
        
    Else
     
'         subAnnotationLayerCreate
     
    End If   ' funcFileExists(strDestinationFileName)



Exit Sub

ERROR_HANDLER:

    funcWriteToDebugLog Me.name, "subAnnotationLayerLoad ERROR: " & Err.Number & " - " & Err.Description & " SOURCE: " & Err.Source
    
   
End Sub


Public Sub subAnnotationLayerCreate()
    
    funcWriteToDebugLog Me.name, "ENTER: subAnnotationLayerCreate()"

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    If SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId) > 1 Then
        funcWriteToDebugLog Me.name, "Number of Layers = " & SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId) & " -- EXIT SUB"
        Exit Sub
    End If
    
    'Create a New Layer
    Dim docContents As IDocContents
    Dim lPageID As Long
    Dim lNewLayerID As Long
    Dim iNumLayers As Integer
    Dim lNewObjectID As Long
    
    '*** Jacob 4/8/2008 - IGNORE ERRORS
    On Error Resume Next
    
'    Call ISpicerMarkup_BindToDocumentControl
'    Call ISpicerMarkup_BindToViewControl
    Call BindControls
    
    
    ' Get the page ID for current page of active document
    lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
    ' Set object variable for IDocContents interface to doc ctrl object
    Set docContents = MainMDIForm.ActiveForm.SpicerDoc1.object
    ' Add a new layer
'    docContents.NewLayer lPageID, IN_LAYER_FULLEDIT
    docContents.NewLayer lPageID, IN_LAYER_ANNOTATION
    
    ' Get the id of the last layer
    iNumLayers = docContents.NumberOfLayers(lPageID)
    lNewLayerID = docContents.LayerID(lPageID, iNumLayers)
    lNewObjectID = MainMDIForm.ActiveForm.SpicerDoc1.NewestObjectID
    
    'Save the Layer ID
    arrPageLayerID(txtPageNumber) = SpicerDoc1.NewestObjectID
    
    ' Display the layerID of the newly imported layer
    '   If lNewLayerID = lNewObjectID Then
    '      MsgBox "ID of the new layer: " + Str(lNewLayerID), vbInformation
    '   Else
    '      MsgBox "ID of the new layer: " + Str(lNewLayerID), vbCritical
    '   End If
       ' De-initialize object var
    Set docContents = Nothing
    
    funcWriteToDebugLog Me.name, "EXIT: subAnnotationLayerCreate()"


End Sub

Public Sub subAnnotationLayerSaveCheck()
    
    If bolObjectLaunched = True Then
        Exit Sub
    End If
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    

    On Error GoTo ErrorHandler
    
    If bolRelatedImagesLoaded = False And InStr(1, Me.Caption, "BATCH:") = 0 Then
        Exit Sub
    End If
        
    
    ' Save Annotations Here?
    If bolAnnotationAdded = True Then
        
        If SpicerDoc1.NumberOfLayers(SpicerView1.ActivePageId) > 1 Then
            result = MsgBox("You have placed Annotations on THIS PAGE..." & vbCrLf & "Would you like to save them?", vbYesNo, "Save Annotations?")
            If result = vbYes Then
                subAnnotationLayerSave
            End If
            bolAnnotationAdded = False
        End If
    
    End If

ErrorHandler:
    'On Error just get out of here!

End Sub



Public Sub SelectTool(ByVal tool As TOOL_TYPE)
    

    
    funcWriteToDebugLog Me.name, "ENTER Sub SelectTool()... TOOL_TYPE = " & tool
    
         'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        
    On Error Resume Next
    
    funcWriteToDebugLog Me.name, "BindControls"
    BindControls
    
    funcWriteToDebugLog Me.name, "SpicerMarkup1.ActiveLayer = arrPageLayerID(" & txtPageNumber & ")"
    SpicerMarkup1.ActiveLayer = arrPageLayerID(txtPageNumber)
    
    funcWriteToDebugLog Me.name, "SpicerMarkup1.ActiveTool = IN_TOOL_NOTOOL"
    SpicerMarkup1.ActiveTool = IN_TOOL_NOTOOL
    DoEvents

    funcWriteToDebugLog Me.name, "SpicerMarkup1.ActiveTool = " & tool
    SpicerMarkup1.ActiveTool = tool
    DoEvents
    
    funcWriteToDebugLog Me.name, "SpicerView1.MinimizeAnnotations"
    SpicerView1.MinimizeAnnotations
    DoEvents
        
    funcWriteToDebugLog Me.name, "EXIT Sub SelectTool()"

End Sub

' Perform all necessary binding for activeX controls
Private Sub BindControls()
    
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
        On Error Resume Next
    
    
    SpicerView1.BindToDocumentControl SpicerDoc1.object
    SpicerEdit1.BindToDocumentControl SpicerDoc1.object
    SpicerMarkup1.BindToDocumentControl SpicerDoc1.object
    SpicerMarkup1.BindToViewControl SpicerView1.object

    SpicerView2.BindToDocumentControl SpicerDoc2.object
    SpicerEdit2.BindToDocumentControl SpicerDoc2.object
    
'    detailWin.SpicerDetail1.BindToViewControl SpicerView1.object
'    thumbWin.SpicerThumbnail1.BindToViewControl SpicerView1.object
'    refWin.SpicerReference1.BindToViewControl SpicerView1.object
'    layersWin.SpicerLayersWin1.BindToViewControl SpicerView1.object
End Sub


Public Sub subResetAnnotationButtons()

    frmAnnotate.cmdHighlightArea.BackColor = vbButtonFace
    frmAnnotate.cmdRedactBlack.BackColor = vbButtonFace
    frmAnnotate.cmdAnnotate.BackColor = vbButtonFace
    frmAnnotate.cmdCopy.BackColor = vbButtonFace
    frmAnnotate.cmdCut.BackColor = vbButtonFace
    frmAnnotate.cmdDrawFreehand.BackColor = vbButtonFace
    frmAnnotate.cmdLine.BackColor = vbButtonFace
    frmAnnotate.cmdMoveResize.BackColor = vbButtonFace
    frmAnnotate.cmdPaste.BackColor = vbButtonFace
    frmAnnotate.cmdSelect.BackColor = vbButtonFace
    frmAnnotate.cmdText.BackColor = vbButtonFace
    
    Call subAnnotationLayerCreate


End Sub


Public Sub subAnnotationLayerShowHide()

    'This example finds whether the first layer on the active page is displayed or hidden.
    '  If it is hidden, the layer is displayed.


   Dim LayerDisplay As ILayerDisplay
   Dim lLayerID As Long
   Dim bIsVisible As Boolean
   
        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear

    'IGNORE Any Errors
    On Error Resume Next
    
    
   ' Set the object variable for the ILayerDisplay interface to the View Control object
   Set LayerDisplay = SpicerView1.object
   ' Get the layerID for layer 1 of the active page
   lLayerID = arrPageLayerID(txtPageNumber)
   
   If lLayerID > 0 Then
       ' Find out whether the layer is shown or hidden
       bIsVisible = LayerDisplay.Visible(lLayerID)
    
       ' If the layer is hidden, display it
       If MainMDIForm.chkViewAnnotations = vbChecked Then
           LayerDisplay.Visible(lLayerID) = True
        Else
           LayerDisplay.Visible(lLayerID) = False
       End If
   End If
       'De-initialize the object variable
   Set LayerDisplay = Nothing

End Sub

Private Sub subDebugChildForm(strLocation As String)

    If Not bolDebug Then
        Exit Sub
    End If
    
'    On Error GoTo ErrorHandler
    On Error Resume Next
    
    funcWriteToDebugLog Me.name, ""
    funcWriteToDebugLog Me.name, strLocation
    funcWriteToDebugLog Me.name, "WINDOW DESCRIPTION            : " & Me.hwnd & "-" & Me.Caption
    funcWriteToDebugLog Me.name, "UBound(arrPageRotation)       : " & UBound(arrPageRotation)
    
    If txtPageNumber = "txtPageNumber" Then
        txtPageNumber = 0
    End If
    
    funcWriteToDebugLog Me.name, "txtPageNumber:                : " & txtPageNumber
    funcWriteToDebugLog Me.name, "arrPageRotation(txtPageNumber): " & arrPageRotation(txtPageNumber)
    funcWriteToDebugLog Me.name, "arrPageFileName(txtPageNumber): " & arrPageFileName(txtPageNumber)
    
Exit Sub

ErrorHandler:

    funcWriteToDebugLog Me.name, "ERROR: " & Err.Number & " - " & Err.Description & " SOURCE: " & Err.Source
    Resume Next
    
End Sub

Public Sub subRasterizeDocument()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    
   Dim RasterTools As IRasterTools
   Dim ActivePage As IActivePage
   Dim lPageID As Long
   Dim iRasterizeType As RASTERIZE_TYPE
   Dim iXResolution As Integer
   Dim iYResolution As Integer
   Dim bColor As Boolean
   Dim bDither As Boolean
   Dim iLighten As Integer
   
   'This flag is to disable the Form Load and Activate subs on the ChildForm
   bolRasterizingDocument = True
   
        ' Set the control to create a NEW Raster Document
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = Me.SpicerConfiguration1.object
        'Set to NOT automatically overwite document on rasterization.
        CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = False
        'Deinitialize the object variable
        Set CFGDocument = Nothing

    BindControls
   
   ' Set the object variable for the IRasterTools interface to the Edit Control object
   Set RasterTools = Me.SpicerEdit1.object
   ' Set the object variable for the IActivePage interface to the View Control object
   Set ActivePage = Me.SpicerView1.object
   
   ' Set the rasterize options
   lPageID = ActivePage.ActivePageId   ' Identifier of the active page
   If Me.SpicerDoc1.NumberOfPages > 1 Then
        iRasterizeType = IN_RASTERIZE_MULTIPAGE   ' Rasterize ALL pages
   Else
'        iRasterizeType = IN_RASTERIZE_DOCUMENT   ' Rasterize entire page
        iRasterizeType = IN_RASTERIZE_AS_DISPLAYED
   End If
   
   iXResolution = 0   ' Keep the original resolution
   iYResolution = 0
   bColor = True   ' Rasterize to color
   bDither = False   ' Do not rasterize to dither
   iLighten = 0
   
   ' Rasterize the active page
   RasterTools.Rasterize lPageID, iRasterizeType, iXResolution, _
               iYResolution, bColor, bDither, iLighten
               
   ' De-initialize the object variables
   Set RasterTools = Nothing
   Set ActivePage = Nothing
   
   bolRasterizingDocument = False

End Sub

Public Sub subRasterizeBatchEX()

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
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
   
   'This flag is to disable the Form Load and Activate subs on the ChildForm
   bolRasterizingDocument = True
   
        ' Set the control to create a NEW Raster Document
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = Me.SpicerConfiguration1.object
        'Set to NOT automatically overwite document on rasterization.
        CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = True
'        'Set to Remove the original
        CFGDocument.RasterOperations(IN_RASTERIZE_REMOVE) = True
        'Deinitialize the object variable
        Set CFGDocument = Nothing


    BindControls
   
   ' Set the object variable for the IRasterBatch interface to the Edit Control object
   Set RasterBatch = Me.SpicerEdit1.object
   
   ' Set the rasterize options
   lObjectID = Me.SpicerDoc1.RootID
   iXResolution = 0   ' Keep the original resolution
   iYResolution = 0
   iColor = IN_COLORTYPE_COLOR ' Rasterize to 256 color
   iBrightness = -50 ' Maximum darkness for Bilevel Dithered, Enhanced, and CAD
   iThreshold = 255  ' Maximum darkness for bilevel
   iOrientation = IN_ORIENTATION_NONE ' Use original resolution
   lXSize = 0 ' Keep the original size

   lYSize = 0
   iUnit = IN_UNITS_INCH
   
   ' Rasterize the entire document
   RasterBatch.RasterizeBatchEx lObjectID, iXResolution, iYResolution, iColor, iBrightness, iThreshold, iOrientation, lXSize, lYSize, iUnit
   
   ' De-initialize the object variables
   Set RasterBatch = Nothing
   
   bolRasterizingDocument = False

End Sub

Public Sub subLaunch(Optional strLaunchOption As String)

    Dim txtLaunchFilePath As String
    Dim txtLaunchFileName As String
    Dim txtLaunchFileExtension As String
    Dim intPositionOfLastSlash As Integer
    Dim intPositionOfLastPeriod As Integer
    
    On Error GoTo LAUNCH_ERROR
    
    
    txtPageFileName = arrPageFileName(txtPageNumber)
    
    txtLaunchFileName = txtPageFileName
    
    txtLaunchFileFullPath = txtPageFileName
    
    '*** 2022-06-07 - Jacob - Check if Launch File Exists
    If Dir(txtLaunchFileFullPath) = "" Then
    
            funcQuickMessage "Show", "ChildForm1.subLaunch | Sorry!!!  I was NOT able to find the file to Launch: " & _
                                vbCrLf & vbCrLf & txtLaunchFileFullPath & _
                                vbCrLf & vbCrLf & "Please contact IT for support."
            Exit Sub
    End If
    
        
    '*** Check if the Original document is to be Launched
    If strLaunchOption = "Edit" Then
        

        
        If Not bolAllowModificationOfOrigDocsMessageDisplayed Then
            funcQuickMessage "Show", "You have been granted the ability to Edit ORIGINAL Documents.  " & _
                                "Any changes to Edited documents CANNOT be undone!" & _
                                vbCrLf & vbCrLf & "Imaging101 cannot control changes to Edited documents and, therefore, is NOT responsible or liable for any changes or modifications to it."
                                
            bolAllowModificationOfOrigDocsMessageDisplayed = True
        End If
        
        
        MainMDIForm.ActiveForm.txtChildFormMessage.Visible = True
        MainMDIForm.ActiveForm.txtChildFormMessage.Text = "OBJECT LAUNCHED for EDIT!"
        MainMDIForm.ActiveForm.StatusBar1.Panels(1).Text = "Object Launched for EDIT"
        
        MainMDIForm.ActiveForm.lblLoadingPages.Visible = False
        
        'Close the document to release for edit in
        SpicerDoc1.CloseDocument False
    
        MainMDIForm.ActiveForm.lstPageList.Visible = False
        MainMDIForm.ActiveForm.SpicerView1.Visible = False
        

    End If
    
    ' De-initialize the object variable
'    Set docSave = Nothing
    
    MainMDIForm.ActiveForm.StatusBar1.Panels(1).AutoSize = sbrContents
    
    
    
                '***********************************************************
                '*** LAUNCH THE FILE
            
                funcWriteToDebugLog Me.name, "LAUNCHING FILE: " & txtLaunchFileFullPath
                Me.txtChildFormMessage.Text = "LAUNCHING FILE: " & vbCrLf & vbCrLf & txtLaunchFileFullPath

                Dim strAutoLaunchTo As String
                strAutoLaunchTo = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Control", "ID=1", "AutoLaunchTo")
                
                
                Dim lRet&
                Dim strShellCommand As String
                Dim strProgramPath As String
                
                strProgramPath = App.Path & "\I101DocumentViewer\I101DocumentViewer.exe"
                
                '*** Check if program exists in AppPath, otherwise assume we're in IDE
                If Dir(strProgramPath) <> "" Then
                    '*** Change Directory to the Application's directory.
                    ChDir App.Path
                    strShellCommand = strProgramPath & " """ & txtLaunchFileFullPath & """"
                Else
                    '*** Change Directory to the Development Release directory.
                    strProgramPath = "C:\VS2022\I101DocumentViewer\bin\Release\I101DocumentViewer.exe"
                    ChDir App.Path
                    strShellCommand = strProgramPath & " """ & txtLaunchFileFullPath & """"
                End If
                
                Err.Clear
                
                '**************************************************************************************
                '*** 2021-10-20 - Jacob - Check if Windows is already open using Wildcards
                
                Dim lngWindowExists As Long
                lngWindowExists = funcFindWindowLike("*" & txtLaunchFileFullPath)
                 
                If lngWindowExists <> 0 Then
                
                        bolObjectLaunched = True
                        funcShowImage = -1
                        'Bring the window we found to the top.
                        Call FormOnTop(lngWindowExists, True)
                
                Else
                
                        
                        If strAutoLaunchTo = "Associated Windows Application" Then
                                Me.txtChildFormMessage.Text = "ChildForm1.subLaunch() | strAutoLaunchTo = " & strAutoLaunchTo & " | shelldoc(" & txtLaunchFileFullPath & ")"

                                funcWriteToDebugLog Me.name, Me.txtChildFormMessage.Text
                        
                                'Launch in Associated Windows Application
                                Call shelldoc(txtLaunchFileFullPath)
                                
                                If Err.Number <> 0 Then
                                                Me.txtChildFormMessage.Text = "LAUNCH FAILED!" & vbCrLf & vbCrLf & _
                                                                                                "REASON: " & vbCrLf & vbCrLf & _
                                                                                                "ERROR: [" & Err.Number & " - " & Err.Description & "] " & vbCrLf & vbCrLf & _
                                                                                                  "shelldoc(" & txtLaunchFileFullPath & ")"
                
                                                funcWriteToDebugLog Me.name, Me.txtChildFormMessage.Text
                                                Exit Sub
                                End If

                                
                        Else
                                'Launch in NEW Imaging101 Viewer
                                '*** RUN the Application
                                Me.txtChildFormMessage.Text = "ChildForm1.subLaunch() | strAutoLaunchTo = " & strAutoLaunchTo & " | Shell(" & strShellCommand & ")"
                                funcWriteToDebugLog Me.name, Me.txtChildFormMessage.Text
                                
                                lRet = Shell(strShellCommand, vbNormalFocus)
                                
                                If Err.Number <> 0 Then
                                                Me.txtChildFormMessage.Text = "LAUNCH FAILED!" & vbCrLf & vbCrLf & _
                                                                                                "REASON: " & vbCrLf & vbCrLf & _
                                                                                                "ERROR: [" & Err.Number & " - " & Err.Description & "] " & vbCrLf & vbCrLf & _
                                                                                                 strShellCommand
                
                                                funcWriteToDebugLog Me.name, Me.txtChildFormMessage.Text
                                                Exit Sub
                                End If
                                
                        End If
               
                End If
  
    
    
    
 
    
Exit Sub
    
LAUNCH_ERROR:

    MsgBox "ERROR: " & Err.Number & " - " & Err.Description & _
    vbCrLf & "Unable to Launch document & vbCrLf & " & txtLaunchFileName


End Sub

Public Sub subSendTo()



    Err.Clear
'    On Error GoTo ERROR_HANDLER
'
'
'
'    '*********************************************************************
'    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - BEGIN
'
'    Dim txtAttachmentFileName As String
'    txtAttachmentFileName = InputBox("What would you like to call THIS document?" & vbCrLf & "Press [Cancel] or the [Esc] key to cancel the send.", "Attachment Name", "Imaging101 Document")
'
'    'Check if the user Canceled or entered no filename
'    If Trim(txtAttachmentFileName) = "" Then
'        Exit Sub
'    End If
''    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".TIF"
'    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".PDF"
'
''''    '*** ZAP Pages we DON'T want to Send
''''    Dim intPagesToSend As Integer
''''    txtAttachmentFileName = InputBox("What would you like to call THIS document?" & vbCrLf & "Press [Cancel] or the [Esc] key to cancel the send.", "Attachment Name", "Imaging101 Document")
''''
''''    'Check if the user Canceled or entered no filename
''''    If Trim(txtAttachmentFileName) = "" Then
''''        Exit Sub
''''    End If
''''    For i = 1 To totalpages
''''
''''    Next
''''
''''    Me.SpicerDoc1.DeleteObject
'
'
'
'    '*** SAVE the Pages
'    Dim docSave As IDocSave
'
'
'    ' Save the modified pages in the Spicer Document format
'    If Me.txtPageCount > 1 Then
'        '*** Rasterize the Pages before sending
''         me.subRasterizeBatch
'        Me.subRasterizeBatchEX
'        ' Set the object variable for the IDocSave interface to the Document Control object
'        ' that was saved by the Rasterize sub
'        Set docSave = Me.SpicerDoc1.object
''        docSave.SaveAsDialog False
''        docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, txtAttachmentFileName
'        'PDF - Flate, 619, R/W, Raster only Multipage PDF subset - bilevel, 8 bit, 24 bit
'        docSave.Save 0, False, 619, txtAttachmentFileName, txtAttachmentFileName
'        ' De-initialize the object variable
'        Set docSave = Nothing
'    Else
'        '*** Rasterize the Pages before sending
'         Me.subRasterizeBatchEX
'        ' Set the object variable for the IDocSave interface to the Document Control object
'        ' that was saved by the Rasterize sub
'        Set docSave = Me.SpicerDoc1.object
''        docSave.SaveAsDialog False
''        docSave.Save 0, False, API_FF_TIFFM, txtAttachmentFileName, txtAttachmentFileName
'        'PDF - Flate, 101, R/W, Raster only PDF subset - bilevel, 8 bit, 24 bit
'        docSave.Save 0, False, 101, txtAttachmentFileName, txtAttachmentFileName
'    End If
'
'
'
'    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - END
'    '*********************************************************************
'
'
'
'    '***********************************************************
'    ' Quick Outlook Send
''    Shell "C:\Program Files\Microsoft Office\Office11\OUTLOOK.EXE " & "/a " & txtFullPathName
'
'
'    '***********************************************************
'    '*** SEND USING Microsoft CDO - BEGIN
'      Dim objSession As MAPI.Session
'      Dim objmessage As MAPI.Message
'      Dim objRecipient As MAPI.Recipient
'      Dim objAttach As MAPI.Attachment
'
'
'      'Create the Session Object
'      Set objSession = CreateObject("MAPI.Session")
'''''''        Set objSession = CreateObject("Redemption.SafeMailItem")
'
'      'Logon using the session object
'      'Specify a valid profile name if you want to
'      'Avoid the logon dialog box
'      'If you don't include a profilename then a dialog is popup requesting one
'      objSession.Logon profileName:=""
'
'
'      'Add a new message object to the OutBox
'      Set objmessage = objSession.Outbox.Messages.Add
'
'      'Set the properties of the message object
'      objmessage.Subject = "Files from Imaging101 - "
'      objmessage.Text = ""
'
'      'Popup the global addresslist to select your recipients
'      'Force the resoluion of the named recipients
'      Set objmessage.Recipients = objSession.AddressBook(, "Select Recipients", , True, , ">>")
'
'      'To add attachments
'      Set objAttach = objmessage.Attachments.Add
'            objAttach.name = txtAttachmentFileName 'Pass the filename
'            objAttach.Source = txtAttachmentFileName  'Pass in your own filename
'            objAttach.Type = CdoFileData 'This is the Default
'
'      'Send the message
''      objmessage.Send showDialog:=False
'      objmessage.Send showDialog:=True
'
'       '*** DELETE THE TEMP FILE
'       Kill txtAttachmentFileName
'
'      'Logoff using the session object
'      objSession.Logoff
'
'      '*** SEND USING Microsoft CDO - END
'
'
'SEND_TO_EXIT:
'
'
'    '*********************************************************************
'    '*** TEMPORARY FIX FOR RASTERIZE PROBLEM
'    Unload Me
'    '*********************************************************************
'
'
'Exit Sub
'
'ERROR_HANDLER:
'    'Only show error message if it is not a User CANCEL
'    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
'        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
'    End If
'
'    '*********************************************************************
'    '*** TEMPORARY FIX FOR RASTERIZE PROBLEM
'    '***    Close the document viewer!!!
'    Unload Me
'    '*********************************************************************

End Sub


Public Sub subSendToOutlook()
  On Error GoTo ErrorHandle
  
  
    Dim txtAttachmentFileName As String
    txtAttachmentFileName = InputBox("What would you like to call THIS document?" & vbCrLf & "Press [Cancel] or the [Esc] key to cancel the send.", "Attachment Name", "Imaging101 Document")
    
    'Check if the user Canceled or entered no filename
    If Trim(txtAttachmentFileName) = "" Then
        Exit Sub
    End If
  
  '***2022-03-22 - Jacob - Changed form "Dim OutApp As Outlook.Application"
  '                               to "Dim OutApp As Object"
  '                               THIS stopped error "ACTIVEX Object NOT Registered."
'  Dim OutApp As Outlook.Application
  Dim OutApp As Object
  Dim OutMail As Object
   
  Set OutApp = CreateObject("Outlook.Application")
  Set OutMail = OutApp.CreateItem(0)
   
    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - BEGIN
    
'    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".TIF"
    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".PDF"
    
'''    '*** ZAP Pages we DON'T want to Send
'''    Dim intPagesToSend As Integer
'''    For i = 1 To totalpages
'''
'''    Next
'''
'''    Me.SpicerDoc1.DeleteObject

    
    
    '*** SAVE the Pages
    Dim docSave As IDocSave

    
    ' Save the modified pages in the Spicer Document format
    If Me.txtPageCount > 1 Then
        '*** Rasterize the Pages before sending
'         me.subRasterizeBatch
        Me.subRasterizeBatchEX
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        Set docSave = Me.SpicerDoc1.object
        'PDF - Flate, 619, R/W, Raster only Multipage PDF subset - bilevel, 8 bit, 24 bit
        docSave.Save 0, False, 619, txtAttachmentFileName, txtAttachmentFileName
        ' De-initialize the object variable
        Set docSave = Nothing
    Else
        '*** Rasterize the Pages before sending
         Me.subRasterizeBatchEX
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        Set docSave = Me.SpicerDoc1.object
        'PDF - Flate, 101, R/W, Raster only PDF subset - bilevel, 8 bit, 24 bit
        docSave.Save 0, False, 101, txtAttachmentFileName, txtAttachmentFileName
    End If

    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - END
    '*********************************************************************
  
  
  With OutMail
'    .Recipients =
    .To = ""
    .CC = ""
    .BCC = ""
    .Subject = "Files from Imaging101 Document Imaging... "
    .Body = ""
    .Attachments.Add txtAttachmentFileName
    .Display 'or Send
  End With
   
       
       
    '*** DELETE THE TEMP FILE
    Kill txtAttachmentFileName
    
   
   
ErrorExit:
  Set OutMail = Nothing
  Set OutApp = Nothing
  
    'Only show error message if it is not a User CANCEL
    If Err.Number <> 0 And InStr(1, UCase(Err.Description), "CANCEL") = 0 Then
        MsgBox "Error: " & Err.Number & "  Description: " & Err.Description
    End If
  Exit Sub
   
ErrorHandle:
  Resume ErrorExit
End Sub


Public Sub subSendToSMTP()


    On Error GoTo ERROR_HANDLER
    
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    
    strCommandSource = "subSendToSMTP"
    
    funcWriteToDebugLog Me.name, txtDocumentRECID
    
    
    Dim txtAttachmentFileName As String
    txtAttachmentFileName = InputBox("What would you like to call THIS document?" & vbCrLf & "Press [Cancel] or the [Esc] key to cancel the send.", "Attachment Name", "Imaging101 Document")
    
    'Check if the user Canceled or entered no filename
    If Trim(txtAttachmentFileName) = "" Then
        Exit Sub
    End If
    
    
'    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".TIF"
    txtAttachmentFileName = Environ("TEMP") & "\" & txtAttachmentFileName & ".PDF"
    
'''    '*** ZAP Pages we DON'T want to Send
'''    Dim intPagesToSend As Integer
'''    For i = 1 To totalpages
'''
'''    Next
'''
'''    Me.SpicerDoc1.DeleteObject

    
    
    '*** SAVE the Pages
    Dim docSave As IDocSave

    
    ' Save the modified pages in the Spicer Document format
    If Me.txtPageCount > 1 Then
        '*** Rasterize the Pages before sending
'         me.subRasterizeBatch
        Me.subRasterizeBatchEX
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        Set docSave = Me.SpicerDoc1.object
        'PDF - Flate, 619, R/W, Raster only Multipage PDF subset - bilevel, 8 bit, 24 bit
        docSave.Save 0, False, 619, txtAttachmentFileName, txtAttachmentFileName
        ' De-initialize the object variable
        Set docSave = Nothing
    Else
        '*** Rasterize the Pages before sending
         Me.subRasterizeBatchEX
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        Set docSave = Me.SpicerDoc1.object
        'PDF - Flate, 101, R/W, Raster only PDF subset - bilevel, 8 bit, 24 bit
        docSave.Save 0, False, 101, txtAttachmentFileName, txtAttachmentFileName
    End If

    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - END
    '*********************************************************************
      
           
    'Prepare to send the attachment
'    frmMain.Show
    frmSMTPeMailForm.Show
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
    SpicerDoc1.CloseDocument (False)
    DoEvents

    '*** Close the printer and clear the buffer.
    Printer.EndDoc
    
    '********************************
    '*** DELETE the Temporary File
    On Error Resume Next
    Kill txtAttachmentFileName
                    
    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub



Private Sub StatusBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

        'Bypass the Spicer Error
        SpicerConfiguration1.BatchMessageMode = True
        
        'CLEAR the ERROR to make sure it doesn't remember a previous error.
        Err.Clear
    
    On Error Resume Next
    
    '*** 2020-05-27 - Jacob - Added Right Click to show document Properties
    If Button = vbRightButton Then
        On Error Resume Next
        SpicerDoc1.DocumentPropertiesDialog
    End If
    
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As ComctlLib.Panel)
    
    If gsecRightsAdminSystem = vbChecked Then
        funcQuickMessage "SHOW", txtPageFileName
    End If

End Sub

Public Function funcGetPageRotation(ByVal intPageNumber As Integer) As Integer

    funcGetPageRotation = arrPageRotation(txtPageNumber)

End Function

Public Function funcLockRotation()

    'Set the Rotation to be Locked for ALL Pages of the Batch
    intLockedRotation = funcGetPageRotation(1)

End Function

Public Function funcCheckForOpenOrImportErrors(ErrNumber As Integer, _
                                                                                    ErrDescription As String, _
                                                                                    txtFullPathName As String)
        
        
        'Bypass the Spicer Error if file cannot be opened
        Me.SpicerConfiguration1.BatchMessageMode = True

        On Error Resume Next
        
        '*** 2022-07-18 - Jacob - Moved Variable Definitions to Top
        Dim txtLaunchFileName As String
        Dim txtLaunchPath As String
        Dim txtLaunchFileFullPath As String
        Dim intPositionOfLastBackslash As Integer
        Dim txtDocumentPagePlaceHolder As String
        
        txtPageFileName = txtFullPathName
        txtLaunchFileFullPath = txtFullPathName
        
        intPositionOfLastPeriod = InStrRev(txtPageFileName, ".")
        txtLaunchFileExtension = Right(txtPageFileName, Len(txtPageFileName) - intPositionOfLastPeriod)
        
        intPositionOfLastBackslash = InStrRev(txtPageFileName, "\")
        txtLaunchFileName = Right(txtPageFileName, Len(txtPageFileName) - intPositionOfLastBackslash)
        txtLaunchFilePath = Left(txtPageFileName, intPositionOfLastBackslash)
        
        '*** 2022-07-18 - Jacob - Appended Tilde (~) to make it a "Temp" file.
        txtDocumentPagePlaceHolder = txtLaunchFilePath & "~" & txtLaunchFileName & ".txt"
       
        Set docContents = Me.SpicerDoc1.object
                
                
                
        '*** 2020-05-14 - Jacob - MOVED THIS SECTION UP from AFTER the OpenFile to allow for AutoLaunch in case of an error.
        '***********************************************************************************************************************************************
        '*** IF ERROR OCCURED OR FILE EXTENSION IS FLAGGED FOR AUTOLAUNCH THEN LAUNCH THE FILE / OBJECT
        

        
        '*** 2020-05-14 - Jacob - ADDED CHECK TO ONLY OPEN IF AutoLaunch NOT Set
        Dim strErrorNumber As String
        Dim strErrorDescription As String
        
        If bolLaunchThisFileType = False Then
        
                funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | bolLaunchThisFileType = FALSE | Prepare to OPEN File."
                  
                '****************************************************************************************************
                '***     PREPARE TO OPEN the FILE
            
                'Set the View Control NOT Visible to prevent the Flash from Full view to zoomed.
                Me.SpicerView1.Visible = False
                
                
                '*** 2020-04-22 - Jacob - Clear the txtChildFormMessage and make it NOT Visible, so it doesn't show for each subsequent Open.
                Me.txtChildFormMessage.Text = ""
                Me.txtChildFormMessage.Visible = False
        

                
                'Bypass the Spicer Error if file cannot be opened
                Me.SpicerConfiguration1.BatchMessageMode = True
                
                
                 '****************************************************************************************************
                 '*** 2021-08-10 - Jacob - Check for Open errors.
                 If ErrNumber = 0 Then
                 
                            '*** 2020-05-14 - Jacob - Moved this logic to make it more consistent
                            
                            'CLEAR the ERROR to make sure it doesn't remember a previous error.
                            Err.Clear
                
                            '*** Clear the Page List
                            funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | NO ERROR OPENING PAGE - Will Test for Other Errors"
                            
                            
                            '*** 2021-07-29 - Jacob - Cascading Error Detection
                            
                            Dim lLayerID As Long
                            Dim lPageID As Long
                            
                             'Set the ACTIVE PAGE only if there were NO Errors Opening the File
                            funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | Set ActivePage = Me.SpicerView1.object"
                            Set ActivePage = Me.SpicerView1.object
                            
                            If Err.Number = 0 Then
                                    funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | ActivePage.BindToDocumentControl Me.SpicerDoc1.object"
                                   ActivePage.BindToDocumentControl Me.SpicerDoc1.object
                                    
                                    If Err.Number = 0 Then
                                            funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | lPageID = Me.SpicerView1.ActivePageId"
                                            lPageID = Me.SpicerView1.ActivePageId
                                    
                                            If Err.Number = 0 Then
                                                funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | lLayerCount = docContents.NumberOfLayers(lPageID)"
                                                lLayerCount = docContents.NumberOfLayers(lPageID)
                                    
                                                    If Err.Number = 0 Then
                                                        funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | lLayerID = Me.SpicerDoc1.LayerID(lPageID, 1)"
                                                        lLayerID = Me.SpicerDoc1.LayerID(lPageID, 1)
                                    
                                                            If Err.Number = 0 Then
                                                                    funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | Me.SpicerView1.ZoomLevel(0) = IN_ZOOM_SCALETOFIT"
                                                                    Me.SpicerView1.ZoomLevel(0) = IN_ZOOM_SCALETOFIT
                                                                    
                                                                    If Err.Number = 0 Then
                                                                                ' Zoom to Saved factor  -  If an error occurs during Zoom then bolErrorOccured will be set to True
                                                                                '                                             and will return the error that occured in strZoomResult.
                                                                                Dim strZoomResult As String
                                                                                funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors()  |  strZoomResult = MainMDIForm.funcZoomToSavedFactor"
                                                                                strZoomResult = MainMDIForm.funcZoomToSavedFactor
                                                                                sErrMessage = strZoomResult
                                                                                funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | sErrMessage = strZoomResult | strZoomResult= " & strZoomResult
                                                                                'MsgBox "funcShowImage() | funcZoomToSavedFactor() ERROR:" & vbCrLf & " While attempting to set Saved ZOOM Factor." & vbCrLf & vbCrLf & strZoomResult & vbCrLf & vbCrLf & "CLICK OK AND I WILL ATTEMPT TO LAUNCH IT.", vbCritical
                                                                            
                                                                    Else
                                                                        bolErrorOccured = True
                                                                        sErrMessage = "funcCheckForOpenOrImportErrors() | ERROR While attempting Me.SpicerView1.ZoomLevel(0) = IN_ZOOM_SCALETOFIT"
                                                                    End If
                                                                
                                                            Else
                                                                bolErrorOccured = True
                                                                sErrMessage = "funcCheckForOpenOrImportErrors() | ERROR While attempting lLayerID = Me.SpicerDoc1.LayerID(lPageID, 1)"
                                                            End If
                                    
                                                    Else
                                                        bolErrorOccured = True
                                                        sErrMessage = "funcCheckForOpenOrImportErrors() | ERROR While attempting lLayerCount = docContents.NumberOfLayers(lPageID)"
                                                    End If
                            
                                            Else
                                                bolErrorOccured = True
                                                sErrMessage = "funcShofuncCheckForOpenOrImportErrorswImage() | ERROR While attempting lPageID = Me.SpicerView1.ActivePageId"
                                            End If
                            
                                    Else
                                            bolErrorOccured = True
                                            sErrMessage = "funcCheckForOpenOrImportErrors() | ERROR While attempting ActivePage.BindToDocumentControl Me.SpicerDoc1.object."
                                     End If
                            
                            Else
                                    bolErrorOccured = True
                                    sErrMessage = "funcCheckForOpenOrImportErrors() | ERROR While attempting Set ActivePage = Me.SpicerView1.object"
                            End If



                                            

                            '*** 2021-08-10 - Jacob - Added this If to handle errors during the Cascading Error traps above...
                            
                            If bolErrorOccured = True Then
                                    
                                     'CAPTURE ERROR(S)
                                    strErrorNumber = Err.Number
                                    strErrorDescription = Err.Description

                                    funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | bolErrorOccured = True | sErrMessage = " & sErrMessage & " | strErrorNumber = " & strErrorNumber & " | strErrorDescription= " & strErrorDescription

'                                    bolLaunchThisFileType = True
                                    'DO NOT Exit Function... so the Launch code can be processed.
                                    
                            Else
                            
                                    funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | bolErrorOccured = False | *** DOCUMENT OPENED SUCCESSFULLY..."
                                    '8/2/2005 Jacob - Commented the above ScaleToGray to make the code more consistent
                                    funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | Me.cmdScaleToGray"
                                    Me.cmdScaleToGray


                                    Set docContents = Nothing
                                    Set frmViewForm = Nothing
                                    
'''                                     Me.SpicerView1.Visible = True
                                
                                     'GET OUT OF HERE NOW
                                     Exit Function
                            End If
                            
                 Else
                 
                            '*** CAPTURE ERROR(S) TO FORCE LAUNCH
                            strErrorNumber = ErrNumber
                            strErrorDescription = ErrDescription
                            
                           ' Error #'s ABOVE 0 are Visual Basic errors.
                           ' BELOW 0 are Spicer ActiveX Control errors.
                                                            
                 End If
                 
        Else
                 funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | bolLaunchThisFileType = True"
        
        End If

        
                
               
                
                
                '*** 2020-04-22 - Jacob - Moved Panel(1) up from after messages to show File info, in case OS asks for APP to Launch into.
                '                                             and added Panes (2) and (3) to show what we're launching
                
                
                Me.StatusBar1.Panels(1).Text = "Object Launched"
                Me.StatusBar1.Panels(2).Text = txtLaunchFileFullPath
                Me.StatusBar1.Panels(3).Text = txtLaunchFileExtension
                Me.StatusBar1.Visible = True
                

                '***********************************************************
                '*** LAUNCH THE FILE
                
            '*** IF NOT ONE OF THE DEFINED FILE TYPES TO LAUNCH
            '      THEN SOMETHING WENT WRONG
            '     DISPLAY ERROR OPENING FILE - AND LAUNCH THE FILE
            '     OTHERWISE, SIMPLY LAUNCH THE FILE
            
                funcWriteToDebugLog Me.name, "LAUNCHING FILE: " & txtLaunchFileFullPath
                Me.txtChildFormMessage.Text = "LAUNCHING FILE: " & vbCrLf & vbCrLf & txtLaunchFileFullPath

                Dim strAutoLaunchTo As String
                strAutoLaunchTo = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Control", "ID=1", "AutoLaunchTo")
                
                
                Dim lRet&
                Dim strShellCommand As String
                Dim strProgramPath As String
                
                strProgramPath = App.Path & "\I101DocumentViewer\I101DocumentViewer.exe"

                '*** Check if program exists in AppPath, otherwise assume we're in IDE
                If Dir(strProgramPath) <> "" Then
                    '*** Change Directory to the Application's directory.
                    ChDir App.Path
                    strShellCommand = strProgramPath & " """ & txtLaunchFileFullPath & """"
                Else
                    '*** Change Directory to the Development Release directory.
                    strProgramPath = "C:\VS2019\I101DocumentViewer\bin\Release\I101DocumentViewer.exe"
                    ChDir App.Path
                    strShellCommand = strProgramPath & " """ & txtLaunchFileFullPath & """"
                End If
                
                Err.Clear
                
                
                '**************************************************************************************
                '*** 2021-10-20 - Jacob - Check if Windows is already open using Wildcards
                
                Dim lngWindowExists As Long
                lngWindowExists = funcFindWindowLike("*" & txtLaunchFileName & "*")
                 
                If lngWindowExists <> 0 Then
                
                        bolObjectLaunched = True
                        funcShowImage = -1
                        'Bring the window we found to the top.
                        Call FormOnTop(lngWindowExists, True)
                
                Else
                
                        
                        If strAutoLaunchTo = "Associated Windows Application" Then
                                'Launch in Associated Windows Application
                                Call shelldoc(txtLaunchFileFullPath)
                                
                        Else
                                'Launch in NEW Imaging101 Viewer
                                '*** RUN the Applicatnoi
                                lRet = Shell(strShellCommand, vbNormalFocus)
                                
                                If Err.Number <> 0 Then
                                                Me.txtChildFormMessage.Text = "LAUNCH FAILED!" & vbCrLf & vbCrLf & _
                                                                                                "REASON: " & vbCrLf & vbCrLf & _
                                                                                                "ERROR: [" & Err.Number & " - " & Err.Description & "] " & vbCrLf & vbCrLf & _
                                                                                                 strShellCommand
                
                                                funcWriteToDebugLog Me.name, Me.txtChildFormMessage.Text
                                                Exit Function
                                End If
                                
                        End If
               
                End If
                
                If bolLaunchThisFileType = True Then
                        Me.txtChildFormMessage.Text = "PAGE LAUNCHED!" & vbCrLf & vbCrLf & _
                                                                                "REASON: " & vbCrLf & vbCrLf & _
                                                                                "File Extension [" & txtLaunchFileExtension & "] Configured for Auto-Launch"
                Else
                         '*** 2020-05-14 - Jacob - EDITED the message text to make more understandable..
                       Me.txtChildFormMessage.Text = "PAGE LAUNCHED!" & vbCrLf & vbCrLf & _
                                                                                strErrorNumber & " - " & strErrorDescription & vbCrLf & vbCrLf & _
                                                                                sErrMessage & vbCrLf & vbCrLf & _
                                                                                "SOMETHING IN THIS FILE CAUSED AN ERROR" & vbCrLf & _
                                                                                "THAT PREVENTED ME FROM OPENING IT." & vbCrLf & vbCrLf & _
                                                                                "I have 'Launched' it" & vbCrLf & _
                                                                                "so that you can view it" & vbCrLf & _
                                                                                "in another Program or Viewer" & vbCrLf & _
                                                                                "that is 'associated' with" & vbCrLf & _
                                                                                "File Extension '." & txtLaunchFileExtension & "' "
                                                                                
                         If bolAIM_Command_AddFile = True _
                         Or txtModuleIndex = gI101ModuleIndex Then
                         
                                      Me.txtChildFormMessage.Text = Me.txtChildFormMessage.Text & vbCrLf & vbCrLf & _
                                                                                                                    "You can CONTINUE to INDEX it." & vbCrLf & _
                                                                                                                    "and it WILL be SAVED properly."
                                                                                                                    
                        End If
                    
                                                                                                                    
                End If
                
                Me.txtChildFormMessage.Text = Me.txtChildFormMessage.Text & vbCrLf & vbCrLf & txtLaunchFileFullPath
                
                
                funcWriteToDebugLog Me.name, Me.txtChildFormMessage.Text
                funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | bolErrorOccured = True | sErrMessage = " & sErrMessage & " | strErrorNumber = " & strErrorNumber & " | strErrorDescription= " & strErrorDescription


                  
                  '***************************************************************
                  '*** Create The Document Page PlaceHolder Text File
                  
                  Dim iFileNo As Integer
                  iFileNo = FreeFile

                  Open txtDocumentPagePlaceHolder For Output As #iFileNo
                  Print #iFileNo, Me.txtChildFormMessage.Text
                  Close #iFileNo
                  
                  '***************************************************************
                  '*** 2022-03-23 - Jacob - Try to insert Dummy Text Page as a Place-Holder.
                  
'                  Dim dblNumberOfPagesBeforeImport As Double
'                  Dim dblNumberOfPagesAfterImport As Double
                  
                  
                  
                    If lstPageList.ListCount = 0 Then
                            docContents.CloseDocument False
                            docContents.OpenFile txtDocumentPagePlaceHolder
                    Else
                            docContents.ImportPage 0, 0, IN_NEWPAGE_END, txtDocumentPagePlaceHolder, txtDocumentPagePlaceHolder
                    End If
                    
                    
'                    '*** Show only ONE (1) Page in the list for the Launched document
'                    dblNumberOfPagesAfterImport = dblNumberOfPagesBeforeImport + 1
                             
                    funcWriteToDebugLog Me.name, "funcCheckForOpenOrImportErrors() | dblNumberOfPagesAfterImport = " & dblNumberOfPagesAfterImport
                             
'                    '*** Update the Arrays to reflect the ACTUAL number of Pages Loaded
'                    '     just in case the FIRST document loaded was a MULTI-PAGE file!
'                    funcWriteToDebugLog Me.name, "subUpdatePageArrays dblNumberOfPagesBeforeImport = " & dblNumberOfPagesBeforeImport & ", dblNumberOfPagesAfterImport = " & dblNumberOfPagesAfterImport
'                    subUpdatePageArrays dblNumberOfPagesBeforeImport, dblNumberOfPagesAfterImport
                    
                  
                  Set docContents = Nothing
                  
                    bolObjectLaunched = True


End Function


Function funcCopyFileToLocalTemp(txtFileName As String, txtFullPathName As String, txtRECID As String)

                 '*** 2022-03-29 - Jacob - Replaced the Copy File to LocalTemp from MainMDIForm.funcShowImage
                 
                 Dim txtCopyFilePath As String
                 Dim txtCopyFileFullPath As String
                 Dim bCopySuccess As Boolean
                 
                strLocalTempDir = funcGetTempDir()
                funcWriteToDebugLog Me.name, "Windows returned TEMPDIR = " & strLocalTempDir
               
                txtCopyFilePath = strLocalTempDir & "Imaging101\"
                funcWriteToDebugLog Me.name, " txtCopyFilePath = " & txtCopyFilePath
                
                txtPageFileName = txtFullPathName
                funcWriteToDebugLog Me.name, "txtPageFileName = " & txtPageFileName
                
                 
                 
                ' *********************************************************************
                ' *** 2023-01-03 - Jacob - Added  txtDetailRECID (DetailReCID or BatchPageRECID) File PREFIX as a Unique Identifier
                '                                          in case different batch pages have the same names.
                ' *********************************************************************
                 
                 
                txtCopyFileFullPath = txtCopyFilePath & txtDetailRECID & "_" & txtFileName   'strBatchPageFileName
                 
                 
                 
                 
                Me.txtChildFormMessage.Text = "COPYING FILE TO LOCAL TEMP DIRECTORY..."
                funcWriteToDebugLog Me.name, "APIFileCopy(" & txtFullPathName & ", " & txtCopyFileFullPath
                 
                '*** 2021-08-11 - Jacob - Check if Launch File Already Exist in Temp dir.
                If Dir(txtCopyFileFullPath) = "" Then
                        bCopySuccess = APIFileCopy(txtFullPathName, txtCopyFileFullPath, False)

                        If bCopySuccess = False Then
                                    funcWriteToDebugLog Me.name, "ERROR COPYING FILE !!!"
                                    Me.txtChildFormMessage.Text = "ERROR COPYING FILE TO LOCAL TEMP DIRECTORY!!!" & vbCrLf & vbCrLf & _
                                                                                                                    "This could mean that either:" & vbCrLf & vbCrLf & _
                                                                                                                    "The file " & vbCrLf & txtLaunchFileName & vbCrLf & "is ALREADY open OR in-use." & vbCrLf & vbCrLf & _
                                                                                                                    "OR" & vbCrLf & vbCrLf & _
                                                                                                                    "There is a problem copying the file to directory: " & vbCrLf & txtCopyFilePath
                                    funcShowImage = -1
                                    bolObjectLaunched = False
                                    
                                    funcCopyFileToLocalTemp = ""
                                    
                                    Exit Function
                        End If
                Else
                
                        funcWriteToDebugLog Me.name, "FILE ALREADY EXISTS AT DESTINATION: " & txtCopyFileFullPath
                        Me.txtChildFormMessage.Text = "FILE ALREADY EXISTS AT DESTINATION: " & vbCrLf & vbCrLf & txtCopyFileFullPath
                
                End If

        funcCopyFileToLocalTemp = txtCopyFileFullPath

End Function


Function funcCheckIfFileShouldBeLaunched(txtFullPathName)
        
        
        On Error Resume Next
                             
        '**********************************************************************************************************
        '*** funcCheckIfFileShouldBeLaunched(txtFullPathName) - Check if File should be Launched
        
        'Load Auto-Launch File Types
        Dim strAutoLaunchFileTypes As String

        '*** 2021-07-14 - Jacob - Moved  Dim bolLaunchThisFileType to Module "VariableDeclarations"
        'Dim bolLaunchThisFileType As Boolean
        Dim intPositionOfPeriod As Integer
        Dim strFileExtension As String
        
        '*** DIM variables for GdPicture
        Dim dblDocHeight As Double
        Dim dblDocWidth As Double
        Dim gdpStatus As GdPictureStatus

        
        '*** 2020-04-22 - Jacob - Added UCase's to make the File Extensions Case-Insensitive
        intPositionOfPeriod = InStrRev(txtFullPathName, ".")
        strFileExtension = UCase(Right(txtFullPathName, Len(txtFullPathName) - intPositionOfPeriod))
        bolLaunchThisFileType = False
        
        strAutoLaunchFileTypes = UCase(funcGetFieldFromDB(RegImaging101ConnectionString, "I101Control", "ID=1", "AutoLaunchFileTypes"))
        
        '*** Load LaunchFileTypes into an Array
        Dim lstAutoLaunchFileTypes() As String
        lstAutoLaunchFileTypes = Split(strAutoLaunchFileTypes, ",")
        
        'Check if this file extension is in the Launch types array
        For i = 0 To UBound(lstAutoLaunchFileTypes)
                If lstAutoLaunchFileTypes(i) = strFileExtension Then
                        bolLaunchThisFileType = True
                        Exit For
                End If
        Next
        
        
        
        Err.Clear
    
        
        '*** Check If any file types have been configured for AutoLaunch
        If bolLaunchThisFileType = False Then
            
            funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | *** bolLaunchThisFileType = False | GdViewer1.CloseDocument"

            '*** Convert EML and MSG files to PDF
'            m_GdPictureDocumentConverter.EmailPageHeight = 635 ' // Letter page size
'            m_GdPictureDocumentConverter.EmailPageWidth = 822 '  // Letter page size
'            m_GdPictureDocumentConverter.EmailPageMarginTop = 10
'            m_GdPictureDocumentConverter.EmailPageMarginBottom = 10
'            m_GdPictureDocumentConverter.EmailPageMarginLeft = 10
'            m_GdPictureDocumentConverter.EmailPageMarginRight = 10
                
            Err.Clear
    
                
            Select Case strFileExtension
            
                Case "EML"
                        
                        If Dir(txtFullPathName & ".pdf") <> "" Then
                                
                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'EML' |   " & txtFullPathName & ".pdf" & " ALREADY EXISTS... SKIP LOAD AND SAVE."
                        
                        Else
         
                                    funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'EML' |   gdpStatus = gdpStatus = m_GdPictureDocumentConverter.LoadFromFile(" & txtFullPathName & ", DocumentFormat_DocumentFormatEML) "
            
                                    gdpStatus = m_GdPictureDocumentConverter.LoadFromFile(txtFullPathName, DocumentFormat_DocumentFormatEML)
                                    
                                     If gdpStatus = GdPictureStatus_OK Then
                                            
                                            funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'EML' - GdPictureStatus_OK |  gdpStatus = m_GdPictureDocumentConverter.SaveAsPDF(" & txtFullPathName & ".pdf" & ", PdfConformance_PDF1_5)"
                                            gdpStatus = m_GdPictureDocumentConverter.SaveAsPDF(txtFullPathName & ".pdf", PdfConformance_PDF1_5)

                                             If gdpStatus = GdPictureStatus_OK Then
                                                    funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'EML' - GdPictureStatus_OK"
                                            Else
                                                    Err.Raise -10104, "funcCheckIfFileShouldBeLaunched()", " | Case 'EML' - ERROR | SaveAsPDF() FAILED "
                                                    funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                                    funcQuickMessage "SHOW", "funcCheckIfFileShouldBeLaunched() | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                                     Exit Function
                                                     
                                            End If
            
                                     Else
                                            
                                        Err.Raise -10103, "funcCheckIfFileShouldBeLaunched()", " | Case 'EML' - ERROR | DocumentConverter.LoadFromFile() FAILED "
                                        funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                        funcQuickMessage "SHOW", "funcCheckIfFileShouldBeLaunched() |  Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                        Return
                                        
                                    End If 'gdpStatus = GdPictureStatus_OK
                               
                         End If ' File Exists
                                     
                                    '*** 2022-07-15 - Jacob - In the PDF world, 1 point = 1/72 inch
                                    dblDocHeight = 11
                                    dblDocWidth = 8.5
                        
                Case "MSG"
                        
                        Err.Clear

                        If Dir(txtFullPathName & ".pdf") <> "" Then
                                
                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'MSG' |   " & txtFullPathName & ".pdf" & " ALREADY EXISTS... SKIP LOAD AND SAVE."
                        
                        Else

                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'MSG' |   gdpStatus = gdpStatus = m_GdPictureDocumentConverter.LoadFromFile(" & txtFullPathName & ", DocumentFormat_DocumentFormatMSG) "
                                
                                gdpStatus = m_GdPictureDocumentConverter.LoadFromFile(txtFullPathName, DocumentFormat_DocumentFormatMSG)
                                
                                  If gdpStatus = GdPictureStatus_OK Then
                                 
                                        funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'MSG' - GdPictureStatus_OK  | gdpStatus = m_GdPictureDocumentConverter.SaveAsPDF(" & txtFullPathName & ".pdf" & ", PdfConformance_PDF1_5)"
                                        gdpStatus = m_GdPictureDocumentConverter.SaveAsPDF(txtFullPathName & ".pdf", PdfConformance_PDF1_5)
                                        
                                         If gdpStatus = GdPictureStatus_OK Then
                                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Case 'MSG' - GdPictureStatus_OK |   gdpStatus = m_GdPictureDocumentConverter.SaveAsPDF(" & txtFullPathName & ".pdf" & ", PdfConformance_PDF1_5)"
                                        Else
                                                Err.Raise -10104, "funcCheckIfFileShouldBeLaunched()", " | Case 'MSG' - ERROR | SaveAsPDF() FAILED "
                                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                                funcQuickMessage "SHOW", "funcCheckIfFileShouldBeLaunched() | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                                 Exit Function
                                                 
                                        End If
        
                                 Else
                                        
                                        Err.Raise -10103, "funcCheckIfFileShouldBeLaunched()", " | Case 'MSG' - ERROR | DocumentConverter.LoadFromFile() FAILED "
                                        funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                        funcQuickMessage "SHOW", "funcCheckIfFileShouldBeLaunched() |  Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                                        Exit Function
                                        
                                End If 'gdpStatus = GdPictureStatus_OK
                               
                         End If ' File Exists
                               
                                '*** 2022-07-15 - Jacob - In the PDF world, 1 point = 1/72 inch
                                dblDocHeight = 11
                                dblDocWidth = 8.5
                        
                Case Else
                
                        '*** Open Document in GdViewer and Get Page sizes
                        'GdViewer1.CloseDocument
            
                        funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() |   gdpStatus = GdViewer1.DisplayFromFile(" & txtFullPathName & ")  | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " "
                        
                        gdpStatus = GdViewer1.DisplayFromFile(txtFullPathName)
                        
                        
                        funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched()  | Err.Number = " & Err.Number & "  | Err.Description=" & Err.Description & " | GdPictureStatus=" & gdpStatus & " | GdViewer1.GetDocumentType() = " & GdViewer1.GetDocumentType() & " | DocumentType_DocumentTypePDF=" & DocumentType_DocumentTypePDF
                        
                        If gdpStatus = GdPictureStatus_OK Then
                            
                            funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() |   GdPictureStatus_OK"
                            
                             '*** GdViewer handles PDF's different than other file types
                             If GdViewer1.GetDocumentType() = DocumentType_DocumentTypePDF Then
                             
                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() |  DocumentType PDF   |   Get PDF Height & Width"
                                
                                '*** 2022-07-15 - Jacob - In the PDF world, 1 point = 1/72 inch
                                dblDocHeight = Format(GdViewer1.PdfGetPageHeight() / 72, "0.##")
                                dblDocWidth = Format(GdViewer1.PdfGetPageWidth() / 72, "0.##")
                                
                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | DocumentType PDF   |  DocWidth=" & dblDocWidth & "   DocHeight=" & dblDocHeight
                                
                            Else
                            
                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() |  DocumentType NOT PDF  |   Get PDF Height & Width"
                                
                                '*** 2022-07-15 - Jacob - Changed Pixel divisor from 72 to GdViewer1.HorizontalResolution
                                dblDocHeight = Format(GdViewer1.PageHeight() / GdViewer1.VerticalResolution, "0.##")
                                dblDocWidth = Format(GdViewer1.PageWidth() / GdViewer1.HorizontalResolution, "0.##")
                                
                                funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | DocumentType NOT PDF |  DocWidth=" & dblDocWidth & "   DocHeight=" & dblDocHeight
                            
                            End If
                            

                            
                        Else
                            
                            Err.Raise -10102, "funcCheckIfFileShouldBeLaunched()", "DisplayFromFile() FAILED | GdPictureStatus=" & gdpStatus & " "
                            funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | DisplayFromFile() FAILED | GdPictureStatus=" & gdpStatus & " "
                            Exit Function
                            
                        End If
                        
                    End Select
                    
                    If Err.Number = 0 Then
                        
                             '*** 2022-07-19 - Jacob - Added Arrays for Page Width and Page Height
                            txtDocHeight = dblDocHeight
                            txtDocWidth = dblDocWidth
                            
                            funcWriteToDebugLog Me.name, "funcCheckIfFileShouldBeLaunched() | BEFORE FILE SIZE CHECK | Err.Number = " & Err.Number & "  Err.Description=" & Err.Description
            
                            '*** Check for Oversized document - The SPICER / OPENTEXT Viewer can open files up to 22" x 22"
                            '     Force LAUNCH if larger than this.
                            If (GdViewer1.GetDocumentType() <> DocumentType_DocumentTypeBitmap) And (dblDocHeight > 22 Or dblDocWidth > 22) Then
                                
        '                        m_GdPicturePDFReducer.ProcessDocument txtFullPathName, txtFullPathName & "_reduced.pdf"
            
                                '*** Raise Error to FORCE  LAUNCH
                                 Err.Raise -10101, "funcCheckIfFileShouldBeLaunched()", "Unable to Display.  Document is larger than Viewer can handle (Width= " & dblDocWidth & "in.  Height= " & dblDocHeight & "in. |  Err.Number = -10101"
                                 funcWriteToDebugLog Me.name, "*** RAISE ERROR -10101 |  Unable to Display. Document is larger than Viewer can handle (Width= " & dblDocWidth & "in.  Height= " & dblDocHeight & "in.  |  Err.Number = -10101"
                            
                            End If
                            
                End If  'Err.Number = 0
                    
        End If

End Function
