VERSION 5.00
Begin VB.Form frmEzScan 
   Caption         =   "EZ Scan"
   ClientHeight    =   9900
   ClientLeft      =   1935
   ClientTop       =   2130
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9900
   ScaleWidth      =   10875
   Begin VB.Frame Frame2 
      Caption         =   "Bar Codes"
      Height          =   1704
      Left            =   5364
      TabIndex        =   17
      Top             =   108
      Width           =   4152
      Begin VB.CommandButton cmdClearBarcodes 
         Caption         =   "Clear"
         Height          =   372
         Left            =   3132
         TabIndex        =   19
         Top             =   288
         Width           =   840
      End
      Begin VB.ListBox lstBarcodes 
         Height          =   645
         Left            =   144
         TabIndex        =   18
         Top             =   288
         Width           =   2712
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1116
      Left            =   84
      TabIndex        =   10
      Top             =   708
      Width           =   5004
      Begin VB.Label lblSaveFileName 
         Caption         =   "Save"
         Height          =   204
         Left            =   1584
         TabIndex        =   16
         Top             =   816
         Width           =   3036
      End
      Begin VB.Label lblScanFileName 
         Caption         =   "Scan"
         Height          =   204
         Left            =   1584
         TabIndex        =   15
         Top             =   528
         Width           =   2988
      End
      Begin VB.Label lblImageFileName 
         Caption         =   "Image"
         Height          =   204
         Left            =   1584
         TabIndex        =   14
         Top             =   240
         Width           =   2988
      End
      Begin VB.Label Label3 
         Caption         =   "SaveFileName   :"
         Height          =   204
         Left            =   144
         TabIndex        =   13
         Top             =   816
         Width           =   1308
      End
      Begin VB.Label Label2 
         Caption         =   "ScanFileName   :"
         Height          =   204
         Left            =   144
         TabIndex        =   12
         Top             =   528
         Width           =   1308
      End
      Begin VB.Label Label1 
         Caption         =   "ImageFileName :"
         Height          =   204
         Left            =   144
         TabIndex        =   11
         Top             =   240
         Width           =   1308
      End
   End
   Begin VB.Frame Frame4 
      Height          =   588
      Left            =   96
      TabIndex        =   4
      Top             =   48
      Width           =   4980
      Begin VB.CommandButton cmdLastPage 
         Caption         =   ">>"
         Height          =   252
         Left            =   2856
         TabIndex        =   6
         Top             =   192
         Width           =   852
      End
      Begin VB.CommandButton cmdFirstPage 
         Caption         =   "<<"
         Height          =   252
         Left            =   72
         TabIndex        =   5
         Top             =   192
         Width           =   852
      End
      Begin VB.PictureBox vsPageNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4128
         ScaleHeight     =   285
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   144
         Width           =   492
      End
      Begin VB.PictureBox cmdNextPage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   8
         Top             =   192
         Width           =   852
      End
      Begin VB.PictureBox cmdPrevPage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   984
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   7
         Top             =   192
         Width           =   852
      End
   End
   Begin VB.PictureBox ezPageView 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6612
      Left            =   5184
      ScaleHeight     =   6555
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   3012
      Width           =   4812
   End
   Begin VB.PictureBox ezTreeView 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6612
      Left            =   192
      ScaleHeight     =   6555
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   3012
      Width           =   4572
   End
   Begin VB.Label lblPageView 
      Caption         =   "PageView"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5160
      TabIndex        =   3
      Top             =   2364
      Width           =   1332
   End
   Begin VB.Label lblTreeView 
      Caption         =   "TreeView"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2364
      Width           =   1380
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScan 
      Caption         =   "&Scan"
      Begin VB.Menu mnuScanSelectScanner 
         Caption         =   "&Select Scanner..."
      End
      Begin VB.Menu mnuScanSetScanParams 
         Caption         =   "Set Scan &Parameters..."
      End
      Begin VB.Menu mnuScanPreviewPage 
         Caption         =   "Pre&view Page"
      End
      Begin VB.Menu mnuScanScanBatch 
         Caption         =   "Scan &Batch"
         Begin VB.Menu mnuScanBatchInsert 
            Caption         =   "&Insert Into Current Batch..."
         End
         Begin VB.Menu mnuScanBatchOverwrite 
            Caption         =   "&Overwrite Current Batch"
         End
         Begin VB.Menu mnuScanBatchCreate 
            Caption         =   "Create &New Batch..."
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmEzScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private boolSaveDocument As Boolean
Private Const DEFLT_SCANFILE = "prevue.tif"
Const DebugMode = False
    ' True - Debug_Print messages go to debug window
    ' false - puts up a message box
    ' set this to false while building a stand alone executable
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


'*********************
' Form event callbacks
'*********************

Private Sub Form_Load()
    ezTreeView.DocHandle = ezPageView.DocHandle
    ezTreeView.ViewMode = 1                           'tree mode
    
    ' set imagecontrol properties so that these controls are
    ' hooked up the particular image control irrespective of
    ' who has the EnvFocus (which is default behaviour)
    
    cmdPrevPage.ImageControl = ezTreeView
    cmdNextPage.ImageControl = ezTreeView
    vsPageNo.ImageControl = ezTreeView
    
    UpdateMenuItems
    
    ezPageView.ScanFileName = MakeTempFileName(ezPageView)
    On Error Resume Next
    Kill ezPageView.ScanFileName  ' delete the temp file on entry
    boolSaveDocument = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ezPageView.Scanning Or (mnuFile.Enabled = False) Then
    Cancel = 1
    MsgBox ("You must stop scanning before closing this application")
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        ezTreeView.Left = Me.width * 0.05
        ezTreeView.width = Me.width * 0.4
        If ((Me.Height - ezTreeView.Top) * 0.8) > 100 Then
            ezTreeView.Height = (Me.Height - ezTreeView.Top) * 0.8
        End If
        
        ezPageView.width = Me.width * 0.4
        ezPageView.Left = Me.width - (Me.width * 0.05) - ezPageView.width
        ezPageView.Height = ezTreeView.Height
        
        lblTreeView.Left = ezTreeView.Left
        lblPageView.Left = ezPageView.Left
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mnuFileExit_Click
End Sub

'***********************
'EzImage event callbacks
'***********************
Private Sub ezPageView_AfterScanSave()
   UpdateBarcodeList
End Sub

Private Sub ezPageView_SaveWarnOverwrite(ByVal FileName As String, ToOverwrite As Long)
   ToOverwrite = frmWarnOnOverWrite.StartForm(FileName)
End Sub

' This event is received after every ScanPage.
' We give the user the optionof
'         Save the page into a temp document
'         forget it and cancel
'         Rescan the page with a chance to change scanner options.

Private Sub ezPageView_AfterScan(nAction As Integer)
    nAction = frmPreviewPage.StartForm(ezPageView)
    
    DoEvents
    If nAction And PREVUE_PIXEZ_SCAN_SAVE Then
        boolSaveDocument = True
    End If
End Sub

Private Sub ezPageView_FeederEmpty(NumPages As Long, StopScan As Long, UseFeeder As Long, ScanBacks As Long)
    
    ezTreeView.Refresh
    ezPageView.Refresh
    
    'PixSLForm.ShowPixSLDialog ezPageView, NumPages, StopScan, UseFeeder, ScanBacks
    frmContinueBatchScan.ShowPrepareScannerDialog ezPageView, NumPages, StopScan, UseFeeder, ScanBacks
    If StopScan Then
        If frmScanBatchCancel.Visible = True Then
            frmScanBatchCancel.Hide
        End If
    End If

End Sub
Private Sub ezPageView_Yield()
    DoEvents                    'to prevent scanbatch from hogging the system.
End Sub

Private Sub ezPageView_Change()
    UpdateMenuItems
    
    If boolSaveDocument Then
        'save here if we ever want a saveas functionality
        boolSaveDocument = False
    End If
End Sub
'dynamic name change implementation
Private Sub ezPageView_ScanBatchControl(Action As Integer)
    Dim s As String
    s = Trim(Str(Int((Rnd() * 9999) + 1)) + ".tif")
    'make up a name according to requirement. e.g. It could be from BarCode value etc.
    'This sample uses a random number
    'Following line should be uncommented to see Dynamic Name changing.
    
    'ezPageView.ScanFileName = s
End Sub


'User clicked on a thumbnail page in the treeview.
'Bring that page up in the pageview.

Private Sub ezTreeView_ThumbSelect(ByVal PageNo As Long)
    ezPageView.PageIndex = PageNo
End Sub

Private Sub ezTreeView_Change()
    'sync page indices
    
    If ezPageView.PageCount > 1 Then
        If ezPageView.PageIndex <> ezTreeView.PageIndex Then
            On Error Resume Next
            ezPageView.PageIndex = ezTreeView.PageIndex
        End If
    End If

End Sub

'*************
'Menu handlers
'*************

Private Sub mnuHelpAbout_Click()
    frmAbout.StartForm ezPageView
End Sub

Private Sub mnuFileOpen_Click()
    frmOpenDialog.StartForm ezPageView
End Sub
Private Sub mnuFileClose_Click()
    ezPageView.Close
    ezPageView.ImageFileName = ""
End Sub
Private Sub mnuFileExit_Click()
    ezTreeView.Close
    ezPageView.Close
    End
End Sub


Private Sub mnuScanBatchCreate_Click()
    Dim rc As Long
    
    ezPageView.Close
    ezPageView.ScanInsertMode = 2                   ' overwrite
    
    rc = frmScannedDocument.StartForm(ezPageView)  'get the ScanFileName
    If rc = 0 Then
        Exit Sub                                   ' user pressed cancel
    End If
    
    
    frmScanBatchCancel.StartForm ezPageView
    
    mnuFile.Enabled = False
    mnuScan.Enabled = False
    
    DoEvents                                  'give time for the dialogs to draw and undraw properly
    
    On Error GoTo handler
    ezPageView.ScanBatch
    frmScanBatchCancel.Hide
    
    mnuFile.Enabled = True
    mnuScan.Enabled = True
    
    Exit Sub
    
handler:
    mnuFile.Enabled = True
    mnuScan.Enabled = True
    

    frmScanBatchCancel.Hide
    Debug_Print "ScanBatchCreate ", Err
End Sub
'Insert batch into an existing image
'No action if there is no open image
Private Sub mnuScanBatchInsert_Click()
    Dim rc As Long
    
    ezPageView.ScanAllowTurnover = 1               'on
    ezPageView.ScanInsertMode = 1                  'insert
    
    
    rc = frmPreInsertCurrentBatch.StartForm(ezPageView)
    If rc = 0 Then
        Exit Sub                                   ' user pressed cancel
    End If
    
    frmScanBatchCancel.StartForm ezPageView
    
    mnuFile.Enabled = False
    mnuScan.Enabled = False
    
    DoEvents                                  'give time for the dialogs to draw and undraw properly
    
    On Error GoTo handler
    ezPageView.ScanBatch
    frmScanBatchCancel.Hide
    
    mnuFile.Enabled = True
    mnuScan.Enabled = True
    
    Exit Sub
    
handler:
    mnuFile.Enabled = True
    mnuScan.Enabled = True

    frmScanBatchCancel.Hide
    Debug_Print "ScanBatchInsertIntoCurrentDocument ", Err
End Sub
'Overwrite batch scan into existing document
'No action if there is no open image
Private Sub mnuScanBatchOverwrite_Click()
    Dim rc As Long
    
    ezPageView.ScanInsertMode = 2
        
    rc = frmPreInsertCurrentBatch.StartForm(ezPageView)
    If rc = 0 Then
        Exit Sub                                   ' user pressed cancel
    End If
    
    frmScanBatchCancel.StartForm ezPageView
    
    mnuFile.Enabled = False
    mnuScan.Enabled = False
    
    DoEvents
    
    On Error GoTo handler
    ezPageView.ScanBatch
    frmScanBatchCancel.Hide
    
    mnuFile.Enabled = True
    mnuScan.Enabled = True
    
    Exit Sub
    
handler:
    mnuFile.Enabled = True
    mnuScan.Enabled = True

    frmScanBatchCancel.Hide
    Debug_Print "ScanBatchOverWrite ", Err

End Sub

' For previews we create a temporary file in the disk.
' This gets deleted for each session.

Private Sub mnuScanPreviewPage_Click()
    
    mnuFile.Enabled = False
    mnuScan.Enabled = False
    ezPageView.Close
    ezPageView.ImageFileName = ""
    DoEvents
    
    ezPageView.ScanFileName = MakeTempFileName(ezPageView)
    On Error GoTo handler
    ezPageView.ScanSingle
    UpdateBarcodeList
    mnuFile.Enabled = True
    mnuScan.Enabled = True
    Exit Sub
handler:
    Debug_Print "Scan PreviewPage  ", Err
    mnuFile.Enabled = True
    mnuScan.Enabled = True
End Sub
Private Sub mnuScanSelectScanner_Click()
    On Error GoTo handler
    ezPageView.ScanSelect
    Exit Sub
handler:
   Debug_Print "SelectScanner  ", Err
End Sub
Private Sub mnuScanSetScanParams_Click()
    On Error GoTo handler
    frmScannerSettings.StartForm ezPageView
    Exit Sub
handler:
    Debug_Print "Scanner Parameters", Err
End Sub

'*******************************
'Standardcontrol event callbacks
'*******************************

Private Sub cmdClearBarcodes_Click()
   lstBarcodes.Clear
End Sub

Private Sub cmdFirstPage_Click()
    ezTreeView.PageIndex = 1
End Sub
Private Sub cmdLastPage_Click()
    ezTreeView.PageIndex = ezPageView.PageCount
End Sub


'**********************************************************************
'////////////////////// HELPER FUNCTIONS //////////////////////////////
'**********************************************************************
Private Sub PrintImageValues(image As PixEzImage)
    Debug.Print "ImageFileName   :", image.ImageFileName
    Debug.Print "ImageFileRoot   :", image.ImageFileRoot
    Debug.Print "ImageFileDir    :", image.ImageFileDir
    
    Debug.Print "ScanFileName    :", image.ScanFileName; ""
    Debug.Print "ScanFileRoot    :", image.ScanFileRoot
    Debug.Print "ScanFileDir     :", image.ScanFileDir
    
    Debug.Print "SaveFileName    :", image.SaveFileName
    Debug.Print "SaveFileRoot    :", image.SaveFileRoot
    Debug.Print "SaveFileDir     :", image.SaveFileDir
    Debug.Print "---------------------------------------"
    
End Sub
Private Sub UpdateMenuItems()
    If ezPageView.PageCount > 0 Then
        mnuScanBatchInsert.Enabled = True
        mnuScanBatchOverwrite.Enabled = True
        mnuFileClose = True
    Else
        mnuScanBatchInsert.Enabled = False
        mnuScanBatchOverwrite.Enabled = False
        mnuFileClose = False
    End If
    lblImageFileName = ezPageView.ImageFileName
    lblScanFileName = ezPageView.ScanFileName
    lblSaveFileName = ezPageView.SaveFileName
End Sub

' This is a diagnostic helper function.
' Sends the Error.Description string to debug window or a message box depending
' on the value of global DebugMode variable

Public Sub Debug_Print(errString As String, errObj As Object)
    If DebugMode Then
        Debug.Print errString, errObj.Description
    Else
        MsgBox errObj.Description, vbOKOnly + vbExclamation, "Error " & Str(errObj) & " in " & errString
    End If
End Sub

' makes up a temporary filename to be used in PreViewPage function.
Public Function MakeTempFileName(image As PixEzImage) As String
    On Error Resume Next
    Dim strPath As String * 300
    Dim rc As Long
    Dim strExt As String
    
    strExt = ".tif"
    
    rc = GetTempPath(300, strPath)
    
    If (rc > 0) Then
        MakeTempFileName = Trim(Left(strPath, rc)) & "prevue" & strExt
    Else
        MakeTempFileName = "prevue" & strExt
    End If
    
End Function
Public Sub UpdateBarcodeList()
    Dim i As Integer
    
    For i = 1 To ezPageView.BarcodeCount
        On Error GoTo handler
        lstBarcodes.AddItem (ezPageView.BarcodeList(i).Text)
    Next i
    Exit Sub
handler:
    Debug_Print "Updating Barcodes", Err
End Sub

