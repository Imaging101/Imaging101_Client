VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEcaptureBatchList 
   BackColor       =   &H00008000&
   Caption         =   "eCapture Batch List"
   ClientHeight    =   5325
   ClientLeft      =   330
   ClientTop       =   1125
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   11145
   Begin VB.TextBox txtImaging101BatchRootDir 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3960
      Width           =   5295
   End
   Begin VB.TextBox txtPagesImported 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   10080
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   8880
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdImportEcaptureBatch 
      Caption         =   "&Import Batch"
      Default         =   -1  'True
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtEcaptureBatchRootDir 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   5295
   End
   Begin VB.TextBox txtecID 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtecFileCabinetID 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdBatchListDataGrid 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcEcaptureBatchList 
      Height          =   375
      Left            =   -120
      Top             =   4320
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
      Left            =   3120
      Top             =   4320
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
      Left            =   5760
      Top             =   4320
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages Imported"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblApplicationRECID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblApplicationName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Importing into Application:  "
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmEcaptureBatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImportEcaptureBatch_Click()
    
''Print grdBatchListDataGrid.Columns(0).Caption
''ecFileCabinetID
        ' Get Database Connections settings from the registry
        
        Screen.MousePointer = vbHourglass
        txtPagesImported = 0
                
'''''''        On Error Resume Next
'''''''        RegImaging101BatchListConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionType", RegFileName)
'''''''        RegImaging101BatchListConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionString." & RegImaging101BatchListConnectionType, RegFileName)
'''''''        txtBatchRootDir = VBGetPrivateProfileString("IMAGING101", "frmImaging101BatchList.txtBatchRootDir", RegFileName)
'''''''        On Error GoTo 0
        
''''        On Error GoTo BATCH_IMPORT_ERRORS


        ' Establish the Imaging101 Batch List Connection
        txtActionBeforeError = "Establish the Imaging101 Batch List Connection"
        Dim conn As ADODB.Connection
        Dim cmd As ADODB.Command
        Set conn = New ADODB.Connection
        Set cmd = New ADODB.Command
        conn.ConnectionString = RegImaging101BatchListConnectionString
        conn.ConnectionTimeout = 60
        conn.Mode = adModeReadWrite
        conn.Open
        Set cmd.ActiveConnection = conn
        
        ' Establish the Captovation eCapture Batch DB Connection
        txtActionBeforeError = "Establish the Captovation eCapture Batch DB Connection"
        Dim connb As ADODB.Connection
        Dim cmdb As ADODB.Command
        Set connb = New ADODB.Connection
        Set cmdb = New ADODB.Command
        connb.ConnectionString = RegEcaptureBatchListConnectionString
        connb.ConnectionTimeout = 60
        connb.Mode = adModeReadWrite
        connb.Open
        Set cmdb.ActiveConnection = connb

        'User Transaction Tracking to prevent partial imports!
        conn.BeginTrans
        connb.BeginTrans
        
        '***  BEGIN:  BUILD ECAPTURE VALUES LIST
        txtActionBeforeError = "BUILD ECAPTURE VALUES LIST"
        
        strEcaptureBatchFields = ""
            
         strEcaptureBatchFields = strEcaptureBatchFields & "'" & lblApplicationRECID & "', "
        
        For intColumnsLoop = 0 To grdBatchListDataGrid.Columns.Count - 1
            strEcaptureBatchFields = strEcaptureBatchFields & "'" & grdBatchListDataGrid.Columns(intColumnsLoop) & "', "
        Next
    
        'Add Batch Directory
        strEcaptureBatchFields = strEcaptureBatchFields & "'" & txtImaging101BatchRootDir & "\" & grdBatchListDataGrid.Columns(1) & "', "
        'Add Generic Description
        strEcaptureBatchFields = strEcaptureBatchFields & "'Imported from eCapture',"
        'Add Total Page Count
        strEcaptureBatchFields = strEcaptureBatchFields & "'" & Int(grdBatchListDataGrid.Columns(3)) + Int(grdBatchListDataGrid.Columns(4)) & "', "
        'Add Scan User
        strEcaptureBatchFields = strEcaptureBatchFields & "'" & gsecUserID & "', "
        
        ' GET NEXT Batch #
        txtBatchRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchRECID")

        'Add BatchRECID
        strEcaptureBatchFields = strEcaptureBatchFields & "'" & txtBatchRECID & "'"

    '***  END:  BUILD ECAPTURE VALUES LIST
    
    strBatchFields = "ApplicationRECID, BatchApplication, BatchName, BatchDesc, BatchPagesCommitted, BatchPagesNotCommitted, BatchStatus, BatchPriority, BatchScanDate, BatchDirectory, BatchNotes, BatchPagesTotal, BatchScanUser, BatchRECID "
                    'ecFileCabinetID , ecName, ecID, ecCommittedPageCount, ecUncommittedPageCount, ecStatus, ecPriority, ecdatetime
    
    cmd.CommandText = "INSERT INTO Batches (" & strBatchFields & ") VALUES (" & strEcaptureBatchFields & ")"
    
    
    ''''  $$$$$  TRAP FOR ERRORS WITH DUPLICATE ENTRIES $$$$$
    ''''  $$$$$  CHECK FIRST OR TRANSACTION ROLLBACK $$$$$
    
    
    txtActionBeforeError = "INSERT eCapture BATCH HEADER into Imaging101"
    cmd.Execute , , adCmdText
    
    
    '*** ADD BATCH PAGES  ***
    For intLoop = 0 To List1.ListCount - 1
        '* Using List1 to get the file Extensions in correct Sorted order
        txtBatchPageFileName = txtecID & "." & Format(List1.List(intLoop), "###")
        txtBatchPageRECID = funcGetNextControlNumber(RegImaging101BatchListConnectionString, "I101Control", "BatchPageRECID")
        cmd.CommandText = "INSERT INTO " & lblApplicationName & "_BatchPage" & " (BatchPageRECID, BatchRECID, BatchPageFileName, BatchPageOrder) VALUES ('" & txtBatchPageRECID & "', '" & txtBatchRECID & "', '" & txtBatchPageFileName & "', '" & intLoop + 1 & "')"
        
        txtActionBeforeError = "INSERT Page" & txtBatchPageFileName & " into Imaging101"
        
        cmd.Execute , , adCmdText
        
        
        '*** BEGIN - Move Batch Pages to the Imaging101 Batch Directory
        
        
            ' CREATE the Directory Structure for storing this Batch
            funcCreateDirectoryStructure txtImaging101BatchRootDir & "\" & grdBatchListDataGrid.Columns(1)
            
            FileCopy txtEcaptureBatchRootDir & "\" & txtecFileCabinetID & "\" & txtecID & "\" & txtBatchPageFileName, txtImaging101BatchRootDir & "\" & grdBatchListDataGrid.Columns(1) & "\" & txtBatchPageFileName
        
        
        
        
        '*** END - Move Batch Pages
        
        
        txtPagesImported = txtPagesImported + 1
        DoEvents
    Next

    ' *** UPDATE THE eCapture Batch Record - Set the ecStatus to "Imported to Imaging101"
    '       using the ecID from the DBGRID
''    Dim lngecID As Long
''    lngecID = ConvertBase(grdBatchListDataGrid.Text, ebBase36, ebDecimal, , , 0)

    cmdb.CommandText = "UPDATE ecBatches SET ecStatus = 'Imported to Imaging101' WHERE ecID = " & grdBatchListDataGrid.Columns(2).Text

'''' MsgBox cmdb.CommandText, vbOKOnly
    
    txtActionBeforeError = "UPDATE ecBatches SET ecStatus = 'Imported to Imaging101'"
    
    cmdb.Execute , , adCmdText
    
    

    ' Commit the successful transactions
    txtActionBeforeError = "COMMIT ALL TRANSACTIONS"
    conn.CommitTrans
    connb.CommitTrans
    
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    
    connb.Close
    Set cmdb = Nothing
    Set connb = Nothing
    
    Screen.MousePointer = vbDefault

    MsgBox "Import of Batch '" & grdBatchListDataGrid.Columns(1) & "' SUCCESSFUL!", vbOKOnly
    
    
    Exit Sub
    
BATCH_IMPORT_ERRORS:
        MsgBox "eCapture IMPORT Error: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch NOT Imported]", vbExclamation
        
        If conn.BeginTrans = True Then
            conn.RollbackTrans
        End If
        If conn.BeginTrans = True Then
            conn.RollbackTrans
        End If
        Screen.MousePointer = vbDefault

    
End Sub



Private Sub Form_Activate()
    ' Get Batch Root Directory setting from the registry
    On Error Resume Next
    Me.txtEcaptureBatchRootDir = VBGetPrivateProfileString(RegAppname, "frmEcaptureBatchList.txtBatchRootDir", RegFileName)
    Me.txtImaging101BatchRootDir = VBGetPrivateProfileString(RegAppname, "frmImaging101BatchList.txtBatchRootDir", RegFileName)
    On Error GoTo 0
    
    lblApplicationName = frmImaging101BatchList.cmbApplicationList.Text
    lblApplicationRECID = frmImaging101BatchList.cmbApplicationList.ItemData(frmImaging101BatchList.cmbApplicationList.ListIndex)
    
End Sub

Private Sub Form_Load()
    Dim RegConnectString As String
    Dim RegImaging101ConnectionType As String
    
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmEcaptureBatchList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmEcaptureBatchList.Left", RegFileName)
    Me.Width = VBGetPrivateProfileString(RegAppname, "frmEcaptureBatchList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmEcaptureBatchList.Height", RegFileName)
''    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmEcaptureBatchList.Caption", RegFileName)
    On Error GoTo 0
    
'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegEcaptureBatchListConnectionType = VBGetPrivateProfileString("DATABASE", "frmEcaptureBatchList.AdodcEcaptureBatchList.ConnectionType", RegFileName)
'''''''    RegEcaptureBatchListConnectionString = VBGetPrivateProfileString("DATABASE", "frmEcaptureBatchList.AdodcEcaptureBatchList.ConnectionString." & RegEcaptureBatchListConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    AdodcEcaptureBatchList.ConnectionString = RegEcaptureBatchListConnectionString
    AdodcEcaptureBatchList.LockType = adLockOptimistic
    AdodcEcaptureBatchList.Mode = adModeRead
    AdodcEcaptureBatchList.CursorType = adOpenForwardOnly
    AdodcEcaptureBatchList.RecordSource = "SELECT ecFileCabinetID, ecName, ecID, ecCommittedPageCount, ecUncommittedPageCount, ecStatus, ecPriority, ecdatetime  FROM ecBatches WHERE ecName LIKE '" & Left(frmImaging101BatchList.cmbApplicationList.Text, 2) & "%' ORDER BY ecName"
    ''''      ecStatus <> 'Imported to Imaging101'
    
    ' The "Set" line is to prevent error "non-nullable column cannot be updated to Null".
    '  must empty the DataSource property of the db grid at design time
    Set Me.grdBatchListDataGrid.DataSource = Me.AdodcEcaptureBatchList
    AdodcEcaptureBatchList.Refresh
    
    '*** Set SQL wildcard string
'    RegConnectionWildcard = "%"
        
    
    '*** Connect to DB for Drop Down Lists
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmEcaptureBatchList.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmEcaptureBatchList.Left", Me.Left, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmEcaptureBatchList.Width", Me.Width, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmEcaptureBatchList.Height", Me.Height, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmEcaptureBatchList.Caption", Me.Caption, RegFileName)
    End If
    frmImaging101BatchList.Show
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  Frame1.Width = Me.ScaleWidth
  grdBatchListDataGrid.Height = Me.ScaleHeight - 2600 - Frame1.Height
  grdBatchListDataGrid.Width = Me.Width - 300
  txtFullPathName.Top = Me.ScaleHeight - 300
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub


Private Sub grdBatchListDataGrid_Click()
    On Error Resume Next
    grdBatchListDataGrid.Col = 0
    txtecFileCabinetID = ConvertBase(grdBatchListDataGrid.Text, ebDecimal, ebBase36, , , 3)
    txtecFileCabinetID = Format(txtecFileCabinetID, "0##")
    grdBatchListDataGrid.Col = 2
    ' Use the BASECONV module to convert the eCapture Batch ID (ecID)
    '   from Decimal to Base36 padded to 8 characters
    txtecID = ConvertBase(grdBatchListDataGrid.Text, ebDecimal, ebBase36, , , 8)
End Sub

Private Sub txtecID_Change()
        
    
    If txtecFileCabinetID <> "" And txtecID <> "" Then
    
        '* DEBUG MessageBox
        result = MsgBox("Ready to Import: [" & txtEcaptureBatchRootDir & "\" & txtecFileCabinetID & "\" & txtecID & "] to [" & txtImaging101BatchRootDir & "]", vbOKCancel)
        If result = vbCancel Then
            Exit Sub
        End If
        
        On Error GoTo BATCH_SELECT_ERRORS
        
        txtActionBeforeError = "Assign File1.Path"
        File1.Path = txtEcaptureBatchRootDir & "\" & txtecFileCabinetID & "\" & txtecID
        txtActionBeforeError = "Assign File1.Pattern"
        File1.Pattern = "0*.*"
        txtActionBeforeError = "File1.Refresh"
        File1.Refresh
        
        List1.Clear
        
        txtActionBeforeError = "Create File List"
        For intLoop = 0 To File1.ListCount - 1
            List1.AddItem Format(Right(File1.List(intLoop), Len(File1.List(intLoop)) - InStrRev(File1.List(intLoop), ".")), "0##")
        Next
    End If
    
Exit Sub

BATCH_SELECT_ERRORS:
        MsgBox "eCapture SELECT Error: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch NOT Imported]", vbExclamation
        
        Screen.MousePointer = vbDefault
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    ' Set the ListView to Sort ListView by column clicked
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Do the sort now
    ListView1.Sorted = True
End Sub


