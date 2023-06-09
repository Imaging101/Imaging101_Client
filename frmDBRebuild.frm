VERSION 5.00
Begin VB.Form frm101DBRebuild 
   BackColor       =   &H000040C0&
   Caption         =   "Imaging101 DB Rebuild Utility"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatusMessage 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2880
      Width           =   10335
   End
   Begin VB.ComboBox cmbTables 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.ComboBox cmbConnection 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmDBRebuild.frx":0000
      Left            =   6240
      List            =   "frmDBRebuild.frx":0010
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton cmdRebuildIndexes 
      Caption         =   "Re-&Build Indexes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Type of Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DB Re-Build Utility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frm101DBRebuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    ' Set a GENERIC Connection String variable to handle ALL Select Statements
    Dim RegGenericConnectionString As String

Private Sub cmbConnection_Click()
    cmbTables.Clear
    
    Dim CatDB As ADOX.Catalog
    Set CatDB = New ADOX.Catalog
    
    
    'open the database
    Select Case cmbConnection.Text
        Case "Application DB"
            ' Get Database Connections settings from the registry
            On Error Resume Next
            RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
            RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
            On Error GoTo 0

            RegGenericConnectionString = RegImaging101ConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
        
        Case "Batch DB"
            On Error Resume Next
            RegBatchListConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionType", RegFileName)
            RegImaging101BatchListConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionString." & RegBatchListConnectionType, RegFileName)
            On Error GoTo 0
            RegGenericConnectionString = RegImaging101BatchListConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
        
        Case "Document Type DB"
            ' Get Database Connections settings from the registry
            On Error Resume Next
            RegDocTypeListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionType", RegFileName)
            RegDocTypeListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcDocTypeList.ConnectionString." & RegDocTypeListConnectionType, RegFileName)
            On Error GoTo 0
            RegGenericConnectionString = RegDocTypeListConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
        
        Case "Client DB"
            ' Get Database Connections settings from the registry
            On Error Resume Next
            RegClientListConnectionType = VBGetPrivateProfileString("DATABASE", "AdodcLookupList.ConnectionType", RegFileName)
            RegClientListConnectionString = VBGetPrivateProfileString("DATABASE", "AdodcLookupList.ConnectionString." & RegClientListConnectionType, RegFileName)
            On Error GoTo 0
            RegGenericConnectionString = RegClientListConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
    
    End Select
    
    'Fill Tables Combo
    For inttableindex = 0 To CatDB.Tables.Count - 1
        ' Show User TABLES ONLY... Don't show System Tables or Queries
        If CatDB.Tables.item(inttableindex).Type = "TABLE" Then
            cmbTables.AddItem CatDB.Tables.item(inttableindex).name
        End If
    Next
    
    Set CatDB = Nothing
    
    Me.Caption = cmbConnection.Text & " - Imaging101 DB Maintenance"
    
End Sub

Private Sub cmbTables_Click()
    On Error Resume Next
        Set datPrimaryRS = Nothing
    On Error GoTo 0
    
    Me.Caption = cmbTables & "->" & cmbConnection.Text & "- Imaging101 DB Maintenance"
    
    If cmbConnection.Text = "Batch DB" Then
        If cmbTables.Text = "Batches" Then
                cmdRebuildIndexes.Visible = True
        End If
        If Right(cmbTables.Text, 8) = "_BatchPage" Then
                cmdRebuildIndexes.Visible = True
        End If
    End If
    
    
End Sub

Private Sub cmdRebuildIndexes_Click()
    If cmbConnection.Text = "Batch DB" Then
        ' Batches Table
        If cmbTables.Text = "Batches" Then
            funcRebuildIndexes "BatchRECID", cmbTables.Text
            funcRebuildIndexes "ApplicationRECID", cmbTables.Text
            funcRebuildIndexes "BatchName", cmbTables.Text
            funcRebuildIndexes "BatchPriority", cmbTables.Text
            funcRebuildIndexes "BatchStatus", cmbTables.Text
        End If
        ' Batch Page Tables
        If Left(cmbTables.Text, 8) = "_BatchPage" Then
            funcRebuildIndexes "BatchRECID", cmbTables.Text
            funcRebuildIndexes "BatchPageRECID", cmbTables.Text
            funcRebuildIndexes "BatchPageOrder", cmbTables.Text
            funcRebuildIndexes "BatchPageIndexed", cmbTables.Text
            funcRebuildIndexes "BatchPageStatus", cmbTables.Text
        End If
    End If
End Sub
Function funcRebuildIndexes(strFieldName As String, strTableName As String)
    
''    On Error GoTo REBUILD_ERROR
    
    ' Create the INDEXes for each Field by adding an "I" to the end of the fieldname
    
            '*** Declarations
            Dim rsb As ADODB.Recordset
            Dim Conb As ADODB.Connection
            Dim ssqlb As String
        
            '*** Set Object Types
            Set Conb = New ADODB.Connection
            Set rsb = New ADODB.Recordset
                    
            '*** Set Connection Modes
            Conb.Mode = adModeReadWrite
            
            '*** Set Lock Types
            rsb.LockType = adLockPessimistic
                    
''            On Error GoTo REBUILD_ERROR
    
    ssqlb = "------ BEGIN Rebuild Field: " & strFieldName & " On Table: " & strTableName & " ------"
    txtStatusMessage = txtStatusMessage & Now & ": " & ssqlb & vbCrLf
    DoEvents
    
            '*** OPEN Connections
            On Error Resume Next
            Conb.Open RegImaging101BatchListConnectionString
            If Err.Number <> 0 Then
                ssqlb = "I Can't open the Database in Exclusive Mode!  Trying to Re-Build the Index for Field (" & strFieldName & ") TableName (" & strTableName & ") -- [ " & Err.Description & " ] Please make sure NO ONE has this DB Open."
                MsgBox ssqlb
                txtStatusMessage = txtStatusMessage & Now & ": " & ssqlb & vbCrLf
                DoEvents
                ssqlb = ""
                Exit Function
            End If
                    
            'Begin SQL Transaction to make sure is doesn't zap the record if we
            '  get an error while deleting the DB TABLE
''            Conb.BeginTrans
                    
                'sql statement
                On Error Resume Next
                
                'Drop OLD Style Index
                ssqlb = "DROP INDEX " & strFieldName & " ON " & strTableName
                txtStatusMessage = txtStatusMessage & Now & ": " & ssqlb & vbCrLf
                DoEvents
                rsb.Open ssqlb, Conb
               
               'Drop New Style Index
                ssqlb = "DROP INDEX " & strFieldName & "I" & " ON " & strTableName
                txtStatusMessage = txtStatusMessage & Now & ": " & ssqlb & vbCrLf
                DoEvents
                rsb.Open ssqlb, Conb
                
                On Error GoTo REBUILD_ERROR
                
                ssqlb = "CREATE INDEX " & strFieldName & "I" & " ON " & strTableName & " (" & strFieldName & ")"
                txtStatusMessage = txtStatusMessage & Now & ": " & ssqlb & vbCrLf
                DoEvents
                rsb.Open ssqlb, Conb
            
''            Conb.CommitTrans
        
            'Close connection and the recordset
            'Again... for some reason the rs was already closed???
        '    rs.Close
            Set rsb = Nothing
''            Con.Close
            Set Conb = Nothing
            
    ssqlb = "------ END Rebuild Field: " & strFieldName & " On Table: " & strTableName & " ------"
    txtStatusMessage = txtStatusMessage & Now & ": " & ssqlb & vbCrLf
    DoEvents

    Exit Function

REBUILD_ERROR:
    
    MsgBox "RE-BUILD INDEXES ERROR: " & Err.Number & " - " & Err.Description & "[ TRANSACTION ROLLED BACK ]"
''    If Conb.State = adStateOpen Then
''        Conb.RollbackTrans
''        Conb.Close
''    End If
''    Set rsb = Nothing
''    Set Conb = Nothing
    Resume Next

End Function

Private Sub Form_Load()
    cmdRebuildIndexes.Visible = False
    
End Sub
