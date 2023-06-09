VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm101DBMaint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imaging101 DB Maintenance"
   ClientHeight    =   6225
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm101DBMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   10860
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
      Left            =   9000
      Picture         =   "frm101DBMaint.frx":0442
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   13
      Top             =   0
      Width           =   1572
   End
   Begin VB.ComboBox cmbConnection 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frm101DBMaint.frx":0AD5
      Left            =   6150
      List            =   "frm101DBMaint.frx":0AE5
      TabIndex        =   8
      Top             =   600
      Width           =   4335
   End
   Begin VB.ComboBox cmbTables 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frm101DBMaint.frx":0B20
      Left            =   6150
      List            =   "frm101DBMaint.frx":0B22
      TabIndex        =   7
      Top             =   900
      Width           =   4335
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
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
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10860
      TabIndex        =   0
      Top             =   5595
      Width           =   10860
      Begin VB.CommandButton cmdRebuildIndexes 
         Caption         =   "&Re-&Build Indexes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8040
         TabIndex        =   12
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4675
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3521
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2367
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1213
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   59
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5895
      Visible         =   0   'False
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
      Caption         =   " "
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
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Bindings        =   "frm101DBMaint.frx":0B24
      Height          =   4335
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         Name            =   "Arial"
         Size            =   7.5
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
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   9000
      TabIndex        =   14
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DB Maintenance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   945
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Type of Database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   645
      Width           =   2655
   End
End
Attribute VB_Name = "frm101DBMaint"
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
            RegGenericConnectionString = RegImaging101ConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
        
        Case "Batch DB"
            RegGenericConnectionString = RegImaging101BatchListConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
        
        Case "Document Type DB"
            ' Get Database Connections settings from the registry
            RegGenericConnectionString = RegDocTypeListConnectionString
            CatDB.ActiveConnection = RegGenericConnectionString
        
        Case "Client DB"
            ' Get Database Connections settings from the registry
            RegGenericConnectionString = RegLookupListConnectionString
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
    
    datPrimaryRS.ConnectionString = RegGenericConnectionString
    
    If cmbTables.Text = "I101TableLookupFields" Then
        datPrimaryRS.RecordSource = "select * FROM " & cmbTables.Text & " ORDER BY ApplicationRECID, TableLookupRECID, DisplayOrder"
    Else
        datPrimaryRS.RecordSource = "select * FROM " & cmbTables.Text
    End If
    
    datPrimaryRS.Refresh
    
    
End Sub



Private Sub Form_Load()

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

' Nothing really special happens here!
    Me.Show
    cmdRebuildIndexes.Visible = False
    
End Sub

Function funcOpenApplicationDB(strTableName As String, strOrderByFields As String)
'    datPrimaryRS.ConnectionString = RegImaging101ConnectionString
'    datPrimaryRS.RecordSource = "select * FROM " & strTableName & " ORDER BY " & strOrderByFields
'    datPrimaryRS.Refresh
End Function
Private Sub subOpenBatchDB(strTableName As String, strOrderByFields As String)

End Sub
Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - 30 - picButtons.Height - grdDataGrid.Top
  grdDataGrid.width = Me.ScaleWidth - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'  'This is where you would put error handling code
'  'If you want to ignore errors, comment out the next line
'  'If you want to trap them, add code here to handle them
'  MsgBox "Data error event hit err:" & Description
'End Sub
'
'Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'This will display the current record position for this recordset
'  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
'End Sub
'
'Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'This is where you put validation code
'  'This event gets called when the following actions occur
'  Dim bCancel As Boolean
'
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.MoveLast
  grdDataGrid.SetFocus
  SendKeys "{down}"

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
    
    With datPrimaryRS.Recordset
''      While Not datPrimaryRS.Recordset.EOF
''        grdDataGrid.SetFocus
''        If grdDataGrid.Bookmark = True Then
            .Delete
            .MoveNext
              
      If .EOF Then .MoveLast
    End With
  
  Exit Sub
DeleteErr:
  MsgBox "DBMaint Delete Error: " & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

End Sub

