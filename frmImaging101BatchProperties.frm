VERSION 5.00
Begin VB.Form frmImaging101BatchProperties 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Batch Properties"
   ClientHeight    =   7335
   ClientLeft      =   5790
   ClientTop       =   2850
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImaging101BatchProperties.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   7890
   Begin VB.ComboBox cmbBatchManager 
      Height          =   315
      ItemData        =   "frmImaging101BatchProperties.frx":0442
      Left            =   1800
      List            =   "frmImaging101BatchProperties.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   4440
      Width           =   3135
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
      Left            =   6480
      Picture         =   "frmImaging101BatchProperties.frx":0446
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   26
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtApplicationRECID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtBatchBoxNumber 
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdUnlockBatch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Unlock Batch"
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
      Left            =   2880
      Picture         =   "frmImaging101BatchProperties.frx":0AD9
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtBatchNotesPrevious 
      Height          =   1605
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   2040
      Width           =   5295
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
      Height          =   360
      ItemData        =   "frmImaging101BatchProperties.frx":17A3
      Left            =   1800
      List            =   "frmImaging101BatchProperties.frx":17AD
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5640
      Width           =   3135
   End
   Begin VB.ComboBox cmbBatchQueue 
      Height          =   288
      ItemData        =   "frmImaging101BatchProperties.frx":17C1
      Left            =   1800
      List            =   "frmImaging101BatchProperties.frx":17C3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3720
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
      Left            =   5400
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtBatchName 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancelChanges 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel changes"
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
      Picture         =   "frmImaging101BatchProperties.frx":17C5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1875
   End
   Begin VB.CommandButton cmdSaveChanges 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save changes"
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
      Left            =   6000
      Picture         =   "frmImaging101BatchProperties.frx":1D4F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1875
   End
   Begin VB.ComboBox cmbBatchStatus 
      Height          =   315
      ItemData        =   "frmImaging101BatchProperties.frx":2091
      Left            =   1800
      List            =   "frmImaging101BatchProperties.frx":2093
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4800
      Width           =   4215
   End
   Begin VB.ComboBox cmbBatchPriority 
      Height          =   315
      ItemData        =   "frmImaging101BatchProperties.frx":2095
      Left            =   1800
      List            =   "frmImaging101BatchProperties.frx":2097
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5160
      Width           =   2535
   End
   Begin VB.ComboBox cmbBatchOwner 
      Height          =   315
      ItemData        =   "frmImaging101BatchProperties.frx":2099
      Left            =   1800
      List            =   "frmImaging101BatchProperties.frx":209B
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox txtBatchNotes 
      Height          =   765
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox txtBatchDesc 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Route To Manager"
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
      TabIndex        =   29
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   6480
      TabIndex        =   27
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Batch Properties"
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
      TabIndex        =   24
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label38 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Box #"
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
      TabIndex        =   23
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Notes"
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
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblBatchQueue 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblBatchStatus 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblBatchPriority 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblBatchOwner 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      Top             =   6600
      Width           =   7935
   End
End
Attribute VB_Name = "frmImaging101BatchProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bolFormLoaded As Boolean
Dim intHoldBatchQueueListIndex As Integer
Dim bolListsLoaded As Boolean
Dim bolEditingBatchName As Boolean


Private Sub cmbBatchQueue_Click()
    '*** If the BatchQueue was Changed, CLEAR the BatchOwner by selecting the Blank/Top list item
    '    but first make sure the Lists have been Loaded, because this event fires for each
    '    function call to funcFindItemInComboBox
    If bolListsLoaded = True Then
        If cmbBatchQueue.ListIndex <> intHoldBatchQueueListIndex Then
            cmbBatchOwner.ListIndex = 0
        End If
    End If

End Sub

Private Sub cmbBatchQueue_DropDown()

    '*** If the BatchQueue was Changed, CLEAR the BatchOwner by selecting the Blank/Top list item
    intHoldBatchQueueListIndex = cmbBatchQueue.ListIndex

End Sub


Private Sub cmdCancelChanges_Click()
    Unload Me
End Sub

Private Sub cmdSaveChanges_Click()
    
    On Error GoTo UPDATE_BATCH_RECORD_ERROR
    
    Dim strQuickMessage As String
    
    txtActionBeforeError = "Open Batches Table, Find Batch Record."
    Dim connImaging101Batch As ADODB.Connection
    Set connImaging101Batch = New ADODB.Connection
    connImaging101Batch.Open RegImaging101BatchListConnectionString
    
    Dim rsImaging101Batch As ADODB.Recordset
    Set rsImaging101Batch = New ADODB.Recordset
    Set rsImaging101Batch.ActiveConnection = connImaging101Batch

    rsImaging101Batch.CursorType = adOpenDynamic
    rsImaging101Batch.LOCKTYPE = adLockOptimistic
    
    rsImaging101Batch.Open "Select * FROM I101Batches where BatchRECID = " & txtBatchRECID, connImaging101Batch
    
'    'Only update the Notes if not Empty
'    If Trim(txtBatchNotes) <> "" Then
'        txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & txtBatchNotes & vbCrLf & rsImaging101Batch("BatchNotes")
'
'        'Check if notes are too long
'        If Len(txtBatchNotes) > rsImaging101Batch("BatchNotes").DefinedSize Then
'            result = MsgBox("The Notes you have entered exceed the maximum allowed field size..." & _
'                                vbCrLf & "Would you like to Add your notes and cut off the oldest notes?", vbYesNo, "Notes Size Exceeded")
'            If result = vbYes Then
'                txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
'            Else
'                'Don't cut off oldest notes
'                txtBatchNotes = txtBatchNotesPrevious
'            End If
'        End If
'    Else
'        'Notes are empty
'        txtBatchNotes = txtBatchNotesPrevious
'    End If
    
    '*** FORCE RECORD LOCK - set the BatchLocked equal to itself
    txtActionBeforeError = "Force Lock"
    rsImaging101Batch.Fields("BatchLocked") = rsImaging101Batch.Fields("BatchLocked")
    rsImaging101Batch.MoveFirst
    
    txtActionBeforeError = "Assign Variables to Fields"
    
    DoEvents
    
    'Only update the Notes if not Empty
    'Commented the Code above -- Will NOT Ask user if he wants to truncate
    If Trim(txtBatchNotes) <> "" Then
            txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & txtBatchNotes & vbCrLf & rsImaging101Batch("BatchNotes")
            txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
            rsImaging101Batch("BatchNotes") = txtBatchNotes
    End If
    
    If gOpenBatchInReadOnlyMode = False Then
        'Only update these fields if Batch NOT in Read-Only mode
        rsImaging101Batch("BatchName") = txtBatchName   ''txtBatchPrefix & txtBatchName & txtBatchSuffix
        rsImaging101Batch("BatchDesc") = txtBatchDesc
        rsImaging101Batch("BatchBoxNumber") = txtBatchBoxNumber
                
    End If
    
    Dim strBatchQueueCurrent As String
    Dim strBatchOwnerCurrent As String
    Dim strBatchManagerCurrent As String
    
    strBatchQueueCurrent = rsImaging101Batch("BatchQueue") & ""
    strBatchOwnerCurrent = rsImaging101Batch("BatchOwner") & ""
    strBatchManagerCurrent = rsImaging101Batch("BatchManager") & ""
    
    rsImaging101Batch("BatchStatus") = cmbBatchStatus
    rsImaging101Batch("BatchPriority") = cmbBatchPriority
    rsImaging101Batch("BatchGroup") = cmbBatchGroup
    
    If (Trim(cmbBatchQueue) <> Trim(strBatchQueueCurrent)) _
    Or (Trim(cmbBatchOwner) <> Trim(strBatchOwnerCurrent)) _
    Or (Trim(cmbBatchManager) <> Trim(strBatchManagerCurrent)) _
    Then
        'Initialize the Batch Route Count if null
        Dim txtBatchRouteCount As String
        txtBatchRouteCount = rsImaging101Batch("BatchRouteCount") & ""
        
        If txtBatchRouteCount = "" Then
            rsImaging101Batch("BatchRouteCount") = 0
        End If
        
        rsImaging101Batch("BatchRouteCount") = rsImaging101Batch("BatchRouteCount") + 1
        rsImaging101Batch("BatchInQueueDate") = Now()
        
        'Check How many times this Batch has been Routed.
        ' If the current user is NOT a Batch Administrator, Route the Batch to the users' Supervisor
        Dim intRouteMaxCount As Integer
        intRouteMaxCount = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & rsImaging101Batch("ApplicationRECID"), "RouteMaxCount")
        
        If rsImaging101Batch("BatchRouteCount") > intRouteMaxCount _
        And gsecRightsBatchAdministration <> vbChecked Then
            strQuickMessage = "  Attempted Route #: " & rsImaging101Batch("BatchRouteCount") & "  Queue: " & cmbBatchQueue & "  UserID: " & cmbBatchOwner
            strQuickMessage = strQuickMessage & vbCrLf & "  Batch '" & txtBatchName & "' has been routed more than " & rsImaging101Batch("BatchRouteCount") & " times... " & vbCrLf & "  it has been re-routed to your Supervisor '" & gsecUserSupervisor & "' !"
            funcQuickMessage "SHOW", strQuickMessage
            'Leave the Queue as it is
            rsImaging101Batch("BatchOwner") = gsecUserSupervisor
            'Add a Note and make sure it does NOT exceed the defined field size
            txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & strQuickMessage & vbCrLf & rsImaging101Batch("BatchNotes")
            txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
            rsImaging101Batch("BatchNotes") = txtBatchNotes
        Else
            strQuickMessage = "  Route #: " & rsImaging101Batch("BatchRouteCount") & "  Queue: " & cmbBatchQueue & "  UserID: " & cmbBatchOwner
            txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & strQuickMessage & vbCrLf & rsImaging101Batch("BatchNotes")
            txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
            rsImaging101Batch("BatchQueue") = cmbBatchQueue
            rsImaging101Batch("BatchOwner") = cmbBatchOwner
            rsImaging101Batch("BatchManager") = cmbBatchManager
            rsImaging101Batch("BatchNotes") = txtBatchNotes
        End If
    End If
        
        
        
    
    txtActionBeforeError = "Update Values"
    rsImaging101Batch.Update
    
    'Close connection and the recordset
    rsImaging101Batch.Close
    Set rsImaging101Batch = Nothing
    connImaging101Batch.Close
    Set connImaging101Batch = Nothing


    '*** CREATE BATCH AUDIT RECORD
    funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Edit Properties"
    
    
    Screen.MousePointer = vbDefault

    Unload Me
    
Exit Sub
    
UPDATE_BATCH_RECORD_ERROR:
    MsgBox "UPDATE_BATCH_RECORD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") ", vbExclamation
     
    'Close connection and the recordset
    On Error Resume Next
    rsImaging101Batch.Close
    Set rsImaging101Batch = Nothing
    connImaging101Batch.Close
    Set connImaging101Batch = Nothing
       

End Sub

Private Sub cmdUnlockBatch_Click()

    'Force Batch Unlock
    subUnLockBatch
    Unload Me
    
End Sub



Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    bolListsLoaded = False
    
    '*** LOAD UserID's
    
    frmImaging101BatchList.Hide
    
    txtActionBeforeError = "Connect to Imaging101 DB"
    
    txtApplicationRECID = frmImaging101BatchList.txtApplicationRECID
    
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
'''''''    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    On Error GoTo FORM_LOAD_ERROR
    
    '***************************************
    '*** LOAD USERS LIST DROP-DOWN
    
    txtActionBeforeError = "Populate UserID List"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select UserName from I101Security ORDER BY UserName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    rs.Open
    
    
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
    
    
    '***************************************
    '*** LOAD Batch Managers LIST DROP-DOWN
    
    txtActionBeforeError = "Populate BatchManager List"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select DISTINCT BatchManager from I101Batches ORDER BY BatchManager"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    rs.Open
    
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbBatchOwner.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchManager.AddItem rs.Fields("BatchManager") & ""
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
              
    txtActionBeforeError = "Populate BatchManager List"
        
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
    '****************************

    '***************************************
    '*** LOAD BATCH STATUS LIST DROP-DOWN
        
    txtActionBeforeError = "Populate BatchStatus List"
        
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
        
    txtActionBeforeError = "Populate BatchPriority List"
        
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



    '**********************************************************
    '*** ADD SPECIAL OPTIONS TO cmbBatchGroup LIST DROP-DOWN
        
    Select Case frmImaging101BatchList.cmbApplicationList.Text
        Case "TTC"
            cmbBatchGroup.AddItem "TTC PRINTED"
            cmbBatchGroup.AddItem "TTC RECEIVED"
    End Select




    '****************************
    '***  Bring In Key fields
    
    txtActionBeforeError = "Open Batches Table, Find Batch Record."
    
    Dim connImaging101Batch As ADODB.Connection
    Set connImaging101Batch = New ADODB.Connection
    connImaging101Batch.Open RegImaging101BatchListConnectionString
    
    Dim rsImaging101Batch As ADODB.Recordset
    Set rsImaging101Batch = New ADODB.Recordset
    Set rsImaging101Batch.ActiveConnection = connImaging101Batch

    

    rsImaging101Batch.CursorType = adOpenDynamic
    rsImaging101Batch.LOCKTYPE = adLockOptimistic
    rsImaging101Batch.Open "Select * FROM I101Batches where BatchRECID = " & frmImaging101BatchList.txtBatchRECID, connImaging101Batch
    
    txtBatchName = rsImaging101Batch("BatchName")
    txtBatchRECID = rsImaging101Batch("BatchRECID")
    txtBatchDesc = rsImaging101Batch("BatchDesc") & ""
    txtBatchNotesPrevious = rsImaging101Batch("BatchNotes") & ""
    txtBatchBoxNumber = rsImaging101Batch("BatchBoxNumber") & ""
    
    '*** 8/6/2009 - Jacob - Changed logic so that instead of dissabling ALL fields if in
    '***                    Read-only mode... now handle more granularly based on Route or Index rights
    
    'Set Field Default States
    txtBatchName.Enabled = False
    txtBatchDesc.Enabled = False
    txtBatchNotes.Enabled = True
    cmbBatchQueue.Enabled = False
    cmbBatchOwner.Enabled = False
    cmbBatchStatus.Enabled = False
    cmbBatchPriority.Enabled = False
    cmbBatchGroup.Enabled = False
    txtBatchBoxNumber.Enabled = False
    
    'Allow Batch Administrators to UNLOCK the Current Batch
    If (gsecRightsBatchAdministration = vbChecked) And ((rsImaging101Batch("BatchLocked") & "") = "Y") Then
        cmdUnlockBatch.Visible = True
    Else
        cmdUnlockBatch.Visible = False
    End If
    
    
        
    If (gsecRightsBatchRoute = True) Or (gsecRightsBatchScan = True) Then
        txtBatchName.Enabled = True
        txtBatchDesc.Enabled = True
        lblBatchOwner.Enabled = True
        cmbBatchOwner.Enabled = True
        lblBatchQueue.Enabled = True
        cmbBatchQueue.Enabled = True
        lblBatchPriority.Enabled = True
        cmbBatchPriority.Enabled = True
        lblBatchStatus.Enabled = True
        cmbBatchStatus.Enabled = True
        cmbBatchGroup.Enabled = True
    End If
    
    If gsecRightsBatchIndex = True Then
        txtBatchName.Enabled = True
        txtBatchDesc.Enabled = True
        lblBatchPriority.Enabled = True
        cmbBatchPriority.Enabled = True
        lblBatchStatus.Enabled = True
        cmbBatchStatus.Enabled = True
        cmbBatchGroup.Enabled = True
        txtBatchBoxNumber.Enabled = True
    End If
    

    txtActionBeforeError = "Find and Select Items in Combo Boxes"

    funcFindItemInComboBox Me.cmbBatchStatus, frmImaging101BatchList.txtBatchStatus
    funcFindItemInComboBox Me.cmbBatchPriority, frmImaging101BatchList.txtBatchPriority
    funcFindItemInComboBox Me.cmbBatchOwner, frmImaging101BatchList.txtBatchOwner
    funcFindItemInComboBox Me.cmbBatchManager, frmImaging101BatchList.txtBatchManager
    funcFindItemInComboBox Me.cmbBatchQueue, frmImaging101BatchList.txtBatchQueue
    funcFindItemInComboBox Me.cmbBatchGroup, frmImaging101BatchList.txtBatchGroup
    
    '*** Hold the Queue ListIndex to CLEAR the BatchOwner If the BatchQueue is Changed
    intHoldBatchQueueListIndex = cmbBatchQueue.ListIndex
    
    
    'If Batch is in Read-Only Mode, DO NOT Allow chaning these fields
    'Override other security settings
    If (gOpenBatchInReadOnlyMode = True) _
    Or (InStr(frmImaging101BatchList.txtBatchCommitStatus, "-FULL") > 0) _
    Then
        txtActionBeforeError = "Batch is in Read-Only Mode or Commited/Split FULL, DO NOT Allow chaning these fields"
        txtBatchName.Enabled = False
        txtBatchDesc.Enabled = False
        txtBatchBoxNumber.Enabled = False
    End If
    
    
    
    bolListsLoaded = True
    
    'Close connection and the recordset
    rsImaging101Batch.Close
    Set rsImaging101Batch = Nothing
    connImaging101Batch.Close
    Set connImaging101Batch = Nothing
    bolFormLoaded = True

Exit Sub
    
FORM_LOAD_ERROR:
        MsgBox "FORM_LOAD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") ", vbExclamation
        

End Sub

Private Sub subSaveBatchProperties()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
'    If gOpenBatchInReadOnlyMode = True Then
'        gOpenBatchInReadOnlyMode = False
'    End If
    
    'Refresh the Calling Form to reflect changes
    Select Case txtCurrentModule
        Case "frmImaging101BatchList"
            ' Only Unlock the Batch if NOT in Read-Only mode
            If gOpenBatchInReadOnlyMode = False Then
                ' UNLOCK the Batch
                strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)
        
                If Left(strReturn, 5) = "ERROR" Then
'                    strBatchModeOption = MsgBox(strReturn & vbCrLf & vbCrLf & " Error Unlocking the Batch.", vbOK, "Error Locking Batch")
                End If
            End If
            frmImaging101BatchList.Show
            frmImaging101BatchList.subListBatches
            frmImaging101BatchList.SetFocus
        Case "frmIndex"
            'Don't UNLOCK because the user is still IN the Batch
            frmIndex.txtBatchQueue = cmbBatchQueue
            frmIndex.txtBatchStatus = cmbBatchStatus
            frmIndex.txtBatchGroup = cmbBatchGroup
            frmIndex.CheckBatchOpenMode
    End Select
    
    Set frmImaging101BatchProperties = Nothing


End Sub

Private Sub subUnLockBatch()

        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        
        Set conn = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rs = New ADODB.Recordset
        
'        On Error GoTo ErrLock
    
        txtActionBeforeError = "Prepare to Open Batch DB Connection"
        '*** Prepare Connection
        With conn
            .ConnectionString = RegImaging101ConnectionString
            .CursorLocation = adUseServer
            .ConnectionTimeout = 120
            .IsolationLevel = adXactReadCommitted
            .mode = adModeReadWrite
            txtActionBeforeError = "Open Batch DB Connection"
            .Open
            .Execute "SET LOCK_TIMEOUT -1"
        End With
        
        Set cmd.ActiveConnection = conn

        '*** Begin Transaction
        conn.BeginTrans


        conn.Execute "UPDATE I101Batches " & _
            " SET  BatchLocked = null, BatchLockedBy=null, BatchLockedDate = null " & _
            " WHERE BatchRECID = " & txtBatchRECID

        conn.CommitTrans

        conn.Errors.Clear
        
        conn.Close
        
        Set rs = Nothing
        Set cmd = Nothing
        Set conn = Nothing
        
Exit Sub

ErrLock:

    If conn.Errors.item(0).NativeError = 1222 Then  ' Lock Timeout
        conn.RollbackTrans
        MsgBox "Sorry... I was unable to Unlock this batch due to a Database Timeout!", vbCritical, "Unlock Batch Timeout"
        conn.Errors.Clear
        Set rs = Nothing
        Set cmd = Nothing
        Set conn = Nothing
        Exit Sub
    End If
    

End Sub



Private Sub txtBatchName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'2004-04-02 Jacob Russo - Changed from MsgBox to InputBox to try to correct a problem at CCCS
'                         where sometimes Clicking OK does nothing!
'2004-04-07 Jacob Russo - Changed from InputBox to the NEW Form frmYesNo
'                         with function funcYesNo
'                         The InputBox had the same problem.
'2010-08-03 Jacob Russo - Added  bolEditingBatchName to prevent the question after
'                         the user answered it the first time.


    'Check to make sure the Form is completelly loaded.
    If bolFormLoaded = True And bolEditingBatchName = False Then
        
        bolEditingBatchName = True
        
        gYesNo = vbNo
        frmYesNo.lblYesNoMessage = "Are you sure you want to Re-name this Batch?"
        frmYesNo.Top = Me.Top
        frmYesNo.Left = Me.Left
        frmYesNo.Show vbModal, Me
        
        
        If gYesNo <> vbYes Then
            txtBatchDesc.SetFocus
            Exit Sub
        Else
            txtBatchName.SetFocus
            txtBatchName.SelStart = 0
            txtBatchName.SelLength = Len(txtBatchName)
        End If
    End If
End Sub

Private Sub txtBatchName_Validate(Cancel As Boolean)

    bolEditingBatchName = False
    
End Sub
