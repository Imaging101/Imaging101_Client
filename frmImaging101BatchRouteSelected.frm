VERSION 5.00
Begin VB.Form frmImaging101BatchRouteSelected 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Batch Route Selected"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
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
      Left            =   6000
      Picture         =   "frmImaging101BatchRouteSelected.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   25
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtApplicationRECID 
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
      Left            =   6120
      TabIndex        =   23
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbApplicationList 
      Height          =   288
      Left            =   2760
      TabIndex        =   21
      Top             =   1320
      Width           =   4335
   End
   Begin VB.ComboBox cmbBatchGroup 
      Height          =   288
      ItemData        =   "frmImaging101BatchRouteSelected.frx":0693
      Left            =   2760
      List            =   "frmImaging101BatchRouteSelected.frx":06A0
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Routing Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   7455
      Begin VB.TextBox txtBatchRECID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBatchDesc 
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
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtBatchName 
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
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label20 
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
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbBatchQueue 
      Height          =   288
      ItemData        =   "frmImaging101BatchRouteSelected.frx":06B6
      Left            =   2760
      List            =   "frmImaging101BatchRouteSelected.frx":06B8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancelChanges 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel Routing"
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
      Picture         =   "frmImaging101BatchRouteSelected.frx":06BA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1875
   End
   Begin VB.CommandButton cmdRouteSelectedBatches 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Route Batches"
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
      Left            =   5520
      Picture         =   "frmImaging101BatchRouteSelected.frx":0C44
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1875
   End
   Begin VB.ComboBox cmbBatchStatus 
      Height          =   288
      ItemData        =   "frmImaging101BatchRouteSelected.frx":0F86
      Left            =   2760
      List            =   "frmImaging101BatchRouteSelected.frx":0F88
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   3375
   End
   Begin VB.ComboBox cmbBatchPriority 
      Height          =   288
      ItemData        =   "frmImaging101BatchRouteSelected.frx":0F8A
      Left            =   2760
      List            =   "frmImaging101BatchRouteSelected.frx":0F8C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   2535
   End
   Begin VB.ComboBox cmbBatchOwner 
      Height          =   288
      ItemData        =   "frmImaging101BatchRouteSelected.frx":0F8E
      Left            =   2760
      List            =   "frmImaging101BatchRouteSelected.frx":0F90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   6000
      TabIndex        =   26
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   " Only change the fields that are necessary."
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
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label lblApplicationList 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Route To Application"
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
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "R O U T E   S E L E C T E D   B A T C H E S"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label22 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label32 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "NOTE:  Any fields LEFT BLANK will leave all selected Batches with their EXISTING values!  "
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Width           =   7455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   4800
      Width           =   7455
   End
End
Attribute VB_Name = "frmImaging101BatchRouteSelected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFormLoaded As Boolean
Dim intHoldBatchQueueListIndex As Integer
Dim bolListsLoaded As Boolean


Private Sub cmbBatchQueue_Click()
    '*** If the BatchQueue was Changed, CLEAR the BatchOwner by selecting the Blank/Top list item
    '    but first make sure the Lists have been Loaded, because this event fires for each
    '    function call to funcFindItemInComboBox
    If bolListsLoaded = True Then
        If cmbBatchQueue.ListIndex <> intHoldBatchQueueListIndex Then
            result = MsgBox("You have changed the Batch QUEUE." & _
            vbCrLf & vbCrLf & "Would you like to CLEAR the USER/Owner" & _
            vbCrLf & "for the Selected Batches ?", vbYesNo, "Batch Queue Changed")
            If result = vbYes Then
                'Set the ListIndex to the SECOND Item to select the "* Clear User *" option
                cmbBatchOwner.ListIndex = 1
            End If
        End If
    End If

End Sub

Private Sub cmbBatchQueue_DropDown()

    '*** If the BatchQueue was Changed, CLEAR the BatchOwner by selecting the Blank/Top list item
    intHoldBatchQueueListIndex = cmbBatchQueue.ListIndex

End Sub



Private Sub cmbApplicationList_Click()

   If cmbApplicationList.Text <> frmImaging101BatchList.cmbApplicationList.Text Then
        result = MsgBox("*** WARNING ***" & vbCrLf & _
                            "Moving Batches to A Different Application" & vbCrLf & _
                            "will result in ALL Indexed FIELDS being Lost" & vbCrLf & _
                            "for the selected Batches!" & vbCrLf & _
                            "ARE YOU SURE THIS IS WHAT YOU WISH TO DO?", vbYesNo, "Move Batches to Different Application")
        If result = vbNo Then
            cmbApplicationList.Text = frmImaging101BatchList.cmbApplicationList.Text
            Exit Sub
        End If
    End If
    
    ' Get the Application to Commit Batches to
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
        
    rs.Source = "Select ApplicationRECID from I101Applications WHERE ApplicationName= '" & cmbApplicationList.Text & "'"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    
    If Not (rs.EOF Or rs.BOF) Then
        txtApplicationRECID = rs!ApplicationRECID
    End If
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
End Sub

Private Sub cmdCancelChanges_Click()
    Unload Me
End Sub

Private Sub subUpdateBatchRecord()

    On Error GoTo UPDATE_BATCH_RECORD_ERROR
    
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
    
    txtActionBeforeError = "Assign Variables to Fields"
    
    DoEvents
    
    Dim strQuickMessage As String

    'Initialize the Batch Routed flag
    Dim bolBatchRouted As Boolean
    bolBatchRouted = False
    
    If cmbApplicationList.Text <> frmImaging101BatchList.cmbApplicationList.Text _
    Or Trim(cmbBatchOwner) <> "" _
    Or Trim(cmbBatchQueue) <> "" Then
        bolBatchRouted = True
    End If
    
    'BEGIN Transaction
    connImaging101Batch.BeginTrans
    
    If bolBatchRouted = True Then
    
            'Initialize the Batch Route Count if null
            Dim txtBatchRouteCount As String
            txtBatchRouteCount = rsImaging101Batch("BatchRouteCount") & ""
            
            If txtBatchRouteCount = "" Then
                rsImaging101Batch("BatchRouteCount") = 0
            End If
            
            rsImaging101Batch("BatchRouteCount") = rsImaging101Batch("BatchRouteCount") + 1
            rsImaging101Batch("BatchInQueueDate") = Now()

            'Check if this is the THIRD time Routing and user is NOT a Batch Administrator
            If rsImaging101Batch("BatchRouteCount") > 3 _
            And gsecRightsBatchAdministration <> vbChecked Then
                strQuickMessage = "  Attempted Route #: " & rsImaging101Batch("BatchRouteCount") & "  Queue: " & cmbBatchQueue & "  UserID: " & cmbBatchOwner
                strQuickMessage = strQuickMessage & vbCrLf & "  Batch '" & txtBatchName & "' has been routed more than " & rsImaging101Batch("BatchRouteCount") & " times... " & vbCrLf & "  it has been re-routed to your Supervisor '" & gsecUserSupervisor & "' !"
                funcQuickMessage "SHOW", strQuickMessage
                'DO NOT CHANGE the BatchQueue
                rsImaging101Batch("BatchOwner") = gsecUserSupervisor
                'Add a Note and make sure it does NOT exceed the defined field size
                txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & strQuickMessage & vbCrLf & rsImaging101Batch("BatchNotes")
                txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
                rsImaging101Batch("BatchNotes") = txtBatchNotes
            Else
                strQuickMessage = "  Route #: " & rsImaging101Batch("BatchRouteCount") & "  Queue: " & cmbBatchQueue & "  UserID: " & cmbBatchOwner
                txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & strQuickMessage & vbCrLf & rsImaging101Batch("BatchNotes")
                txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
        
                If Trim(cmbBatchOwner) = "* Clear User *" Then
                    rsImaging101Batch("BatchOwner") = ""
                    bolBatchRouted = True
                Else
                    If Trim(cmbBatchOwner) <> "" Then
                        rsImaging101Batch("BatchOwner") = cmbBatchOwner
                        bolBatchRouted = True
                    End If
                End If
            
                If Trim(cmbBatchQueue) <> "" Then
                    rsImaging101Batch("BatchQueue") = cmbBatchQueue
                End If
                
                rsImaging101Batch("BatchNotes") = txtBatchNotes
            End If

    End If
    
    'Check if the Batch is being MOVED to another Application
    If cmbApplicationList.Text <> frmImaging101BatchList.cmbApplicationList.Text Then
        Dim strSQL As String
        
        'Change the Application RECID
        rsImaging101Batch("ApplicationRECID") = txtApplicationRECID
        
        
        'Prepare the Record Set for the Batch MOVE
        Dim rsImaging101BatchMove As ADODB.Recordset
        Set rsImaging101BatchMove = New ADODB.Recordset
        Set rsImaging101BatchMove.ActiveConnection = connImaging101Batch
    
        rsImaging101BatchMove.CursorType = adOpenDynamic
        rsImaging101BatchMove.LOCKTYPE = adLockOptimistic
        
        'Copy the BatchPage Records from the Source to the Destination Application
        strSQL = ""
        strSQL = strSQL & "INSERT into " & cmbApplicationList.Text & "_BatchPage"
        strSQL = strSQL & " (BatchRECID, BatchPageRECID, BatchPageFileName, BatchPageOrder, BatchPageIndexed, "
        strSQL = strSQL & "   BatchPageIsSeparator, BatchPageNote, BatchDocDesc, BatchPageStatus, BatchPageCommitDate, "
        strSQL = strSQL & "   BatchPageCommitUser, BatchPageQCDate, BatchPageQCUser, BatchPageIndexDate, "
        strSQL = strSQL & "   BatchPageIndexUser, BatchPagePageCount) "
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & "  BatchRECID, BatchPageRECID, BatchPageFileName, BatchPageOrder, BatchPageIndexed, "
        strSQL = strSQL & "  BatchPageIsSeparator, BatchPageNote, BatchDocDesc, BatchPageStatus, BatchPageCommitDate, "
        strSQL = strSQL & "  BatchPageCommitUser, BatchPageQCDate, BatchPageQCUser, BatchPageIndexDate, "
        strSQL = strSQL & "  BatchPageIndexUser , BatchPagePageCount "
        strSQL = strSQL & " FROM " & frmImaging101BatchList.cmbApplicationList.Text & "_BatchPage "
        strSQL = strSQL & " WHERE BatchRECID = " & txtBatchRECID
    
        txtActionBeforeError = "Move BatchPage Records to Another Application: " & strSQL
        
        rsImaging101BatchMove.Open strSQL, connImaging101Batch
        'No Need to Close the rs, the INSERT Command automatically Closes it after execution
        
        'DELETE the BatchPage Records from the Source Application
        strSQL = ""
        strSQL = strSQL & "DELETE FROM " & frmImaging101BatchList.cmbApplicationList.Text & "_BatchPage"
        strSQL = strSQL & " WHERE BatchRECID = " & txtBatchRECID
        
        txtActionBeforeError = "Delete BatchPage Records from Source Application: " & strSQL
        rsImaging101BatchMove.Open strSQL, connImaging101Batch
        'No Need to Close the rs, the DELETE Command automatically Closes it after execution
        
        'Add NOTE stating the MOVE was completed
        strQuickMessage = "  MOVED BATCH from Application: " & frmImaging101BatchList.cmbApplicationList.Text & "  to  " & cmbApplicationList.Text
        txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & strQuickMessage & vbCrLf & rsImaging101Batch("BatchNotes")
        txtBatchNotes = Left(txtBatchNotes, rsImaging101Batch("BatchNotes").DefinedSize - 1)
        rsImaging101Batch("BatchNotes") = txtBatchNotes
    End If
    
    If Trim(cmbBatchStatus) <> "" Then
        rsImaging101Batch("BatchStatus") = cmbBatchStatus
    End If
    
    If Trim(cmbBatchPriority) <> "" Then
        rsImaging101Batch("BatchPriority") = cmbBatchPriority
    End If
    
    If Trim(cmbBatchGroup) <> "" Then
        rsImaging101Batch("BatchGroup") = cmbBatchGroup
    End If

    txtActionBeforeError = "Update Values"
    rsImaging101Batch.Update
    
    'COMMIT Transaction
    connImaging101Batch.CommitTrans
    
    

    '*** CREATE BATCH AUDIT RECORD
    funcCreateBatchAuditRecord RegImaging101BatchListConnectionString, gsecUserName, txtBatchRECID, "Route Selected Batches"


Exit Sub
    
UPDATE_BATCH_RECORD_ERROR:
        funcQuickMessage "SHOW", "UPDATE_BATCH_RECORD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - TRANSACTION ROLLED BACK."
        On Error Resume Next
        connImaging101Batch.RollbackTrans

End Sub

Private Sub cmdRouteSelectedBatches_Click()

    result = MsgBox("Are you SURE you wish to ROUTE the SELECTED Batches?", vbYesNo)

    If result = vbYes Then
    
        For i = 1 To frmImaging101BatchList.ListView1.ListItems.Count
            If frmImaging101BatchList.ListView1.ListItems(i).Selected = True Then
                Debug.Print frmImaging101BatchList.ListView1.ListItems(i).Text
                frmImaging101BatchList.ListView1.ListItems(i).Selected = True   ' Force item selection
                ' Select the Batch to commit
                frmImaging101BatchList.subGetBatchHeaderInfo
                
                    'Bring In Key fields for Status Display
                    txtBatchName = frmImaging101BatchList.txtBatchName
                    txtBatchRECID = frmImaging101BatchList.txtBatchRECID
                    txtBatchDesc = frmImaging101BatchList.txtBatchDesc
    
                    ' LOCK the Batch
                    strReturn = frmImaging101Winsock.funcSendData("LOCK BATCH" & "|" & txtBatchRECID)
                    
                    If Left(strReturn, 5) = "ERROR" Then
                        MsgBox strReturn & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Server Communication Failure"
'                        Exit Sub
                    Else
                        '*** Update the appropriate Fields in the Batch Record
                        '       ONLY if there were NO Errors.
                        subUpdateBatchRecord
                        
                        ' UNLOCK the Batch
                        strReturn = frmImaging101Winsock.funcSendData("UNLOCK BATCH" & "|" & txtBatchRECID)
                        
                        If Left(strReturn, 5) = "ERROR" Then
'                            MsgBox strReturn & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Server Communication Failure"
                        End If
                    End If
    
            End If
        Next
    
    End If
    
    Screen.MousePointer = vbDefault

    frmImaging101BatchList.SetFocus
        ' Refresh the Batch List
    frmImaging101BatchList.subListBatches

    Unload Me
    
End Sub




Private Sub Form_Load()

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    bolListsLoaded = False

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
    '*** LOAD APPLICATION LIST DROP-DOWN
    
    txtActionBeforeError = "Populate Application List"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
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
    

    
    '***************************************
    '*** LOAD USERS LIST DROP-DOWN
    
    txtActionBeforeError = "Populate UserID List"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101Security ORDER BY UserName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    rs.Open
    
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbBatchOwner.AddItem ""
    cmbBatchOwner.AddItem "* Clear User *"
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchOwner.AddItem rs.Fields("UserName")
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
        
    txtActionBeforeError = "Populate Batch Queues List"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "SELECT * from I101BatchQueues " & _
                " WHERE ApplicationRECID = " & frmImaging101BatchList.txtApplicationRECID & _
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
        
    txtActionBeforeError = "Populate Batch Status List"
    
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
    cmbBatchQueue.AddItem ""
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
        
    txtActionBeforeError = "Populate Batch Priority List"
    
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
    cmbBatchQueue.AddItem ""
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
    
    '*** Set the Default Value for the Batch Group List
    'cmbBatchGroup.Text = "REGULAR"
    
    

    '**********************************************************
    '*** ADD SPECIAL OPTIONS TO cmbBatchGroup LIST DROP-DOWN
        
    Select Case frmImaging101BatchList.cmbApplicationList.Text
        Case "TTC"
            cmbBatchGroup.AddItem "TTC PRINTED"
            cmbBatchGroup.AddItem "TTC RECEIVED"
    End Select


    
    '************************************************************************
    '*** Enable the ApplicationList ONLY if user has BatchRoute Rights
    cmbApplicationList.Text = frmImaging101BatchList.cmbApplicationList.Text
    txtApplicationRECID.Text = frmImaging101BatchList.txtApplicationRECID.Text
    If gsecRightsBatchRoute = vbChecked Then
        lblApplicationList.enabled = True
        cmbApplicationList.enabled = True
    Else
        lblApplicationList.enabled = False
        cmbApplicationList.enabled = False
    End If
    
    '*** Hold the Queue ListIndex to CLEAR the BatchOwner If the BatchQueue is Changed
    intHoldBatchQueueListIndex = cmbBatchQueue.ListIndex
    
    bolListsLoaded = True
    
    
    strFormLoaded = True
    
Exit Sub
    
FORM_LOAD_ERROR:
        MsgBox "FORM_LOAD_ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") ", vbExclamation
        

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmImaging101BatchRouteSelected = Nothing

End Sub


