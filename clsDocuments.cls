VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Private DoingRequery As Boolean
Public Event MoveComplete()


Private Sub Class_Initialize()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
''  db.Open "PROVIDER=MSDASQL;dsn=Imaging101s;uid=sa;pwd=sa;database=Documents;"
''  db.Open "dsn=Imaging101s;uid=sa;pwd=sa;database=Imaging101;"
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Password=;Data Source=C:\WorkArea\Jacob\Program\Imaging101\Imaging101.MDB"

  
  Set adoControlRS = New Recordset
  adoControlRS.Open "select BatchID, DocumentID From Control Where ID = 1", db, adOpenStatic, adLockOptimistic
  DataMembers.Add "Control"

  
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select BatchID, DocumentID , UNCFilePath, Filename, Fileroom, Filecabinet, Folder, FolderDescription, DocumentDate, DocumentType, DocumentSubType, DateAdded, DocumentExpireDate, DocumentNote, Field7, Field8, Field9, Field10, Field11, Field12, Field13, Field14, Field15, Field16, Field17, Field18, Field19, Field20 from Documents Order by Fileroom", db, adOpenStatic, adLockOptimistic
  DataMembers.Add "Primary"
    
  
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Select Case DataMember
  Case "Primary"
    Set Data = adoPrimaryRS
  End Select
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  RaiseEvent MoveComplete
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Public Property Get EditingRecord() As Boolean
  EditingRecord = (adoPrimaryRS.EditMode <> adEditNone)
End Property

Public Property Get AbsolutePosition() As Long
  AbsolutePosition = adoPrimaryRS.AbsolutePosition
End Property

Public Sub AddNew()

 '         adoPrimaryRS.AddNew
          
        'These assignments where necessary because VB
        '  displays errors if we try to assign the
        '  form fields directly to the adoPrimaryRS fields.
 '       BatchID = frmImport.txtBatchID
 '       DocumentID = ""
 '       UNCFilePath = frmImport.txtUNCFilePath
 '       FileName = frmImport.txtImageFileName
 '       Fileroom = frmImport.txtFileRoom
 '       Filecabinet = frmImport.txtFileCabinet
 '       DateAdded = frmImport.txtDateAdded
 '       DocumentDate = frmImport.txtDocumentDate
 '       DocumentType = frmImport.txtDocumentType
'        DocumentSubType = frmImport.txtDocumentSubType
 '
'        DocumentExpireDate = frmImport.txtDocumentExpireDate
'        DocumentNote = frmImport.txtDocumentNote
        
'        adoPrimaryRS!BatchID = BatchID
'        adoPrimaryRS!DocumentID = DocumentID
'        adoPrimaryRS!UNCFilePath = UNCFilePath
'        adoPrimaryRS!FileName = FileName
'        adoPrimaryRS!Fileroom = Fileroom
'        adoPrimaryRS!Filecabinet = Filecabinet
'        adoPrimaryRS!DateAdded = DateAdded
'        adoPrimaryRS!DocumentDate = DocumentDate
'        adoPrimaryRS!DocumentType = DocumentType
'        adoPrimaryRS!DocumentSubType = DocumentSubType
'        adoPrimaryRS!DocumentExpireDate = DocumentExpireDate
'        adoPrimaryRS!DocumentNote = DocumentNote
'    adoPrimaryRS.Update
    
'    frmDocumentList.grdDataGrid.Columns.Item(0).Visible = False
'    frmDocumentList.grdDataGrid.Columns.Item(1).Visible = False
'    frmDocumentList.grdDataGrid.Columns.Item(2).Visible = False
'    frmDocumentList.grdDataGrid.Columns.Item(3).Visible = False
    
'    txtFileSource = UNCFilePath & "\" & FileName
'    txtFileDestination = txtRootDirToStoreObjects & "\" &
'    FileCopy
    
'    booRecordAdded = True

End Sub

Public Sub Delete()
  adoPrimaryRS.Delete
    MoveNext
End Sub

Public Sub Requery()
    

End Sub

Public Sub Resort(ByVal ColIndex As Integer)
    
End Sub

Public Sub Update()
  With adoPrimaryRS
    .UpdateBatch adAffectAll
    If .EditMode = adEditAdd Then
      MoveLast
    End If
  End With
End Sub

Public Sub Cancel()
  With adoPrimaryRS
    .CancelUpdate
    If .EditMode = adEditAdd Then
      MoveFirst
    End If
  End With
End Sub

Public Sub MoveFirst()
  adoPrimaryRS.MoveFirst
End Sub

Public Sub MoveLast()
  adoPrimaryRS.MoveLast
End Sub

Public Sub MoveNext()
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
End Sub

Public Sub MovePrevious()
  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
End Sub

