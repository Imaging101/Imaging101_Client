Attribute VB_Name = "basDBofADO"
' DBofADO.bas
'
Option Explicit

Public Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
    
Public gFileSpec As String               ' Filespec of MDB
Public gTableName As String              ' Table name of selected MDB
Public gstrFields() As String
Public gstrFieldsOrig() As String

Public gcdg As Object

Public gAcnn As adodb.Connection
Public gstrCNN As String



Sub DBFilesMDBproc()
    On Error GoTo errhandler
    
    ' Obtain gFileSpec
    Dim i As Integer
    If GetFileSpec("(*.mdb)|*.mdb") = True Then
         If UCase(Right(gFileSpec, 4)) <> ".MDB" Then
             MsgBox "Please select a .MDB file"
             Exit Sub
         End If
         
         Set gAcnn = New adodb.Connection
         gAcnn.CursorLocation = adUseClient
         gstrCNN = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & gFileSpec & ";"

            ' Only gAcnn, not gRcnn
         If (gAcnn.Errors.Count > 0) Then
             ' Just Display The First Error In The Collection
            MsgBox "Error: " & gAcnn.Errors(0).Description, _
                 0, "Connect Error!"
            Exit Sub
         End If
         
         frmDBofADO.Show
    End If
    Exit Sub
    
  ' Provided a way to exit, if error occurred in called form
  ' forcing it to be closed
errhandler:
    ErrMsgProc "basMain DBFilesMDBProc"
End Sub




Function GetFileSpec(ByVal strFilter As String) As Boolean
    On Error GoTo errhandler
 
    Dim tmpfile As String
    tmpfile = gFileSpec
   
    Do
        frmFrame.CommonDialog1.CancelError = True
        frmFrame.CommonDialog1.FileName = tmpfile
        frmFrame.CommonDialog1.Filter = strFilter
        frmFrame.CommonDialog1.ShowOpen
        
        If frmFrame.CommonDialog1.FileName = "" Then
            Exit Do
        End If
    
        tmpfile = frmFrame.CommonDialog1.FileName
        
        If IsFileThere(tmpfile) Then
            Exit Do
        End If
        
        MsgBox "File specification not found.  Please re-try"
    Loop
    
    If tmpfile <> "" Then
        gFileSpec = tmpfile
        GetFileSpec = True
    Else
        GetFileSpec = False
    End If
    
    Exit Function
    
errhandler:
   GetFileSpec = False
   If Err.Number <> 32755 Then
       ErrMsgProc "basMain GetFileSpec"
   End If
End Function



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub



' Convert the numeric value returned by DB to Enum, so
' that at least the user could have a guess of what it is.
Function ConvType(ByVal TypeVal As Long) As String
  Select Case TypeVal
      Case adBigInt                    ' 20
         ConvType = "adBigInt"
      Case adBinary                    ' 128
         ConvType = "adBinary"
      Case adBoolean                   ' 11
         ConvType = "adBoolean"
      Case adBSTR                      ' 8 i.e. null terminated string
         ConvType = "adBSTR"
      Case adChar                      ' 129
         ConvType = "adChar"
      Case adCurrency                  ' 6
         ConvType = "adCurrency"
      Case adDate                      ' 7
         ConvType = "adDate"
      Case adDBDate                    ' 133
         ConvType = "adDBDate"
      Case adDBTime                    ' 134
         ConvType = "adDBTime"
      Case adDBTimeStamp               ' 135
         ConvType = "adDBTimeStamp"
      Case adDecimal                   ' 14
         ConvType = "adDecimal"
      Case adDouble                    ' 5
         ConvType = "adDouble"
      Case adEmpty                     ' 0
         ConvType = "adEmpty"
      Case adError                     ' 10
         ConvType = "adError"
      Case adGUID                      ' 72
         ConvType = "adGUID"
      Case adIDispatch                 ' 9
         ConvType = "adIDispatch"
      Case adInteger                   ' 3
         ConvType = "adInteger"
      Case adIUnknown                  ' 13
         ConvType = "adIUnknown"
      Case adLongVarBinary             ' 205
         ConvType = "adLongVarBinary"
      Case adLongVarChar               ' 201
         ConvType = "adLongVarChar"
      Case adLongVarWChar              ' 203
         ConvType = "adLongVarWChar"
      Case adNumeric                  ' 131
         ConvType = "adNumeric"
      Case adSingle                    ' 4
         ConvType = "adSingle"
      Case adSmallInt                  ' 2
         ConvType = "adSmallInt"
      Case adTinyInt                   ' 16
         ConvType = "adTinyInt"
      Case adUnsignedBigInt            ' 21
         ConvType = "adUnsignedBigInt"
      Case adUnsignedInt               ' 19
         ConvType = "adUnsignedInt"
      Case adUnsignedSmallInt          ' 18
         ConvType = "adUnsignedSmallInt"
      Case adUnsignedTinyInt           ' 17
         ConvType = "adUnsignedTinyInt"
      Case adUserDefined               ' 132
         ConvType = "adUserDefined"
      Case adVarBinary                 ' 204
         ConvType = "adVarBinary"
      Case adVarChar                   ' 200
         ConvType = "adVarChar"
      Case adVariant                   ' 12
         ConvType = "adVariant"
      Case adVarWChar                  ' 202
         ConvType = "adVarWChar"
      Case adWChar                     ' 130
         ConvType = "adWChar"
   End Select
End Function



Function ConvAttr(ByVal mAttr As Long) As String
    Dim tmp As String
    tmp = ""
    If (mAttr And adFldMayDefer) Then
        tmp = tmp & "adFldMayDefer "             '2
    End If
    If (mAttr And adFldUpdatable) Then
        tmp = tmp & "adFldUpdatable "            '4
    End If
    If (mAttr And adFldUnknownUpdatable) Then
        tmp = tmp & "adFldUnknownUpdatable "     '8
    End If
    If (mAttr And adFldFixed) Then
        tmp = tmp & "adFldFixed "           '16
    End If
    If (mAttr And adFldIsNullable) Then
        tmp = tmp & "adFldIsNullable "      '32
    End If
    If (mAttr And adFldMayBeNull) Then
        tmp = tmp & "adFldMayBeNull "       '64
    End If
    If (mAttr And adFldLong) Then
        tmp = tmp & "adFldLong "            '128
    End If
    If (mAttr And adFldRowID) Then
        tmp = tmp & "adFldRowID "           '256
    End If
    If (mAttr And adFldRowVersion) Then
       tmp = tmp & "adFldRowVersion "       '512
    End If
    If (mAttr And adFldCacheDeferred) Then
        tmp = tmp & "adFldCacheDeferred "   '4096
    End If
    If tmp = "" Then
        tmp = "Unknown"
    End If
    ConvAttr = tmp
End Function



Function ConvLockType(ByVal mLockType) As String
    Select Case mLockType
       Case (mLockType And adLockReadOnly)
           ConvLockType = "adLockReadOnly"           ' 1
       Case (mLockType And adLockPessimistic)
           ConvLockType = "adLockPessimistic"        ' 2
       Case (mLockType And adLockOptimistic)
           ConvLockType = "adLockOptimistic"         ' 3
       Case (mLockType And adLockBatchOptimistic)
           ConvLockType = "adLockBatchOptimistic"    ' 4
       Case Else
           ConvLockType = "(Unknown)"
    End Select
End Function



Function ConvEditMode(ByVal mEditMode) As String
    Select Case mEditMode
       Case (mEditMode And adEditNone)
           ConvEditMode = "adEditNone"               ' 0
       Case (mEditMode And adEditInProgress)
           ConvEditMode = "adEditInProgress"         ' 1
       Case (mEditMode And adEditAdd)
           ConvEditMode = "adEditAdd"                ' 2
       Case Else
           ConvEditMode = "(Unknown)"
    End Select
End Function




Function ConvState(ByVal mState) As String
    Select Case mState
       Case (mState And adStateClosed)
           ConvState = "adStateClosed"           ' 0, default
       Case (mState And adStateOpen)
           ConvState = "adStateOpen"             '
       Case (mState And adStateConnecting)
           ConvState = "adStateConnecting"
       Case (mState And adStateExecuting)
           ConvState = "adStateExecuting"
       Case (mState And adStateFetching)
           ConvState = "adStateFetching"
       Case Else
           ConvState = "(Unknown)"
    End Select
End Function



'Returns a sum of one or more of the RecordStatusEnum values.
'Use the Status property to see what changes are pending for records
'modified during batch updating. You can also use the Status property
'to view the status of records that fail during bulk operations, such
'as when you call the Resync, UpdateBatch, or CancelBatch methods on
'a Recordset object, or set the Filter property on a Recordset object
'to an array of bookmarks. With this property, you can determine how
'a given record failed and resolve it accordingly.
Function ConvStatus(ByVal mStatus) As String
    ' Because one or more values can be present, accumulate the string
    Dim tmp As String
    tmp = ""
    Select Case mStatus
       Case (mStatus And adRecOK)
          ConvStatus = "adRecOK"           ' 0 Record was successfully update
       Case (mStatus And adRecNew)
          ConvStatus = "adRecNew"          ' 1 Is new
       Case (mStatus And adRecModified)
          ConvStatus = "adRecModified"     ' 2 Was modified.
       Case (mStatus And adRecDeleted)
          ConvStatus = "adRecDeleted"      ' 4 Was deleted.
       Case (mStatus And adRecUnmodified)
          ConvStatus = "adRecUnmodified"   ' 8 Was not modified.
       Case (mStatus And adRecInvalid)
          ConvStatus = "adRecInvalid"      ' 16 Was not saved because its bookmark is invalid.
       Case (mStatus And adRecMultipleChanges)
          ConvStatus = "adRecMultipleChanges"  ' 64 Not saved because it would have affected multiple records.
       Case (mStatus And adRecPendingChanges)
          ConvStatus = "adRecPendingChanges"   ' 128 Was not saved because it refers to a pending insert.
       Case (mStatus And adRecCanceled)
          ConvStatus = "adRecCanceled"         ' 256 Was not saved because the operation was canceled.
       Case (mStatus And adRecCantRelease)
          ConvStatus = "adRecCantRelease"      ' 1024 Was not saved because of existing record locks.
       Case (mStatus And adRecConcurrencyViolation)
          ConvStatus = "adRecConcurrencyViolation"   ' 2048 Was not saved because optimistic concurrency was in use.
       Case (mStatus And adRecIntegrityViolation)
          ConvStatus = "adRecIntegrityViolation"     ' 4096 Was not saved because the user violated integrity constraints.
       Case (mStatus And adRecMaxChangesExceeded)
          ConvStatus = "adRecMaxChangesExceeded"     ' 8192 Was not saved because there were too many pending changes.
       Case (mStatus And adRecObjectOpen)
          ConvStatus = "adRecObjectOpen"             ' 16384 Was not saved because of a conflict with an open storage object.
       Case (mStatus And adRecOutOfMemory)
          ConvStatus = "adRecOutOfMemory"            ' 32768 Was not saved because the computer has run out of memory.
       Case (mStatus And adRecPermissionDenied)
          ConvStatus = "adRecPermissionDenied"       ' 65536 Was not saved because the user has insufficient permissions.
       Case (mStatus And adRecSchemaViolation)
          ConvStatus = "adRecSchemaViolation"        ' 131072 Was not saved because it violates structure of underlying database.
       Case (mStatus And adRecDBDeleted)
          ConvStatus = "adRecDBDeleted"              ' 262144 The record has already been deleted from the data source.
       Case Else
          ConvStatus = "A combination of serveral status present"
   End Select
End Function



Function ConvCursorType(ByVal mCursorType) As String
    Select Case mCursorType
       Case (mCursorType And adOpenForwardOnly)
           ConvCursorType = "adOpenForwardOnly"      ' 0
       Case (mCursorType And adOpenKeyset)
           ConvCursorType = "adOpenKeyset"           ' 1
       Case (mCursorType And adOpenDynamic)
           ConvCursorType = "adOpenKynamic"          ' 2
       Case (mCursorType And adOpenStatic)
           ConvCursorType = "adOpenStatic"           ' 3
       Case Else
           ConvCursorType = "(Unknown)"
    End Select
End Function




Function IsFldTypeUnAllowedForSort(ByVal inType As Long) As Boolean
     ' We disallow these types of fields, adBSTR (null-terminated string), adBinary,
     ' adVarBinary, adLongBinary(205) for OLE object
     ' (adLongBinary 6 is to be allowed, e.g. for currency.  adLongVarChar 201 should
     ' be allowed when for memo type)
    Const ExclFldTypes = "XX8/128/204/205"
    Dim s As String
    Dim inS As String
    inS = LTrim(Trim(CStr(inType)))
    If Len(inS) = 1 Then
         s = "XX" & inS
    ElseIf Len(inS) = 2 Then
         s = "X" & inS
    Else
         s = inS
    End If
    IsFldTypeUnAllowedForSort = (InStr(ExclFldTypes, s) = 0)
End Function





Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    Dim mFile As String
    mFile = LongToShort(inFileSpec)
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function



Private Function LongToShort(inSpec) As String
    Dim i
    Dim ShortSpec As String
    Dim mBuffer As String
    Dim mBufLen As Long
    mBufLen = 164
    mBuffer = String(mBufLen, 0)
    i = GetShortPathName(inSpec, mBuffer, mBufLen)
    LongToShort = Left$(mBuffer, i)
End Function




