Attribute VB_Name = "FunctionDeclarations"
' LISTVIEW
Public Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Public Const LVM_SETCOLUMNORDERARRAY = &H1000 + 58
Public Const LVM_GETCOLUMNORDERARRAY = &H1000 + 59
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
        
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal msg As Long, ByVal wParam As Long, _
                ByVal lparam As Long) As Long
                
'*** timeGetTime returns time in Milliseconds
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
        
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" ( _
        ByVal lpszPath As String, _
        ByVal lpPrefixString As String, _
        ByVal wUnique As Long, _
        ByVal lpTempFileName As String) As Long


'*** STAY ON TOP Constants
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'*** BROWSE FOR DIRECTORY FUNCTION DECLARATIONS
Private Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lparam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

'*** 2020-04-14 - Jacob - Added the PathMatchSpecW Windows API call to Match the FilePattern more accurately
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long

'*** 2021-02-12 - Jacob - Added Copy Files Using the Win32 API to try to Speed  Up Indexing Commits
'                                              Implemented this via the APIFileCopy() Function
Private Declare Function CopyFile Lib "kernel32" _
  Alias "CopyFileA" (ByVal lpExistingFileName As String, _
  ByVal lpNewFileName As String, ByVal bFailIfExists As Long) _
  As Long
  
 
    Dim strLogPath As String
    Dim strTempDir As String
    

    
        




                
Function Availability(iAvailable, MenuName)
    '******* DUMMY FUNCTION TO PREVENT THE SPICER ERROR
    '*******   ON FUNCTION CALL:   Availability iAvailable, MenuName
End Function
Function ScaleX(intSize, Millimeters, Pixels)
    '******* DUMMY FUNCTION TO PREVENT THE funcListView_SetColumnWidth ERROR
    
End Function
Function TextWidth(txt)
    '******* DUMMY FUNCTION TO PREVENT THE funcListView_SetColumnWidth ERROR

End Function

Function funcDirectoryExists(strDirectoryNameToCheck) As Boolean

    '*** 2022-12-01 - Jacob - Added On Error Resume Next and funcDirectoryExists = False as default
    On Error Resume Next
    
    funcDirectoryExists = False
    
    Const ATTR_DIRECTORY = 16
    
    If Dir$(strDirectoryNameToCheck, ATTR_DIRECTORY) <> "" Then
        funcDirectoryExists = True
    Else
        funcDirectoryExists = False
    End If

End Function

        

Function funcFileExists(strFileNameToCheck As String) As Boolean
        'Dir returns the first file name that matches pathname.
        'When no more file names match, Dir returns a zero-length string ("").
        'Once a zero-length string is returned,
        'you must specify pathname in subsequent calls or an error occurs
        
        '*** 2021-03-16 - Jacob - Added check for Empty String sent, because Dir() returns the first file if string is empty.
        Dim strFileNameFound As String
        
        If strFileNameToCheck = "" Then
            funcFileExists = False
        Else
        
            ' Returns "WIN.INI" if it exists.
            strFileNameFound = Dir(strFileNameToCheck)
            If strFileNameFound <> "" Then
                funcFileExists = True
            Else
                funcFileExists = False
            End If
            
        End If
        
      
End Function

Function IsFileOpen(strFileSpec As String) As Boolean

    ' Purpose   Test to see if the file has been locked by another process
    '                or if the file is Read-Only
    
    ' If the file is flagged as Read Only return In-User and Get Out NOW
    txtActionBeforeError = "If GetAttr(" & strFileSpec & ") And vbReadOnly Then"
    funcWriteToDebugLog App.Title, txtActionBeforeError
    
    If GetAttr(strFileSpec) And vbReadOnly Then
        txtActionBeforeError = "File Flagged as READ-ONLY:  " & strFileSpec
        funcWriteToDebugLog App.Title, txtActionBeforeError
        IsFileOpen = True
        Exit Function
    End If
        
    
    Dim intFn As Integer, lngErrNum As Long, strErrDescr As String

    On Error Resume Next        ' No Error checking
    intFn = FreeFile()          ' Get a free file number.
    ' Attempt to open the file and lock it.
'    Open strFileSpec For Input Lock Read As intFn

    txtActionBeforeError = "IsFileOpen: Open " & strFileSpec & " For Binary Access Read Write Lock Read Write As " & intFn
'    txtActionBeforeError = "IsFileOpen: Open " & strFileSpec & " For Binary Access Read  Lock Read  As " & intFn

    funcWriteToDebugLog App.Title, txtActionBeforeError
    
   Open strFileSpec For Binary Access Read Write Lock Read Write As intFn
'   Open strFileSpec For Binary Access Read Lock Read As intFn

    
    Close intFn                 ' Close the file.
    
    lngErrNum = Err.Number      ' Save the error number that occurred.
    strErrDescr = Err.Description
    
    ' Report the error
    Select Case lngErrNum
        Case 0
'            'Now Test if File can be MOVED/RENAMED
'            Dim fso As New FileSystemObject
'            Set fso = New Scripting.FileSystemObject
'            Err.Clear
'
'            txtActionBeforeError = "IsFileOpen:  fso.MoveFile " & strFileSpec & "," & strFileSpec & "_TEST"
'            funcWriteToDebugLog App.Title, txtActionBeforeError
'
'            fso.MoveFile strFileSpec, strFileSpec & "_TEST"
'            lngErrNum = Err.Number      ' Save the error number that occurred.
'            strErrDescr = Err.Description
'            If lngErrNum = 0 Then
'                ' No error occurred - File is NOT already open by another user.
'                IsFileOpen = False
'                'Rename file BACK to it's Original name
'                fso.MoveFile strFileSpec & "_TEST", strFileSpec
'            Else
'                IsFileOpen = True
'            End If
        Case 70
            ' Error number for "Permission Denied. - File is opened by another user.
            IsFileOpen = True
        Case 75
            ' Error number for "Permission Denied. - File is opened by another user.
            IsFileOpen = True
        Case Else
            ' Yikes some other error occurred.
            IsFileOpen = True
'            MsgBox "Error! " & vbCrLf & "Error Number " & lngErrNum & vbCrLf & "Error Descr " & strErrDescr, vbExclamation, "IsFileOpen()"
    End Select

    On Error GoTo 0             ' Turn error checking back on.

End Function



Function UnloadAllForms()
    Dim i As Integer
        While Forms.Count > 2
           ' Find first form besides the "Main/First" to unload
           i = 0
           While Forms(i).Caption = Forms(0).Caption Or Forms(i).Caption = Forms(1).Caption
              i = i + 1
           Wend
           Unload Forms(i)
        Wend

      ' Last thing to be done...
''      Unload Forms(0)
''      End
End Function

Function HideAllForms()

    Dim i As Integer
    
    For i = i To Forms.Count - 1
       If funcIsFormLoaded2(Forms(i).name) Then
            Forms(i).Hide
            DoEvents
       End If
    Next

End Function
 
Function ShowAllForms()

    Dim i As Integer
    
    For i = i To Forms.Count - 1
       If funcIsFormLoaded2(Forms(i).name) Then
            ''''''Debug.Print Forms(i).name
            'This select is to only show forms NOT in Case statements
            Select Case Forms(i).name
                Case "frmImaging101Winsock"
                Case "frmLogin"
                Case "frmConfig"
                Case "frmMessageForm"
                Case "frmImaging101Status"
                Case "frmPopUpNotifyForm"
                Case Else
                    Forms(i).Show
                    TimePause 0.25
            End Select
       End If
    Next

End Function

Function TimePause(PauseTime)
' PauseTime parameter must be in Seconds
Dim Start, Finish, TotalTime, OptikaAppID
    Start = Timer   ' Set start time.
    
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
       
End Function


Function funcGetNextControlNumber(strConnectionString As String, strTableName As String, strFieldName As String)

    funcGetNextControlNumber = frmImaging101Winsock.funcSendData("GET NEXT RECID" & "|" & strConnectionString & "|" & strTableName & "|" & strFieldName)
    
'        '*** CONNECT TO BATCH CONTROL DATABASE TABLE
'        Dim conn As ADODB.Connection
'        Dim rs As ADODB.Recordset
'        Set conn = New ADODB.Connection
'        Set rs = New ADODB.Recordset
'
'        conn.ConnectionString = strConnectionString
'        conn.ConnectionTimeout = 120
'        conn.Mode = adModeReadWrite
'        conn.Open
'
'        ' Begin Transaction
'        conn.BeginTrans
'
'        rs.Open strTableName, conn, adOpenDynamic, adLockPessimistic, adCmdTable
'
'        ' Get first Record... should only be ONE record in Control DB
'        rs.MoveFirst
'
'        '*** GET NEXT BATCHRECID
'        Dim lngControlNumber As Long
'        ' This section is to prevent an error if the field is Null
'        On Error Resume Next
'        lngControlNumber = rs.Fields(strFieldName)
'        If Err.Number <> 0 Then
'            lngControlNumber = 0
'            Err.Clear
'        End If
'        On Error GoTo 0
'
'        rs.Fields(strFieldName) = lngControlNumber + 1
'
'        rs.Update
'
'
'        funcGetNextControlNumber = rs.Fields(strFieldName)
'
'        ' Commit Transaction
'        conn.CommitTrans
'
'
'        rs.Close
'        conn.Close
'        Set rs = Nothing
'        Set conn = Nothing

End Function
Function funcGetFieldFromDB(strConnectionString As String, strTableName As String, strSearchFilter As String, strFieldName As String)
        
        '*** THIS FUNCTION IS FOR GETTING QUICK INDIVIDUAL VALUES FROM A DATABASE TABLE
        '***   AND AVOIDS ALL THE LOGIC INVOLVED
        
        On Error GoTo ERROR_HANDLER
        
        '*** Make sure we received a Search Filter
        If Trim(strSearchFilter) = "" Then
            funcWriteToDebugLog "FunctionDeclarations", "[funcGetFieldFromDB ERROR] Sorry... No Search Filter was Provided!"
            Exit Function
        End If
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        conn.ConnectionString = strConnectionString
        conn.ConnectionTimeout = 120
        conn.Mode = adModeRead
        conn.Open
        
        ssql = "SELECT * FROM " & strTableName & " WHERE " & strSearchFilter
        
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
            .Source = ssql
        End With

        rs.Open
        
       '2017-09-19 - Jacob - Added "If" to prevent unnecessary errors.
        If Not rs.EOF Then
            
            'Return the Found Value
            '*** 2020-04-23 - Jacob - Added the Ampersand blank (&" ") to prevent "Invalid Use of Null" error
            funcGetFieldFromDB = rs.Fields(strFieldName) & ""
            
        Else
            funcGetFieldFromDB = ""
        End If
    
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
Exit Function

ERROR_HANDLER:

        funcWriteToDebugLog "[funcGetFieldFromDB", " ERROR] - " & _
                                                            Err.Number & " -  " & Err.Description & vbCrLf & ssql
        funcGetFieldFromDB = ""
        'Clear the Error to Prevent a "Cascading" errors upon exit
        Err.Clear
        
        '*** 2021-04-02 - Jacob - Added these to handle specifically in the Sub or Function thei rs and Con objects are created.
        'Close connection and the recordset
        rs.Close
        Set rs = Nothing
        Con.Close
        Set Con = Nothing
                        
        
        Resume Next
        
End Function



Function funcSaveFieldToDB(strConnectionString As String, strTableName As String, strSearchFilter As String, strFieldName As String, strFieldValue As Variant)
        
        '*** THIS FUNCTION IS FOR SAVING / SETTING QUICK INDIVIDUAL VALUES TO A DATABASE TABLE
        '***   AND AVOIDS ALL THE LOGIC INVOLVED
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        
        Set conn = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rs = New ADODB.Recordset
        
        On Error GoTo ErrLock
    
        txtActionBeforeError = "Prepare to Open Batch DB Connection"
        '*** Prepare Connection
        With conn
            .ConnectionString = strConnectionString
            .CursorLocation = adUseServer
            .ConnectionTimeout = 120
            .IsolationLevel = adXactReadCommitted
            .Mode = adModeWrite
            txtActionBeforeError = "Open Batch DB Connection"
            .Open
            .Execute "SET LOCK_TIMEOUT -1"
        End With
        
        Set cmd.ActiveConnection = conn


        '*** Prepare Result Set
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
        End With
        
        rs.Source = "UPDATE " & strTableName & _
                            "   SET " & strFieldName & " = '" & strFieldValue & "' " & _
                            " WHERE " & strSearchFilter

        conn.Errors.Clear
        rs.Open

        
'
       ' rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
Exit Function

ErrLock:

    If conn.Errors.item(0).NativeError = 1222 Then  ' Lock Timeout
        conn.RollbackTrans
        GboolTrans = False
'        funcUnLockBatch = "ERROR: Record Lock Timeout Expired... "
        
        conn.Errors.Clear
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

        Exit Function
    End If
    

        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
    
End Function

Function funcGetUserCountFromDB(strConnectionString As String)

        
        '*** THIS FUNCTION IS FOR GETTING QUICK USER COUNT
        
        On Error GoTo ERROR_HANDLER
        
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        conn.ConnectionString = strConnectionString
        conn.ConnectionTimeout = 120
        conn.Mode = adModeRead
        conn.Open
        
        
        ssql = "SELECT Count(*) AS UserCount FROM I101Security "
        ssql = ssql & " WHERE userid not like '%imaging101%'  and userid not like '%i101%'  "
        ssql = ssql & " and username not like '%imaging101%' and username not like '%Jacob Russo%'  "
       
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
            .Source = ssql
        End With

        rs.Open
        

        'Return the Found Value
        funcGetUserCountFromDB = rs.Fields("UserCount")
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
Exit Function

ERROR_HANDLER:

        funcWriteToDebugLog "[funcGetFieldFromDB", " ERROR] - " & _
                                                            Err.Number & " -  " & Err.Description & vbCrLf & ssql
        funcGetUserCountFromDB = ""
        'Clear the Error to Prevent a "Cascading" errors upon exit
        Err.Clear
        Resume Next
        
End Function


Function funcTableExists() As Boolean
    Dim oCat As ADOX.Catalog
    Dim oTable As ADOX.Table
    Dim sTable As String
    Dim bFoundTable As Boolean
    
    sTable = "customers"
    
    Set oCat = New ADOX.Catalog
    oCat.ActiveConnection = RegImaging101ConnectionString
    
    bFoundTable = False
    For Each oTable In oCat.Tables
        If UCase(oTable.name) = UCase(sTable) Then
            bFoundTable = True
            Exit For
        End If
    Next
    
    If bFoundTable Then
        '"Table Found"
        funcTableExists = True
    Else
        '"Table not found"
        funcTableExists = False
    End If
    

End Function


            
Public Function funcGetDirectoryLocation(InitDir As String) As String
'    oShell.BrowseForFolder CONSTANTS
'    ---------------------------------
'    ssfALTSTARTUP = 0x1d,
'    ssfAPPDATA = 0x1a,
'    ssfBITBUCKET = 0xa,
'    ssfCOMMONALTSTARTUP = 0x1e,
'    ssfCOMMONAPPDATA = 0x23,
'    ssfCOMMONDESKTOPDIR = 0x19,
'    ssfCOMMONFAVORITES = 0x1f,
'    ssfCOMMONPROGRAMS = 0x17,
'    ssfCOMMONSTARTMENU = 0x16,
'    ssfCOMMONSTARTUP = 0x18,
'    ssfCONTROLS = 0x3,
'    ssfCOOKIES = 0x21,
'    ssfDESKTOP = 0x0,
'    ssfDESKTOPDIRECTORY = 0x10,
'    ssfDRIVES = 0x11,
'    ssfFAVORITES = 0x6,
'    ssfFONTS = 0x14,
'    ssfHISTORY = 0x22,
'    ssfINTERNETCACHE = 0x20,
'    ssfLOCALAPPDATA = 0x1c,
'    ssfMYPICTURES = 0x27,
'    ssfNETHOOD = 0x13,
'    ssfNETWORK = 0x12,
'    ssfPERSONAL = 0x5,
'    ssfPRINTERS = 0x4,
'    ssfPRINTHOOD = 0x1b,
'    ssfPROFILE = 0x28,
'    ssfPROGRAMFILES = 0x26,
'    ssfPROGRAMS = 0x2,
'    ssfRECENT = 0x8,
'    ssfSENDTO = 0x9,
'    ssfSTARTMENU = 0xb,
'    ssfSTARTUP = 0x7,
'    ssfSYSTEM = 0x25,
'    ssfTEMPLATES = 0x15,
'    ssfWINDOWS = 0x24
    
      Dim oShell As New Shell
      Dim oFolder As Folder
      Dim oFolderItem As FolderItem
     
      On Error Resume Next
      
      Set oFolder = oShell.BrowseForFolder(0, "Select a Folder", ssfDRIVES And ssfNETWORK And ssfNETHOOD And ssfPERSONAL)
       
      Set oFolderItem = oFolder.Items.item
      If Not (oFolderItem Is Nothing) Then
        funcGetDirectoryLocation = oFolderItem.Path
      End If
      



End Function

Function funcCreateDirectoryStructure(txtFullDirectoryToCreate As String) As Boolean

        Dim fso As New FileSystemObject
        Set fso = New Scripting.FileSystemObject
        
            
    Dim txtNewDir As String
    Dim txtFullDirectoryWithoutNetworkPath As String
    Dim txtNetworkPath As String
    Dim arrFullDirSections() As String

     '********************************************************
     '*** CREATE DESTINATION DIRECTORY STRUCTURE FOR FILE
     '********************************************************

     On Error Resume Next ' Ignore errors

    '*** Create the Root Directory Structure

     ' Remove Right Backslash if necessary
     txtFullDirectoryWithoutNetworkPath = Trim(txtFullDirectoryToCreate)
     If Right(txtFullDirectoryToCreate, 1) = "\" Then
         txtFullDirectoryWithoutNetworkPath = Left(txtFullDirectoryToCreate, Len(txtFullDirectoryToCreate) - 1)
     End If

     txtNewDir = ""

     '*** SEE IF DIRECTORY EXISTS... Only create if it does not exist.
     If Trim(Dir(txtFullDirectoryWithoutNetworkPath, vbDirectory)) = "" Then

             '*** SPLIT THE DIRECTORY STRUCTURE INTO EACH SECTION
              '    because the MKDIR statement can't make a subdirectory if
              '    its root doesn't exist.
    
            '*** CHECK IF IT IS A NETWORK PATH
            If Left(txtFullDirectoryToCreate, 2) = "\\" Then
                txtNewDir = Left(txtFullDirectoryToCreate, InStr(3, txtFullDirectoryToCreate, "\"))
                txtFullDirectoryWithoutNetworkPath = Right(txtFullDirectoryToCreate, Len(txtFullDirectoryToCreate) - Len(txtNewDir))
            End If
    
            arrFullDirSections = Split(txtFullDirectoryWithoutNetworkPath, "\")
    
              '-- Loop through array to Build each Subdirectory
              For iCounter = LBound(arrFullDirSections) To UBound(arrFullDirSections)
              
                    txtNewDir = txtNewDir & arrFullDirSections(iCounter) & "\"
                    
                    If Not fso.FolderExists(txtNewDir) Then
                    
                            fso.CreateFolder txtNewDir
                            
                            '*** 2021-08-11 - Jacob - Added check for errors with return.
                            If Err.Number = 0 Then
                                    funcCreateDirectoryStructure = True
                            Else
                                    funcCreateDirectoryStructure = False
                            End If
                            
                    End If
                   
              Next
        
     Else
             funcCreateDirectoryStructure = True
     End If

    Set fso = Nothing
    
     '***  RETURN TO NORMAL ERROR TRAPPING
     'On Error GoTo 0

End Function

Function funcGetSetUserSettings(strGetOrSet As String, strUserSettingsName As String, strUserSettingsValue As String)
        
        '***************************************************************************
        '*** THIS FUNCTION IS FOR QUICKLY GETTING OR SETTING USER-SPECIFIC VALUES
        '***   AND AVOIDS ALL THE LOGIC INVOLVED
        '***************************************************************************
        
        'Check that we don't exceed the Value field size.
        If UCase(strGetOrSet) = "GET" And Len(strUserSettingsValue) > 250 Then
            funcWriteToDebugLog "FunctionDeclarations", "Text for [" & strUserSettingsName & "] must be less than 250 characters.  Value NOT Saved!"
        End If
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim ssql As String

        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        conn.ConnectionString = RegImaging101ConnectionString
        conn.ConnectionTimeout = 120
        conn.Mode = adModeReadWrite
        conn.Open
        
        ssql = "SELECT * FROM I101UserSettings WHERE SecurityRECID = " & gsecSecurityRECID & " AND UserSettingsName = '" & strUserSettingsName & "'"
        
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Source = ssql
        End With

        rs.Open

        If UCase(strGetOrSet) = "GET" Then
            If Not rs.EOF Then
                'Return the Found Value
                funcGetSetUserSettings = rs.Fields("UserSettingsValue")
            Else
                funcGetSetUserSettings = ""
            End If
        Else  ' SET
            If rs.EOF Then
                rs.AddNew
                rs.Fields("UserSettingsRECID") = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "UserSettingsRECID")
            End If
            'Save the new value
            rs.Fields("SecurityRECID") = gsecSecurityRECID
            rs.Fields("UserSettingsName") = strUserSettingsName
            rs.Fields("UserSettingsValue") = Left(strUserSettingsValue, 250)
            rs.Update
        End If
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

        
End Function

Function funcGetDetailSubdirectoryString(DetailRECID As Double) As String
    Dim intIndex As Integer
    Dim strFormattedDetailRECID As String
    Dim strSubDirectory As String
     
    strFormattedDetailRECID = Format(DetailRECID, "0000000000")
    
    For intIndex = 1 To 7
        strSubDirectory = strSubDirectory & "\" & Mid(strFormattedDetailRECID, intIndex, 1)
    Next
    
    funcGetDetailSubdirectoryString = strSubDirectory
    
End Function


Public Function funcMakeTopMost(hwndWindowHandle As Object, TrueOrFalse As Boolean)
    ' Make this window the top most window so that it appears in
    ' front of the ecIndex window
    Dim lReturn As Long
    
    ' The SetWindowPos API call will do this.
    If TrueOrFalse = True Then
        lReturn = SetWindowPos(hwndWindowHandle.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    Else
        lReturn = SetWindowPos(hwndWindowHandle.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    End If
    
End Function



Public Function funcFindItemInComboBox(ListObject As ComboBox, TextToFind As String) As String
        ' Walk down the Application list... there was no easier way to set the
        '   ListIndex to the right value to Trigger the "cmbApplicationList_Click" event
        For i = 0 To ListObject.ListCount - 1
            If TextToFind = ListObject.List(i) Then
                ' This will Trigger the ListObject's "_Click" Event
                '   if any to take any required actions
                ListObject.ListIndex = i
                Exit For
            End If
        Next i

End Function

Public Function funcFindItemInComboBoxPartial(ListObject As ComboBox, TextToFind As String) As String
        ' Walk down the Application list... there was no easier way to set the
        '   ListIndex to the right value to Trigger the "cmbApplicationList_Click" event
        For i = 0 To ListObject.ListCount - 1
            If TextToFind = Left(ListObject.List(i), Len(TextToFind)) Then
                ' This will Trigger the ListObject's "_Click" Event
                '   if any to take any required actions
                ListObject.ListIndex = i
                Exit For
            End If
        Next i

End Function
Public Function funcFindItemInListBox(ListObject As ListBox, TextToFind As String) As String

        ' Walk down the Application list... there was no easier way to set the
        '   ListIndex to the right value to Trigger the "cmbApplicationList_Click" event
        For i = 0 To ListObject.ListCount - 1
            If TextToFind = ListObject.List(i) Then
                ' This will Trigger the ListObject's "_Click" Event
                '   if any to take any required actions
                ListObject.ListIndex = i
                Exit For
            End If
        Next i
        
End Function



Public Function funcIsFormLoaded(frm As Form) As Boolean
    
    
    Dim i As Integer
    funcIsFormLoaded = False

    For i = 0 To Forms.Count - 1

        If Forms(i) Is frm Then
            funcIsFormLoaded = True
            Exit Function
        End If
    Next
End Function




Public Function funcIsFormLoaded2(FormName As String) As Boolean

    Dim objForm As Form
    
    funcIsFormLoaded2 = False
    
    For Each objForm In Forms
        If objForm.name = Trim(FormName) Then
            funcIsFormLoaded2 = True
            Exit Function
        End If
    Next objForm

End Function




Function funcCreateBatchAuditRecord(strConnectionString As String, UserID As String, BatchRECID As Double, BatchAuditAction As String) As String

        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim rsAudit As ADODB.Recordset
        
        Set conn = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rs = New ADODB.Recordset
        Set rsAudit = New ADODB.Recordset
        
        On Error GoTo ErrLock
    
        txtActionBeforeError = "Prepare to Open Batch DB Connection"
        '*** Prepare Connection
        With conn
            .ConnectionString = strConnectionString
            .CursorLocation = adUseServer
            .ConnectionTimeout = 120
            .IsolationLevel = adXactReadCommitted
            .Mode = adModeReadWrite
            txtActionBeforeError = "Open Batch DB Connection"
            .Open
            .Execute "SET LOCK_TIMEOUT -1"
        End With
        
        Set cmd.ActiveConnection = conn

'        '*** Begin Transaction
'        conn.BeginTrans

        '*** Prepare Result Set
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
        End With
        
        rs.Source = "SELECT * " & _
                    " FROM I101Batches " & _
                    " WHERE BatchRECID = " & BatchRECID
        
        conn.Errors.Clear
        rs.Open

        'Position the cursor on the rowset
        rs.MoveFirst
        
        '*** CREATE the NEW RECORD
        With rsAudit
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
        End With
        
        '*** Prepare Result Set
        rsAudit.Source = "SELECT * " & _
                    " FROM I101BatchAudit" & _
                    " WHERE 0 = 1"
        
        rsAudit.Open
        rsAudit.AddNew
        
        '*************************************************************
        
        rsAudit.Fields("BatchAuditDate") = Now()
        rsAudit.Fields("BatchAuditUser") = UserID
        rsAudit.Fields("BatchAuditAction") = BatchAuditAction
        
        Dim intIndex As Integer
        
        For intIndex = 0 To rs.Fields.Count - 1
            ''''''Debug.Print rsAudit.Fields(intIndex + 3).name & " -> " & rs.Fields(intIndex).name
            rsAudit.Fields(intIndex + 3) = rs.Fields(intIndex)
        Next
                
        '*************************************************************
        
        rsAudit.Update
        
        
        rs.Close
        rsAudit.Close
        conn.Close
        
        Set rs = Nothing
        Set rsAudit = Nothing
        Set cmd = Nothing
        Set conn = Nothing
        
Exit Function

ErrLock:

    If conn.Errors.item(0).NativeError = 1222 Then  ' Lock Timeout
        conn.RollbackTrans
        GboolTrans = False
'        funcUnLockBatch = "ERROR: Record Lock Timeout Expired... "
        
        conn.Errors.Clear
        Exit Function
    End If
    
    'Get BatchError and Return the full value
''    funcUnLockBatch = BatchErrHandler(conn)
'    Resume Next


End Function



Public Function funcQuickMessage(HideShow As String, Message As String)

    Select Case UCase(HideShow)
        Case "SHOW"
            Dim fMessageForm As New frmMessageForm
            fMessageForm.Show
            'fMessageForm.txtMessage = Message
            fMessageForm.subDisplayMessage Message
            DoEvents
            FormOnTop fMessageForm.hwnd, True
            DoEvents
        Case "HIDE"
            Unload fMessageForm 'frmMessageForm
    End Select
    
    'Scroll Text to the TOP
    fMessageForm.txtMessage.SelStart = 1
    
End Function


Public Function funcFillList(ListObject As Object, ConnectionString As String, TableName As String, FieldName As String, WhereClause As String, AddBlank As Boolean, ClearList As Boolean)

    On Error GoTo ERROR_HANDLER
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim Con As ADODB.Connection
    

    Dim ssql As String

    Set Con = New ADODB.Connection
    
    Con.ConnectionTimeout = 120
    Con.CommandTimeout = 600
        
    Set rs = New ADODB.Recordset
    Con.Open ConnectionString
    
    
    'sql statement to select items on the drop down list
    ssql = "SELECT DISTINCT " & FieldName & " FROM " & TableName & " "
    
    If WhereClause <> "" Then
        ssql = ssql & " WHERE " & WhereClause
    End If
    
    ssql = ssql & " ORDER BY " & FieldName
    
    ''''''Debug.Print ssql
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = ssql
    End With
    
    rs.Open
    
    'Hold the current List item value
    Dim strListObjectText As String
    strListObjectText = ListObject.Text
    
    If ClearList Then
        ' Reset the List
        ListObject.Clear
    End If
    
    If AddBlank Then
        'Add blank item at top
        ListObject.AddItem ""
    End If
    
    ''''''Debug.Print rs.RecordCount
    
    'Only loop if we found records matching the criteria
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
            If IsNull(rs.Fields(0)) Then
                ListObject.AddItem ""
            Else
                ListObject.AddItem rs.Fields(0)
            End If
            rs.MoveNext
        Loop
    End If
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing

    'Restore the value to the Object
    ListObject.Text = strListObjectText
    
Exit Function

ERROR_HANDLER:
    funcWriteToDebugLog "FunctionDeclarations", "ERROR Filling List: " & Err.Number & "  Description: " & Err.Description
    
End Function


Public Function funcKillFileIfSmallerThan(FullFilePath As String, FileMinimumSize As Long) As Boolean
    
    Dim lngFileSize As Long
    
    'FileLen function returns the file size in BYTES.
    lngFileSize = FileLen(FullFilePath)
    
    funcKillFileIfSmallerThan = False
    
    If lngFileSize < FileMinimumSize Then
         funcWriteToDebugLog "FunctionDeclarations", "KILL  " & FullFilePath & "   Size: " & lngFileSize & " Min: " & FileMinimumSize
        Kill FullFilePath
    
        'Set to true to notify the calling routing that the image WAS deleted
        funcKillFileIfSmallerThan = True

    End If
    
End Function


Public Function funcGetFullPathForAnnotation(ApplicationRECID As Double, DetailRECID As Double) As String

    '******************************************************
    '*** Get the Root Directory to Store Image Annotations
    
'    RegRootDirectoryPathForImageAnnotations = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & MainMDIForm.ActiveForm.txtApplicationRECID, "RootDirectoryPathForImageAnnotations") & ""
    RegRootDirectoryPathForImageAnnotations = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Applications", "ApplicationRECID = " & ApplicationRECID, "RootDirectoryPathForImageAnnotations") & ""
    
    funcGetFullPathForAnnotation = RegRootDirectoryPathForImageAnnotations & "\" & _
                                    Format(CStr(ApplicationRECID), "0000") & _
                                    funcGetDetailSubdirectoryString(DetailRECID)
    

End Function


Public Function funcCountListViewItemsSelected(ListViewObject As Object) As Double

    For i = 1 To ListViewObject.ListItems.Count
        If ListViewObject.ListItems(i).Selected = True Then
            funcCountListViewItemsSelected = funcCountListViewItemsSelected + 1
        End If
    Next

End Function

Public Function funcGetTempName() As String

    Dim ReturnVal As Long
    Dim PathBuffSize As Long
    Dim PathBuff As String
    Dim Prefix As String
    Dim FileName As String
    
    PathBuffSize = 255
    PathBuff = Space$(255)
    Prefix = "TMP"
    FileName = Space$(255)
    
    ReturnVal = GetTempPath(PathBuffSize, PathBuff)
    
    If ReturnVal = 0 Then
         funcWriteToDebugLog "FunctionDeclarations", "Error getting TEMP Directory."
        Exit Function
    End If
    
    PathBuff = TrimNull(PathBuff)
    
    ReturnVal = GetTempFileName(PathBuff, Prefix, 0&, FileName)
    
    If ReturnVal = 0 Then
         funcWriteToDebugLog "FunctionDeclarations", "Error getting TEMP Filename."
    Else
        funcGetTempName = TrimNull(FileName)
    End If
    
End Function

Public Function funcGetTempDir() As String

    Dim ReturnVal As Long
    Dim PathBuffSize As Long
    Dim PathBuff As String
    Dim Prefix As String
    
    PathBuffSize = 255
    PathBuff = Space$(255)
    
    ReturnVal = GetTempPath(PathBuffSize, PathBuff)
    
    If ReturnVal = 0 Then
        funcWriteToDebugLog "FunctionDeclarations", "Error getting TEMP Directory."
        Exit Function
    Else
        funcGetTempDir = TrimNull(PathBuff)
    End If
    
End Function

Public Function TrimNull(item As String)

   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function

Function funcRunSQLCommand(strConnectionString As String, strSQLCommand As String)
        
        '*** THIS FUNCTION IS FOR SENDING A SQL COMMAND
        '***   AND AVOIDS ALL THE LOGIC INVOLVED
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        
        Set conn = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rs = New ADODB.Recordset
        
        On Error GoTo ErrLock
    
        txtActionBeforeError = "funcRunSQLCommand - Prepare to Open DB Connection"
        funcWriteToDebugLog "funcRunSQLCommand", txtActionBeforeError

        '*** Prepare Connection
        With conn
            .ConnectionString = strConnectionString
            .CursorLocation = adUseServer
            .ConnectionTimeout = 120
            .IsolationLevel = adXactReadCommitted
            .Mode = adModeWrite
            txtActionBeforeError = "funcRunSQLCommand - Open DB Connection"
            funcWriteToDebugLog "funcRunSQLCommand", txtActionBeforeError
            .Open
            .Execute "SET LOCK_TIMEOUT -1"
        End With
        
        funcWriteToDebugLog "funcRunSQLCommand", strSQLCommand
        Set cmd.ActiveConnection = conn

'        '*** Begin Transaction
'        conn.BeginTrans

        '*** Prepare Result Set
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
        End With
        
        rs.Source = strSQLCommand
        
        conn.Errors.Clear
        
        '*** EXECUTE The SQL Command
        rs.Open

        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        
Exit Function

ErrLock:

    If conn.Errors.item(0).NativeError = 1222 Then  ' Lock Timeout
        conn.RollbackTrans
        GboolTrans = False
        funcWriteToDebugLog "FunctionDeclarations", "ERROR: Record Lock Timeout Expired... "
        
        conn.Errors.Clear
        Exit Function
    End If
    
    funcWriteToDebugLog "FunctionDeclarations", "funcRunSQLCommand ERROR: " & Err.Number & " - " & Err.Description
        
End Function


Function funcCleanFileName(strFileName As String) As String
    '*** Remove INVALID Characters from the File Name

    Dim i As Integer
    Dim strBadChars As String
    
'    strBadChars = "`~!@#$%^&*()+={}[]|\/?<>:;"
    
    strBadChars = "\:/;*?<>|" & Chr(34)
    
    For i = 1 To Len(strBadChars)
        strFileName = Replace(strFileName, Mid(strBadChars, i, 1), "-")
    Next
    
    funcCleanFileName = strFileName
    
End Function


' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Function funcBrowseForDirectory() As String

    Dim browse_info As BrowseInfo
    Dim item As Long
    Dim dir_name As String
    
    Const BIF_RETURNONLYFSDIRS = 1
    Const BIF_NEWDIALOGSTYLE = &H40 '<-- added the constant

   browse_info.hWndOwner = hwnd
   browse_info.pidlRoot = 0
   browse_info.sDisplayName = Space$(260)
   browse_info.sTitle = "Select Directory"
'   browse_info.ulFlags = 1 ' Return directory name.
   browse_info.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE  ' Return directory name & show NEW button
   browse_info.lpfn = 0
   browse_info.lparam = 0
   browse_info.iImage = 0
   
   item = SHBrowseForFolder(browse_info)
   If item Then
       dir_name = Space$(260)
       If SHGetPathFromIDList(item, dir_name) Then
           funcBrowseForDirectory = Left(dir_name, _
               InStr(dir_name, Chr$(0)) - 1)
       Else
           funcBrowseForDirectory = ""
       End If
   End If
   
End Function


''''Function funcBurnToCD()
''''
''''    Dim objCD As CDBurn
''''    Set objCD = New CDBurn
''''
''''    If objCD.HasRecordableDrive Then
''''    Dim strDrive As String
''''    Dim buff As String
''''    Dim buffsize As Long
''''    Dim lngResult As Long
''''    Const S_OK As Long = 0
''''
''''    buff = Space$(4)
''''    buffsize = Len(buff)
''''
''''    If objCD.GetRecorderDriveLetter(buff, buffsize) = S_OK Then
''''
''''    lngResult = objCD.Burn(Me.hwnd)
''''
''''    Else
''''    MsgBox "no drive letter returned"
''''    End If
''''    Else
''''    MsgBox "No CD writer found", vbOKOnly + vbExclamation, ""
''''    End If
''''
''''End Function
''''


Function funcWriteToDebugLog(ByVal strSourceForm As String, ByVal strLogMessage As String)
            
    If Not bolDebug And Not bolDebugService Then
        Exit Function
    End If
    
    On Error GoTo ERROR_HANDLER
    
    txtActionBeforeError = "funcRemovePasswordFromText(strLogMessage)"
    strLogMessage = funcRemovePasswordFromText(strLogMessage)
    
'    strTempDir = funcGetTempDir()
    
    'Get the AppData directories
    Dim strNewLocationForINI As String
    strLocalAppDataDir = Environ$("LocalAppData")
    strNewLocationForINI = strLocalAppDataDir & "\Imaging101"

    strLogPath = strNewLocationForINI & "\" & App.EXEName & "_DEBUG_" & Format(Now(), "yyyy-mm-dd") & ".LOG"
    
    txtActionBeforeError = "Open " & strLogPath & " For Append As #99"
    Open strLogPath For Append As #99
    txtActionBeforeError = "Print #99, " & Now() & " ~ " & strSourceForm & " ~ " & strLogMessage
    Print #99, Now() & " ~ " & Format$(strSourceForm, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " ~ " & strLogMessage
    Close #99
        
    
Exit Function

ERROR_HANDLER:

    If funcServiceIsRunning("Imaging101AutoExport") Then
        App.LogEvent "funcWriteToDebugLog ERROR: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ")", LogEventTypeConstants.vbLogEventTypeError
    Else
        App.LogEvent "funcWriteToDebugLog ERROR: " & Err.Number & " - " & Err.Description & " DURING ACTION: (" & txtActionBeforeError & ")", LogEventTypeConstants.vbLogEventTypeError
    End If
    
End Function

Function funcServiceIsRunning(sServiceName)

    Dim objInst, objSet
    funcServiceIsRunning = False
    Set objSet = GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_Service")
    For Each objInst In objSet
        If (UCase(sServiceName) = UCase(objInst.name)) And (UCase(objInst.State) = UCase("Running")) Then
            funcServiceIsRunning = True
        End If
    Next
    
End Function


Function funcRemovePasswordFromText(ByVal strTextToRemovePasswordFrom As String) As String

    Dim intPasswordLocation As Integer
    Dim intPasswordLength As Integer
    
'    strTextToRemovePasswordFrom = "Provider=SQL.1;Password=abcd;userid=1234"
    If InStr(1, UCase(strTextToRemovePasswordFrom), "PASSWORD=") Then
        intPasswordLocation = InStr(1, UCase(strTextToRemovePasswordFrom), "PASSWORD=") + 9
        intPasswordLength = InStr(intPasswordLocation, UCase(strTextToRemovePasswordFrom), ";") - intPasswordLocation
        strTextToRemovePasswordFrom = Left(strTextToRemovePasswordFrom, intPasswordLocation - 1) _
                                        & "***" & Right(strTextToRemovePasswordFrom, Len(strTextToRemovePasswordFrom) - (intPasswordLocation + intPasswordLength) + 1)
    End If
    
    funcRemovePasswordFromText = strTextToRemovePasswordFrom

End Function
Function funcWriteToSystemEventLog(ServiceControl As Control, ServiceEventType As Integer, ByVal ServiceMessageText As String)
    
    'Write to System Event Log
    If bolDebug Or bolDebugService Then
     
        Call ServiceControl.LogEvent(ServiceEventType, svcMessageInfo, ServiceMessageText & " - " & Now() & "... Version " & App.Major & "." & App.Minor & "." & App.Revision)
    
    End If
    
    'Write to DEBUG Log File
    funcWriteToDebugLog "funcWriteToSystemEventLog", ServiceMessageText
    DoEvents

End Function

Sub LoadText(objTextBox As TextBox, strFilePath As String)

    'Call LoadText (Text1,"C:\Windows\System\Saved.txt")
    
    On Error GoTo ERROR_HANDLER
    
    Dim strInputLine As String
    Dim strText As String
    
    Open strFilePath For Input As #1
    
    Do While Not EOF(1)
    
                Line Input #1, strInputLine
    
                strText = strText + strInputLine + Chr$(13) + Chr$(10)
    
            Loop
    
            objTextBox.Text = strText
    
    Close #1
    
    Exit Sub
    
ERROR_HANDLER:
    
    funcWriteToDebugLog "FunctionDeclarations", "File Not Found: [" & strFilePath & "]"

End Sub



Public Function IsArrayEmpty(aArray As Variant) As Boolean

    On Error Resume Next

    
    IsArrayEmpty = UBound(aArray)
    If Err.Number = 9 Then
            IsArrayEmpty = True
    End If
    
'   IsArrayEmpty = Err ' Error 9 (Subscript out of range)
    
End Function


Public Function funcGetSecurityRights(ByVal txtSecurityRECID As String, ByVal txtSecurityApplicationRECID As String)

    '*** LOAD SECURITY RIGHTS BY ROLE / USER / APPLICATION
    
    On Error GoTo ERROR_HANDLER
    
    '*** CLEAR Security Variables
         gsecRightsAdminApplication = 0
         gsecRightsRetrieveImages = 0
         gsecRightsBatchScan = 0
         gsecRightsBatchIndex = 0
         gsecRightsBatchAdministration = 0
         gsecRightsBatchView = 0
         gsecRightsBatchCommit = 0
         gsecRightsBatchRoute = 0
         gsecRightsBatchChangeOrder = 0
         gsecRightsImportFromFile = 0
         gsecRightsImportFromEcapture = 0
         gsecRightsDeleteDocuments = 0
         gsecRightsModifyIndexes = 0
         gsecRightsDeleteBatches = 0
         gsecRightsSendMail = 0
         gsecRightsLaunchDoc = 0
         gsecRightsPrint = 0
         gsecRightsAnnotate = 0
         gsecRightsThumbnails = 0
         gsecRightsScannerSettings = 0
         gsecRightsDocPackage = 0
         gsecRightsExport = 0
         gsecBatchMode = ""
         gsecUserSupervisor = ""
         gsecBatchListOrder = ""
         gsecBatchDefaultQueue = ""
'         gsecBatchDefaultApplication = rs.Fields!BatchDefaultApplication & ""
         gsecBatchQueueNotificationFrequency = "0"
         gsecViewResetImagesOnFind = 0
         gsecAllowModificationOfOrigDocs = 0
         gsecRightsBatchFindRestricted = 0
         gsecRightsBatchFindRestrictToQueue = 0
         gsecRightsBatchFindRestrictToOwner = 0
         gsecRightsBatchChangeQueue = 0
         gsecRightsBatchChangeOwner = 0
         gsecRightsBatchAllowDocTypeEdit = 0
         gsecAdvancedSearch = 0
         
         '*** 2020-05-15 - Jacob - Added ability to Edit Search Templates by User / by Application
        gsecRightsEditSearchTemplates = 0
         
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim Con As ADODB.Connection
    Dim ssql As String

    Set Con = New ADODB.Connection
    
    Con.ConnectionTimeout = 120
    Con.CommandTimeout = 600
    
    Set rs = New ADODB.Recordset
    Con.Open RegImaging101ConnectionString
    
    'SQL Statement to Select the Security settings by Role/User/Application
    ssql = "Select * from I101SecurityRoleApp where SecurityRECID = " & _
            txtSecurityRECID & _
            " AND ApplicationRECID = " & _
            txtSecurityApplicationRECID
    
    rs.Open ssql, Con
    
    If rs.EOF Then
        funcQuickMessage "SHOW", "SECURITY for this Application Has NOT been set.  Please contact a supervisor to correct this."
        bolErrorOccured = True
        Exit Function
    End If
    
'         gsecDocumentGroup = rs.Fields!DocumentGroup & ""
'         gsecUserGroups = rs.Fields!UserGroups & ""
        
         'Get Rights Flags and set to UpperCase just in case...
'         gsecRightsAdminSystem = UCase(rs.Fields!RightsAdminSystem & "")
         gsecRightsAdminApplication = UCase(rs.Fields!RightsAdminApplication & "")
         gsecRightsRetrieveImages = UCase(rs.Fields!RightsRetrieveImages & "")
         gsecRightsBatchScan = UCase(rs.Fields!RightsBatchScan & "")
         gsecRightsBatchIndex = UCase(rs.Fields!RightsBatchIndex & "")
         gsecRightsBatchAdministration = UCase(rs.Fields!RightsBatchAdministration & "")
         gsecRightsBatchView = UCase(rs.Fields!RightsBatchView & "")
         gsecRightsBatchCommit = UCase(rs.Fields!RightsBatchCommit & "")
         gsecRightsBatchRoute = UCase(rs.Fields!RightsBatchRoute & "")
         gsecRightsBatchChangeOrder = UCase(rs.Fields!RightsBatchChangeOrder & "")
         gsecRightsImportFromFile = UCase(rs.Fields!RightsImportFromFile & "")
         gsecRightsImportFromEcapture = UCase(rs.Fields!RightsImportFromEcapture & "")
         gsecRightsDeleteDocuments = UCase(rs.Fields!RightsDeleteDocuments & "")
         gsecRightsModifyIndexes = UCase(rs.Fields!RightsModifyIndexes & "")
         gsecRightsDeleteBatches = UCase(rs.Fields!RightsDeleteBatches & "")
         gsecRightsSendMail = UCase(rs.Fields!RightsSendMail & "")
         gsecRightsLaunchDoc = UCase(rs.Fields!RightsLaunchDoc & "")
         gsecRightsPrint = UCase(rs.Fields!RightsPrint & "")
         gsecRightsAnnotate = UCase(rs.Fields!RightsAnnotate & "")
         gsecRightsThumbnails = UCase(rs.Fields!RightsThumbnails & "")
         gsecRightsScannerSettings = UCase(rs.Fields!RightsScannerSettings & "")
         gsecRightsDocPackage = UCase(rs.Fields!RightsDocPackage & "")
         gsecRightsExport = UCase(rs.Fields!RightsExport & "")
         
         gsecBatchMode = rs.Fields!BatchMode & ""
         gsecUserSupervisor = rs.Fields!UserSupervisor & ""
         gsecBatchListOrder = rs.Fields!BatchListOrder & ""
         gsecBatchDefaultQueue = rs.Fields!BatchDefaultQueue & ""
'         gsecBatchDefaultApplication = rs.Fields!BatchDefaultApplication & ""
         gsecBatchQueueNotificationFrequency = rs.Fields!BatchQueueNotificationFrequency & ""
         
         If gsecBatchQueueNotificationFrequency = "" Then
             gsecBatchQueueNotificationFrequency = "0"
         End If
         
         
         gsecViewResetImagesOnFind = rs.Fields!ViewResetImagesOnFind & ""
         
         gsecAllowModificationOfOrigDocs = rs.Fields!AllowModificationOfOrigDocs & ""
         
         gsecRightsBatchFindRestricted = rs.Fields!RightsBatchFindRestricted & ""
         gsecRightsBatchFindRestrictToQueue = rs.Fields!RightsBatchFindRestrictToQueue & ""
         gsecRightsBatchFindRestrictToOwner = rs.Fields!RightsBatchFindRestrictToOwner & ""
         gsecRightsBatchChangeQueue = rs.Fields!RightsBatchChangeQueue & ""
         gsecRightsBatchChangeOwner = rs.Fields!RightsBatchChangeOwner & ""
         
         gsecRightsBatchAllowDocTypeEdit = rs.Fields!RightsBatchAllowDocTypeEdit & ""
         
         gsecAdvancedSearch = rs.Fields!RightsAdvancedSearch & ""
         
         
         gsecRightsEditSearchTemplates = rs.Fields!RightsEditSearchTemplates & ""
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing
    

    
Exit Function

ERROR_HANDLER:

    funcWriteToDebugLog "FunctionDeclarations", "funcGetSecurityRights ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    Resume Next


End Function


Public Function funcGetMenuRights(ByVal txtSecurityRECID As String)

    '*** LOAD SECURITY RIGHTS BY ROLE / USER / APPLICATION
    
    On Error GoTo ERROR_HANDLER
    
        
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim Con As ADODB.Connection
    Dim ssql As String

    Set Con = New ADODB.Connection
    
    Con.ConnectionTimeout = 120
    Con.CommandTimeout = 600
        
    Con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    
    'SQL Statement to Select the Security settings by Role/User/Application
    ssql = "SELECT MAX(RightsAdminApplication) As RightsAdminApplication"
    ssql = ssql & ",MAX(RightsAdminSystem) AS RightsAdminSystem"
    ssql = ssql & ",MAX(RightsBatchAdministration) AS RightsBatchAdministration"
    ssql = ssql & ",MAX(RightsBatchCommit) AS RightsBatchCommit"
    ssql = ssql & ",MAX(RightsBatchIndex) AS RightsBatchIndex"
    ssql = ssql & ",MAX(RightsBatchRoute) AS RightsBatchRoute"
    ssql = ssql & ",MAX(RightsBatchScan) AS RightsBatchScan"
    ssql = ssql & ",MAX(RightsBatchView) AS RightsBatchView"
    ssql = ssql & ",MAX(RightsImportFromFile) AS RightsImportFromFile"
    ssql = ssql & ",MAX(RightsRetrieveImages) AS RightsRetrieveImages"
    ssql = ssql & " FROM I101SecurityRoleApp"
    ssql = ssql & " WHERE SecurityRECID = " & txtSecurityRECID
    ssql = ssql & " GROUP BY SecurityRECID"
    
    rs.Open ssql, Con
    
    If rs.EOF Then
        Exit Function
    End If
    
       
         'Get Rights Flags and set to UpperCase just in case...
'         gsecRightsAdminSystem = UCase(rs.Fields!RightsAdminSystem & "")
         gsecRightsAdminApplication = UCase(rs.Fields!RightsAdminApplication & "")
         gsecRightsRetrieveImages = UCase(rs.Fields!RightsRetrieveImages & "")
         gsecRightsBatchScan = UCase(rs.Fields!RightsBatchScan & "")
         gsecRightsBatchIndex = UCase(rs.Fields!RightsBatchIndex & "")
         gsecRightsBatchAdministration = UCase(rs.Fields!RightsBatchAdministration & "")
         gsecRightsBatchView = UCase(rs.Fields!RightsBatchView & "")
         gsecRightsBatchCommit = UCase(rs.Fields!RightsBatchCommit & "")
         gsecRightsBatchRoute = UCase(rs.Fields!RightsBatchRoute & "")
         gsecRightsImportFromFile = UCase(rs.Fields!RightsImportFromFile & "")
         
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing
    
        
Exit Function

ERROR_HANDLER:

    funcWriteToDebugLog "FunctionDeclarations", "funcGetMenuRights ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    Resume Next


End Function


Public Function funcExportSaveDocument(strDirectory As String, _
                                                ctlDocControl As Control, _
                                                ctlViewControl As Control, _
                                                ctlEditControl As Control, _
                                                Optional strSaveFileName As String, _
                                                Optional strSaveFormat As String) As String


    Dim strSaveFileNameExtension As String
    Dim txtAttachmentFileName As String
    
    strDirectory = Trim(strDirectory)
    If Right(strDirectory, 1) <> "\" Then
        strDirectory = strDirectory & "\"
    End If
    
    'Define Filename ONLY if NOT Passed
    If Trim(strSaveFileName) = "" Then
        strSaveFileName = "TempFileName_" & Now()
    End If
    
    'Remove INVALID Characters from the File Name
    strSaveFileName = funcCleanFileName(strSaveFileName)
    
        txtAttachmentFileName = strDirectory & Trim(strSaveFileName) & "." & strSaveFormat
    
        
        Dim docSave As IDocSave
    
    
        Debug.Print
        
        '*** Rasterize the Pages before sending
'         me.subRasterizeBatch
        funcRasterizeBatchEX ctlDocControl, ctlViewControl, ctlEditControl
        DoEvents
        ' Set the object variable for the IDocSave interface to the Document Control object
        ' that was saved by the Rasterize sub
        Set docSave = ctlDocControl.object
'        docSave.SaveAsDialog False
    
    
        '***  Save the modified pages in the Spicer Document format
        '     The FORMAT is Different for Single-page VS Multi-Page PDF
        If ctlDocControl.NumberOfPages > 1 Then

            If strSaveFormat = "PDF" Or Trim(strSaveFormat) = "" Then
                'PDF - Flate, 619, R/W, Raster only Multipage PDF subset - bilevel, 8 bit, 24 bit
                docSave.Save 0, False, 619, txtAttachmentFileName, txtAttachmentFileName
            Else
                docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, txtAttachmentFileName
            End If

            DoEvents
            ' De-initialize the object variable
            Set docSave = Nothing
        Else
            
            If strSaveFormat = "PDF" Then
                'PDF - Flate, 101, R/W, Raster only PDF subset - bilevel, 8 bit, 24 bit
                docSave.Save 0, False, 101, txtAttachmentFileName, txtAttachmentFileName
            Else
                docSave.Save 0, False, API_MPAGE_TIFF, txtAttachmentFileName, txtAttachmentFileName
            End If
            
            DoEvents
        End If

'    End If
    
    'Discard the docSave object
    Set docSave = Nothing
    Set ctlDocControl = Nothing
    Set ctlEditControl = Nothing
    Set ctlViewControl = Nothing
    
    DoEvents
    
    
    
    '*** RASTERIZE DOCUMENT AND SAVE ALL PAGES TO ATTACH - END
    '*********************************************************************
    
    funcExportSaveDocument = txtAttachmentFileName
    
Exit Function

ERROR_HANDLER:

    funcWriteToDebugLog "FunctionDeclrations", "funcExportSelectedSaveDocument ERROR: " & Err.Number & " - " & Err.Description
    
End Function


Public Function funcRasterizeBatchEX(ctlDocControl As Control, _
                                                            ctlViewControl As Control, _
                                                            ctlEditControl As Control)

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
   
   On Error GoTo ERROR_HANDLER
   
   'This flag is to disable the Form Load and Activate subs on the ChildForm
   bolRasterizingDocument = True
   
        ' Set the control to create a NEW Raster Document
        Dim CFGDocument As ICFGDocument
        'Set the object variable for the ICFGDocument interface to the configuration control object
        Set CFGDocument = MainMDIForm.ActiveForm.SpicerConfiguration1.object
        'Set to NOT automatically overwite document on rasterization.
        CFGDocument.RasterOperations(IN_RASTERIZE_OVERWRITE) = True
'        'Set to Remove the original
        CFGDocument.RasterOperations(IN_RASTERIZE_REMOVE) = True
        'Deinitialize the object variable
        Set CFGDocument = Nothing


    ctlViewControl.BindToDocumentControl ctlDocControl.object
    ctlEditControl.BindToDocumentControl ctlDocControl.object
   
   ' Set the object variable for the IRasterBatch interface to the Edit Control object
   Set RasterBatch = ctlEditControl.object
   
    'Remarks
    '
    'The threshold value is ignored unless the MergeColor value is set to bilevel output.
    'The brightness value is ignored unless MergeColor is set to bilevel (enhanced), bilevel (CAD), or bilevel (diffusion dithered).
    'Changing the orientation is a permanent change to the data, not just a display or header change.
    'The xSize and ySize values must either both be set to 0, or both be set to a size value. You cannot use 0 for one and a size value for the other. When both are set to 0, the units value is ignored.
    '
    'The raster image created by this method either replaces the original image, or must be placed into another Document Control, depending on the ReplaceCurrentDocWhenRasterizing property. For details, see Placing new raster images in a Document Control.
   
    'IN_COLORTYPE_COLOR  0   produces a color image containing up to 16 million colors
    'IN_COLORTYPE_BILEVEL    1   produces a black-and-white image with no dithered sections
    'IN_COLORTYPE_BILEVEL_ENHANCED   2   Windows NT/2000/XP only. Produces a monochrome image depending on the format of the original document. For a list of results, click Bilevel Enhanced.
    'IN_COLORTYPE_BILEVEL_CAD    3   Windows NT only. Produces a monochrome image depending on the format of the original document. For a list of results, click Bilevel CAD.
    'IN_COLORTYPE_BILEVEL_DITHER 4   produces a bilevel image in which dithering--generating pixel patterns--is used to simulate gray or color areas of the original image
    'IN_COLORTYPE_24_BIT_COLOR   5   produces a color image containing up to 16 million colors.
    'IN_COLORTYPE_GRAYSCALE  6   produces a grayscale image.
    
   ' Set the rasterize options
   lObjectID = ctlDocControl.RootID
   iXResolution = 0   ' Keep the original resolution
   iYResolution = 0
   iColor = IN_COLORTYPE_COLOR    ' Rasterize to COLOR
   iBrightness = 100 ' Defines the lightness or darkness of the
                     '   bilevel (enhanced), bilevel (CAD), or bilevel (diffusion dithered) profile.
                     '   Values can range from -50 (Dark) to 150 (Light)
   iThreshold = 200  ' Defines the lightness or darkness of the
                      '  bilevel profile.
                     '   Values can range from 0 (Light) to 255 (Dark).
   iOrientation = IN_ORIENTATION_NONE ' Use original orientation
   lXSize = 0 ' Keep the original size
   lYSize = 0
   iUnit = IN_UNITS_INCH

   lYSize = 0
   iUnit = IN_UNITS_INCH
   
   ' Rasterize the entire document
   RasterBatch.RasterizeBatchEx lObjectID, iXResolution, iYResolution, iColor, iBrightness, iThreshold, iOrientation, lXSize, lYSize, iUnit
   
   ' De-initialize the object variables
   Set RasterBatch = Nothing
'   Set ctlDocControl = Nothing
'   Set ctlEditControl = Nothing
'   Set ctlViewControl = Nothing
   
   bolRasterizingDocument = False

Exit Function

ERROR_HANDLER:
        ' De-initialize the object variables
        Set RasterBatch = Nothing
   
        funcQuickMessage "SHOW", "subRasterizeBatchEX: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [DOCUMENT NOT EXPORTED]"
        

End Function

Public Function funcUncommitBatch(strBatchRECID As String, _
                                                        strApplicationName As String, _
                                                        strBatchCommitStatus As String)

'    If (InStr(strBatchCommitStatus, "Committed") > 0) _
'    And gsecRightsAdminSystem = vbChecked Then
    
'        result = MsgBox("Un-commit Batch?", vbYesNo, "Un-Commit Batch")
'
'        If result = vbYes Then

                Dim strSQLCommand As String
                
                'Uncommit the Batch
                Dim txtBatchNotes As String
                txtBatchNotes = ""
                txtBatchNotes = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Batches", "BatchRECID = " & strBatchRECID, "BatchNotes")
                txtBatchNotes = Now() & " - " & gsecUserName & vbCrLf & "   Uncommitted Batch" & vbCrLf & txtBatchNotes
                
                strSQLCommand = "UPDATE I101Batches SET BatchCommitStatus = '', BatchPagesCommitted=0, BatchPagesNotCommitted=BatchPagesTotal, BatchNotes='" & txtBatchNotes & "'   WHERE BatchRECID = " & strBatchRECID
                funcRunSQLCommand RegImaging101ConnectionString, strSQLCommand
                
                '*** 2022-01-13 - Jacob - Also clear BatchPageCommitDate and BatchPageCommitUser
                'Uncommit the Batch Pages
                strSQLCommand = "UPDATE " & strApplicationName & "_BatchPage SET BatchPageStatus = '', BatchPageCommitDate = NULL, BatchPageCommitUser = NULL   WHERE BatchRECID = " & strBatchRECID
                funcRunSQLCommand RegImaging101ConnectionString, strSQLCommand
                
                 'Uncommit the Commited Documents
                 'NOTE:  This does NOT actually DELETE the documents... it simply FLAGS the DocumentLocked as "D" for Deleted
                strSQLCommand = "UPDATE " & strApplicationName & "  SET DocumentLocked = 'D', DocumentLockedBy = '" & gsecUserName & "', DocumentLockedDate = '" & Now() & "' WHERE DocumentBatchRECID = " & strBatchRECID
                funcRunSQLCommand RegImaging101ConnectionString, strSQLCommand
               
'        End If
    
'    End If
    

End Function


Public Function funcBuildFTPUploadFileName(strApplicationRECID As String, WorkingForm As Form)

    funcWriteToDebugLog "FunctionDeclarations", "ENTERING: funcBuildFTPUploadFileName()"

    Dim rs As ADODB.Recordset
    Dim Con As ADODB.Connection
    Dim ssql As String

    Set Con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Applications where ApplicationRECID = " & Int(strApplicationRECID)
    rs.Open ssql, Con
    
    On Error Resume Next
    
    Dim strFTPUploadFileName As String
    Dim intFTPMaxFieldIndex As Integer
        
    strFTPUploadFileName = ""
    intFTPMaxFieldIndex = 3
    
    
    For intFTPFieldIndex = 0 To intFTPMaxFieldIndex
    
            For intIndex = 0 To WorkingForm.lblFieldDescription.Count - 1
            
                ' Set field default value
                If UCase(Trim(WorkingForm.lblFieldDescription.item(intIndex).Caption)) = _
                    UCase(Trim(rs.Fields("FTPFileNameField" & intFTPFieldIndex))) _
                    Then
                    
                            strFTPUploadFileName = strFTPUploadFileName & _
                                                                WorkingForm.mebIndexValues(intIndex).Text

                            Exit For
                    
                End If
                
            Next
            
            'DO NOT ADD THE DELIMITER FOR THE LAST FIELD
            If intFTPFieldIndex < intFTPMaxFieldIndex Then
                'Add Delimiter
                strFTPUploadFileName = strFTPUploadFileName & rs.Fields("FTPFileNameDelimiter" & intFTPFieldIndex)
            End If
            
        Next


    On Error GoTo ERROR_HANDLER
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing

    funcBuildFTPUploadFileName = strFTPUploadFileName
    
Exit Function

ERROR_HANDLER:
        ' De-initialize the object variables
        Set RasterBatch = Nothing
   
        funcQuickMessage "SHOW", "funcBuildFTPUploadFileName: " & Err.Number & " - " & Err.Description & "  DURING ACTION: (" & txtActionBeforeError & ") - [DOCUMENT NOT EXPORTED]"
        
    
End Function

Public Function funcFindFile(ByVal sFol As String, sFile As String, _
                                            nDirs As Long, nFiles As Long, arrFilesFound() As String) As Currency
   
    Dim fso As New FileSystemObject
    Dim fld 'As Folder

    Dim tFld 'As Folder
    Dim tFil As File
    Dim FileName As String
    
    On Error GoTo Catch
    Set fld = fso.GetFolder(sFol)
    
    FileName = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                   vbHidden Or vbSystem Or vbReadOnly)
    
    While Len(FileName) <> 0
       funcFindFile = funcFindFile + FileLen(fso.BuildPath(fld.Path, _
       FileName))
       nFiles = nFiles + 1
       ReDim Preserve arrFilesFound(nFiles)
       arrFilesFound(nFiles) = fso.BuildPath(fld.Path, FileName)  ' Load ListBox
       FileName = Dir()  ' Get next file
       DoEvents
    Wend
    
    Label1 = "Searching " & vbCrLf & fld.Path & "..."
    nDirs = nDirs + 1
    If fld.SubFolders.Count > 0 Then
       For Each tFld In fld.SubFolders
          DoEvents
          funcFindFile = funcFindFile + funcFindFile(tFld.Path, sFile, nDirs, nFiles, arrFilesFound)
       Next
    End If
   
    
   Exit Function
   
Catch:  FileName = ""
       Resume Next
End Function

Public Function funcFindSubdirectories(ByVal sFol As String, nDirs As Long, arrSubDirectoriesFound() As String) As Currency

    Dim fso As New FileSystemObject
    Dim fld 'As Folder

    Dim tFld 'As Folder
    Dim tFil As File
    Dim FileName As String
    
    On Error GoTo Catch
    Set fld = fso.GetFolder(sFol)
   
    Label1 = "Searching " & vbCrLf & fld.Path & "..."
    
    nDirs = nDirs + 1
    
    If fld.SubFolders.Count > 0 Then
    
       For Each tFld In fld.SubFolders
          
            '2022-11-18 - Jacob - Added IF Special Directories

            If InStr(tFld.Path, "\BAD") = 0 _
            And InStr(tFld.Path, "\PROCESSED") = 0 _
            And InStr(tFld.Path, "\FAILED") = 0 _
            And InStr(tFld.Path, "(TEMP)_") = 0 Then
                    
                    ReDim Preserve arrSubDirectoriesFound(nDirs)
                    arrSubDirectoriesFound(nDirs) = tFld.Path
                    funcFindSubdirectories = funcFindSubdirectories + funcFindSubdirectories(tFld.Path, nDirs, arrSubDirectoriesFound)
                    
            End If
            
       Next
    End If
         
    Exit Function
    
Catch:  FileName = ""
       Resume Next
End Function


Public Function funcGetFilesCount(ByRef sFolderPath, ByRef sImportFilePattern)

      '*** 2020-04-14 - Jacob - Get Count of Files in Entire Sub-directory under the sFolderPath, based on sImportFilePattern

    Dim oParentFolder, oSubFolder
    Dim fso 'As Folder
    Dim ParentFile
    Dim SubFolderFile
    
   Set fso = New FileSystemObject

    'funcGetFilesCount = 0

   'On Error Resume Next
    Set oParentFolder = fso.GetFolder(sFolderPath)
    
'    For Each ParentFile In oParentFolder.files
'
'        If InStr(oParentFolder.Path, "\BAD'\") = 0 _
'        And InStr(oParentFolder.Path, "\PROCESSED") = 0 _
'        And InStr(oParentFolder.Path, "\FAILED\") = 0 Then
'
'            If Left(ParentFile.name, 1) <> "~" _
'            And InStr(UCase(sImportFilePattern), UCase(fso.GetExtensionName(ParentFile))) Then
'                funcGetFilesCount = funcGetFilesCount + 1
'                ''''''Debug.Print funcGetFilesCount & " - PARENT - " & ParentFile.name
'
'            End If
'
'        End If
'
'    Next


    For Each oSubFolder In oParentFolder.SubFolders
        '2022-11-17 - Jacob - Added Check for "(TEMP)_"
        If InStr(oSubFolder.Path, "\BAD") = 0 _
        And InStr(oSubFolder.Path, "\PROCESSED") = 0 _
        And InStr(oSubFolder.Path, "\FAILED") = 0 _
        And InStr(oSubFolder.Path, "(TEMP)_") = 0 Then
'*TO-DO:  See if possible to speed up directory / file list
             '*** 2020-04-14 - Jacob - Added the PathMatchSpecW Windows API call to Match the FilePattern more accurately
            For Each SubFolderFile In oSubFolder.Files
                If Left(SubFolderFile.name, 1) <> "~" _
                And PathMatchSpecW(StrPtr(SubFolderFile.name), StrPtr(sImportFilePattern)) Then
                    funcGetFilesCount = funcGetFilesCount + 1
                    ''''''Debug.Print funcGetFilesCount & " - SUBFOL - " & SubFolderFile.name
                End If
            Next
            
            '2022-11-17 - Jacob - Moved INTO the IF, so we DON'T count files INSIDE these directories
            funcGetFilesCount = funcGetFilesCount + funcGetFilesCount(oSubFolder.Path, sImportFilePattern)
            
        End If
        
    Next
    
End Function




Public Sub funcDeleteEmptyFolders(ByVal strRootFolderPathNotToDelete As String, ByVal strFolderPath As String)
   Dim fsoSubFolders 'As Folders
   Dim fsoFolder 'As Folder
   Dim fsoSubFolder 'As Folder
   
   Dim strPaths()
   Dim lngFolder As Long
   Dim lngSubFolder As Long
      
   DoEvents
   
   Set m_fsoObject = New FileSystemObject
   If Not m_fsoObject.FolderExists(strFolderPath) Then
        Exit Sub
   End If
   
   Set fsoFolder = m_fsoObject.GetFolder(strFolderPath)
   
   On Error Resume Next
   
   'Has sub-folders
   If fsoFolder.SubFolders.Count > 0 Then
        lngFolder = 1
        ReDim strPaths(1 To fsoFolder.SubFolders.Count)
        'Get each sub-folders path and add to an array
        For Each fsoSubFolder In fsoFolder.SubFolders
            strPaths(lngFolder) = fsoSubFolder.Path
            lngFolder = lngFolder + 1
        Next fsoSubFolder
        
        lngSubFolder = 1
        'Recursively call the function for each sub-folder
        Do While lngSubFolder < lngFolder
           Call funcDeleteEmptyFolders(strRootFolderPathNotToDelete, strPaths(lngSubFolder))
           lngSubFolder = lngSubFolder + 1
        Loop
    End If
   
    'Delete if NOT the ROOT Folder and Path has No sub-folders or files
    If Trim(fsoFolder.Path) <> Trim(strRootFolderPathNotToDelete) Then
        If fsoFolder.Files.Count = 0 _
        And fsoFolder.SubFolders.Count = 0 Then
                funcWriteToDebugLog "funcDeleteEmptyFolders", "*** DELETING EMPTY FOLDER; " & fsoFolder.Path
                fsoFolder.Delete
        End If
    Else
            funcWriteToDebugLog "funcDeleteEmptyFolders", "*** SKIPPING DELETE OF ROOT FOLDER; " & fsoFolder.Path
    
    End If
End Sub


Function funcFieldExists(ByVal rs, ByVal FieldName) As Boolean

    On Error Resume Next
    funcFieldExists = rs.Fields(FieldName).name <> ""
    If Err <> 0 Then
        funcFieldExists = False
    End If
    Err.Clear

End Function

Public Function APIFileCopy(Src As String, Dest As String, Optional FailIfDestExists As Boolean) As Boolean

'*** 2021-02-12 - Jacob - PURPOSE: COPY FILES via Windows 32-Bit API

'PARAMETERS: src: Source File (FullPath)
                            'dest: Destination File (FullPath)
                            'FailIfDestExists (Optional):
                            'Set to true if you don't want to
                            'overwrite the destination file if
                             'it exists
            
            'Returns (True if Successful, false otherwise)
            
'EXAMPLE:
  'dim bSuccess as boolean
  'bSuccess = APIFileCopy ("C:\MyFile.txt", "D:\MyFile.txt")
  
    '*** 2022-06-07 - Jacob - Added
    Dim strDirectoryStructureToCreate As String
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

  
  strDirectoryStructureToCreate = fso.GetParentFolderName(Dest)
  
  funcWriteToDebugLog "FunctionDeclararions.APIFileCopy()", "funcCreateDirectoryStructure(" & strDirectoryStructureToCreate & ")"
  funcCreateDirectoryStructure strDirectoryStructureToCreate

    Dim lRet As Long
    lRet = CopyFile(Src, Dest, FailIfDestExists)
    APIFileCopy = (lRet > 0)

End Function


Public Function funcRunningInIDE() As Boolean
    
    '*** 2021-02-12 - Jacob - PURPOSE: Check if program is running in the IDE

  On Error Resume Next
  Debug.Print 0 / 0
  funcRunningInIDE = Err.Number <> 0
  
End Function



Function funcGetFileNameFromPath(strFullPath As String) As String

    funcGetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
    
End Function



Function funcIniAndCfgFileSetup()

    
    '*** LOCATE AND CORRECT Imaging101Client.INI LOCATION
    
   '*** 2022-11-08 - Jacob - Moved from FormLoad in both Imaging101 and AutoImport
   '                                        to Standardize the moving and confiiguration of Imaging101Client.ini and Image.ini Files.
   '                                        Also to
   
   funcWriteToDebugLog "funcIniAndCfgFileSetup", "Locating 'Imaging101Client.ini' File"
    
    Dim strAppDataDir As String
    Dim strLocalAppDataDir As String
    
    Dim datCurrentFileDate As Date
    Dim datNewestFileDate As Date
    Dim strNewestIniFilePath As String
    Dim strNewLocationForFiles As String
    
    Dim arrImaging101ClientIniFilesFound() As String
    Dim arrImageIniFilesFound() As String
    Dim arrImageCfgFilesFound() As String

    Dim nDirs As Long
    Dim nFiles As Long
    Dim lSize As Currency

    'Get the AppData directories
    strAppDataDir = Environ$("AppData")
    strLocalAppDataDir = Environ$("LocalAppData")
    strNewLocationForFiles = strLocalAppDataDir & "\Imaging101"
    

    'Set the NEW Location for the Imaging101Client.INI file
    RegFileName = strNewLocationForFiles & "\Imaging101Client.ini"

    
     
     'Write variable contents to debug log
    funcWriteToDebugLog "funcIniAndCfgFileSetup", "  "
    funcWriteToDebugLog "funcIniAndCfgFileSetup", "strAppDataDir = " & strAppDataDir
    funcWriteToDebugLog "funcIniAndCfgFileSetup", "strLocalAppDataDir = " & strLocalAppDataDir
    funcWriteToDebugLog "funcIniAndCfgFileSetup", "strNewLocationForFiles = " & strNewLocationForFiles
    funcWriteToDebugLog "funcIniAndCfgFileSetup", "RegFileName = " & RegFileName
    funcWriteToDebugLog "funcIniAndCfgFileSetup", " "

        
    Dim ofs As Scripting.FileSystemObject
    Dim ofile
    Dim intIniFileLoop
    
    
    '*********************************************************************************************
    'Check if the Imaging101Client.Ini file is already in the Local AppData Dir
    'Define FileSystemObject
    
    
    Set ofs = New Scripting.FileSystemObject
    
    If Not ofs.FolderExists(strNewLocationForFiles) Then
        funcWriteToDebugLog "funcIniAndCfgFileSetup", "* Creating Directory: " & strNewLocationForFiles
        funcCreateDirectoryStructure strNewLocationForFiles
    End If
    

    If Not ofs.FileExists(RegFileName) Then
 

        
        funcWriteToDebugLog "funcIniAndCfgFileSetup", "* Locating copies of Imaging101Client.INI"
    
         lSize = funcFindFile(strAppDataDir, "Imaging101Client.ini", nDirs, nFiles, arrImaging101ClientIniFilesFound)
       
        On Error Resume Next
        
        Debug.Print UBound(arrImaging101ClientIniFilesFound())
        Debug.Print Err.Number & " - " & Err.Description
        
        nFiles = nFiles + 1
        
        'Add the WINDOWS directory instance to the Array
        ReDim Preserve arrImaging101ClientIniFilesFound(nFiles)
        arrImaging101ClientIniFilesFound(nFiles) = "C:\Windows\Imaging101Client.INI"
        
        'Find the last modified file
        For intIniFileLoop = 1 To nFiles
        
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "** Found: " & arrImaging101ClientIniFilesFound(intIniFileLoop)
            
            If ofs.FileExists(arrImaging101ClientIniFilesFound(intIniFileLoop)) Then
                 Set ofile = ofs.GetFile(arrImaging101ClientIniFilesFound(intIniFileLoop))
                datCurrentFileDate = ofile.DateLastModified
                If datCurrentFileDate > datNewestFileDate Then
                    datNewestFileDate = ofile.DateLastModified
                    strNewestIniFilePath = ofile.Path
                    'Move the last modified INI to the Local dir.
                    ofs.CopyFile strNewestIniFilePath, RegFileName
                End If
            End If
            
            If UCase(arrImaging101ClientIniFilesFound(intIniFileLoop)) <> UCase("C:\Windows\Imaging101Client.INI") _
            And arrImaging101ClientIniFilesFound(intIniFileLoop) <> RegFileName Then
                ofs.DeleteFile arrImaging101ClientIniFilesFound(intIniFileLoop), True
            End If
            
            
        Next
        
    
    End If 'Not ofs.FileExists(RegFileName)
    
    Set ofs = Nothing
    
    
    '*********************************************************************************************

  
    
    
        '*********************************************************************************************
        '*** CORRECT PROBLEM with the Image.INI file.
        
        '*** 2022-11-08 - Jacob - Changed Method to Detect if Running in IDE, and Added Writes to DEBUG LOG
        
        funcWriteToDebugLog "funcIniAndCfgFileSetup", "Configure image.ini File"
        
        Dim strIImageIniFilePath As String
        
        strIImageIniFilePath = "C:\Open Text\Imagenation\image.ini"
        
        '*** Check if Image.ini File Exists in AppPath, otherwise assume we're in IDE
        If Dir(strIImageIniFilePath) <> "" Then

            'Running in IDE
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "Found " & strIImageIniFilePath & " | " & "Running in IDE"
            
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "Set 'System Dir' in 'C:\Open Text\Imagenation\image.ini' to 'C:\Open Text\Imagenation'"
            result = WritePrivateProfileString("System", "System Dir", "C:\Open Text\Imagenation\", "C:\Open Text\Imagenation\image.ini")
            
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "Set 'Interface Profile File' in 'C:\Open Text\Imagenation\image.ini' to 'C:\Open Text\Imagenation\image.cfg'"
            result = WritePrivateProfileString("System", "Interface Profile File", "C:\Open Text\Imagenation\image.cfg", "C:\Open Text\Imagenation\image.ini")
            
        Else
        
            'Running as EXE
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "Did NOT Find " & strIImageIniFilePath & " | " & "Running as EXE"
            
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "Set 'System Dir' in '" & App.Path & "\image.ini' to '" & App.Path & "'"
            result = WritePrivateProfileString("System", "System Dir", App.Path, App.Path & "\image.ini")
            
            funcWriteToDebugLog "funcIniAndCfgFileSetup", "Set 'Interface Profile File' in '" & App.Path & "\image.ini' to '" & App.Path & "\image.cfg'"
            result = WritePrivateProfileString("System", "Interface Profile File", App.Path & "\image.cfg", App.Path & "\image.ini")
            
        End If

End Function

