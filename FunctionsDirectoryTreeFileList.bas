Attribute VB_Name = "FunctionsDirectoryTreeFileList"
'VB6 standard module [PathLog] 12-25-99 9:00 >1-19-00 15:09
' * PROGRAMMER  : John Almand
' * WEB SITE    : WWW.SHALLIZAR.COM
' * E-MAIL      : WEBMASTER@SHALLIZAR.COM
'
'This procedure has been excerpted from the PathLogger.bas
'module of the MVIEW project. It is the core algorithm of
'Pathlogger.bas. It does not have the functionality and
'features of the full PathLogger module. For more information
'on the complete PathLogger.bas module, contact
'www.shallizar.com
'
'PURPOSE:
' LOG THE FILES IN A BRANCH OF A DIRECTORY TREE
'
' Elucidation:
' Make an array of strings of the names of all the files
' in a folder and all of its subfolders.
' Make a seperate array of strings of the subfolder names,
' including the parent folder.
' Make an array of pointers linking the two string arrays.
'
'SPECIAL CHARACTERISTICS:
' NON-RECURSIVE
' This procedure iterates through the subfolders without
' repeatedly calling itself. It avoids the extra overhead
' and potential stack overflow of recursive procedures.
'
'REQUIRES:
' Reference to Microsoft Scripting Runtime
'
'#########################################################
Option Explicit

Private objFSO As FileSystemObject
'Private objFiles As Files
'Private objFile As File
'Private objFolders As Folders
'Private objFolder As Folder

Private objFiles As Object
Private objFile As Object
Private objFolders As Object
Private objFolder As Object

Private lngFNAMEScntr As Long





Public Sub PathLogInit()
'PURPOSE: Create & initialize objFSO as a
'FileSystemObject object.
'Call this procedure from the LOAD event of the Startup form.
 Set objFSO = New FileSystemObject
End Sub

Public Sub LogPath(strPARENT As String, strFNAMES() As String, _
                   lngPptrs() As Long, strPaths() As String)
'PURPOSE: Make a list of all the files in a Folder and all of
'its subfolders.

'Typically used with a DirListBox control on a form. A command
'button should be provided to run this procedure. After selecting
'a folder in the DirListBox, the user would click the button to
' log the folder and its subfolders (which is the Path).

'The calling form or module passes a string of the path,
'a variable array of strings for the subfolder names, a
'variable array of long integers for the pointers to the
'folder names, & a variable array of strings for the filenames.

'EXAMPLE:
' LogPath Dir1.Path, arrayFolders, arrayPointers, arrayFilenames

'The result in arrayFilenames can be used to fill a listbox.
'A filename selected in a listbox can be fully referenced by
'combining it with its parent folder name,
' thus: arrayFolders(arrayPointers(I)) & arrayFilenames(I)

 Dim lngTopIndex As Long
 Dim lngPathIndex As Long
 Dim strNextPath As String
 
Screen.MousePointer = MousePointerConstants.vbArrowHourglass


'if path is invalid, exit
 If Not objFSO.FolderExists(strPARENT) Then
     Screen.MousePointer = MousePointerConstants.vbDefault

    Exit Sub
 End If
 
' "seed" the loop
 lngTopIndex = 0
 lngPathIndex = 0
 lngFNAMEScntr = 0
 ReDim strPaths(0)
 strPaths(0) = IFBACKSLASH(strPARENT)

'reset the filename and pathpointer arrays just in case no
'filenames match the wildcard pattern
 ReDim strFNAMES(0)
 ReDim lngPptrs(0)
 
 Do

' Add subfolders, if any, to array
  Set objFolders = objFSO.GetFolder(strPaths(lngPathIndex)).SubFolders
   For Each objFolder In objFolders
    lngTopIndex = lngTopIndex + 1
    ReDim Preserve strPaths(lngTopIndex)
    strPaths(lngTopIndex) = strPaths(lngPathIndex) & _
                            objFolder.name & "\"
    DoEvents
   Next
  
' Add filenames, if any, to array
  Set objFiles = objFSO.GetFolder(strPaths(lngPathIndex)).Files
  For Each objFile In objFiles
    ReDim Preserve strFNAMES(lngFNAMEScntr)
    strFNAMES(lngFNAMEScntr) = objFile.name
    ReDim Preserve lngPptrs(lngFNAMEScntr)
    lngPptrs(lngFNAMEScntr) = lngPathIndex
    lngFNAMEScntr = lngFNAMEScntr + 1
    DoEvents
  Next
    
' Point to next entry in subfolder array
  lngPathIndex = lngPathIndex + 1
  
  DoEvents
  
' If there are no more subfolders, exit
  Loop Until lngPathIndex > lngTopIndex
     
  Screen.MousePointer = MousePointerConstants.vbDefault
  
End Sub

Public Property Get Count() As Long
' Get the number of filenames found
 Count = lngFNAMEScntr
End Property

Public Sub Terminate()
' Clear all the objects from memory
 Set objFSO = Nothing
 Set objFiles = Nothing
 Set objFile = Nothing
 Set objFolders = Nothing
 Set objFolder = Nothing
End Sub

Private Function IFBACKSLASH(strX As String) As String
' function for fixing the DOS path of a root directory
 IFBACKSLASH = IIf(Right(strX, 1) = "\", strX, strX & "\")
End Function



