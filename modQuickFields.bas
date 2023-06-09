Attribute VB_Name = "modQuickFields"
' *******************************************************************************
' *** CHANGE CONTROL SECTION - BEGIN
' *******************************************************************************
' 7/31/98 JR:  Added "GetWindowsDirectory" Function Declare
' 7/31/98 JR:  Added Public Function ConvertCtoVBString to convert the returned
'              "Buffer" from the GetWindowsDirectory Function to a VB String

' *******************************************************************************
' *** CHANGE CONTROL SECTION - END
' *******************************************************************************

'**********************************
'**  Function Declarations:
'**  Declare necessary API routines:

Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long

Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
         ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal wFlags As Long) As Long

Declare Function GetCurrentDirectory& Lib "kernel32" Alias "GetCurrentDirectoryA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String)

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long) As Long
         
         
'*** 2021-10-20 - Jacob - Added declares to find windows that CONTAIN text
Private Declare Function EnumWindows Lib "user32" _
   (ByVal lpEnumFunc As Long, _
    ByVal lparam As Long) As Long

Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" _
   (ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long

'Custom structure for passing in the parameters in/out of the hook enumeration function
'Could use global variables instead, but this is nicer.
Private Type FindWindowParameters

    strTitle As String  'INPUT
    hwnd As Long     'OUTPUT

End Type
         
         
         
         
' Public Objects
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Global AppWinTitle As String
Global WindowDetected As Boolean   ' Flag for final release.
Global hwnd As Long
Global ForcedByIndexTimer As Boolean
Global numTotalSubfolders As Integer
Global NumIndexCount As Integer
Global txtLocalScan As String



'

Sub DetectWindow(AppWinTitle)
' Procedure detects a running Window and registers it.
' The "AppWinTitle" string can be either the Window Title or the Class
'   this function will check both just in case.

    Const WM_USER = 1024
    On Error Resume Next    ' Defer error trapping.

    ' If Window is running this API call returns its handle.
    ' First try using the Window Title
    hwnd = FindWindow(vbNullString, AppWinTitle)
    If hwnd = 0 Then
        ' We didn't find the window Try using CLASS instead
        hwnd = FindWindow(AppWinTitle, vbNullString)
    End If
    If hwnd = 0 Then    ' 0 means FPmulti not running.
        WindowDetected = False
        Exit Sub
    Else
        ' FPmulti is running so use the SendMessage API
        ' function to enter it in the Running Object Table.
        WindowDetected = True
        Err.Clear   ' Clear Err object in case error occurred.
        SendMessage hwnd, WM_USER + 18, 0, 0
    End If
End Sub

Function DetectWindowWait(AppWinTitle)
   WindowDetected = False
   txtMessageText = "Looking for " & AppWinTitle
    
    ' Display the Message Window "Always ON TOP"
    FormOnTop frmMessageForm.hwnd, True
   
   While Not WindowDetected
        DetectWindow AppWinTitle
        TimePause 2
        numMessageCounter = numMessageCounter + 1
   Wend
   frmMessageForm.Hide
End Function

    
'Public Function TimePause(PauseTime)
'' PauseTime parameter must be in Seconds
'Dim Start, Finish, TotalTime, OptikaAppID
'    Start = Timer   ' Set start time.
'    Do While Timer < Start + PauseTime
'        DoEvents    ' Yield to other processes.
'    Loop
'
'End Function
   
    
Public Sub FormOnTop(HANDLE As Long, OnTop As Boolean)
    Dim wFlags As Long, PosFlag As Long
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or _
        SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Select Case OnTop
        Case True
            PosFlag = HWND_TOPMOST
        Case False
            PosFlag = HWND_NOTOPMOST
    End Select
    SetWindowPos HANDLE, PosFlag, 0, 0, 0, 0, wFlags
End Sub


Public Function ConvertCtoVBString(InString As String) As String
    InString = Trim(InString)
    ConvertCtoVBString = Left(InString, Len(InString) - 1)
End Function


Public Function GetSysDir() As String
    Dim Temp As String * 256
    Dim X As Integer
    X = GetSystemDirectory(Temp, Len(Temp)) ' Make API Call (Temp will hold return value)
    GetSysDir = Left$(Temp, X)              ' Trim Buffer and return string
End Function

Public Function GetWinDir() As String
    Dim Temp As String * 256
    Dim X As Integer
    X = GetWindowsDirectory(Temp, Len(Temp)) ' Make API Call (Temp will hold return value)
    GetWinDir = Left$(Temp, X)               ' Trim Buffer and return string
End Function




Public Function funcFindWindowLike(strWindowTitle As String) As Long

    'We'll pass a custom structure in as the parameter to store our result...
    Dim Parameters As FindWindowParameters
    Parameters.strTitle = strWindowTitle ' Input parameter
    
    Call EnumWindows(AddressOf EnumWindowProc, VarPtr(Parameters))
    
    funcFindWindowLike = Parameters.hwnd
    
End Function

Private Function EnumWindowProc(ByVal hwnd As Long, _
                               lparam As FindWindowParameters) As Long
   
   Dim strWindowTitle As String

   strWindowTitle = Space(260)
   Call GetWindowText(hwnd, strWindowTitle, 260)
   strWindowTitle = TrimNull(strWindowTitle) ' Remove extra null terminator
                                          
   '*** 2022-04-13 - Jacob - Added check for (Not strWindowTitle Like "Viewer*") to ignore the ChildForm in the Imaging101 Viewer
'   If (Not strWindowTitle Like "Viewer*") And (strWindowTitle Like lparam.strTitle) Then
   If (strWindowTitle Like lparam.strTitle) Then
   
        lparam.hwnd = hwnd 'Store the result for later.
        EnumWindowProc = 0 'This will stop enumerating more windows
   
   Else

        EnumWindowProc = 1

   End If
                           
End Function

Private Function TrimNull(strNullTerminatedString As String)

    Dim lngPos As Long

    'Remove unnecessary null terminator
    lngPos = InStr(strNullTerminatedString, Chr$(0))
   
    If lngPos Then
        TrimNull = Left$(strNullTerminatedString, lngPos - 1)
    Else
        TrimNull = strNullTerminatedString
    End If
   
End Function
