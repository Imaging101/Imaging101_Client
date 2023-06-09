Attribute VB_Name = "modQuickIndex"
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

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
         (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, lParam As Long) As Long
          
Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
         ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal wFlags As Long) As Long

Declare Function GetCurrentDirectory& Lib "kernel32" Alias "GetCurrentDirectoryA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String)

 Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long) As Long
         
         
' Public Objects
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


'Global objects
Global OraSession As Object
Global OraDatabase As Object
Global SqlDynaset As Object

' Not required to be constant.
Global SqlQuery$        '= "select * from emp"
Global DatabaseName$    '= "Exampledb"
Global Connect$         '= "scott/tiger"

Global Const strMultiApplicationName$ = "HOSPITAL"

Global Const WarnFirstEmp$ = "You are already on the first employee."
Global Const WarnLastEmp$ = "You are already on the last employee."

Global Const DUPLICATE_KEY = 1000


Global ErrMsg$
Global ErrNum As Integer

'From ORACONST.TXT
Global Const ORADYN_NOCACHE = &H8&

' Global Const CHR = &H8&

Global DoUpdate As Boolean
Global DoAdd As Boolean

Global NoChanged As Boolean

Global AppWinTitle As String
Global WindowDetected As Boolean   ' Flag for final release.
Global hwnd As Long
Global ForcedByIndexTimer As Boolean
Global numTotalSubfolders As Integer
Global HostFieldsDidChange As Boolean
Global FPmultiPreviousScreenLocation As String
Global MultiLoadSubFoldersComplete As Boolean
Global NumIndexCount As Integer
Global txtLocalScan As String
'Global Buffer As String      ' Defined in the QAPI_Declarations Module
Global LongVar As Long
Global txtOptikaIniLocation As String
Global txtOptikaAppDrive As String
Global txtOptikaAppDir As String
Global txtDatabaseDriveDir As String
Global txtOptikaStartIn As String
Global Result As String
Global SessionFieldMapFormLoaded As Boolean


'*** Set Defaults for Registry Entries ***
Global Const RegAppname As String = "QuickIndex"
Global Const RegSectionName As String = "Settings"
Global Const RegFileName As String = "QuickIndex.INI"

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

    
Function TimePause(PauseTime)
' PauseTime parameter must be in Seconds
Dim Start, Finish, TotalTime, OptikaAppID
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
       
End Function
   
    
Public Sub FormOnTop(Handle As Long, OnTop As Boolean)
    Dim wFlags As Long, PosFlag As Long
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or _
        SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Select Case OnTop
        Case True
            PosFlag = HWND_TOPMOST
        Case False
            PosFlag = HWND_NOTOPMOST
    End Select
    SetWindowPos Handle, PosFlag, 0, 0, 0, 0, wFlags
End Sub


Public Function ConvertCtoVBString(InString As String) As String
    InString = Trim(InString)
    ConvertCtoVBString = Left(InString, Len(InString) - 1)
End Function

