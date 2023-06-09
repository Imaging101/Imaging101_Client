Attribute VB_Name = "mod_ShellExecute"
'Example by Joel (crashcode6@hotmail.com)
'This example requires a command button (Command1)
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Const SE_ERR_NOASSOC = 31
Const sOperation As String = "open"     ' Constants for shell operations
Const sRun As String = "RUNDLL32.EXE"
Const sParameters As String = "shell32.dll,OpenAs_RunDLL "
Public Function shelldoc(ByVal sFile As String)
    Dim sPath As String, RetVal As Long, _
    lRet As Long
    lRet = ShellExecute(GetDesktopWindow(), sOperation, sFile, _
                        vbNullString, vbNullString, SW_SHOWNORMAL)
    If lRet = SE_ERR_NOASSOC Then ' No association exists
        'Create a buffer
        sPath = Space(255)
        'Get the system directory
        RetVal = GetSystemDirectory(sPath, 255)
        'Remove all unnecessary chr$(0)'s
        'and move on the stack
        sPath = Left$(sPath, RetVal)
    
        lRet = ShellExecute(GetDesktopWindow(), "open", sRun, _
                            sParameters + sFile, sPath, SW_SHOWNORMAL)
    End If
End Function


