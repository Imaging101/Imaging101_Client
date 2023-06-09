Attribute VB_Name = "PrivateProfileModule"
' This first line is the declaration from win32api.txt
' Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName$, _
         ByVal lpszKey$, _
         ByVal lpszDefault$, _
         ByVal lpszReturnBuffer$, _
         ByVal cchReturnBuffer&, _
         ByVal lpszFile$)
        
'   ***********************************************
'   ****** WritePrivateProfileString Notes ********
'
'   *** To DELETE/REMOVE a "KeyName" use vbNullString for the Key Value:
'   Call WritePrivateProfileString(sSection, _
                                  sKeyName, _
                                  vbNullString, _
                                  sIniFile)
'   *** To DELETE/REMOVE an Entire "Section" use vbNullString
'       for both Section and Key Values:
'   Call WritePrivateProfileString(sSection, _
                                  vbNullString, _
                                  vbNullString, _
                                  sIniFile)
Public Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As String, _
         ByVal lpString As String, _
         ByVal lpFileName As String) As Long
          
' VBGetPrivateProfileString
'   An example of modular programming - This provides a safer
'   interface to GetPrivateProfileString
'
Function VBGetPrivateProfileString(section$, Key$, File$) As String
    Dim KeyValue$
    #If Win32 Then
        Dim characters As Long
    #Else
        Dim characters As Integer
    #End If
    
    '3/23/2005 Jacob - Changed String$ from 128 to 4096 to allow loading entire Section
    KeyValue$ = String$(4096, 0)
    
    '3/23/2005 Jacob - Changed ReturnBuffer Size from 127 to 4095 to allow loading entire Section
    characters = GetPrivateProfileStringByKeyName(section$, Key$, "", KeyValue$, 4095, File$)

'    If characters > 0 Then
        If Left(KeyValue$, 1) <> Chr(0) Then
            KeyValue$ = Left$(KeyValue$, characters)
        Else
            KeyValue$ = ""
        End If
'    End If
        
    VBGetPrivateProfileString = KeyValue$
    

End Function

