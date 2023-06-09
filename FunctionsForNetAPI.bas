Attribute VB_Name = "FunctionsForNetAPI"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Windows type used to call the Net API
Private Type USER_INFO_10
   usr10_name          As Long
   usr10_comment       As Long
   usr10_usr_comment   As Long
   usr10_full_name     As Long
End Type

'private type to hold the actual strings displayed
Private Type USER_INFO
   Name          As String
   fullname     As String
   Comment       As String
   usrcomment   As String
End Type

Private Const ERROR_SUCCESS As Long = 0&
Private Const MAX_COMPUTERNAME As Long = 15
Private Const MAX_USERNAME As Long = 256

Private Declare Function NetUserGetInfo Lib "netapi32" _
   (lpServer As Byte, _
   username As Byte, _
   ByVal Level As Long, _
   lpBuffer As Long) As Long
   
Private Declare Function NetApiBufferFree Lib "netapi32" _
  (ByVal Buffer As Long) As Long

Private Declare Function GetUserName Lib "advapi32" _
   Alias "GetUserNameA" _
  (ByVal lpBuffer As String, _
   nSize As Long) As Long
   
Private Declare Function GetComputerName Lib "kernel32" _
   Alias "GetComputerNameA" _
  (ByVal lpBuffer As String, _
   nSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nBytes As Long)

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long

Private Declare Function StrLen Lib "kernel32" _
   Alias "lstrlenW" _
  (ByVal lpString As Long) As Long


Public Function rgbGetComputerName() As String

  'return the name of the computer
   Dim tmp As String
   
   tmp = Space$(MAX_COMPUTERNAME + 1)
    
   If GetComputerName(tmp, Len(tmp)) <> 0 Then
      rgbGetComputerName = TrimNull(tmp)
   End If
   
End Function


Private Function TrimNull(item As String)

   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function


Public Function rgbGetUserName() As String

  'return the name of the user
   Dim tmp As String
   
   tmp = Space$(MAX_USERNAME)
   
   If GetUserName(tmp, Len(tmp)) Then
      rgbGetUserName = TrimNull(tmp)
   End If

End Function


Private Function GetUserNetworkInfo(bServername() As Byte, _
                                    bUsername() As Byte) As USER_INFO
   
   Dim usrapi As USER_INFO_10
   Dim buff As Long
   
   If NetUserGetInfo(bServername(0), bUsername(0), 10, buff) = ERROR_SUCCESS Then
      
     'copy the data from buff into the
     'API user_10 structure
      CopyMemory usrapi, ByVal buff, Len(usrapi)
      
     'extract each member and return
     'as members of the UDT
      GetUserNetworkInfo.Name = GetPointerToByteStringW(usrapi.usr10_name)
      GetUserNetworkInfo.fullname = GetPointerToByteStringW(usrapi.usr10_full_name)
      GetUserNetworkInfo.Comment = GetPointerToByteStringW(usrapi.usr10_comment)
      GetUserNetworkInfo.usrcomment = GetPointerToByteStringW(usrapi.usr10_usr_comment)
   
      NetApiBufferFree buff
   
   End If
   
End Function


Private Function GetPointerToByteStringW(lpString As Long) As String
  
   Dim buff() As Byte
   Dim nSize As Long
   
   If lpString Then
   
     'its Unicode, so mult. by 2
      nSize = lstrlenW(lpString) * 2
      
      If nSize Then
         ReDim buff(0 To (nSize - 1)) As Byte
         CopyMemory buff(0), ByVal lpString, nSize
         GetPointerToByteStringW = buff
     End If
     
   End If
   
End Function
'--end block--'




