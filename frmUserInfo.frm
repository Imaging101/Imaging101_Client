VERSION 5.00
Begin VB.Form frmUserInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Get User Name"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frmUserInfo.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Network User Info"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2265
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Network User Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   0
      Width           =   2625
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   4920
      TabIndex        =   14
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Full Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Computer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   name          As String
   fullname     As String
   comment       As String
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
  (ByVal buffer As Long) As Long

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


Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

   Text1.Text = rgbGetUserName()
   Text2.Text = rgbGetComputerName()
   
  Command1_Click
  
End Sub


Private Sub Command1_Click()

   Dim usr As USER_INFO
   Dim bUsername() As Byte
   Dim bServername() As Byte
   Dim tmp As String
  
  'This assures that both the server
  'and user params have data.  Must
  'perform actual comparisons to 0 since
  'results such as "2 And 5" equate to False!
   If Len(Text1.Text) > 0 And _
      Len(Text2.Text) > 0 Then
   
      bUsername = Text1.Text & Chr$(0)
   
     'This demo uses the current machine as the
     'server param, which works on NT4 and Win2000.
     'If connected to a PDC or BDC, pass that
     'name as the server, instead of the return
     'value from GetComputerName().
      tmp = Text2.Text
   
     'assure the server string is properly formatted
      If Len(tmp) > 0 Then
      
         If InStr(tmp, "\\") Then
               bServername = tmp & Chr$(0)
         Else
            bServername = "\\" & tmp & Chr$(0)
         End If
      
      End If

     'Return the user information for the passed
     'user. The return values are assigned directly
     'to the non-API USER_INFO data type that we
     'defined (I prefer UDTs). Alternatively, if
     'you're a 'classy' sort of guy,  the return
     'values could be assigned directly to properties
     'in the function.
      usr = GetUserNetworkInfo(bServername(), bUsername())
      
      Text3.Text = usr.name
      
     'The call may or may not return the
     'full name, comment or usrcomment
     'members, depending on the user's
     'listing in User Manager.
      Text4.Text = usr.fullname
      Text5.Text = usr.comment
      Text6.Text = usr.usrcomment
   
   End If

End Sub


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
      GetUserNetworkInfo.name = GetPointerToByteStringW(usrapi.usr10_name)
      GetUserNetworkInfo.fullname = GetPointerToByteStringW(usrapi.usr10_full_name)
      GetUserNetworkInfo.comment = GetPointerToByteStringW(usrapi.usr10_comment)
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



