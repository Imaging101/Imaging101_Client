VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmI101AIM 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCommandLine 
      Height          =   1575
      Left            =   120
      LinkTopic       =   "Imaging101.EXE|I101AIM"
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmI101AIM.frx":0000
      Top             =   1200
      Width           =   4215
   End
   Begin MSWinsockLib.Winsock WinsockI101AIM 
      Left            =   3840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Imaging101 Application Integration"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmI101AIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub I101AIM_Winsock_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo ERROR_HANDLER
    
    '*** This Event is triggered by a "Winsock1.Connect" request from the Client
    '***   INSTEAD OF the "socket_DataArrival" event
    
    'Now accept the new connection
    'A connection was requested from the server.
    'Make sure this is control 0 in the array.
    'This is the only one that can accept connections.
    If Index = 0 Then
    
        'Open a NEW Socket
        iSocketIndex = iSocketIndex + 1
        Load Socket(iSocketIndex)
            
        Socket(iSocketIndex).LocalPort = txtLocalPort
        Socket(iSocketIndex).Accept requestID
        Socket(iSocketIndex).SendData "Connection Accepted... Request ID = " & requestID
    
    End If
    
Exit Sub

ERROR_HANDLER:

        MsgBox = "ERROR: " & "|" & " Socket_ConnectionRequest " & "|" & Err.Number & "|" & Err.Description & "|"
    
End Sub

Private Sub I101AIM_Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub txtCommandLine_Change()
 DoEvents
    
    
''    Dim objClientSearch As New ClientSearch
    
    'Split Command Line Arguments into Subcommands
    strCommandArray() = Split(txtCommandLine, "-")
    
    'Split First Sub-command argument to get Search Name
    strSubCommandArray() = Split(strCommandArray(1), "=")
    
    ' Add the command line to the TOP of the Outputn Text AFTER we Split it!
    txtCommandLine.Text = "Command Line:" & vbCrLf & "   " & txtCommandLine.Text
    
    ' Add the command breakdown text
    txtCommandLine.Text = txtCommandLine.Text & vbCrLf & vbCrLf & "Command Parameter Breakdown:"
    
    ' Check for UserID
    If UCase(strSubCommandArray(0)) = "U" And strSubCommandArray(1) <> "" Then
        txtCommandLine.Text = txtCommandLine.Text & vbCrLf & "    UserID = " & Trim(strSubCommandArray(1))
    Else
        MsgBox "User ID NOT found!"
        Exit Sub
    End If
    DoEvents
    
    For intCommandLoop = 2 To UBound(strCommandArray(), 1)
        'Split Each Sub-command argument after Search Name to get Field and Value
        strSubCommandArray() = Split(strCommandArray(intCommandLoop), "=")
        txtCommandLine.Text = txtCommandLine.Text & vbCrLf & "    Parameter= " & Trim(strSubCommandArray(0)) & ",    Value= " & Trim(strSubCommandArray(1))
        DoEvents
    Next
    
''    objClientSearch.Execute True
    DoEvents
    
''    objClientSearch.Clear
    
'    Can now execute another query
End Sub
