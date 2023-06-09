VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmFTP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   5700
   ClientWidth     =   13560
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   13560
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11880
      Picture         =   "FTP.frx":0A02
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox lstStatus 
      Height          =   3660
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   13335
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   11880
      TabIndex        =   4
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FTP File Transfer Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblRESPONSE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   8235
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msCurrentFile As String
Private msNewFile As String
Private msFTPProcessState As String

  
Friend Sub FTPFile(ByVal sFTPServer As String, _
                               ByVal sFTPCommand As String, _
                               ByVal sFTPUser As String, _
                               ByVal sFTPPwd As String, _
                               ByVal sFTPSrcFileName As String, _
                               ByVal sFTPTgtFileName As String, _
                               ByVal sFTPDeleteSourceFile As Boolean)
 
 Dim ofs As Scripting.FileSystemObject
 Dim sURL As String
 Dim intRetryCount As Integer
 
 On Error GoTo FTP_ERROR

 
 Me.Show
 intRetryCount = 0
 
 Me.HRG True
 
' msCurrentFile = ""
 
 Set ofs = New Scripting.FileSystemObject
  
FTP_RETRY:

'    sURL = "ftp://" & sFTPUser & ":" & sFTPPwd & "@" & sFTPServer
    
    Inet1.Protocol = icFTP
    Inet1.RequestTimeout = 120
    Inet1.RemotePort = 21
    Inet1.AccessType = icDirect
    Inet1.URL = sFTPServer
    Inet1.username = sFTPUser   '  "i101ftptest@imaging101.com"
    Inet1.Password = sFTPPwd   '  "ftptest"
    
   msNewFile = sFTPTgtFileName
   
   sURL = Inet1.URL
   
    On Error Resume Next
   
   Select Case sFTPCommand
 
      Case "PUT"
      
            msCurrentFile = sFTPSrcFileName
            If ofs.FileExists(sFTPSrcFileName) = False Then
                Err.Raise 10101, "Source File NOT found"
                GoTo FTP_ERROR
            End If
            
            msFTPProcessState = "UPLOADING File:" & Space(5) & msNewFile
            '*** Must use the Quotation Marks  - Chr(34)
            '    because the Inet control DOES NOT LIKE SPACES in Filenames or Directories
            Inet1.Execute sURL, sFTPCommand & Space(1) & Chr(34) & sFTPSrcFileName & Chr(34) & " " & Chr(34) & sFTPTgtFileName & Chr(34)
            
        
      Case "GET"
      
            msCurrentFile = sFTPTgtFileName
            msFTPProcessState = "DOWNLOADING File:" & Space(5) & sFTPSrcFileName
            If ofs.FileExists(sFTPTgtFileName) = True Then ofs.DeleteFile sFTPTgtFileName, True
            Inet1.Execute sURL, sFTPCommand & Space(1) & sFTPSrcFileName & " " & sFTPTgtFileName
      
      Case "DELETE"
      
            msCurrentFile = sFTPTgtFileName
            msFTPProcessState = "DELETING File:" & Space(5) & sFTPSrcFileName
            Inet1.Execute sURL, sFTPCommand & Space(1) & sFTPSrcFileName

      Case "MKDIR"
      
            msCurrentFile = sFTPTgtFileName
            msFTPProcessState = "MAKING Directory:" & Space(5) & sFTPSrcFileName
            Inet1.Execute sURL, sFTPCommand & Space(1) & sFTPSrcFileName
      
      Case "CWD"
      
            msCurrentFile = sFTPTgtFileName
            msFTPProcessState = "CHANGING Working Directory to:" & Space(5) & sFTPSrcFileName
            Inet1.Execute sURL, sFTPCommand & Space(1) & sFTPSrcFileName

    End Select
    
'        'Trap any Inet1 errors from other modules or procedures
'        If blnFTPError = True Then
'            GoTo FTP_ERROR
'        End If
    
    Me.WaitForResponse
    msFTPProcessState = "CLOSE Connection"
    Inet1.Execute sURL, "quit"
    Me.WaitForResponse
    
    '*** TRAP ERRORS
    If blnFTPError Then
        GoTo FTP_ERROR
    End If
    
    
    'Zap the Sourcefile if flag is set to True
    If sFTPDeleteSourceFile Then
        If ofs.FileExists(sFTPSrcFileName) = True Then ofs.DeleteFile sFTPSrcFileName, True
    End If
 
    Me.HRG False
    
Exit Sub
    
FTP_ERROR:
  blnFTPError = True
  Me.lblRESPONSE.Caption = "*** ERROR:" & Err.Number & " - " & Err.Description
  Me.lblRESPONSE.Refresh
  
  If intRetryCount < 3 Then
    intRetryCount = intRetryCount + 1
    Me.lblRESPONSE.Caption = "*** RETRYING Transfer after error... Retry #" & intRetryCount
    GoTo FTP_RETRY
  End If
  
  Set ofs = Nothing
  HRG False
  
End Sub
 

Friend Sub WaitForResponse()

  Dim fWait As Boolean
  
  ''On Error GoTo ErrHandler
 On Error GoTo 0
 
  fWait = True
  Do Until fWait = False
        DoEvents
        fWait = Inet1.StillExecuting
  Loop
 
ErrHandler:
  Err.Clear
End Sub
 




Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
 
    On Error Resume Next
'    On Error GoTo 0
    
    Select Case State
            Case icNone
            Case icResolvingHost:
                Me.lblRESPONSE.Caption = "Resolving Host"
            Case icHostResolved:
                Me.lblRESPONSE.Caption = "Host Resolved"
            Case icConnecting:
                Me.lblRESPONSE.Caption = "Connecting..."
            Case icConnected:
                Me.lblRESPONSE.Caption = "Connected"
            Case icResponseReceived:
''                Me.lblRESPONSE.Caption = "Transferring File:" & Space(5) & msCurrentFile
                Me.lblRESPONSE.Caption = msFTPProcessState
'            Case icIncorrectUserName:
'                Me.lblRESPONSE.Caption = "INCORRECT USERNAME!"
'            Case icIncorrectPassword:
'                Me.lblRESPONSE.Caption = "INCORRECT PASSWORD!"
'            Case icBadUrl:
'                Me.lblRESPONSE.Caption = "BAD URL!"
            Case icDisconnecting:
                Me.lblRESPONSE.Caption = "Disconnecting..."
            Case icDisconnected:
                Me.lblRESPONSE.Caption = "Disconnected"
            Case icError:
                'Flag Global error flag and Raise and Error Code
                blnFTPError = True
                Me.lblRESPONSE.Caption = "*** FTP ERROR: " & Inet1.ResponseCode & " " & Inet1.ResponseInfo & " [" & Err.Description & "]"
                Err.Raise 10101, "frmFTP", "*** FTP ERROR: " & Inet1.ResponseCode & " " & Inet1.ResponseInfo
                
            Case icResponseCompleted:
                Me.lblRESPONSE.Caption = "Process Complete."
     End Select
      
''''''     Me.lblRESPONSE.Refresh
''''''     Me.txtMessageLog.SelStart = Len(txtMessageLog.Text)
     DoEvents
     
      Err.Clear
     
End Sub

Friend Sub HRG(fShowHourGlass As Boolean)

   If fShowHourGlass = True Then
      Screen.MousePointer = 11
   Else
      Screen.MousePointer = 0
   End If
   
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  Set frmFTP = Nothing
End Sub


Private Sub lblRESPONSE_Change()

    ' vbSendMail 'Status Event'
    lstStatus.AddItem lblRESPONSE.Caption
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1
    funcWriteToDebugLog Me.name, lblRESPONSE.Caption
    DoEvents
    
End Sub


