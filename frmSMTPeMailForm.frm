VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMTPeMailForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Send via SMTP"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frmSMTPeMailForm.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1455
      TabIndex        =   24
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtAttachmentFilePath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   2880
      Width           =   6975
   End
   Begin VB.TextBox txtSMTPSendBCC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   6495
   End
   Begin VB.TextBox txtSMTPSendCC 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   6495
   End
   Begin VB.TextBox txtSMTPSendTo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   6495
   End
   Begin VB.CheckBox cboSMTPSendCcToSender 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send me a copy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtSMTPFromEmailDisplayName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Jacob Russo"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtSMTPFromEmailAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "jacob@imaging101.com"
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtSMTPEmailSubject 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3720
      Width           =   6975
   End
   Begin VB.TextBox txtSMTPEmailMessage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4560
      Width           =   6975
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SEND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Picture         =   "frmSMTPeMailForm.frx":0693
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Picture         =   "frmSMTPeMailForm.frx":135D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewAttachment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View Attachment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "frmSMTPeMailForm.frx":18E7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   7080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6840
      Width           =   7095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Attachments:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Separate multiple entries with a Semicolon "";"" "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   600
      TabIndex        =   19
      Top             =   1560
      Width           =   3972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BCC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   2295
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   2055
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1815
      Width           =   375
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
      TabIndex        =   15
      Top             =   480
      Width           =   1845
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM Display Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM eMail Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   855
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   1932
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   14
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   1932
   End
End
Attribute VB_Name = "frmSMTPeMailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim txtAttachmentFilePath As String

'******************************************
' SMTP email Module Declarations

Dim bolEmailErrorOccured As Boolean
Dim bolEmailErrorOccuredGlobal As Boolean


'******************************************
' vbSendMail Module Declarations
' Reference to vbSendMail (SMTP) DLL
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
' misc vbSendMail local vars
Dim bolRequiresAuthentication  As Boolean
Dim bolUsePOP3Auth             As Boolean
Dim bolHtml           As Boolean
Dim MyEncodeType      As ENCODE_METHOD
Dim etPriority        As MAIL_PRIORITY
Dim bolReceipt        As Boolean

Dim txtSMTPHost As String
Dim txtSMTPPOP3Host As String
Dim cboSMTPRequiresAuthentication As String
Dim cboSMTPUsePOP3Auth  As String
Dim txtSMTPAuthenticationUserID As String
Dim txtSMTPAuthenticationPassword As String
Dim txtSMTPPort As String

Dim strErrMsg As String




Private Sub cmdSend_Click()

    Screen.MousePointer = vbHourglass

    On Error GoTo ERROR_HANDLER

        '***********************************************************
        '*** GET the SMTP User Settings
        funcSaveFieldToDB RegImaging101ConnectionString, "I101Security", "UserID = '" & gsecUserID & "'", "SMTPFromEmailAddress", txtSMTPFromEmailAddress
        funcSaveFieldToDB RegImaging101ConnectionString, "I101Security", "UserID = '" & gsecUserID & "'", "SMTPFromEmailDisplayName", txtSMTPFromEmailDisplayName
        funcSaveFieldToDB RegImaging101ConnectionString, "I101Security", "UserID = '" & gsecUserID & "'", "SMTPSendCcToSender", cboSMTPSendCcToSender
    
    
    
    
    
    '*** Set vbSendMail Boolean values
    bolRequiresAuthentication = False
    bolUsePOP3Auth = False
    
    If cboSMTPRequiresAuthentication = vbChecked Then
        bolRequiresAuthentication = True
    End If

    If cboSMTPUsePOP3Auth = vbChecked Then
        bolUsePOP3Auth = True
    End If
    
    With poSendMail

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = False                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MIME_ENCODE                  ' Optional, default = MIME_ENCODE
        .Priority = PRIORITY_NORMAL                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = False                                 ' Optional, default = FALSE
        .UseAuthentication = bolRequiresAuthentication   ' Optional, default = FALSE
        .UsePopAuthentication = bolUsePOP3Auth           ' Optional, default = FALSE
        .username = txtSMTPAuthenticationUserID    ' Optional, default = Null String
        .Password = txtSMTPAuthenticationPassword  ' Optional, default = Null String, value is NOT saved
        .POP3Host = txtSMTPPOP3Host
        .MaxRecipients = 100                                           ' Optional, default = 100, recipient count before error is raised
        .SMTPHost = txtSMTPHost
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 20                      ' Optional, default = 10
        ' .ConnectRetry = 10                         ' Optional, default = 5
        ' .MessageTimeout = 120                      ' Optional, default = 60
        ' .PersistentSettings = False                 ' Optional, default = TRUE
         .SMTPPort = txtSMTPPort                  ' Optional, default = 25
        
        
        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        
        ' **************************************************************************
        ' BASIC properties for sending email
        ' **************************************************************************
        .SMTPHost = txtSMTPHost                                     ' Required the fist time, optional thereafter
        .From = txtSMTPFromEmailAddress                             ' Required the fist time, optional thereafter
        .FromDisplayName = txtSMTPFromEmailDisplayName              ' Optional, saved after first use
        .Recipient = txtSMTPSendTo                                  ' Required, separate multiple entries with delimiter character
'        .RecipientDisplayName = Me.txtAddress.Text                  ' Optional, separate multiple entries with delimiter character
        
        'Check if CC should be sent back to Sender
        If cboSMTPSendCcToSender = vbChecked Then
            If Trim(Me.txtSMTPSendCC.Text) <> "" Then
                Me.txtSMTPSendCC.Text = Me.txtSMTPSendCC.Text & " ; " & Me.txtSMTPFromEmailAddress
            Else
                Me.txtSMTPSendCC.Text = Me.txtSMTPFromEmailAddress
            End If
        End If
        
        .CcRecipient = Me.txtSMTPSendCC.Text                            ' Optional, separate multiple entries with delimiter character
        
'        .CcDisplayName = txtCcName                                     ' Optional, separate multiple entries with delimiter character
'        .ReplyToAddress = txtFromEmailAddress                          ' Optional, used when different than 'From' address
        
        .BccRecipient = txtSMTPSendBCC                                  ' Optional, separate multiple entries with delimiter character
        
        .Subject = Trim(Me.txtSMTPEmailSubject.Text)                    ' Optional
        .Message = Trim(Me.txtSMTPEmailMessage.Text)                    ' Optional
        
        .Attachment = Trim(txtAttachmentFilePath)                       ' Optional, separate multiple entries with delimiter character


        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        
        '.Connect                                  ' Optional, use when sending bulk mail
        .Send                                       ' Required
        '.Disconnect                               ' Optional, use when sending bulk mail
        
        
'        txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
'                                                    '   MX look up was used to find a host    End With
        
        'RETRY Processing for eMail FAILURES
        If bolEmailErrorOccured Then
            'Reset eMail Error Flag
            bolEmailErrorOccured = False
            'Pause for 5 seconds - try 2nd time
            poSendMail_Status "***  PAUSING  30 Seconds for RETRY"
'            frmImaging101AutoExport.NTService1.LogEvent 101, 101, "PAUSING 30 Seconds for RETRY"
            FunctionDeclarations.TimePause 30
            poSendMail_Status "***  RE-TRYING EMAIL SEND - 2nd ATTEMPT"
'            frmImaging101AutoExport.NTService1.LogEvent 101, 101, "RE-TRYING EMAIL SEND - 2nd ATTEMPT"
            .Send

            If bolEmailErrorOccured Then
                'Reset eMail Error Flag
                bolEmailErrorOccured = False
                'Pause for 5 seconds - try 3rd time
                poSendMail_Status "***  PAUSING  30 Seconds for RETRY"
'                frmImaging101AutoExport.NTService1.LogEvent 101, 101, "PAUSING 30 Seconds for RETRY"
                FunctionDeclarations.TimePause 30
                poSendMail_Status "***  RE-TRYING EMAIL SEND - 3Rd ATTEMPT"
'                frmImaging101AutoExport.NTService1.LogEvent 101, 101, "RE-TRYING EMAIL SEND - 3rd ATTEMPT"
                .Send
                
                'All THREE Attempts FAILED... Set the Global Flags
                If bolEmailErrorOccured Then
                    bolErrorOccured = True
                    bolLineErrorOccured = True
                    bolEmailErrorOccuredGlobal = True
                End If
            
            End If

        End If
    
    End With
            
            
    DoEvents
    
    
    
'    'Close the Document to clear
'    SpicerDoc1.CloseDocument (False)
'    DoEvents
'
'    '*** Close the printer and clear the buffer.
'    Printer.EndDoc
    
    '********************************
    '*** DELETE the Temporary Files
'    On Error Resume Next
    
    Dim arrFilesToDelete() As String
    arrFilesToDelete = Split(txtAttachmentFilePath, ";")
    For i = 0 To UBound(arrFilesToDelete)
        Kill arrFilesToDelete(i)
    Next
                    
    Screen.MousePointer = MousePointerConstants.vbDefault

    Unload Me
    
Exit Sub

ERROR_HANDLER:
    

    bolErrorOccured = True
    strErrMsg = "cmdSend_Click ERROR: Error #: " & Err.Number & " - " & Err.Description
'    subWriteToAuditTraceFile txtTraceFilePath, dblDocumentRECID, dblDetailRECID, txtDestinationFilename, strErrMsg
    
    funcWriteToDebugLog Me.name, strErrMsg
    funcQuickMessage "SHOW", strErrMsg

'    funcWriteToSystemEventLog frmImaging101AutoExport.NTService1, svcMessageError, strErrMsg
    
    '********************************
    '*** DELETE the Temporary File
    On Error Resume Next
    Err.Clear
    arrFilesToDelete = Split(txtAttachmentFilePath, ";")
    For i = 0 To UBound(arrFilesToDelete)
        Kill arrFilesToDelete(i)
    Next
    
    If Err.Number <> 0 Then
        funcQuickMessage "SHOW", "An error occured while trying to DELETE the TEMP files: " & vbCrLf & txtAttachmentFilePath
    End If
    
    Screen.MousePointer = MousePointerConstants.vbDefault

End Sub



Private Sub cmdViewAttachment_Click()

    shelldoc txtAttachmentFilePath

End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

'    StatusBar1.Panels(1).width = Me.ScaleWidth

    '*** Initialize the vbSendMail (SMTP) component
    Set poSendMail = New clsSendMail
    

End Sub

Public Sub subStartup(strAttachmentFilePath As String)

    On Error GoTo ERROR_HANDLER
    
    'Assign the attachment filepath to the Form variable
    '  for the View Attachment button.
    txtAttachmentFilePath = strAttachmentFilePath
    
    txtSMTPFromEmailAddress = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Security", "UserID = '" & gsecUserID & "'", "SMTPFromEmailAddress") & ""
    txtSMTPFromEmailDisplayName = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Security", "UserID = '" & gsecUserID & "'", "SMTPFromEmailDisplayName") & ""
    cboSMTPSendCcToSender = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Security", "UserID = '" & gsecUserID & "'", "SMTPSendCcToSender") & ""

    '***********************************************************
    '*** GET the SMTP System Settings
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ssql As String
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.ConnectionString = RegImaging101ConnectionString
    conn.ConnectionTimeout = 120
    conn.mode = adModeRead
    conn.Open
    
    ssql = "SELECT * FROM I101Control WHERE ID = 1"
    
    With rs
        .ActiveConnection = conn
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LOCKTYPE = adLockReadOnly
        .Source = ssql
    End With
    
    rs.Open

    'Return the Found Value
    txtSMTPHost = rs("SMTPHost") & ""
    txtSMTPPOP3Host = rs("SMTPPOP3Host") & ""
    cboSMTPRequiresAuthentication = rs("SMTPRequiresAuthentication") & ""
    cboSMTPUsePOP3Auth = rs("SMTPUsePOP3Auth") & ""
    txtSMTPAuthenticationUserID = rs("SMTPAuthenticationUserID") & ""
    txtSMTPAuthenticationPassword = rs("SMTPAuthenticationPassword") & ""
    txtSMTPEmailSubject = rs("SMTPDefaultEmailSubject") & ""
    txtSMTPEmailMessage = rs("SMTPDefaultEmailMessage") & ""
    txtSMTPPort = rs("SMTPPort") & ""
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    txtSMTPSendTo.SetFocus


Exit Sub

ERROR_HANDLER:

    If Err.Number = 13 Then  ' Type Mismatch
        Resume Next
    Else
        funcQuickMessage "SHOW", "subGetSMTPeMailSettings ERROR: " & Err.Number & vbCrLf & Err.Description
    End If
    

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'Deinitialize the vbSendMail Object
    Set poSendMail = Nothing

End Sub

' *****************************************************************************************
' The following four poSendMail Subs capture the Events fired by the vbSendMail component
' *****************************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblStatus = lPercentCompete & "%"
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    If IsNumeric(Status) Then
        ProgressBar1.Value = lPercentCompete
    End If
    
End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event
    
    bolEmailErrorOccured = True

'    frmImaging101AutoExport.NTService1.LogEvent 101, 101, "Send Mail FAILED: " & Explanation
    
        'Make sure file is closed... just in case.
'        Close #2
'        Open txtTraceFilePath For Append As #2
'
'        Print #2, "        " & "*** Send Mail FAILED: " & Explanation
'
'        Close #2
    

    
    
    lblStatus = "SEND FAILED!!!"
    
'    StatusBar1.Panels(1).Text = lblStatus
   
    MsgBox "*** Send Mail FAILED: " & Explanation
    
    Screen.MousePointer = vbDefault
'    cmdSend.Enabled = True
    
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
'    MsgBox "Send Successful!"
    poSendMail_Status "Send Successful!"


End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
'    lstStatus.AddItem Status
'
''    If bolDebug Then
'        'Make sure file is closed... just in case.
'        Close #2
'        Open txtTraceFilePath For Append As #2
'
'        Print #2, "        " & Status
'
'        Close #2
''    End If
'
'    lstStatus.ListIndex = lstStatus.ListCount - 1
'    lstStatus.ListIndex = -1

    lblStatus = Status
    
End Sub


Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

End Sub
