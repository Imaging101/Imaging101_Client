VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmPopUpNotifyForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imaging101 Notification"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      Picture         =   "frmPopUpNotifyForm.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   1
      Top             =   0
      Width           =   1572
   End
   Begin VB.Timer timNoficationTimer 
      Interval        =   60000
      Left            =   960
      Top             =   360
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1680
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer timNoficationLengthTimer 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   360
   End
   Begin VB.TextBox txtNotification 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmPopUpNotifyForm.frx":0693
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C07000&
      Height          =   225
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "frmPopUpNotifyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WorkareaLeft As Integer
Dim WorkareaTop As Integer
Dim WorkareaWidth As Integer
Dim WorkareaHeight As Integer

Dim intTimerEventCounter As Integer
Dim RegImaging101ConnectionString As String

'General declarations relating to the display of the popup.
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_SHOWNOACTIVATE = 4

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10


    
Public Sub subShowNotification(strNotificationText As String, intNotificationDurationMilliseconds As Integer)

     With SysInfo1
       WorkareaLeft = .WorkareaLeft '/ Screen.TwipsPerPixelX
       WorkareaTop = .WorkareaTop '/ Screen.TwipsPerPixelY
       WorkareaWidth = .WorkareaWidth '/ Screen.TwipsPerPixelX
       WorkareaHeight = .WorkareaHeight '/ Screen.TwipsPerPixelY
     End With

    txtNotification.Text = strNotificationText
    
'    Me.Left = WorkareaWidth - Me.Width - 10
'    'Position ALL the way to the BOTTOM
    Me.Top = WorkareaTop - Me.Height
    Me.Left = WorkareaWidth - Me.width
    
'    Me.Show

'    'Stay on Top of other windows... AFTER Sliding UP...
'    '  otherwise it slides Up OVER the Windows Task Bar
'    funcMakeTopMost Me, True

    'Show the form without activating it. (This form will
    'never be activated - otherwise the titlebar of the
    'parent window would show its deactivated state.
    ShowWindow hwnd, SW_SHOWNOACTIVATE
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

    DoEvents

    
    'Slide the form DOWN from the TOP
    For i = 1 To Me.Height Step 30
        Me.Top = WorkareaTop - Me.Height + i + 10
        DoEvents
    Next
    
    
    'Set the Timer Duration & Enable the timer
    timNoficationLengthTimer.Interval = intNotificationDurationMilliseconds
    timNoficationLengthTimer.enabled = True
    
End Sub


Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

'    gsecUserName = "Notify"
'
'    If frmImaging101Winsock.Winsock1.State <> sckConnected Then
'            frmImaging101Winsock.cmdConnect_Click
'            DoEvents
'            frmImaging101Winsock.funcWaitForDataToArrive
'    End If
'
'    If frmImaging101Winsock.Winsock1.State <> sckConnected Then
'           MsgBox "The Server did NOT respond with 'Connected'... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Connection Failed"
'            frmImaging101Winsock.cmdClose_Click
'    End If
'
'    RegImaging101ConnectionString = frmImaging101Winsock.funcSendData("GET IMAGING101 CONNECTION STRING")
'        If Left(RegImaging101ConnectionString, 5) = "ERROR" Then
'            MsgBox "The Server did NOT return the Imaging101ConnectionString... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
'            frmImaging101Winsock.cmdClose_Click
'        End If
'
'   'Disconnect
'   frmImaging101Winsock.cmdClose_Click
'
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ShowWindow hwnd, SW_HIDE
    Cancel = True
    DoEvents


End Sub

Private Sub timNoficationLengthTimer_Timer()

    'STOP the Timer
    timNoficationLengthTimer.enabled = False
    'DON'T Stay on Top of other windows...
    '  otherwise it slides Down OVER the Windows Task Bar
'    funcMakeTopMost Me, False
'    DoEvents
    
    'Slide the form UP
    For i = 1 To (WorkareaTop - Me.Height) Step -50
        Me.Top = WorkareaTop + i + 10
        DoEvents
    Next
    
'    Me.Hide
    ShowWindow hwnd, SW_HIDE
    
    DoEvents

End Sub

Private Sub subGetQueueDetails()

        On Error GoTo ERROR_HANDLER
        
     '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String
    Dim strWhere As String
    Dim strQueueItems As String
    Dim intRecordCount As Integer

    Dim strApplicRECID
    strApplicRECID = funcGetFieldFromDB(RegImaging101BatchListConnectionString, "I101Applications", "ApplicationName = '" & gsecBatchDefaultApplication & "'", "ApplicationRECID")
    
    If strApplicRECID = "" Then
        Exit Sub
    End If
    
    
    Set con = New ADODB.Connection
    con.Open RegImaging101BatchListConnectionString
    
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LOCKTYPE = adLockReadOnly
    End With
        
    rs.Source = "SELECT DISTINCT BatchQueue, " & _
                "MAX(DATEDIFF(day, BatchScanDate, getdate())) AS MaxAgeDays " & _
                " FROM I101Batches " & _
                " WHERE "
            
    strWhere = ""
    
    '*** Show Batches for the Logged-In User's OR Unassigned Batches
    strWhere = strWhere & " ( (BatchOwner is NULL) OR (BatchOwner = '') OR (BatchOwner = '" & gsecUserName & "') )"
    
    '*** Show Batches for the Logged-In User's OR Unassigned Batches
    
    strWhere = strWhere & " AND  ( ApplicationRECID = " & strApplicRECID & " )"
    
    '*** Don't show Fully-Committed or Fully-Split Batches
    strWhere = strWhere & " AND ((BatchCommitStatus NOT LIKE 'Committed-FULL%') OR (BatchCommitStatus is null)) "
    strWhere = strWhere & " AND ((BatchCommitStatus + '' not like 'Split-FULL%') OR (BatchCommitStatus is null))"
            
    
    rs.Source = rs.Source & strWhere & " GROUP BY BatchQueue ORDER BY  MaxAgeDays DESC "
            
 
    con.Errors.Clear
    rs.Open
        
    rs.MoveFirst
    
    DoEvents
    
    intRecordCount = rs.RecordCount
    
    strQueueItems = ""
    
    If intRecordCount > 0 Then
    
        strQueueItems = "====ITEMS IN BATCH QUEUES====" & vbCrLf & _
                        "Application: " & gsecBatchDefaultApplication & vbCrLf & _
                        "Days    Batch Queue" & vbCrLf & _
                        "-----------------------------"
        
        While Not rs.EOF
            strQueueItems = strQueueItems & vbCrLf & Format(rs.Fields("MaxAgeDays"), "###0") & "  " & rs.Fields("BatchQueue")
            DoEvents
            rs.MoveNext
        Wend
    End If
    
    rs.Close
    con.Close
    Set rs = Nothing
    Set con = Nothing
    
    'Show Pop Up message ONLY if items found...
    '  Do this AFTER Closing the DB Connecition... in case we decide to add
    '  some other processing during the display.
    If intRecordCount > 0 Then
        frmPopUpNotifyForm.subShowNotification strQueueItems, 4000
    End If
    
Exit Sub

ERROR_HANDLER:

        
        Resume Next
        

End Sub


Private Sub timNoficationTimer_Timer()

    'Exit if the Notification Frequency is Disabled (=0)
    If Trim(gsecBatchQueueNotificationFrequency) = "0" _
    Or Trim(gsecBatchQueueNotificationFrequency) = "" Then
        Exit Sub
    End If
    
    'Max Timer Interval is rougly One Minute
    '  we use the intTimerEventCounter to allow multiple minutes
    If intTimerEventCounter = 0 Or intTimerEventCounter >= CInt(gsecBatchQueueNotificationFrequency) Then
        
        intTimerEventCounter = 1
        
        timNoficationTimer.enabled = False
        
        'Do SELECT for Batch Queues for this User
        subGetQueueDetails
        
        timNoficationTimer.enabled = True
        
    End If
    
    intTimerEventCounter = intTimerEventCounter + 1
    
End Sub


Private Sub txtNotification_Click()

     funcQuickMessage "SHOW", txtNotification.Text

End Sub


