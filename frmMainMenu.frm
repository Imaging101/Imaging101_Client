VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Menu - Imaging101 "
   ClientHeight    =   6225
   ClientLeft      =   3375
   ClientTop       =   3300
   ClientWidth     =   10260
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImageRetrieval 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Image Retrieval"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMainMenu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1515
      UseMaskColor    =   -1  'True
      Width           =   4095
   End
   Begin VB.CommandButton cmdConfiguration 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Configuration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMainMenu.frx":2EC4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3645
      UseMaskColor    =   -1  'True
      Width           =   4095
   End
   Begin VB.CommandButton cmdBatchManagement 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batch Management"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMainMenu.frx":3306
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2580
      UseMaskColor    =   -1  'True
      Width           =   4095
   End
   Begin VB.CommandButton cmdTTCDownloadTables 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Download TTC Tables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5490
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   4770
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   525
      Width           =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C07000&
      BorderWidth     =   24
      X1              =   0
      X2              =   10320
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   315
      Left            =   6795
      TabIndex        =   13
      Top             =   975
      Width           =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C07000&
      BorderWidth     =   24
      X1              =   0
      X2              =   10320
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label lblUsersCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "UsersCount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   8655
      TabIndex        =   15
      Top             =   5610
      Width           =   825
   End
   Begin VB.Label lblUsersCountCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "User Count:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   7470
      TabIndex        =   14
      Top             =   5610
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   7800
      Picture         =   "frmMainMenu.frx":6C6C
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label lblUserID 
      BackStyle       =   0  'Transparent
      Caption         =   "UserID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   1800
      TabIndex        =   11
      Top             =   4905
      Width           =   2400
   End
   Begin VB.Label lblLicencedTo 
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   1800
      TabIndex        =   10
      Top             =   5175
      Width           =   4440
   End
   Begin VB.Label lblLicencedToCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Licenced to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   1440
   End
   Begin VB.Label lblLicenceNumberCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "License #:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   1440
   End
   Begin VB.Label lblLicenseNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "License #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   5400
      Width           =   4440
   End
   Begin VB.Label lblProcessorIDCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Processor ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblProcessorID 
      BackStyle       =   0  'Transparent
      Caption         =   "Processor ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   1800
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Label lblLoggedInUserCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Logged In User:  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   3060
      Left            =   120
      Picture         =   "frmMainMenu.frx":7838
      Stretch         =   -1  'True
      Top             =   1545
      Width           =   4485
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mPopLogout 
         Caption         =   "&Logout"
      End
      Begin VB.Menu Spacer1 
         Caption         =   ""
      End
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopHide 
         Caption         =   "&Hide"
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBatchManagement_Click()
'''    Me.Hide
    On Error Resume Next ' In case I can't open the DB
    Me.WindowState = vbMinimized
    frmImaging101BatchList.Show
    On Error GoTo 0
End Sub

Private Sub cmdConfiguration_Click()
'''    Me.Hide
    Me.WindowState = vbMinimized
    frmConfig.Show

End Sub

Private Sub cmdImageRetrieval_Click()

'''    Me.Hide
    Me.WindowState = vbMinimized
    frmImaging101Search.Show

End Sub

Private Sub cmdScanImport_Click()
    
End Sub



Private Sub cmdTTCDownloadTables_Click()

    Me.Hide
    frmDownloadTTC.Show
    
End Sub


Private Sub MainMenuStart()

        Dim strPicturePath As String
        
        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        
        'If Main Menu Graphic Picture exists... Load it!
        strPicturePath = App.Path & "\I101MainMenuPicture.jpg"
        If funcFileExists(strPicturePath) Then
            Me.Image1.Picture = LoadPicture(strPicturePath)
        End If
        
'        Me.Show
        
'        DoEvents
        
'        DoEvents
        
        
        ' Display UserID and License Info.
        lblUserID.Caption = gsecUserID
        lblLicencedTo.Caption = gsecSiteInformationClientLong
        lblLicenseNumber.Caption = gsecSiteInformationLicenseCode
        lblProcessorID.Caption = gProcessorID
        lblUsersCount.Caption = funcGetUserCountFromDB(RegImaging101ConnectionString)

        'DISABLE the SYSTRAY "Exit" Option if requested
        If bolNoExit Then
            mPopExit.Enabled = False
        End If



End Sub

Private Sub Dir1_Change()

    Dim arrayFolders() As String
    Dim arrayPointers() As Long
    Dim arrayFilenames() As String
    Dim i As Integer
    
    PathLogInit
    LogPath Dir1.Path, arrayFolders, arrayPointers, arrayFilenames
    
    For i = 0 To UBound(arrayFilenames)
        lstFolders.AddItem arrayFolders(i)
        lstPointers.AddItem arrayPointers(i)
        lstFiles.AddItem arrayFilenames(i)
    Next
    
End Sub



Private Sub Form_Activate()

    '*** 2020-09-29 - Jacob - Disabled this code because it seems like it causes FLASHING between the Menu and Retrieval forms.
    '                                                NOT sure why it was done to begin with???
'        If Me.cmdImageRetrieval.Visible Then
'            Me.cmdImageRetrieval.SetFocus
'        End If



End Sub

Private Sub Form_Load()
    
        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
        Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
        
        If bolDebug Then
            Me.Caption = Me.Caption & " - DEBUG"
        End If

        
        ' Enable / Disable Buttons based on Security settings for Logged-In user
        
        cmdTTCDownloadTables.Visible = False
        cmdImageRetrieval.Visible = True
        cmdConfiguration.Visible = False
        cmdBatchManagement.Visible = False
        
        
'* Jacob - 8/28/2006 -  Disabled CUSTOM BUTTONS
'        '*** Set up CUSTOM Button(s)
'        Select Case gsecSiteInformationClientShort
'
'        Case "TTC"
'            cmdTTCDownloadTables.Visible = True
'        Case "JMH"
'            cmdJMHDownloadTables.Visible = True
'        End Select

        '*************************************************
        '*** GET MENU RIGHTS  - Temporary Fix
        '***                    because security rights are assigned based on Application
        
        funcGetMenuRights gsecSecurityRECID
        
    
        '*** SET UP SECURITY
'        If gsecRightsAdminSystem = "" Then
'            gsecRightsAdminSystem = 0
'        End If
        
        If gsecRightsAdminSystem = vbChecked Then
            cmdConfiguration.Visible = True
        Else
            cmdConfiguration.Visible = False
        End If
        

        
'        If gsecRightsDeleteDocuments = vbChecked Then
'            '*
'        End If
        
        '*** Currently This is the chkRightsAdminApplication field
        If gsecRightsAdminSystem = vbChecked Then
            cmdImageRetrieval.Visible = True
            cmdBatchManagement.Visible = True
        End If
        
        If gsecRightsBatchScan = vbChecked _
        Or gsecRightsBatchIndex = vbChecked _
        Or gsecRightsBatchView = vbChecked _
        Or gsecRightsBatchCommit = vbChecked _
        Or gsecRightsBatchRoute = vbChecked _
        Or gsecRightsImportFromFile = vbChecked _
        Or gsecRightsBatchCommit = vbChecked Then
            cmdBatchManagement.Visible = True
        End If
        
        If gsecRightsRetrieveImages = vbChecked Then
            cmdImageRetrieval.Visible = True
        Else
            cmdImageRetrieval.Visible = False
        End If
        
        '*** NOW GET THE MAIN MENU SET UP
        MainMenuStart
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'UNLOADMODE VALUES:
    '   0 - The user has chosen the Close command from the Control-menu box on the form, or hits the big X on the other side.
    '   1 - The Unload method has been invoked from code.
    '   2 - The current Windows-environment session is ending.
    '   3 - The Microsoft Windows Task Manager is closing the application.
    '   4 - An MDI child form is closing because the MDI form is closing.
      
    If UnloadMode < 1 Then
    
        'Hide the Menu if the System Tray option is enabled
        If bolSysTrayActive = True Then
            Me.Hide
            Cancel = True
        End If
      
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmLogin.Show
    
    'Disconnect from Remote Host
    frmImaging101Winsock.cmdClose_Click
    
    'Remove the systray icon
    subSysTrayDeleteIcon
    
    On Error Resume Next
    
    '*** DELETE TEMP FILES
    '*** 2020-05-18 - Jacob - Modified TEMP Directory to eliminate Environ() check, make more standard
    Dim strLocalTempDir As String
    strLocalTempDir = funcGetTempDir()
    strLocalTempDir = strLocalTempDir & "Imaging101\"

    If funcFileExists(strLocalTempDir & "*.*") Then
        'Kill strLocalTempDir & "*.*"
        
        sFilename = Dir(strLocalTempDir)

        Do While sFilename > ""
        
          funcWriteToDebugLog Me.name, "*** frmMainMenu.Form.Unload() |  KILL " & strLocalTempDir & sFilename
          Kill strLocalTempDir & sFilename
          sFilename = Dir()
        
        Loop
    End If
'    On Error GoTo 0
    
    Set frmMainMenu = Nothing
    
    '*** This function is VERY Dangerous!
    '    It will unload all forms except the FIRST one.
''''    UnloadAllForms

End Sub


Public Sub subSysTrayAddIcon()

        'Add the Icon to the System Tray
        With nid
         .cbSize = Len(nid)
         .hwnd = Me.hwnd
         .uId = vbNull
         .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         .uCallBackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon
         .szTip = "Imaging101 Client" & vbNullChar
        End With

'        Call subLogSystemEvent(svcEventInformation, svcMessageInfo, "[Shell_NotifyIcon NIM_ADD, nid ]")
    
        Shell_NotifyIcon NIM_ADD, nid

End Sub

Public Sub subSysTrayDeleteIcon()

     'called when user clicks the popup menu Exit command
        
        'Delete the Icon from the System Tray
        With nid
         .cbSize = Len(nid)
         .hwnd = Me.hwnd
         .uId = vbNull
         .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         .uCallBackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon
         .szTip = "Imaging101 Client" & vbNullChar
        End With

'    Call subLogSystemEvent(svcEventInformation, svcMessageInfo, "[Entering mPopExit_Click Sub]")
       
        Shell_NotifyIcon NIM_DELETE, nid

End Sub

Private Sub lblCompany_Click()

End Sub

Private Sub Image1_Click()

'    frmImaging101QueueManagement.Show
    
    
End Sub

Private Sub lblVersion_Click()

    If UCase(gsecUserID) = "JACOB" Then
            frmDirTreeList.Show
    End If

End Sub

Private Sub mPopExit_Click()
    
    If bolNoExit = True Then
        Exit Sub
    End If
    
    'Delete the Systray Icon
    subSysTrayDeleteIcon
    
    'END Imaging101 NOW!
    End
       
End Sub

Private Sub mPopHide_Click()
    
'    Call subLogSystemEvent(svcEventInformation, svcMessageInfo, "[Entering mPopRestore_Click Sub]")
       
       'called when the user clicks the popup menu Restore command
       Dim result As Long
       Me.WindowState = vbNormal
       result = SetForegroundWindow(Me.hwnd)
       Call HideAllForms
       
End Sub
Private Sub mPopRestore_Click()
    
'    Call subLogSystemEvent(svcEventInformation, svcMessageInfo, "[Entering mPopRestore_Click Sub]")
       
       'called when the user clicks the popup menu Restore command
       Dim result As Long
       Me.WindowState = vbNormal
       result = SetForegroundWindow(Me.hwnd)
'       Me.Show
       Call ShowAllForms
       
End Sub

Private Sub mPopLogout_Click()

    'Delete the Systray Icon
    subSysTrayDeleteIcon
    
    Unload Me

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
     On Error Resume Next
     
      'this procedure receives the callbacks from the System Tray icon.
      Dim result As Long
      Dim msg As Long
      
       'the value of X will vary depending upon the scalemode setting
       '*** NOTE:  The Form "ScaleMode" MUST be set to "Twip" OR "Pixel"
       '           for the Msg to get the correct CONSTANT value!
       If Me.ScaleMode = vbPixels Then
        msg = X
       Else
        msg = X / Screen.TwipsPerPixelX
       End If
       
'       Debug.Print Msg
       
       Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
'            Me.WindowState = vbNormal
'            result = SetForegroundWindow(Me.hwnd)
''            Me.Show
'            ShowAllForms
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            result = SetForegroundWindow(Me.hwnd)
'            Me.Show
            ShowAllForms
        Case WM_RBUTTONUP        '517 display popup menu
            result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupSys
       End Select
       
End Sub

