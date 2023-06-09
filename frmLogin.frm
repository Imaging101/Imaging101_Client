VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login - Imaging101"
   ClientHeight    =   6015
   ClientLeft      =   4230
   ClientTop       =   2775
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0442
   ScaleHeight     =   6015
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Get User Groups"
      Height          =   465
      Left            =   8520
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4815
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   5760
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboDomains 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5775
      TabIndex        =   3
      Top             =   1875
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5775
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2730
      Width           =   3360
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   5640
      Picture         =   "frmLogin.frx":0BF2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdLOGIN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&LOG IN"
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
      Height          =   810
      Left            =   5835
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmLogin.frx":117C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3405
      Width           =   3285
   End
   Begin VB.TextBox txtUserID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5775
      TabIndex        =   0
      Top             =   2355
      Width           =   3345
   End
   Begin VB.CommandButton cmdUserInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Inf&o"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdShowWinsock 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Winsock"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6960
      Top             =   4800
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
      Left            =   6315
      TabIndex        =   19
      Top             =   1260
      Width           =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C07000&
      BorderWidth     =   34
      X1              =   0
      X2              =   10320
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C07000&
      BorderWidth     =   24
      X1              =   0
      X2              =   10320
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Label lblUserID 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   270
      Index           =   0
      Left            =   4800
      TabIndex        =   10
      Top             =   2475
      Width           =   945
   End
   Begin VB.Label lblDomains 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Domain"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   270
      Left            =   4800
      TabIndex        =   16
      Top             =   1950
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   6000
      Picture         =   "frmLogin.frx":1E46
      Top             =   360
      Width           =   3810
   End
   Begin VB.Image Image1 
      Height          =   3075
      Left            =   0
      Picture         =   "frmLogin.frx":348E
      Top             =   1215
      Width           =   4485
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
      Left            =   1650
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Label lblLabels 
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
      Left            =   3705
      TabIndex        =   14
      Top             =   4305
      Visible         =   0   'False
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
      Left            =   1665
      TabIndex        =   13
      Top             =   4800
      Width           =   4440
   End
   Begin VB.Label lblLabels 
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
      TabIndex        =   12
      Top             =   4770
      Width           =   1440
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07000&
      Height          =   270
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   2835
      Width           =   1080
   End
   Begin VB.Label lblLabels 
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
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1440
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
      Left            =   1680
      TabIndex        =   8
      Top             =   4560
      Width           =   4440
   End
   Begin VB.Label lblConnecting 
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   8415
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
    
Dim ConApplication As ADODB.Connection
Dim rsApplication As ADODB.Recordset

Public LoginSucceeded As Boolean
Dim bolConnectionSucceeded As Boolean

Dim txtLoadLogFileName As String



Private Sub cboDomains_Click()

    Timer1.Enabled = False
    lblConnecting.Visible = False
    
    'Get SITE Information
    subGetSiteInfo
    

End Sub

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    End
End Sub

Private Sub cmdLogin_Click()

    subShowStatus "Connecting to I101 Server... "
    DoEvents
    
    cboDomains_Click
    
    
    '*** DON'T attempt LOGIN if Connection to I101Server Failed!
    If bolConnectionSucceeded = False Then
        subShowStatus "Connection to I101 Server FAILED... "
        DoEvents
        
        Timer1.Enabled = False
        Exit Sub
    End If

    subShowStatus "Connection to I101 Server SUCCESSFUL..."
    DoEvents
    
    '*** Set the UserName so we know who is attempting to Log In
    gsecUserName = txtUserID

    '*** Start Timer while waiting for Login
    Timer1.Enabled = True
    
    '*** Disable Buttons to make sure they are not clicked more than once!
    cmdLOGIN.Enabled = False
    cmdCancel.Enabled = False
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    
'''    'Get SITE Information
'''    subGetSiteInfo

'''    If frmImaging101Winsock.Winsock1.State <> sckConnected Then
'''            frmImaging101Winsock.cmdConnect_Click
'''            DoEvents
'''            frmImaging101Winsock.funcWaitForDataToArrive
'''    End If
    
    subShowStatus "Get SQL DB Connection Settings..."
    
    RegImaging101ConnectionString = frmImaging101Winsock.funcSendData("GET IMAGING101 CONNECTION STRING")
        If Left(RegImaging101ConnectionString, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the Imaging101ConnectionString... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
            GoTo LOGIN_EXIT
        End If
    
    RegImaging101BatchListConnectionString = frmImaging101Winsock.funcSendData("GET IMAGING101 BATCHCONNECTION STRING")
        If Left(RegImaging101BatchListConnectionString, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the Imaging101BatchConnectionString... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
            GoTo LOGIN_EXIT
        End If
    
    RegDocTypeListConnectionString = frmImaging101Winsock.funcSendData("GET DOCTYPELIST CONNECTION STRING")
        If Left(RegImaging101ConnectionString, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the DocTypeList Connection String... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
            GoTo LOGIN_EXIT
        End If
    
    RegLookupListConnectionString = frmImaging101Winsock.funcSendData("GET LOOKUP CONNECTION STRING")
        If Left(RegImaging101BatchListConnectionString, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the Lookup Connection String... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
            GoTo LOGIN_EXIT
        End If
        
'    RegRootDirToStoreObjects = frmImaging101Winsock.funcSendData("GET ROOT DIR TO STORE OBJECTS")
'        If Left(RegImaging101BatchListConnectionString, 5) = "ERROR" Then
'            MsgBox "The Server did NOT return the Imaging101BatchConnectionString... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
'            frmImaging101Winsock.cmdClose_Click
'            GoTo LOGIN_EXIT
'        End If
    
    On Error GoTo LOGIN_ERROR_TRAP
    
    
    
    '*******************************************************************************
    '*** UPDATE the SQL Database as NEEDED to ensure we have the
    '*** latest Database modifications.
    '***   NOTE:  This must be done AFTER Getting the Database Connection settings
    '***   7/22/2005 Jacob - Moved before the actual LOGIN to
    '***                     allow mods to the Security Tables
    
    subShowStatus "Update the SQL Database."
    
    subUpdateSQL

    
        
    '********************************************************************************
    '*** NOW LOG-IN
    
    subShowStatus "Connect to Database to Log In User:  " & txtUserID
    
    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101Security where UserID = '" & txtUserID & "'"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    
    con.Errors.Clear
    rs.Open
    
    
    'check for correct password
'    If txtPassword = rs.Fields!Password & "" Then
'*** Changed on 2004-01-03 to Test USERID instead of Password because I added the Domain Validation Code above!
'     Just need to make sure the USERID Exists in Imaging101's DB to be able to get the Security global detail
    
    If UCase(txtUserID) = UCase(rs.Fields!UserID) & "" Then
     
         subShowStatus "Getting Security Settings for  User:  " & txtUserID
             
         gsecSecurityRECID = rs.Fields!SecurityRECID & ""
         gsecUserID = rs.Fields!UserID & ""
         gsecUserName = rs.Fields!username & ""
         gsecPassword = rs.Fields!Password & ""
         
         gsecBatchDefaultApplication = rs.Fields!BatchDefaultApplication & ""
         gsecRightsAdminSystem = rs.Fields!RightsAdminSystem & ""
         
        'Close connection and the recordset
        rs.Close
        Set rs = Nothing
        con.Close
        Set con = Nothing
    
    '******************************************************************************
    '   VALIDATE PASSWORD - BEGIN
    '******************************************************************************
        
   '*** 2020-07-13 - Jacob - Added
   subShowStatus "Validating Password for user:  " & txtUserID

    
    '*** Notes for the SSPValidateUser() function
    '    An Empty/Null Password will Validate as TRUE... We MUST make sure a Password is entered.
    '    Passing an Empty Domain will Validate the Currently logged Domain.
    '    we will Pass in an Empty Domain for NOW until we have time to correct.
    Dim txtDomain As String
    txtDomain = ""
    
    ' If the user Logging in is the same as the user Logged In on the Workstation
    '   bypass the Domain Validation
    If UCase(txtUserID) <> UCase(rgbGetUserName()) Then
    
        If Trim(txtPassword) = "" Then
            MsgBox "You MUST enter a Valid Password...", vbInformation
            Me.Show
            txtPassword.SetFocus
            GoTo LOGIN_EXIT
        End If
            
        subShowStatus "Validate Password for user => " & txtUserID
        
        'If a Password is entered Validate Password on Domain
        If SSPValidateUser(txtUserID, txtDomain, txtPassword) <> True Then
            'If a Password DID NOT Validate on Domain - Check for Imaging101 Password
            If txtPassword <> gsecPassword Then
                MsgBox "Invalid Password, try again!", , "Login Failed"
                txtPassword.SetFocus
                GoTo LOGIN_EXIT
            End If
        End If
    End If

    '******************************************************************************
    '   VALIDATE PASSWORD - END
    '******************************************************************************
        
        subShowStatus "Validation succeeded... Send Login Request to I101 Server."
        
        If frmImaging101Winsock.funcSendData("LOGIN") <> "Logged IN" Then
            '*** 2020-07-13 - Jacob - Added
            subShowStatus "The Server did NOT respond to Login Request"
            MsgBox "The Server did NOT respond to the LOGIN request... Please try again.", vbInformation, "Login Failed"
            GoTo LOGIN_EXIT
        End If
            

        'Clear password for security on return to Login screen.
        frmLogin.txtPassword = ""

        'Unload the Splash Screen
        Unload frmSplash
        
        Me.Hide
        
        'See if Systray option is enabled.
        If bolSysTrayActive Then
            subShowStatus "Systray Enabled"
            'Add Systray Icon
            frmMainMenu.subSysTrayAddIcon
            'Check if Menu should be displayed
            If bolShowMenu Then
                subShowStatus "  Show Menu"
                frmMainMenu.Show
            Else
                subShowStatus "  Hide Menu"
                frmMainMenu.Hide
            End If
        
        Else
            'Don't add SysTray icon and show the Main Menu
            subShowStatus "Systray Disabled - Show Menu"
            frmMainMenu.Show
        End If
        
            '*** 2020-07-13 - Jacob - Added
            subShowStatus "Login Succeeded..."
        
        LoginSucceeded = True
        
        frmImaging101Winsock.TimerHeartBeat.Enabled = True
        subShowStatus "Heartbeat Timer Enabled"
        
        
    Else
        'Close connection and the recordset
        rs.Close
        Set rs = Nothing
        con.Close
        Set con = Nothing
    
        'Unload the Splash Screen
        subShowStatus "Login Failed - Invalid UserID or Password"
        
        MsgBox "Login FAILED" & vbCrLf & vbCrLf & "Invalid UserID or Password" & vbCrLf & vbCrLf & "Please try again.", , "Login Failed"
        
        Unload frmSplash

        Me.Show
        txtPassword.SetFocus
        GoTo LOGIN_EXIT
        
    End If
    
    funcWriteToDebugLog Me.name, "Load frmPopUpNotifyForm"
    
    Load frmPopUpNotifyForm
    
LOGIN_EXIT:
    Timer1.Enabled = False
    lblConnecting.Caption = ""
    cmdLOGIN.Enabled = True
    cmdCancel.Enabled = True
   
    funcWriteToDebugLog Me.name, "Unload frmSplash"

    Unload frmSplash
    
    funcWriteToDebugLog Me.name, "*** EXIT LOGIN"
    
Exit Sub
    
LOGIN_ERROR_TRAP:
    
    
    'Ignore errors
    On Error Resume Next
    If con.State = adStateOpen Then
        'Close connection and the recordset
        If rs.State = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
        con.Close
        Set con = Nothing
        MsgBox "LOGIN Error: " & Err.Number & " - " & Err.Description, vbCritical
    Else
        MsgBox "LOGIN Error:  CANNOT CONNECT TO THE DATABASE - Please contact the system Administrator...  Error [" & Err.Number & "]- " & Err.Description, vbCritical
        
    End If
    
    ' Jump back up to the LOGIN_EXIT to reset the timer,label, etc.
    GoTo LOGIN_EXIT
    
    'This Resume is a DUMMY for Testing only
    Resume Next
    
End Sub


Private Sub cmdShowWinsock_Click()

    frmImaging101Winsock.Show

End Sub

Private Sub cmdUserInfo_Click()
    
    frmUserInfo.Show
    
End Sub


Private Sub Command1_Click()

   ' Call funcGetDomainGroupsAndUsers2
   
   Dim result() As String
    result = AllGroups()
    
End Sub

Private Sub Form_Activate()

    Dim strPicturePath As String
    
   
    
    'If Main Menu Graphic Picture exists... Load it!
    strPicturePath = App.Path & "\I101MainMenuPicture.jpg"
    If funcFileExists(strPicturePath) Then
        Me.Image1.Picture = LoadPicture(strPicturePath)
    End If
    
    
    'Initialize the Index Form Load Complete global variable
    bolIndexFormLoadComplete = False
    
    'Initialize the flag for tracking if we already displayed the
    '  WARNING for modifying Original documents!
    bolAllowModificationOfOrigDocsMessageDisplayed = False
    
    'Stop on the password
    Me.Show
    txtPassword.SetFocus


        

    
End Sub

Private Sub Form_Load()

   funcWriteToDebugLog Me.name, "frmLogin() | *** ENTERING  FormLoad()"

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    'Show the Splash Screen
    frmSplash.Show

    
   'Make sure user CANNOT load multiple instances of Imaging101.EXE
    If App.PrevInstance = True Then
        MsgBox "An instance of Imaging101 Client" & vbCrLf & "is already running..." & vbCrLf & vbCrLf & "If you can't find it," & vbCrLf & vbCrLf & "press Alt-Ctrl-Del to end it.", vbOKOnly
        Unload Me
    End If

    
    '*** SET UP COMMAND PARAMETER VARIABLES
    Dim intSysTrayActive As Integer
    Dim intShowMenu As Integer
    Dim intDebugMode As Integer
    Dim intNoExit As Integer
    Dim intForceDBUpdate As Integer
    Dim intForceLogin As Integer
    
   funcWriteToDebugLog Me.name, "frmLogin() | SET the Command Line Parameter FLAGS | Command= " & Command
    
    'SET the Command Line Parameter FLAGS
    intSysTrayActive = InStr(1, UCase(Command), "/SYSTRAY")
    intShowMenu = InStr(1, UCase(Command), "/SHOWMENU")
    intDebugMode = InStr(1, UCase(Command), "/DEBUG")
    intNoExit = InStr(1, UCase(Command), "/NOEXIT")
    intForceDBUpdate = InStr(1, UCase(Command), "/FORCEDBUPDATE")
    intForceLogin = InStr(1, UCase(Command), "/FORCELOGIN")
    
    'Don't show login form if Systray option is enabled.
    
    bolSysTrayActive = False
    bolDebug = False
    bolShowMenu = False
    bolNoExit = False
    bolForceDBUpdate = False
    bolForceLogin = False
    
    If intSysTrayActive > 0 Then
        bolSysTrayActive = True
    End If
    
    If intShowMenu > 0 Then
        bolShowMenu = True
    End If
    
    If intDebugMode > 0 Then
        bolDebug = True
    End If
    
    If intNoExit > 0 Then
        bolNoExit = True
    End If

    If intForceDBUpdate > 0 Then
        bolForceDBUpdate = True
    End If
    
    If intForceLogin > 0 Then
        bolForceLogin = True
    End If
    
    If bolDebug Then
            Me.Caption = Me.Caption & " - DEBUG"
    End If

    
'    'Initialize the LOG File
'    txtLoadLogFileName = funcGetTempDir & "Imaging101Load.log"
'    Open txtLoadLogFileName For Output As #77
    funcWriteToDebugLog Me.name, " "
    funcWriteToDebugLog Me.name, " "
    funcWriteToDebugLog Me.name, "**************************************************************************"
    funcWriteToDebugLog Me.name, " LOAD IMAGING101 CLIENT - " & lblVersion.Caption
    funcWriteToDebugLog Me.name, " "
    funcWriteToDebugLog Me.name, " Command Line = " & UCase(Command)
    funcWriteToDebugLog Me.name, " "
    
    
'     '**********************************************************************************************
'    '*** 2021-08-10 - Jacob - Added GdPicture.14 Reference and Register Key
'
'    Dim oLicenseManager As New LicenseManager
'    oLicenseManager.RegisterKEY ("21187669691526875171712354586964174544")
'
    
    
    '********************************************************************************
    '*** LOCATE AND CORRECT Imaging101Client.INI LOCATION
    
    Call funcIniAndCfgFileSetup
    
    
    
    
    '**********************************************************
    '*** DOMAIN SELECTION BEGIN
    
        frmSplash.lblMessage = "Load List of Imaging101 Host Domains"

        Dim i As Integer
        Dim strDomain As String
        Dim strHostToLoad As String
        
        lblDomains.Visible = True
        cboDomains.Visible = True
        
        strDomain = VBGetPrivateProfileString(RegAppname, "Imaging101_RemoteHost", RegFileName)
        
        'IF SERVER NAME HAS NOT BEEN SET... AFTER INITIAL INSTALL
        If strDomain = "SERVERNAME" Then
            '2015-09-04 - Jacob - Set the Server name in the Domains ComboBox to use when loggin in below
            strDomain = funcAskUserFor_Imaging101_RemoteHost
            cboDomains.Text = strDomain
       End If
       
        
        'Add and Set the Default Domain First
        cboDomains.AddItem strDomain
        cboDomains.Text = strDomain
        
        On Error Resume Next
        For i = 1 To 99
            strHostToLoad = "Imaging101_RemoteHost" & i
            strDomain = VBGetPrivateProfileString(RegAppname, strHostToLoad, RegFileName)
            If strDomain <> "" Then
                cboDomains.AddItem strDomain
            End If
        Next
        

        frmImaging101Winsock.txtRemoteHost = cboDomains.Text
        
    '*** DOMAIN SELECTION END
    
    



    '*** Get USERID here but rest of info in subGetSiteInfo
    txtUserID = rgbGetUserName()

    frmSplash.lblMessage = "Attempting Login"

    If Not bolForceLogin Then
        ' Attempt to Login Automatically
        cmdLogin_Click
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
''''  UnloadAllForms
    
    'Close the LOG File
'    Close #77
    
    End
    
End Sub

Private Sub subGetSiteInfo()

    lblConnecting.Visible = True
    Timer1.Enabled = True
    cmdLOGIN.Enabled = False
    
    'RESET Site Info Labels
    lblLicencedTo.Caption = "*** NOT CONNECTED ***"
    lblLicenseNumber.Caption = ""
    lblProcessorID = ""
    
    
    frmImaging101Winsock.txtRemoteHost = cboDomains.Text
    
    subShowStatus "CONNECTING TO IMAGING101 SERVER [" & cboDomains.Text & "]..."
    DoEvents
    
    frmImaging101Winsock.cmdConnect_Click
    DoEvents
    
    frmImaging101Winsock.funcWaitForDataToArrive
    
    If frmImaging101Winsock.Winsock1.State <> sckConnected Then
    
                MsgBox "The Server [" & cboDomains.Text & "] did NOT respond with 'Connected'... " & vbCrLf & _
                "Please try again." & vbCrLf & _
                "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "Connection Failed"
                
           
            frmImaging101Winsock.cmdClose_Click
            txtPassword.SelLength = Len(txtPassword)
            bolConnectionSucceeded = False
            Timer1.Enabled = False
            lblConnecting.Visible = False
            cmdLOGIN.Enabled = True
            'Unload the Splash Screen
            Unload frmSplash
            
            Exit Sub
    End If
    
    bolConnectionSucceeded = True
    cmdLOGIN.Enabled = True

    subShowStatus "GET CLIENT SHORTNAME..."
    DoEvents

    gsecSiteInformationClientShort = frmImaging101Winsock.funcSendData("GET CLIENT SHORTNAME")
        If Left(gsecSiteInformationClientShort, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the Client Short Name... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
'            GoTo LOGIN_EXIT
        End If
        
    subShowStatus "GET CLIENT LONGNAME..."
    DoEvents
        
    gsecSiteInformationClientLong = frmImaging101Winsock.funcSendData("GET CLIENT LONGNAME")
        If Left(gsecSiteInformationClientLong, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the Client Long Name... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
'            GoTo LOGIN_EXIT
        End If
    
    subShowStatus "GET CLIENT LICENSE..."
    DoEvents
    
    gsecSiteInformationLicenseCode = frmImaging101Winsock.funcSendData("GET CLIENT LICENSE")
        If Left(gsecSiteInformationLicenseCode, 5) = "ERROR" Then
            MsgBox "The Server did NOT return the Client License... Please try again." & vbCrLf & "If the problem persists, please have the system administrator check the Imaging101 Server.", vbInformation, "GetConnectionString Failed"
            frmImaging101Winsock.cmdClose_Click
'            GoTo LOGIN_EXIT
        End If
        
        
    '*** 2022-11-13 - Jacob - Disabled GET PROCESSOR ID because we don't use it
    '                                        since we disabled BarCode detection
   
'    subShowStatus "GET PROCESSOR ID..."
'    DoEvents
'
'    '*** 2020-07-13 - Jacob - Enabled gProcessorID
'    gProcessorID = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
'
'    '*** 2022-11-13 - Jacob - Added Dim txtComputerName
'    Dim txtComputerName As String
'    txtComputerName = rgbGetComputerName()
        
    lblLicencedTo.Caption = gsecSiteInformationClientLong
    lblLicenseNumber.Caption = gsecSiteInformationLicenseCode
'    lblProcessorID = gProcessorID



End Sub

Private Sub subShowStatus(strStatus As String)


    StatusBar1.Panels(2).Text = strStatus
    frmSplash.strConnectionStatus = strStatus

'    Print #77, Now() & ": " & strStatus
    funcWriteToDebugLog Me.name, strStatus
    
    DoEvents
    
End Sub






Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '*** 2022-05-17 - Jacob - Added Ability to Enable / Disable DEBUG Mode with Ctrl+RightClick on frmLogin Image.

    'Constant (Button) Value Description
    'vbLeftButton      1     Left button is pressed
    'vbRightButton     2     Right button is pressed
    'vbMiddleButton    4     Middle button is pressed
    '
    'Constant (Shift) Value Description
    'vbShiftMask      1     SHIFT key is pressed.
    'vbCtrlMask       2     CTRL key is pressed.
    'vbAltMask        4     ALT key is pressed.

    If Button = vbRightButton And Shift = vbCtrlMask Then
        
        If bolDebug = False Then
                bolDebug = True
                MsgBox "DEBUG ENABLED." & vbCrLf & "When you click OK, I will open the folder" & vbCrLf & "where the DEBUG files will be stored.", vbOKOnly
                
                    'Get the AppData directories
                    Dim strNewLocationForINI As String
                    Dim strLocalAppDataDir As String
                    strLocalAppDataDir = Environ$("LocalAppData")
                    strNewLocationForINI = strLocalAppDataDir & "\Imaging101"
                    'Open the Imaging101 directory
                    Call Shell("Explorer.exe t,""" & strNewLocationForINI & """", vbNormalFocus)

        Else
                MsgBox "DEBUG DISABLED", vbOKOnly
                bolDebug = False
        End If
        
    End If

End Sub

Private Sub Timer1_Timer()

    lblConnecting.Caption = "Connecting... " & Now()
    StatusBar1.Panels(1).Text = "Connecting... " & Now()
    
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)

End Sub

Private Sub txtUserID_GotFocus()
    txtUserID.SelStart = 0
    txtUserID.SelLength = Len(txtUserID)
    
End Sub



Private Sub subUpdateSQL()

    'UPDATE the SQL Database IF NEEDED
    funcWriteToDebugLog Me.name, "UPDATE the SQL Database IF NEEDED"


    Dim bolSkipUpdate As Boolean
    Dim lngDBVersion As Long
    
    '*** GET THE DATABASE VERSION from the Database
    '*** RESUME on Error just in case the DB Version Field has NOT been created yet!
    On Error Resume Next
    
       '*** 2020-07-13 - Jacob - Added
   subShowStatus "GET Imaging101 DATABASE VERSION"
   
    lngDBVersion = funcGetFieldFromDB(RegImaging101ConnectionString, "I101Control", "ID=1", "DataBaseVersionI101Client")
    If lngDBVersion > 0 Then
        '*** 2020-07-13 - Jacob - Added
        subShowStatus "Current Imaging101 DATABASE VERSION = " & lngDBVersion
    Else
        '*** 2020-07-13 - Jacob - Added
        subShowStatus "ERROR getting Current Imaging101 DATABASE VERSION."
        MsgBox "ERROR getting Current Imaging101 DATABASE VERSION! " & vbCrLf & vbCrLf & "This could mean that" & vbCrLf & vbCrLf & "the SQL Server" & vbCrLf & vbCrLf & "is NOT Responding.", vbCritical, "Get DB Version Error"
        Exit Sub
    End If
    
    
    On Error GoTo ERROR_HANDLER
    
    
    '****************************************************************************************
    '*** CHECK DATABASE VERSION
    '****************************************************************************************
    
    '*** The Database Version is set to the "Version Number"
    '*** of the Imaging101 Client
    '*** in the format MMMmmmrrr
    '*** where MMM = Major (WITHOUT Leading Zeros),
    '***       mmm = minor (with Leading Zeros),
    '***       rrr = revision (with Leading Zeros)
    '***
    '*** It MUST BE CHANGED MANUALLY after ANY Change to the DB structure!!!
    '***
    
    'Database Version Format: Major.Minor.Release

    Dim intAppMajor As Integer
    Dim intAppMinor As Integer
    Dim intAppMajorMult As Long
    Dim intAppMinorMult As Long
    Dim intAppRevision As Integer
    Dim fDBversion As Long
    
    intAppMajor = CInt(App.Major)
    intAppMinor = CInt(App.Minor)
    intAppRevision = CInt(App.Revision)
    intAppMajorMult = 1000000
    intAppMinorMult = 1000
    
    'Determine what the DB Version should be for this version of the Client
    fDBversion = (intAppMajor * intAppMajorMult) + (intAppMinor * intAppMinorMult) + intAppRevision
    
            '*** 2020-07-13 - Jacob - Added
    subShowStatus "Check if SQL Database Needs to be UPDATED (DB Version " & fDBversion

    
    If (lngDBVersion < fDBversion) Or (bolForceDBUpdate = True) Then
            
        '*** 2020-07-13 - Jacob - Added
        subShowStatus "Establish Conection to SQL Database"
            
        'Establish Database Connection
        Dim con As ADODB.Connection
        Dim cmd As ADODB.Command
        
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        con.ConnectionTimeout = 120
        
        If bolForceDBUpdate = True Then
            con.CommandTimeout = 3600   '60 Minutes
            cmd.CommandTimeout = 3600
        Else
            con.CommandTimeout = 3600  '60 Minutes
            cmd.CommandTimeout = 3600
        End If
        
        con.Open RegImaging101ConnectionString
        
        Set cmd.ActiveConnection = con
        
        Err.Clear
        con.Errors.Clear
        
        '*** 2020-07-13 - Jacob - Added
        subShowStatus "Begin UPDATE of the SQL Database"
    
        con.BeginTrans
            
        '* Add new Field(s) to I101Fields */
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanAutoDetectPaperOut nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101BatchScannerSettings SET ScanAutoDetectPaperOut = '1'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanMinimumImageSize int NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101BatchScannerSettings SET ScanMinimumImageSize = 1000"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD LookupDBWhereClause nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LookupDBWhereClause = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
        '****************************************************************************************
        '*** 10/12/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FieldToSelectAfterNextPageClick nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FieldToSelectAfterNextPageClick = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
    
        
        '****************************************************************************************
        '*** 11/24/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Batches ADD BatchBoxNumber nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101BatchAudit ADD BatchBoxNumber nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        '****************************************************
        '*** Add BoxNumber field to ALL Applications
        
        Set rsApplication = New ADODB.Recordset
        Set rsApplication.ActiveConnection = con
        
        rsApplication.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    
        rsApplication.CursorLocation = adUseClient
        rsApplication.CursorType = adOpenDynamic
        rsApplication.LOCKTYPE = adLockReadOnly
        
        con.Errors.Clear
        
        rsApplication.Open
        rsApplication.MoveFirst
        
        'Prevent errors if this is a New Installation
        If rsApplication.RecordCount > 0 Then
            For intIndex = 0 To rsApplication.RecordCount - 1
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & " ADD BatchBoxNumber nvarchar (50) NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                
                rsApplication.MoveNext
            Next
        End If
        
        'Close ConApplicationnection and the recordset
        rsApplication.Close
        Set rs = Nothing
    
        '****************************
    
        '****************************************************************************************
        '*** 11/30/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchQueues ADD BatchQueueAllowScanInto nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        '****************************************************************************************
        '*** 12/01/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanBatchBoxNumberRequired nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        '****************************************************************************************
        '*** 12/16/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE DOCTYPES ADD RouteToQueue nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
    
        '****************************************************************************************
        '*** 12/17/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD RouteMaxCount nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET RouteMaxCount = '3'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        '****************************************************************************************
        '*** 12/21/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldRouteToBatchQueue nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldRouteToBatchQueue = '0'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
    
    
        '****************************************************************************************
        '*** 12/27/2004
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FieldToAssignDocumentGroup nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FieldToAssignDocumentGroup = 'Document Group'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FieldToAssignDocumentType nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FieldToAssignDocumentType = 'Document Type'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
        '****************************************************************************************
        '*** 01/03/2005
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD RootDirectoryPathForImageArchive nvarchar (255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET RootDirectoryPathForImageArchive = '" & RegRootDirToStoreObjects & "'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD RootDirectoryPathForImageAnnotations nvarchar (255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET RootDirectoryPathForImageAnnotations = '" & RegRootDirToStoreObjects & "'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
    
        '****************************************************************************************
        '*** 03/24/2005
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanDuplex nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanFileType nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanColorFormat nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchScannerSettings ADD ScanCompression nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
    
    
        '****************************************************************************************
        '*** 06/27/2005
        '****************************************************************************************
        
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD LookupDBTableIsOnSQLServer nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LookupDBTableIsOnSQLServer = '1'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
    
        '****************************************************************************************
        '*** 06/29/2005
        '****************************************************************************************
        
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD RootDirectoryPathForBatches nvarchar (255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
    
        '****************************************************************************************
        '*** 7/19/2005 - ADD I101DocPackage Header Table
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = ""
        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101DocPackage] ("
        cmd.CommandText = cmd.CommandText & "  DocPackageRECID float NULL,"
        cmd.CommandText = cmd.CommandText & "  ApplicationRECID float NULL,"
        cmd.CommandText = cmd.CommandText & "  DocPackageDescription nvarchar (60) NULL,"
        cmd.CommandText = cmd.CommandText & "  DocPackageNotes nvarchar (250) NULL,"
        cmd.CommandText = cmd.CommandText & "  DocPackageSendMode nvarchar (10) NULL,"
        cmd.CommandText = cmd.CommandText & "  DocPackageSendFormat nvarchar (10) NULL "
        cmd.CommandText = cmd.CommandText & ") ON [PRIMARY]"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
           cmd.CommandText = "CREATE  INDEX [DocPackageRECIDI] ON [dbo].[I101DocPackage]([DocPackageRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
           
           cmd.CommandText = "CREATE  INDEX [ApplicationRECIDI] ON [dbo].[I101DocPackage]([ApplicationRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
        
           cmd.CommandText = "CREATE  INDEX [DocPackageDescriptionI] ON [dbo].[I101DocPackage]([DocPackageDescription]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
        End If
    
    
        '****************************************************************************************
        '*** 7/19/2005 - ADD I101DocPackageDetail Table
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = ""
        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101DocPackageDetail] ("
        cmd.CommandText = cmd.CommandText & "  DocPackageDetailRECID float NULL,"
        cmd.CommandText = cmd.CommandText & "  DocPackageRECID float NULL,"
        cmd.CommandText = cmd.CommandText & "  DOCTYPE nvarchar (255) NULL,"
        cmd.CommandText = cmd.CommandText & "  DocPackageDetailOrder int NULL "
        cmd.CommandText = cmd.CommandText & ") ON [PRIMARY]"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
           cmd.CommandText = "CREATE  INDEX [DocPackageDetailRECIDI] ON [dbo].[I101DocPackageDetail]([DocPackageDetailRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
           
           cmd.CommandText = "CREATE  INDEX [DocPackageRECIDI] ON [dbo].[I101DocPackageDetail]([DocPackageRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
           
           cmd.CommandText = "CREATE  INDEX [DocPackageDetailOrderI] ON [dbo].[I101DocPackageDetail]([DocPackageDetailOrder]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
        End If
    
    
        '****************************************************************************************
        '*** 7/19/2005 - ADD RightsDocPackage Field to I101Security Table
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsDocPackage nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Security SET RightsDocPackage = '0'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
        '****************************************************************************************
        '*** 06/29/2005
        '****************************************************************************************
        
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101AutoImport ADD AutoImportMode nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101AutoImport SET AutoImportMode = 'B'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
    
    
    
        '****************************************************************************************
        '*** 12/02/2004 - Created Originally
        '*** 11/04/2005 - Modified and moved to here because it was not working at previous location
        '***               no real explanation as to why though...
        '***               Just moving it to here corrected the problem.
        '****************************************************************************************
        
        '****************************************************
        '*** Add DetailRotation field to ALL Applications
        
        Set rsApplication = New ADODB.Recordset
        Set rsApplication.ActiveConnection = con
        
        rsApplication.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    
        rsApplication.CursorLocation = adUseClient
        rsApplication.CursorType = adOpenDynamic
        rsApplication.LOCKTYPE = adLockReadOnly
        
        con.Errors.Clear
        
        rsApplication.Open
        rsApplication.MoveFirst
        
        'Prevent errors if this is a New Installation
        If rsApplication.RecordCount > 0 Then
            For intIndex = 0 To rsApplication.RecordCount - 1
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & "_Detail ADD DetailRotation int NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                
                rsApplication.MoveNext
            Next
        End If
        
        'Close ConApplicationnection and the recordset
        rsApplication.Close
        Set rs = Nothing
        
        
        '****************************************************
        '*** Add BatchPageRotation field to ALL Applications
        
        Set rsApplication = New ADODB.Recordset
        Set rsApplication.ActiveConnection = con
        
        rsApplication.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    
        rsApplication.CursorLocation = adUseClient
        rsApplication.CursorType = adOpenDynamic
        rsApplication.LOCKTYPE = adLockReadOnly
        
        con.Errors.Clear
        
        rsApplication.Open
        rsApplication.MoveFirst
        
        'Prevent errors if this is a New Installation
        If rsApplication.RecordCount > 0 Then
            For intIndex = 0 To rsApplication.RecordCount - 1
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & "_BatchPage ADD BatchPageRotation int NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                
                rsApplication.MoveNext
            Next
        End If
        
        'Close ConApplicationnection and the recordset
        rsApplication.Close
        Set rs = Nothing
    
    
    
        '****************************************************************************************
        '*** 07/02/2006
        '****************************************************************************************
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsExport nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Security SET RightsExport = '0'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        '****************************************************************************************
        '*** 10/12/2006
        '****************************************************************************************
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchQueues ADD ApplicationRECID Float NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101BatchQueues SET ApplicationRECID = 0"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
    
    
        '****************************************************************************************
        '*** 10/12/2006
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101TableLookupFields ALTER COLUMN LookupTableFieldName nvarchar(250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        
        '****************************************************************************************
        '*** 01/31/2007
        '****************************************************************************************
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101TableLookupFields ADD TreatAsNumeric  nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101TableLookupFields SET TreatAsNumeric = 0"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        
         '****************************************************************************************
        '*** 03/26/2007
        '****************************************************************************************
    
        '****************************************************
        '*** Add RootDirectoryPathForHtmlSource field to ALL Applications
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD RootDirectoryPathForHtmlSource nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        '****************************************************************************************
        '*** 04/27/2007 - Increase Size of UserSettingsValue field form 50 to 250
        '****************************************************************************************
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101UserSettings ALTER COLUMN UserSettingsValue nvarchar(250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
       
        
        '****************************************************************************************
        '*** 06/05/2007 - Increase Size of LookupDBTableName field form 50 to 250
        '****************************************************************************************
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ALTER COLUMN LookupDBTableName nvarchar(250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        '****************************************************************************************
        '*** 8/28/2007 - Add FieldDefaultForBarcodeOnly field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldDefaultForBarcodeOnly nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldDefaultForBarcodeOnly = '0'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        '****************************************************************************************
        '*** 8/28/2007 - Add BatchQueueNotificationFrequency field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Security ADD BatchQueueNotificationFrequency nvarchar (4) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Security SET BatchQueueNotificationFrequency = '10'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        '*** FORCE Disable if the field is NULL
        cmd.CommandText = "UPDATE I101Security SET BatchQueueNotificationFrequency = '0' " & _
                            " WHERE BatchQueueNotificationFrequency IS NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        '****************************************************************************************
        '*** 12/6/2007 - Add FieldIndexLookupDelimiter & FieldSearchCondition fields
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD ApplicationBatchNameDelimiter nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET ApplicationBatchNameDelimiter = '-'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldSearchCondition nvarchar (10) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldSearchCondition = 'Contains'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
    
    
        '****************************************************************************************
        '*** 12/13/2007 - Increase Size of BatchName field form 50 to 250
        '****************************************************************************************
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Batches ALTER COLUMN BatchName nvarchar(255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
    
    
    
        '****************************************************
        '*** Change Size of  Field to ALL Applications
        
'        Set rsApplication = New ADODB.Recordset
'        Set rsApplication.ActiveConnection = con
'
'        rsApplication.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
'
'        rsApplication.CursorLocation = adUseClient
'        rsApplication.CursorType = adOpenDynamic
'        rsApplication.LockType = adLockReadOnly
'
'        con.Errors.Clear
'
'        rsApplication.Open
'        rsApplication.MoveFirst
'
'        'Prevent errors if this is a New Installation
'        If rsApplication.RecordCount > 0 Then
'            For intIndex = 0 To rsApplication.RecordCount - 1
'
'                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & " ALTER COLUMN DocumentBatchName nvarchar(250) NULL"
'                txtActionBeforeError = cmd.CommandText
'                cmd.Execute , , adCmdText
'
'                rsApplication.MoveNext
'            Next
'        End If
'
'        'Close ConApplicationnection and the recordset
'        rsApplication.Close
'        Set rs = Nothing
    
    
    
        '****************************************************************************************
        '*** 04/21/2008 - Add FTPUserID, FTPPassword & FTPSite fields
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPUserID nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field for The Ticket Clinic
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPUserID = 'pcn'" & _
                              " WHERE ApplicationName = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPPassword nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPPassword = 'pcn00'" & _
                              " WHERE ApplicationName = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPSite nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPSite = 'my.ticketclinic.com'" & _
                              " WHERE ApplicationName = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
    
        '****************************************************************************************
        '*** 05/01/2008 - Add ApplicationBatchNameDelimiter field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD MaxItemsToRetrieve nvarchar (6) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET MaxItemsToRetrieve = '500'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        '****************************************************************************************
        '*** 05/01/2008 - Add ApplicationBatchNameDelimiter field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Security ADD AllowModificationOfOrigDocs nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Security SET AllowModificationOfOrigDocs = '0'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
    
    
        '****************************************************************************************
        '*** 05/13/2008 - ADD SMTP EMAIL Configuration Fields to I101Control Table
    
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SendEmailViaSMTP nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Control SET SendEmailViaSMTP = '0' WHERE ID=1"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPHost nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPPOP3Host nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPRequiresAuthentication nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPUsePOP3Auth nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPAuthenticationUserID nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPAuthenticationPassword nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPDefaultEmailSubject nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPDefaultEmailMessage nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
    
    
    
    
        '****************************************************************************************
        '*** 05/13/2008 - ADD SMTP EMAIL Configuration Fields to I101Security Table
    
        cmd.CommandText = "ALTER TABLE I101Security ADD SMTPFromEmailAddress nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Security ADD SMTPFromEmailDisplayName nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        cmd.CommandText = "ALTER TABLE I101Security ADD SMTPSendCcToSender nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
    
    
    
        '****************************************************************************************
        '*** Add DATABASE VERSION Field
        '****************************************************************************************
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Control ADD DataBaseVersionI101Client nvarchar (10) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
    
        '****************************************************************************************
        '*** 11/13/2008 - Increase Size of FieldSearchCondition field form 10 to 15
        '****************************************************************************************
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ALTER COLUMN FieldSearchCondition nvarchar(15) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText

    
        '****************************************************************************************
        '*** 05/13/2008 - ADD FieldSearchCondition Field to I101TableLookupFields Table
    
        cmd.CommandText = "ALTER TABLE I101TableLookupFields ADD FieldSearchCondition nvarchar (15) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101TableLookupFields SET FieldSearchCondition = 'Contains'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
    
        '****************************************************
        '*** Change Size of BatchName and FileName Fields to ALL Applications
        
        Set rsApplication = New ADODB.Recordset
        Set rsApplication.ActiveConnection = con
        
        rsApplication.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    
        rsApplication.CursorLocation = adUseClient
        rsApplication.CursorType = adOpenDynamic
        rsApplication.LOCKTYPE = adLockReadOnly
        
        con.Errors.Clear
        
        rsApplication.Open
        rsApplication.MoveFirst
        
        'Prevent errors if this is a New Installation
        If rsApplication.RecordCount > 0 Then
            For intIndex = 0 To rsApplication.RecordCount - 1
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & " ALTER COLUMN DocumentBatchName nvarchar(255) NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & "_Detail ALTER COLUMN DetailFileName nvarchar(255) NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & "_Detail ALTER COLUMN DetailSubdirectory nvarchar(255) NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText

                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & "_BatchPage ALTER COLUMN BatchPageFileName nvarchar(255) NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
               
               rsApplication.MoveNext
            Next
        End If
        
        'Close ConApplicationnection and the recordset
        rsApplication.Close
        Set rs = Nothing
    
    
    
        '****************************************************************************************
        '*** 07/15/2009 - ADD BatchInQueueDate Field to I101Batches Table
    
        cmd.CommandText = "ALTER TABLE I101Batches ADD BatchInQueueDate DateTime NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Batches SET BatchInQueueDate = BatchScanDate"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
    
    
        cmd.CommandText = "ALTER TABLE I101BatchAudit ADD BatchInQueueDate DateTime NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
    
    
    
        '****************************************************************************************
        '*** 07/16/2009 - ADD Batch Restrictions Fields to I101Batches Table
    

        cmd.CommandText = "ALTER TABLE I101Security ADD RightsBatchFindRestricted nvarchar(2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsBatchFindRestrictToQueue nvarchar(2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsBatchFindRestrictToOwner nvarchar(2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsBatchChangeQueue nvarchar(2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsBatchChangeOwner nvarchar(2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
    
    
        '****************************************************************************************
        '*** 08/10/2009 - Correct Field Type Error introduced in 7/16/2009 (v6.4.1)
    
         cmd.CommandText = "UPDATE I101Security SET RightsBatchFindRestricted = '0' " & _
                           " WHERE (RightsBatchFindRestricted = '') OR (RightsBatchFindRestricted IS NULL) "
         txtActionBeforeError = cmd.CommandText
         cmd.Execute , , adCmdText

         cmd.CommandText = "UPDATE I101Security SET RightsBatchFindRestrictToQueue = '0' " & _
                           " WHERE (RightsBatchFindRestrictToQueue = '') OR (RightsBatchFindRestrictToQueue IS NULL) "
         txtActionBeforeError = cmd.CommandText
         cmd.Execute , , adCmdText
         
         cmd.CommandText = "UPDATE I101Security SET RightsBatchFindRestrictToOwner = '0' " & _
                           " WHERE (RightsBatchFindRestrictToOwner = '') OR (RightsBatchFindRestrictToOwner IS NULL) "
         txtActionBeforeError = cmd.CommandText
         cmd.Execute , , adCmdText
         
         cmd.CommandText = "UPDATE I101Security SET RightsBatchChangeQueue = '0' " & _
                         " WHERE (RightsBatchChangeQueue = '') OR (RightsBatchChangeQueue IS NULL) "
         txtActionBeforeError = cmd.CommandText
         cmd.Execute , , adCmdText
         
         cmd.CommandText = "UPDATE I101Security SET RightsBatchChangeOwner = '0' " & _
                           " WHERE (RightsBatchChangeOwner = '') OR (RightsBatchChangeOwner IS NULL) "
         txtActionBeforeError = cmd.CommandText
         cmd.Execute , , adCmdText
         


        '****************************************************************************************
        '*** 08/11/2009 - ADD Batch Restrictions Fields to I101Batches Table
    
        cmd.CommandText = "ALTER TABLE I101Security ADD RightsBatchAllowDocTypeEdit nvarchar(2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "UPDATE I101Security SET RightsBatchAllowDocTypeEdit = '0' " & _
                          " WHERE (RightsBatchAllowDocTypeEdit = '') OR (RightsBatchAllowDocTypeEdit IS NULL) "
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        
        '****************************************************************************************
        '*** 01/11/2010 - ADD Batch Restrictions Fields to I101Batches Table
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD AutoLookupOnBatchLoad nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET AutoLookupOnBatchLoad = '1'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        
        '****************************************************************************************
        '*** 09/06/2010 - ADD I101SecurityRoleApp Table
        '                 Security by Role by Application
        
        bolSkipUpdate = False
        cmd.CommandText = ""
        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101SecurityRoleApp] ("
        cmd.CommandText = cmd.CommandText & "  [SecurityRoleRECID] [float] NOT NULL,"
        cmd.CommandText = cmd.CommandText & "  [ApplicationRECID] [float] NULL,"
        cmd.CommandText = cmd.CommandText & "  [SecurityRECID] [float] NOT NULL,"
        cmd.CommandText = cmd.CommandText & "  [AllowModificationOfOrigDocs] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [BatchDefaultApplication] [nvarchar](30) NULL,"
        cmd.CommandText = cmd.CommandText & "  [BatchDefaultQueue] [nvarchar](50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [BatchListOrder] [nvarchar](255) NULL,"
        cmd.CommandText = cmd.CommandText & "  [BatchMode] [nvarchar](50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [BatchQueueNotificationFrequency] [nvarchar](4) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsAdminApplication] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsAdminSystem] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsAnnotate] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchAdministration] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchAllowDocTypeEdit] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchChangeOrder] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchChangeOwner] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchChangeQueue] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchCommit] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchFindRestricted] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchFindRestrictToOwner] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchFindRestrictToQueue] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchIndex] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchRoute] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchScan] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsBatchView] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsDeleteBatches] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsDeleteDocuments] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsDocPackage] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsExport] [nvarchar](2) NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsImportFromEcapture] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsImportFromFile] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsLaunchDoc] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsModifyIndexes] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsPrint] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsRetrieveImages] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsScannerSettings] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsSendMail] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [RightsThumbnails] [int] NULL,"
        cmd.CommandText = cmd.CommandText & "  [UserSupervisor] [nvarchar](50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [ViewResetImagesOnFind] [int] NULL"
        cmd.CommandText = cmd.CommandText & "  ) ON [PRIMARY]"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
           'Insert one record for each Application a User is Assigned to
           cmd.CommandText = "INSERT INTO I101SecurityRoleApp"
           cmd.CommandText = cmd.CommandText & "  (SecurityRoleRECID"
           cmd.CommandText = cmd.CommandText & "  ,ApplicationRECID"
           cmd.CommandText = cmd.CommandText & "  ,SecurityRECID"
           cmd.CommandText = cmd.CommandText & "  ,AllowModificationOfOrigDocs"
           cmd.CommandText = cmd.CommandText & "  ,BatchDefaultApplication"
           cmd.CommandText = cmd.CommandText & "  ,BatchDefaultQueue"
           cmd.CommandText = cmd.CommandText & "  ,BatchListOrder"
           cmd.CommandText = cmd.CommandText & "  ,BatchMode"
           cmd.CommandText = cmd.CommandText & "  ,BatchQueueNotificationFrequency"
           cmd.CommandText = cmd.CommandText & "  ,RightsAdminApplication"
           cmd.CommandText = cmd.CommandText & "  ,RightsAdminSystem"
           cmd.CommandText = cmd.CommandText & "  ,RightsAnnotate"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchAdministration"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchAllowDocTypeEdit"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchChangeOrder"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchChangeOwner"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchChangeQueue"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchCommit"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchFindRestricted"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchFindRestrictToOwner"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchFindRestrictToQueue"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchIndex"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchRoute"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchScan"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchView"
           cmd.CommandText = cmd.CommandText & "  ,RightsDeleteBatches"
           cmd.CommandText = cmd.CommandText & "  ,RightsDeleteDocuments"
           cmd.CommandText = cmd.CommandText & "  ,RightsDocPackage"
           cmd.CommandText = cmd.CommandText & "  ,RightsExport"
           cmd.CommandText = cmd.CommandText & "  ,RightsImportFromEcapture"
           cmd.CommandText = cmd.CommandText & "  ,RightsImportFromFile"
           cmd.CommandText = cmd.CommandText & "  ,RightsLaunchDoc"
           cmd.CommandText = cmd.CommandText & "  ,RightsModifyIndexes"
           cmd.CommandText = cmd.CommandText & "  ,RightsPrint"
           cmd.CommandText = cmd.CommandText & "  ,RightsRetrieveImages"
           cmd.CommandText = cmd.CommandText & "  ,RightsScannerSettings"
           cmd.CommandText = cmd.CommandText & "  ,RightsSendMail"
           cmd.CommandText = cmd.CommandText & "  ,RightsThumbnails"
           cmd.CommandText = cmd.CommandText & "  ,UserSupervisor"
           cmd.CommandText = cmd.CommandText & "  ,ViewResetImagesOnFind)"
           cmd.CommandText = cmd.CommandText & "SELECT"
           cmd.CommandText = cmd.CommandText & "  '0'"
           cmd.CommandText = cmd.CommandText & "  ,ApplicationRECID"
           cmd.CommandText = cmd.CommandText & "  ,I101Security.SecurityRECID"
           cmd.CommandText = cmd.CommandText & "  ,AllowModificationOfOrigDocs"
           cmd.CommandText = cmd.CommandText & "  ,BatchDefaultApplication"
           cmd.CommandText = cmd.CommandText & "  ,BatchDefaultQueue"
           cmd.CommandText = cmd.CommandText & "  ,BatchListOrder"
           cmd.CommandText = cmd.CommandText & "  ,BatchMode"
           cmd.CommandText = cmd.CommandText & "  ,BatchQueueNotificationFrequency"
           cmd.CommandText = cmd.CommandText & "  ,RightsAdminApplication"
           cmd.CommandText = cmd.CommandText & "  ,RightsAdminSystem"
           cmd.CommandText = cmd.CommandText & "  ,RightsAnnotate"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchAdministration"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchAllowDocTypeEdit"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchChangeOrder"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchChangeOwner"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchChangeQueue"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchCommit"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchFindRestricted"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchFindRestrictToOwner"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchFindRestrictToQueue"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchIndex"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchRoute"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchScan"
           cmd.CommandText = cmd.CommandText & "  ,RightsBatchView"
           cmd.CommandText = cmd.CommandText & "  ,RightsDeleteBatches"
           cmd.CommandText = cmd.CommandText & "  ,RightsDeleteDocuments"
           cmd.CommandText = cmd.CommandText & "  ,RightsDocPackage"
           cmd.CommandText = cmd.CommandText & "  ,RightsExport"
           cmd.CommandText = cmd.CommandText & "  ,RightsImportFromEcapture"
           cmd.CommandText = cmd.CommandText & "  ,RightsImportFromFile"
           cmd.CommandText = cmd.CommandText & "  ,RightsLaunchDoc"
           cmd.CommandText = cmd.CommandText & "  ,RightsModifyIndexes"
           cmd.CommandText = cmd.CommandText & "  ,RightsPrint"
           cmd.CommandText = cmd.CommandText & "  ,RightsRetrieveImages"
           cmd.CommandText = cmd.CommandText & "  ,RightsScannerSettings"
           cmd.CommandText = cmd.CommandText & "  ,RightsSendMail"
           cmd.CommandText = cmd.CommandText & "  ,RightsThumbnails"
           cmd.CommandText = cmd.CommandText & "  ,UserSupervisor"
           cmd.CommandText = cmd.CommandText & "  ,ViewResetImagesOnFind"
           cmd.CommandText = cmd.CommandText & " FROM  I101SecurityApplications CROSS JOIN I101Security"
           cmd.CommandText = cmd.CommandText & " WHERE (I101SecurityApplications.SecurityRECID = I101Security.SecurityRECID)"
           cmd.CommandText = cmd.CommandText & " ORDER BY ApplicationRECID, i101SECURITY.SecurityRECID"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
        
           'Create Indexes
           cmd.CommandText = "CREATE  INDEX [SecurityRoleRECIDI] ON [dbo].[I101SecurityRoleApp]([SecurityRoleRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
        
           cmd.CommandText = "CREATE  INDEX [ApplicationRECIDI] ON [dbo].[I101SecurityRoleApp]([ApplicationRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
        
           cmd.CommandText = "CREATE  INDEX [SecurityRECIDI] ON [dbo].[I101SecurityRoleApp]([SecurityRECID]) ON [PRIMARY]"
           txtActionBeforeError = cmd.CommandText
           cmd.Execute , , adCmdText
     
        End If
    
        
'        'Initialize the field
'        If Not bolSkipUpdate Then
'            cmd.CommandText = "UPDATE I101TableLookupFields SET FieldSearchCondition = 'Contains'"
'            txtActionBeforeError = cmd.CommandText
'            cmd.Execute , , adCmdText
'        End If
        
        '****************************************************************************************
        '*** 01/19/2011 - Add the DOCSUBTYPE field
        '****************************************************************************************
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE DOCTYPES ADD DOCSUBTYPE nvarchar (255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    

        '****************************************************************************************
        '*** 01/19/2011 - ADD FieldToAssignDocumentSubType to I101Applications Table
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FieldToAssignDocumentSubType nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        '****************************************************************************************
        '*** 04/13/2011 - ADD FTP Commit Fields
    
        cmd.CommandText = "ALTER TABLE I101Applications ADD ApplicationCommitBatchOption nvarchar (20) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
       
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameField0 nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameField1 nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameField2 nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameField3 nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameDelimiter0 nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
       
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameDelimiter1 nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
       
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPFileNameDelimiter2 nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
       
        
        cmd.CommandText = "ALTER TABLE DOCTYPES ADD CommitViaFTP nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
       
        
        bolSkipUpdate = False
        
       'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE DOCTYPES SET CommitViaFTP = 0 WHERE CommitViaFTP IS NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        
            cmd.CommandText = "UPDATE DOCTYPES SET CommitViaFTP = 1 WHERE APPLICATION = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        

        '****************************************************************************************
        '*** 05/05/2011 - ADD CommitViaFTP Field to ALL Applications
    
        
        Set rsApplication = New ADODB.Recordset
        Set rsApplication.ActiveConnection = con
        
        rsApplication.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    
        rsApplication.CursorLocation = adUseClient
        rsApplication.CursorType = adOpenDynamic
        rsApplication.LOCKTYPE = adLockReadOnly
        
        con.Errors.Clear
        
        rsApplication.Open
        rsApplication.MoveFirst
        
        'Prevent errors if this is a New Installation
        If rsApplication.RecordCount > 0 Then
            For intIndex = 0 To rsApplication.RecordCount - 1
                
                cmd.CommandText = "ALTER TABLE " & rsApplication.Fields!ApplicationName & "_BatchPage" & " ADD CommitViaFTP nvarchar (2) NULL"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                
                bolSkipUpdate = False
                
                'Initialize the field
                 If Not bolSkipUpdate Then
                     cmd.CommandText = "UPDATE " & rsApplication.Fields!ApplicationName & "_BatchPage" & " SET CommitViaFTP = 0 WHERE CommitViaFTP IS NULL"
                     txtActionBeforeError = cmd.CommandText
                     cmd.Execute , , adCmdText
                 End If
                
               
               rsApplication.MoveNext
            Next
        End If
        
        'Close ConApplicationnection and the recordset
        rsApplication.Close
        Set rs = Nothing
    
        
        
        '****************************************************************************************
        '*** 05/05/2011 - SET Default  ApplicationCommitBatchOption
        
        cmd.CommandText = "UPDATE I101Applications SET ApplicationCommitBatchOption = 'Application Only'  WHERE (ApplicationCommitBatchOption IS NULL) OR (ApplicationCommitBatchOption = '')"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        
        '****************************************************************************************
        '*** 04/21/2008 - Add FTPPort field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPPort nvarchar (6) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field for The Ticket Clinic
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPPort = '21'" & _
                              " WHERE FTPPort IS NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
       
        
        
        '****************************************************************************************
        '*** 08/24/2011 - Add FTP fields for SITE B
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPUserID_B nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field for The Ticket Clinic
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPUserID_B = 'userid'" & _
                              " WHERE ApplicationName = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPPassword_B nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPPassword_B = 'pwd'" & _
                              " WHERE ApplicationName = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPSite_B nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPSite_B = 'ftp.site.address'" & _
                              " WHERE ApplicationName = 'TTC'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
                bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD FTPPort_B nvarchar (6) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field for The Ticket Clinic
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET FTPPort_B = '21'" & _
                              " WHERE FTPPort IS NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


        '****************************************************************************************
        '*** 08/24/2011 - Add FTP fields for SITE B
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD LookupDBConnectionString_B nvarchar (255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LookupDBConnectionString_B = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD LookupDBTableName_B nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LookupDBTableName_B = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD LookupDBWhereClause_B nvarchar (250) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
    
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LookupDBWhereClause_B = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If

        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD LookupDBTableIsOnSQLServer_B nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LookupDBTableIsOnSQLServer_B = '1'"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD CaseIdCutoff float NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET CaseIdCutoff = 700000"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If

        


        '****************************************************************************************
        '*** 01/13/2012 - Increase size of I101BatchAudit DocumentBatchName field from 50 to 250
        '***                    to match size in I101Batches changed on 12/13/2007

        cmd.CommandText = "ALTER TABLE I101BatchAudit ALTER COLUMN BatchName nvarchar(255) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText


        '****************************************************************************************
        '*** 10/10/2012 - Change FieldType for fields set more than 64 to "LongText"
        '***                    to match size in I101Batches changed on 12/13/2007

        cmd.CommandText = "UPDATE I101Fields SET FieldType = 'LongText'  WHERE FieldSize > 63"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText



        '****************************************************************************************
        '*** 03/01/2013 - ADD AdvancedSearch Field to I101SecurityRoleApp Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101SecurityRoleApp ADD RightsAdvancedSearch int NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
         If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101SecurityRoleApp SET RightsAdvancedSearch = 0 "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        

        '****************************************************************************************
        '*** 05/22/2013 - Add SetUserAsBatchOwnerOnSPLIT field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD SetUserAsBatchOwnerOnSPLIT nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET SetUserAsBatchOwnerOnSPLIT = '0' WHERE SetUserAsBatchOwnerOnSPLIT is NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        

        '****************************************************************************************
        '*** 09/10/2013 - ADD SMTP EMAIL PORT Field to I101Control Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Control ADD SMTPPort nvarchar (6) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Control SET SMTPPort = '25' WHERE ID=1 AND SMTPPort IS NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        '****************************************************************************************
        '*** 04/28/2014 - ADD ckbFieldTableLookupOverridesDefault Field to
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldTableLookupOverridesDefault nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldTableLookupOverridesDefault = '1' WHERE FieldTableLookupOverridesDefault is NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


'        '****************************************************************************************
'        '*** 9/28/2012 - ADD I101eMailDocRules Table
'        '****************************************************************************************
'
'        bolSkipUpdate = False
'        cmd.CommandText = ""
'        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101eMailDocRules] ("
'        cmd.CommandText = cmd.CommandText & "  eMailDocRulesRECID float NULL,"
'        cmd.CommandText = cmd.CommandText & "  ApplicationName nvarchar (255)  NULL,"
'        cmd.CommandText = cmd.CommandText & "  FieldName_1 nvarchar (255) NULL,"
'        cmd.CommandText = cmd.CommandText & "  FieldValue_1 nvarchar (255) NULL, "
'        cmd.CommandText = cmd.CommandText & "  FieldName_2 nvarchar (255) NULL,"
'        cmd.CommandText = cmd.CommandText & "  FieldValue_2 nvarchar (255) NULL, "
'        cmd.CommandText = cmd.CommandText & "  FieldName_3 nvarchar (255) NULL,"
'        cmd.CommandText = cmd.CommandText & "  FieldValue_3 nvarchar (255) NULL, "
'        cmd.CommandText = cmd.CommandText & "  FieldName_4 nvarchar (255) NULL,"
'        cmd.CommandText = cmd.CommandText & "  FieldValue_4 nvarchar (255) NULL, "
'        cmd.CommandText = cmd.CommandText & "  eMailAddress nvarchar (255) NULL"
'
'
'        cmd.CommandText = cmd.CommandText & ") ON [PRIMARY]"
'        txtActionBeforeError = cmd.CommandText
'        cmd.Execute , , adCmdText
'
'        If Not bolSkipUpdate Then
'           cmd.CommandText = "CREATE  INDEX [eMailDocRulesRECIDI] ON [dbo].[I101eMailDocRules]([eMailDocRulesRECID]) ON [PRIMARY]"
'           txtActionBeforeError = cmd.CommandText
'           cmd.Execute , , adCmdText
'
'           cmd.CommandText = "CREATE  INDEX [ApplicationNamei] ON [dbo].[I101eMailDocRules]([ApplicationName]) ON [PRIMARY]"
'           txtActionBeforeError = cmd.CommandText
'           cmd.Execute , , adCmdText
'
'        End If
    


        '****************************************************************************************
        '*** 09/15/2016
        '****************************************************************************************
        
        'Add Route To Batch User Field in I101Fields
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldRouteToBatchUser nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldRouteToBatchUser = '0' "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


        'Add Route To Account Manager Field in I101Fields
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldRouteToBatchManager nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldRouteToBatchManager = '0' "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        

        'Add Route To Batch Manager Field in I101Batches
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Batches ADD BatchManager nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Batches SET BatchManager = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


        'Add Route To Batch Manager Field in I101BatchAudit
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101BatchAudit ADD BatchManager nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Batches SET BatchManager = ''"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


        '****************************************************************************************
        '*** 03/22/2017
        '****************************************************************************************
        
        'Add Route To Batch User Field in I101Fields
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD FieldDropDownListAlsoOnFiler nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET FieldDropDownListAlsoOnFiler = '0' "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


        '****************************************************************************************
        '*** 08/29/2017 - ADD AutoLaunchFileTypes Field to I101Control Table
    
        cmd.CommandText = "ALTER TABLE I101Control ADD AutoLaunchFileTypes nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
            cmd.CommandText = "UPDATE I101Control SET AutoLaunchFileTypes = ' ' WHERE ID=1"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText


        '****************************************************************************************
        '*** 11/07/2017 - ADD AdvancedSearch Field to I101SecurityRoleApp Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101SecurityRoleApp ADD RightsFileDocsViaI101FILER int NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
         If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101SecurityRoleApp SET RightsFileDocsViaI101FILER = 1 "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If



        '****************************************************************************************
        '*** 01/08/2017 - ADD I101SearchTemplate Table
        
        cmd.CommandText = "ALTER TABLE I101Control ADD SearchTemplateRECID float NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Initialize the field
            cmd.CommandText = "UPDATE I101Control SET SearchTemplateRECID = 1 WHERE ID=1"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText


        '****************************************************************************************
        '*** 01/08/2017 - ADD I101SearchTemplate Table
        
        bolSkipUpdate = False
        cmd.CommandText = ""
        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101SearchTemplates] ("
        cmd.CommandText = cmd.CommandText & "  [SearchTemplateRECID] [float] NOT NULL,"
        cmd.CommandText = cmd.CommandText & "  [ApplicationRECID] [float] NULL,"
        cmd.CommandText = cmd.CommandText & "  [SearchTemplateName] [nvarchar](50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [WhereFreehand] [nvarchar](4000) NULL"

        cmd.CommandText = cmd.CommandText & "  ) ON [PRIMARY]"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        'Create Indexes
       cmd.CommandText = "CREATE  INDEX [SearchTemplateRECIDI] ON [dbo].[I101SearchTemplates]([SearchTemplateRECID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
    
       cmd.CommandText = "CREATE  INDEX [ApplicationRECIDI] ON [dbo].[I101SearchTemplates]([ApplicationRECID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
        
        
        
        '****************************************************************************************
        '*** 01/08/2017 - ADD I101SearchTemplateUsers Table
        
        bolSkipUpdate = False
        cmd.CommandText = ""
        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101SearchTemplateUsers] ("
        cmd.CommandText = cmd.CommandText & "  [SearchTemplateRECID] [float] NOT NULL,"
        cmd.CommandText = cmd.CommandText & "  [SecurityRECID] [float] NULL,"

        cmd.CommandText = cmd.CommandText & "  ) ON [PRIMARY]"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText

        'Create Indexes
       cmd.CommandText = "CREATE  INDEX [SearchTemplateRECIDI] ON [dbo].[I101SearchTemplateUsers]([SearchTemplateRECID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
    
       cmd.CommandText = "CREATE  INDEX [SecurityRECIDI] ON [dbo].[I101SearchTemplateUsers]([SecurityRECID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
       
       
        '****************************************************************************************
        '*** 01/12/2018 - Add EnableSearchTemplates field
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications ADD EnableSearchTemplates nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET EnableSearchTemplates = '0' WHERE EnableSearchTemplates is NULL"
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
       

        '****************************************************************************************
        '*** 04/12/2019 - Add HideForSearchIndex field
        '****************************************************************************************
        
        'Add Route To Batch User Field in I101Fields
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Fields ADD HideForSearchIndex nvarchar (2) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Fields SET HideForSearchIndex = '0' "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If


        '****************************************************************************************
        '*** 2020-05-15 - ADD RightsEditSearchTemplates Field to I101SecurityRoleApp Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101SecurityRoleApp ADD RightsEditSearchTemplates int NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
         If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101SecurityRoleApp SET RightsEditSearchTemplates = 0 "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If
        
        
        '****************************************************************************************
        '*** 2020-07-23 - ADD IndexFullText Field to I101Applications Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications  ADD IndexFullText bit NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
         If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET IndexFullText = 0 "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If

        
        
        '****************************************************************************************
        '*** 2021/11/09 - Added New Field for selecting how to Auto-Launch documents.
        
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Control ADD AutoLaunchTo nvarchar (50) NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Control SET AutoLaunchTo = 'New Imaging101 Document Viewer' "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If



        '****************************************************************************************
        '*** 2022-07-21 - ADD I101LogDocumentActions Table
        
        'LogAction possible values:  R=Retrieve, D=Delete, M=Modify, E=Export, P=Purge

        bolSkipUpdate = False
        cmd.CommandText = ""
        cmd.CommandText = cmd.CommandText & "CREATE TABLE [dbo].[I101LogDocumentActions] ("
        cmd.CommandText = cmd.CommandText & "  [LogDate] [datetime] NOT NULL,"
        cmd.CommandText = cmd.CommandText & "  [LogAction] [nvarchar] (50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [UserID] [nvarchar] (50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [UserName] [nvarchar] (50) NULL,"
        cmd.CommandText = cmd.CommandText & "  [DocumentRECID] [float] NULL,"
        cmd.CommandText = cmd.CommandText & "  [LogNotes] [nvarchar] (250) NULL,"

        cmd.CommandText = cmd.CommandText & "  ) ON [PRIMARY]"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText

        'Create Indexes
       cmd.CommandText = "CREATE  INDEX [LogDateI] ON [dbo].[I101LogDocumentActions]([LogDate]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
    
       cmd.CommandText = "CREATE  INDEX [LogActionI] ON [dbo].[I101LogDocumentActions]([LogAction]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText

       cmd.CommandText = "CREATE  INDEX [UserIDI] ON [dbo].[I101LogDocumentActions]([UserID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
       
       cmd.CommandText = "CREATE  INDEX [UserNameI] ON [dbo].[I101LogDocumentActions]([UserName]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText

       cmd.CommandText = "CREATE  INDEX [DocumentRECIDI] ON [dbo].[I101LogDocumentActions]([DocumentRECID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText
       
       
        
        '****************************************************************************************
        '*** 2022-07-21 - ADD LogOpenedDocuments Field to I101Applications Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101Applications  ADD LogOpenedDocuments bit NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
         If Not bolSkipUpdate Then
            cmd.CommandText = "UPDATE I101Applications SET LogOpenedDocuments = 0 "
            txtActionBeforeError = cmd.CommandText
            cmd.Execute , , adCmdText
        End If

        
        
        '****************************************************************************************
        '*** 2022-09-09 - ADD ApplicationRECID Field to I101LogDocumentActions Table
    
        bolSkipUpdate = False
        cmd.CommandText = "ALTER TABLE I101LogDocumentActions  ADD [ApplicationRECID] [float] NULL"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        

        'Create Indexes
       cmd.CommandText = "CREATE  INDEX [ApplicationRECIDI] ON [dbo].[I101LogDocumentActions]([ApplicationRECID]) ON [PRIMARY]"
       txtActionBeforeError = cmd.CommandText
       cmd.Execute , , adCmdText






        '****************************************************************************************
        '****************************************************************************************
        '*** S A V E   C U R R E N T   D A T A B A S E   V E R S I O N
        '****************************************************************************************
        '****************************************************************************************
        
        cmd.CommandText = "UPDATE I101Control SET DataBaseVersionI101Client = '" & fDBversion & "'"
        txtActionBeforeError = cmd.CommandText
        cmd.Execute , , adCmdText
        
        '*** 2020-07-13 - Jacob - Added
        subShowStatus "UPDATED the SQL Database to Version: " & fDBversion








        '****************************************************************************************
        '****************************************************************************************
        '*** C O M M I T   C H A N G E S
        '****************************************************************************************
        '****************************************************************************************
    
        con.CommitTrans
        con.Close
        Set con = Nothing
        Set cmd = Nothing
        
    End If
    
    
SKIP_DB_UPDATE:

Exit Sub

ERROR_HANDLER:
   
   '01/11/2018 JACOB - Added "ALREADY EXISTS" to handle errors with Index or Statistics
    If InStr(1, UCase(con.Errors.item(0).Description), "UNIQUE") _
    Or InStr(1, UCase(con.Errors.item(0).Description), "IS ALREADY AN") _
    Or InStr(1, UCase(con.Errors.item(0).Description), "ALREADY EXISTS") Then

    
        bolSkipUpdate = True
        Resume Next
    End If
    
    '*** 2020-07-13 - Jacob - Added
    subShowStatus "SQL Database Update ERROR: " & con.Errors.item(0).Number & vbCrLf & con.Errors.item(0).Description
    MsgBox "SQL Database Update ERROR: " & con.Errors.item(0).Number & vbCrLf & con.Errors.item(0).Description & vbCrLf & "  DURING ACTION: (" & txtActionBeforeError & ") - [Transaction Rolled Back - Batch NOT Imported]" & vbCrLf & "<Database State: " & con.Errors.item(0).SQLState & ">", vbExclamation
        
    On Error Resume Next
    
    If con.State = adStateOpen Then
        con.RollbackTrans
        con.Close
    End If
    
    Set con = Nothing
    Set cmd = Nothing

End Sub
