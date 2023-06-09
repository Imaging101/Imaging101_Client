VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransactions 
   Caption         =   "Transactions - Teller Log "
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Cancel          =   -1  'True
      Caption         =   "&Cancel Transaction"
      Height          =   495
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTransactions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   7335
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19595265
      CurrentDate     =   29221
   End
   Begin VB.Frame frameFields 
      Height          =   3255
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   11655
      Begin VB.ComboBox cbTRANTYPE 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   915
         Width           =   3135
      End
      Begin MSMask.MaskEdBox mebAMOUNTCASHIN 
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   6240
         TabIndex        =   12
         Top             =   2850
         Width           =   2895
      End
      Begin VB.TextBox txtCTRNUMBER 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7440
         TabIndex        =   11
         Top             =   2325
         Width           =   1815
      End
      Begin VB.TextBox txtCTRREQUIRED 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6240
         TabIndex        =   10
         Top             =   2280
         Width           =   375
      End
      Begin VB.ComboBox cbTRANLOCATION 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   1230
         Width           =   1695
      End
      Begin VB.TextBox txtACCOUNTTITLE 
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Top             =   1050
         Width           =   4095
      End
      Begin VB.TextBox txtACCOUNTNUM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   765
         Width           =   2535
      End
      Begin VB.TextBox txtCIF 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtPERSONCONDUCTINGLASTNAME 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1965
         Width           =   2295
      End
      Begin VB.TextBox txtPERSONCONDUCTINGFIRSTNAME 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   2250
         Width           =   2295
      End
      Begin VB.TextBox txtTRANDATE 
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtTRUSERID 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtTRDATETIME 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtTRRECID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mebAMOUNTCASHOUT 
         Height          =   360
         Left            =   6240
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12632319
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   915
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "GLOBUS Reference #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   47
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "CTR #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   46
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "CTR Required?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label lblAMOUNTCASHOUT 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash OUT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblAMOUNTCASHIN 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash IN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Account TITLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   41
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Account #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   40
         Top             =   765
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "CIF #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "LAST Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "FIRST Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   2250
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Person Conducting Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Trans.UserID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   34
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Entry Date/Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tran. RecID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame frameButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -120
      TabIndex        =   17
      Top             =   -120
      Width           =   9615
      Begin VB.CommandButton cmdSecurity 
         Caption         =   "&Security"
         Height          =   735
         Left            =   7800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdLocations 
         Caption         =   "L&ocations"
         Height          =   735
         Left            =   6960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdTranTypes 
         Caption         =   "&TranTypes"
         Height          =   735
         Left            =   6120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":0CD6
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdConfigure 
         Caption         =   "&Configure"
         Height          =   735
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":1118
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdReports 
         Caption         =   "Re&Ports"
         Height          =   735
         Left            =   8760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":155A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.CommandButton cmdLogOff 
         Caption         =   "&Log-Off"
         Height          =   735
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":199C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   735
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":1DDE
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   735
         Left            =   1920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":2220
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   735
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":2662
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   735
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTransactions.frx":2AA4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.PictureBox picImaging101Logo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   9480
      Picture         =   "frmTransactions.frx":2EE6
      ScaleHeight     =   705
      ScaleWidth      =   2295
      TabIndex        =   15
      Top             =   0
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Display Transactions for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   9720
      TabIndex        =   16
      Top             =   720
      Width           =   1845
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    '****************************
    '*** Declarations
    Dim Con As ADODB.Connection
    Dim sSQL As String
    Dim cmd As ADODB.Command
    
    Dim Mode As String




Private Sub cbTRANTYPE_Change()

'    cbTRANTYPE_Click
    Debug.Print "TEST"
    
End Sub

Private Sub cbTRANTYPE_Click()
        ' This procedure Finds whether the transaction is CashIN or CashOUT
        '  and Enables/Disables the Input fields according to the pre-defined type

    Dim strWhere As String

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With

    rs.Source = "SELECT * FROM TRANTYPES WHERE TRANTYPE = '" & Me.cbTRANTYPE.Text & "' ORDER BY TRANTYPE"

    rs.Open

    If rs.RecordCount > 0 Then

        'Make sure input fields are not blank
        If mebAMOUNTCASHIN.Text = "" Then mebAMOUNTCASHIN.Text = "0"
        If mebAMOUNTCASHOUT.Text = "" Then mebAMOUNTCASHOUT.Text = "0"

        If Trim(rs.Fields("CASHINOUT")) = "IN" Then
            Me.mebAMOUNTCASHIN.Visible = True
            Me.lblAMOUNTCASHIN.Visible = True
            Me.mebAMOUNTCASHOUT.Visible = False
            Me.lblAMOUNTCASHOUT.Visible = False
            If CDec(Trim(Me.mebAMOUNTCASHIN.Text)) = 0 And CDec(Trim(Me.mebAMOUNTCASHOUT.Text)) <> 0 Then
                Me.mebAMOUNTCASHIN.Text = Me.mebAMOUNTCASHOUT.Text
                Me.mebAMOUNTCASHOUT.Text = ""
            End If

        ElseIf Trim(rs.Fields("CASHINOUT")) = "OUT" Then
            Me.mebAMOUNTCASHIN.Visible = False
            Me.lblAMOUNTCASHIN.Visible = False
            Me.mebAMOUNTCASHOUT.Visible = True
            Me.lblAMOUNTCASHOUT.Visible = True
            If CDec(Trim(Me.mebAMOUNTCASHOUT.Text)) = 0 And CDec(Trim(Me.mebAMOUNTCASHIN.Text)) <> 0 Then
                Me.mebAMOUNTCASHOUT.Text = Me.mebAMOUNTCASHIN.Text
                Me.mebAMOUNTCASHIN.Text = ""
            End If

        End If
        
        If IsNull(rs.Fields("CIF_REQUIRED")) Then
            gCIF_REQUIRED = False
        Else
            gCIF_REQUIRED = rs.Fields("CIF_REQUIRED")
        End If
        
        If IsNull(rs.Fields("ACCOUNT_REQUIRED")) Then
            gACCOUNT_REQUIRED = False
        Else
            gACCOUNT_REQUIRED = rs.Fields("ACCOUNT_REQUIRED")
        End If
        
    Else
        
        Dim Result As String
        
        Result = MsgBox("You have typed a Value that may NOT be a VALID option... " & _
                        vbCrLf & "ARE YOU SURE YOU WISH TO KEEP THIS VALUE?" & _
                        vbCrLf & "If you answer YES, then you MUST enter the Transaction Amount" & _
                        vbCrLf & "in the proper Cash IN or Cash OUT field.", vbYesNo, "Manually Entered Value?")
        
        If Result = vbYes Then
            Me.lblAMOUNTCASHIN.Visible = True
            Me.lblAMOUNTCASHOUT.Visible = True
            Me.mebAMOUNTCASHIN.Visible = True
            Me.mebAMOUNTCASHOUT.Visible = True
            gCIF_REQUIRED = False
            gACCOUNT_REQUIRED = False
            Me.mebAMOUNTCASHIN.SetFocus
        Else
            Me.lblAMOUNTCASHIN.Visible = False
            Me.lblAMOUNTCASHOUT.Visible = False
            Me.mebAMOUNTCASHIN.Visible = False
            Me.mebAMOUNTCASHOUT.Visible = False
            cbTRANTYPE.Text = ""
            cbTRANTYPE.SetFocus
        End If
        
    End If
        
    rs.Close
    Set rs = Nothing
        
End Sub

Private Sub cbTRANTYPE_Validate(Cancel As Boolean)

    '*** THIS TEST IS REQUIRED IN CASE THE USER CLICKS [CANCEL]
    '    Otherwise it will always pop the "You have typed a Value that may NOT be a VALID option..." message
    '    in cbTRANTYPE_Click
    If cbTRANTYPE.Text = "" Then
        Exit Sub
    End If
    
    cbTRANTYPE_Click
    
    'The following test is just to make sure the Focus does NOT go the field following mebAMOUNTCASH...
    If mebAMOUNTCASHIN.Visible = False And mebAMOUNTCASHOUT.Visible = False Then
        Cancel = True
    Else
        If mebAMOUNTCASHIN.Visible = True Then
            mebAMOUNTCASHIN.SetFocus
        End If
    End If
    
    
End Sub

Private Sub cmdAdd_Click()

    'Add New Transaction Record
    subClearForm
    Mode = "Add"
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    StatusBar1.Panels(1).Text = "Adding New Transaction - Please enter Data and Click the [Update] button."
    Me.txtTRDATETIME.Text = CStr(Now())
    
'    txtTRRECID.Text = GetNextControlNumber
    txtTRRECID.Text = "Adding"
    bolCanceling = False
    
    Me.txtTRUSERID.Text = gsecUserID
    Me.txtTRANDATE.Text = FormatDateTime(Now(), vbShortDate)
    Me.cbTRANLOCATION.Text = gsecUSERLOCATION
    Me.cbTRANTYPE.SetFocus
    cmdCancel.Visible = True
    cmdUpdate.Enabled = True
    
End Sub

Private Sub cmdCancel_Click()

    bolCanceling = True

    subClearForm
    cmdCancel.Visible = False
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
    
End Sub

Private Sub cmdConfigure_Click()

    frmTRANCONFIG.Show
    Me.Hide

End Sub

Private Sub cmdDelete_Click()
    
    Dim Result As String
    Result = MsgBox("Are you sure you wish to DELETE Transaction #: " & txtTRRECID & " for Account: " & txtACCOUNTTITLE & " ?", vbYesNo, "Delete Transaction")
    If Result = vbNo Then
        StatusBar1.Panels(1).Text = "Delete CANCELLED!"
        Exit Sub
    End If

    subDeleteTransaction
    StatusBar1.Panels(1).Text = "Transaction Deleted."
    subPopulateListView
    
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    
    
End Sub

Private Sub cmdLocations_Click()

    frmTRANLOCATIONS.Show
    Me.Hide
    

End Sub

Private Sub cmdLogOff_Click()
    ' Log Off the Current User
    StatusBar1.Panels(1).Text = "Logging Off Current User."
    frmLogin.Show
    Unload Me
    
End Sub

Private Sub cmdRefresh_Click()

    'Refresh the Transaction List & Combos
    StatusBar1.Panels(1).Text = "Refreshing..."
    subPopulateTRANLOCATIONSCombo
    subPopulateTRANTYPECombo
    subPopulateListView
    StatusBar1.Panels(1).Text = "Refresh Complete"


End Sub

Private Sub Command6_Click()

End Sub

Private Sub cmdReports_Click()

    Dim strWhere As String

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With
        
    Dim dtSelectedDate As Date
    dtSelectedDate = CDate(DTPicker1.Value)
        
    rs.Source = "SELECT " & _
                 "       TRRECID, " & _
                 "       TRDATETIME, " & _
                 "       TRUSERID, " & _
                 "       TRANDATE, " & _
                 "       PERSONCONDUCTINGLASTNAME, " & _
                 "       PERSONCONDUCTINGFIRSTNAME, " & _
                 "       CIF, " & _
                 "       ACCOUNTNUM, " & _
                 "       ACCOUNTTITLE, " & _
                 "       TRANTYPE, " & _
                 "       AMOUNTCASHIN, " & _
                 "       AMOUNTCASHOUT, " & _
                 "       TRANLOCATION, " & _
                 "       CTRNUMBER, " & _
                 "       CTRREQUIRED " & _
                 "FROM TRANREG "


            strWhere = "WHERE TRDATETIME >= '" & FormatDateTime(dtSelectedDate, vbShortDate) & "'" & _
                        "  AND TRDATETIME < '" & dtSelectedDate + 1 & "'"

            
            rs.Source = rs.Source & strWhere & " ORDER BY " & " TRRECID "
            
    Con.Errors.Clear
    rs.Open
  
    '********* REPORT CODE ***************
    
    Dim m_GrandTotal As Currency
    Dim xx As Integer
    
    
    If Not rs.EOF And Not rs.BOF Then
    
        
        Set rptDailyCashTransactionLog.DataSource = rs 'Bind the report to recordset
        
        With rptDailyCashTransactionLog.Sections("Section1") 'loop through the controls in the Detail section
            'There's no count 0 it always start with 1
            'by looping through this section we can access all the controls that we dragged in to this section
            For xx = 1 To .Controls.Count
                If TypeOf .Controls(xx) Is RptTextBox Then 'is it TextBox Control
                    .Controls(xx).DataField = rs.Fields(xx - 1).Name 'bind it
                    .Controls(xx).Width = rs.Fields(xx - 1).DefinedSize * 50
                    
                End If
            Next
        End With
        
        rptDailyCashTransactionLog.Orientation = rptOrientLandscape 'set the Orientation of report
        rptDailyCashTransactionLog.ReportWidth = 12000 'set the size of report
        rptDailyCashTransactionLog.Refresh 'refresh the data report to update the changes
        rptDailyCashTransactionLog.WindowState = vbMaximized
        rptDailyCashTransactionLog.Show 'show the report
    
    End If
    
    rs.Close
    Set rs = Nothing
    Con.Close
    Set Con = Nothing
    Set rptDailyCashTransactionLog = Nothing
    
End Sub

Private Sub Command7_Click()

End Sub

Private Sub cmdSecurity_Click()
    
    frmTRANSECURITY.Show
    Me.Hide
    

End Sub

Private Sub cmdTranTypes_Click()

    frmTRANTYPES2.Show
    Me.Hide
    
End Sub

Private Sub cmdUpdate_Click()

    If Not IsValidForm Then
        Exit Sub
    End If
        
        'Update the Selected Transaction Record
        ' Check mode to determine action
        If Mode = "Add" Then
            subAddTransaction
            StatusBar1.Panels(1).Text = "Transaction Added!"
            Mode = "Update"
        Else
            subUpdateTransaction
            StatusBar1.Panels(1).Text = "Transaction Updated!"
        End If

        'Check for Transactions Over/Under Total
        'Check if over the pre-defined limit to request the CTR
        'if the transaction is not already flagged.
        If txtCTRREQUIRED.Text <> "Y" Then
            subCheckForOverLimit
        End If

        'Only clear the form if CTR is not required
        '  or if a CTR number has been entered
        If txtCTRREQUIRED.Text = "Y" _
        And Trim(txtCTRNUMBER.Text) = "" Then
            StatusBar1.Panels(1).Text = "Transaction Updated... WAITING for CTR Number!  Click [Update] button when done."
            'Disable Add and Delete buttons
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = True
        Else
            subClearForm
            cmdAdd.Enabled = True 'Add
        End If

        subPopulateListView
        
        cmdCancel.Visible = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        
        
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DTPicker1_Change()

    subPopulateListView



End Sub

Private Sub Form_Activate()
        
        Select Case UCase(Trim(gsecUSERLEVEL))
            Case "SYSADMIN"
                cmdConfigure.Visible = True
                cmdSecurity.Visible = True
                cmdTranTypes.Visible = True
                cmdLocations.Visible = True
                
            Case "SUPERVISOR"
                cmdConfigure.Visible = False
                cmdSecurity.Visible = True
                cmdTranTypes.Visible = True
                cmdLocations.Visible = True

            Case Else
                cmdConfigure.Visible = False
                cmdSecurity.Visible = False
                cmdTranTypes.Visible = False
                cmdLocations.Visible = False
        End Select
        
        DTPicker1.Value = Now()
        
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        
        subPopulateListView
        
End Sub

Private Sub Form_Load()

    subPopulateTRANLOCATIONSCombo
    subPopulateTRANTYPECombo
    subPopulateListView

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  frameFields.Top = Me.ScaleHeight - frameFields.Height - StatusBar1.Height
  frameFields.Width = Me.ScaleWidth - frameFields.Left - 50
  ListView1.Height = Me.ScaleHeight - ListView1.Top - frameFields.Height - StatusBar1.Height - 50
  ListView1.Width = Me.ScaleWidth - ListView1.Left - 50
  StatusBar1.Panels(1).Width = StatusBar1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmLogin.Show

End Sub


Public Sub subPopulateListView()

    
    ListView1.ListItems.Clear
        
     '*** Declarations -- MOVED TO MODULE LEVEL TOP
''    Dim rs As adodb.Recordset
''    Dim Con As adodb.Connection
''    Dim ssql As String
    Dim strWhere As String

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With
        
    Dim dtSelectedDate As Date
    dtSelectedDate = CDate(DTPicker1.Value)
        
    rs.Source = "SELECT " & _
                 "       TRRECID, " & _
                 "       TRDATETIME, " & _
                 "       TRUSERID, " & _
                 "       TRANDATE, " & _
                 "       PERSONCONDUCTINGLASTNAME, " & _
                 "       PERSONCONDUCTINGFIRSTNAME, " & _
                 "       CIF, " & _
                 "       ACCOUNTNUM, " & _
                 "       ACCOUNTTITLE, " & _
                 "       TRANTYPE, " & _
                 "       AMOUNTCASHIN, " & _
                 "       AMOUNTCASHOUT, " & _
                 "       TRANLOCATION, " & _
                 "       CTRNUMBER, " & _
                 "       CTRREQUIRED " & _
                 "FROM TRANREG "


            strWhere = "WHERE TRDATETIME >= '" & FormatDateTime(dtSelectedDate, vbShortDate) & "'" & _
                        "  AND TRDATETIME < '" & dtSelectedDate + 1 & "'"
        ''                     "WHERE TRDATETIME = " & PrepareStr(Format(DateTimePicker1.Text, "Short Date"))

            
            rs.Source = rs.Source & strWhere & " ORDER BY " & " TRRECID "
            
 
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenForwardOnly
'    rs.LockType = adLockOptimistic
    
''    On Error GoTo ERROR_TRAP
    
    'sql statement to select items on the drop down list
''    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID
''    rs.Open ssql, Con
    Con.Errors.Clear
    rs.Open
        
    
   
    '*** Setup Up ListView properties - BEGIN
    
    ListView1.Visible = False
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Clear
    '    Column widths are in PIXELS!
    ListView1.MultiSelect = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.AllowColumnReorder = True
    
        '***  SET COLUMN HEADINGS
        For intListIndex = 0 To rs.Fields.Count - 1
            ListView1.ColumnHeaders.Add , , rs.Fields.item(intListIndex).Name, Len(rs.Fields.item(intListIndex).Name) * 150, lvwColumnLeft
        Next
                
    On Error Resume Next
    
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    While Not rs.EOF
            For intListIndex = 0 To rs.Fields.Count - 1
                If intListIndex = 0 Then
                    If Not IsNull(rs.Fields.item(intListIndex).Value) Then
                        Set lstItem = ListView1.ListItems.Add(, , rs.Fields.item(intListIndex).Value)
                    End If
                Else
            
                        '* This null check is to make sure we don't Skip fields caused by an error.
                        If Not IsNull(rs.Fields.item(intListIndex).Value) Then
                            ' Not null... show value
                            Select Case rs.Fields.item(intListIndex).Type
                                Case adDBTimeStamp
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Format(rs.Fields.item(intListIndex).Value, "yyyy/mm/dd AMPM hh:mm:ss "))
'                                Case adInteger
'                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("      " & Format(rs.Fields.item(intListIndex).Value, "##,###"), 6))
'                                Case adNumeric, adDouble
'                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("          " & Format(rs.Fields.item(intListIndex).Value, "##,###,###"), 10))
                                Case adCurrency
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , Right("                " & Format(rs.Fields.item(intListIndex).Value, "$##,###,##0.00"), 14))
                                Case Else
                                    Set lstSubItem = lstItem.ListSubItems.Add(, , rs.Fields.item(intListIndex).Value)
                            End Select
                            
'                            If rs.Fields.item(intListIndex).Type = adDBTimeStamp Then
'                                Set lstSubItem = lstItem.ListSubItems.Add(, , Format(rs.Fields.item(intListIndex).Value, "yyyy/mm/dd AMPM hh:mm:ss "))
'                            Else
'                                Set lstSubItem = lstItem.ListSubItems.Add(, , rs.Fields.item(intListIndex).Value)
'                            End If
                        Else
                            ' Null... show empty string
                            Set lstSubItem = lstItem.ListSubItems.Add(, , "")
                        End If
                        
                        If rs.Fields.item(intListIndex).Name = "BatchNotes" Then
                            lstItem.ListSubItems(intListIndex).ForeColor = vbRed
                       End If
                End If
            Next
        rs.MoveNext
    Wend
    On Error GoTo 0
    
    ' AutoSize ALL Columns
    Dim i As Integer, lParam As Long
    UseHeader = True
    If UseHeader = False Then
        lParam = LVSCW_AUTOSIZE
    Else
        lParam = LVSCW_AUTOSIZE_USEHEADER
    End If
    For i = 0 To ListView1.ColumnHeaders.Count - 1
        SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lParam
    Next
    
    ' Hide the RecordID's
'    ListView1.ColumnHeaders(1).Width = 0
'    ListView1.ColumnHeaders(2).Width = 0
    
'    ' Size the Key fields to a standard size
'    ListView1.ColumnHeaders(3).Width = 3000
    
    ListView1.Visible = True

    '*** Setup Up ListView properties - END
    
    ' Disable Buttons until at least ONE ROW is selected

    rs.Close
    Set rs = Nothing
    

End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
    ' Set the Sort Column
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Sort It!
    ListView1.Sorted = True
    
End Sub

Private Sub AutoSizeColumns(Listview As Listview, Optional ByVal UseHeader As Boolean = False)
  Dim i As Integer, lParam As Long
  If UseHeader = False Then
      lParam = LVSCW_AUTOSIZE
  Else
      lParam = LVSCW_AUTOSIZE_USEHEADER
  End If
  For i = 0 To Listview.ColumnHeaders.Count - 1
      SendMessage Listview.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal lParam
  Next
End Sub


Private Sub subClearForm()

        ' Clear the data entry form.
        txtTRRECID.Text() = ""
        txtTRDATETIME.Text() = ""
        txtTRUSERID.Text() = ""
        txtTRANDATE.Text() = ""
        txtPERSONCONDUCTINGFIRSTNAME.Text() = ""
        txtPERSONCONDUCTINGLASTNAME.Text() = ""
        txtCIF.Text() = ""
        txtACCOUNTNUM.Text() = ""
        txtACCOUNTTITLE.Text() = ""
        cbTRANTYPE.Text() = ""
        mebAMOUNTCASHIN.Text() = ""
        mebAMOUNTCASHOUT.Text() = ""
        cbTRANLOCATION.Text() = ""
        txtCTRREQUIRED.Text() = ""
        txtCTRNUMBER.Text() = ""

End Sub

Private Sub subDeleteTransaction()
    ' This sub is used to delete the product record from the database
    ' when the user clicks the delete button

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With


        ' Build Select statement to query product information from the products
        ' table
        rs.Source = "DELETE FROM TRANREG " & _
                 "WHERE TRRECID = " & CLng(txtTRRECID.Text)

        Err.Clear
        
        rs.Open
        
        If Err.Number > 0 Then
            MsgBox "Delete Failed!  TRRECID = " & txtTRRECID.Text & _
                   " Error #: " & Err.Number & "  Error Description: " & Err.Description, vbCritical, "Delete"
        End If

        ' Close and Clean up objects
        'The Delete call automatically closes the RecordSet (rs)
        Set rs = Nothing
        
End Sub

Private Sub subPopulateForm()
    
    
    
    Dim lstIndex As Long
    
    lstIndex = Me.ListView1.SelectedItem.Index
    ' Get Main Item
    txtTRRECID.Text = Me.ListView1.ListItems(lstIndex).Text

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With


        ' Build Select statement to query product information from the products
        ' table
        rs.Source = "SELECT * " & _
                 "FROM TRANREG " & _
                 "WHERE TRRECID = " & txtTRRECID.Text


    rs.Open
        
    If rs.RecordCount > 0 Then
    
        ' Populate form with the data
        txtTRRECID = rs.Fields("TRRECID")
        txtTRDATETIME = rs.Fields("TRDATETIME") & ""
        txtTRUSERID = rs.Fields("TRUSERID") & ""
        txtTRANDATE = FormatDateTime(rs.Fields("TRANDATE"), vbShortDate) & ""
        txtPERSONCONDUCTINGFIRSTNAME.Text = rs.Fields("PERSONCONDUCTINGFIRSTNAME") & ""
        txtPERSONCONDUCTINGLASTNAME.Text = rs.Fields("PERSONCONDUCTINGLASTNAME") & ""
        txtCIF = rs.Fields("CIF") & ""
        txtACCOUNTNUM = rs.Fields("ACCOUNTNUM") & ""
        txtACCOUNTTITLE = rs.Fields("ACCOUNTTITLE") & ""
        cbTRANTYPE = rs.Fields("TRANTYPE") & ""
        mebAMOUNTCASHIN = rs.Fields("AMOUNTCASHIN")
        mebAMOUNTCASHOUT = rs.Fields("AMOUNTCASHOUT")
        cbTRANLOCATION = rs.Fields("TRANLOCATION") & ""
        txtCTRNUMBER = rs.Fields("CTRNUMBER") & ""
        txtCTRREQUIRED = rs.Fields("CTRREQUIRED") & ""
    End If

    rs.Close
    Set rs = Nothing
    
    cbTRANTYPE_Click
    
    'Update & Delete Button Security
    Select Case Trim(gsecUSERLEVEL)
    
        Case "SYSADMIN"
            cmdUpdate.Enabled = True
            cmdDelete.Enabled = True
            Mode = "Update"
            
        Case "SUPERVISOR"
            cmdUpdate.Enabled = True
            cmdDelete.Enabled = True
            Mode = "Update"
            
        Case "TELLER"
            'Check if record was created by currently logged-in user
            If UCase(Trim(gsecUserID)) = UCase(Trim(txtTRUSERID)) Then
                
                'Are they allowed to modify their own records?
                If gsecMODIFYOWN Then
                    'Can only change transactions on the Date it was entered
                    If FormatDateTime(Now(), vbShortDate) = FormatDateTime(txtTRDATETIME.Text, vbShortDate) Then
                        cmdUpdate.Enabled = True
                        Mode = "Update"
                    Else
                        cmdUpdate.Enabled = False
                        Mode = "View"
                        
                    End If
                Else
                    'Not Allowed to modify own transactions.
                    cmdUpdate.Enabled = False
                    Mode = "View"
                End If
                
                'Are they allowed to Delete their own records?
                If gsecDELETEOWN Then
                    'Can only Delete transactions on the Date it was entered
                    If FormatDateTime(Now(), vbShortDate) = FormatDateTime(txtTRDATETIME.Text, vbShortDate) Then
                        cmdDelete.Enabled = True
                        Mode = "Update"
                    Else
                        cmdDelete.Enabled = False
                        Mode = "View"
                    End If
                Else
                    'Not Allowed to Delete own transactions.
                    cmdDelete.Enabled = False
                    Mode = "View"
                End If
                
                
            Else
                'Record was created by another user...
                'Is this user allowed to modify other users' transactions?
                If gsecMODIFYOTHER Then
                    'Can only change transactions on the Date it was entered
                    If FormatDateTime(Now(), vbShortDate) = FormatDateTime(txtTRDATETIME.Text, vbShortDate) Then
                        cmdUpdate.Enabled = True
                        Mode = "Update"
                    Else
                        cmdUpdate.Enabled = False
                        Mode = "View"
                    End If
                Else
                    'Not Allowed to modify other users' transactions.
                    cmdUpdate.Enabled = False
                    Mode = "View"
                End If
                    
                If gsecDELETEOTHER Then
                    'Can only Delete transactions for the Date it was entered
                    If FormatDateTime(Now(), vbShortDate) = FormatDateTime(txtTRDATETIME.Text, vbShortDate) Then
                        cmdDelete.Enabled = True
                        Mode = "Update"
                    Else
                        cmdDelete.Enabled = False
                        Mode = "View"
                    End If
                Else
                    'Not Allowed to Delete other users' transactions.
                    cmdDelete.Enabled = False
                    Mode = "View"
                End If
            End If
            
        Case Else
            'If UserLevel not specified... disable Update and Delete
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            Mode = "View"
    End Select
    

    Select Case Mode
        
        Case "View"
            StatusBar1.Panels(1).Text = "View Only -- Transaction changes cannot be saved."
        Case "Update"
            StatusBar1.Panels(1).Text = "To Modify this transaction - make the required changes and Click the [Update] button."
        Case Else
            StatusBar1.Panels(1).Text = ""
    End Select
    
End Sub



Private Sub subCheckForOverLimit()

    'Set Up Connection and RecordSet
    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With


    'Set up Variables
     Dim strSQL As String
     Dim strSQLSelect As String
     Dim strSQLFrom As String
     Dim strSQLWhere As String

     Dim strID As String

     Dim CashIn As Currency
     Dim CashOut As Currency
     Dim CashTotal As Currency

     ' Set Vars with fields from Config DB
     Dim intDAYSTOSCANFORDEPOSITS As Integer
     Dim decAMOUNTTFORCTR As Currency


     '********************  LOAD CONFIGURATION VALUES   *************************
                 
     strSQL = "SELECT * FROM TRANCONFIG WHERE TCRECID=1"

    rs.Source = strSQL
    rs.Open
        
    If rs.RecordCount > 0 Then

        'Make sure days is not blank
        If CInt(rs.Fields("DAYSTOSCANFORDEPOSITS")) <> 0 Then
            intDAYSTOSCANFORDEPOSITS = CInt(rs.Fields("DAYSTOSCANFORDEPOSITS"))
        Else
            intDAYSTOSCANFORDEPOSITS = 0
        End If
        'Make sure amount is not blank
        If CDec(rs.Fields("AMOUNTTFORCTR")) <> 0 Then
            decAMOUNTTFORCTR = CDec(rs.Fields("AMOUNTTFORCTR"))
        Else
            decAMOUNTTFORCTR = 0
        End If

    Else
        MsgBox "Error Getting Configuration Values.  Please call technical support!", vbCritical
        Exit Sub
    End If


    rs.Close
    Set rs = Nothing
    

    'Make sure input fields are not blank
    If mebAMOUNTCASHIN.Text = "" Then mebAMOUNTCASHIN.Text = "0"
    If mebAMOUNTCASHOUT.Text = "" Then mebAMOUNTCASHOUT.Text = "0"


    '*********************************************************************
    '********************  SINGLE TRANSACTIONS   *************************
    
    'Single cash transaction of more than the Amount Configured
    If (CDec(mebAMOUNTCASHIN.Text) >= decAMOUNTTFORCTR) _
    Or (CDec(mebAMOUNTCASHOUT.Text) >= decAMOUNTTFORCTR) Then
        MsgBox ("Please fill out a CTR! " & vbCrLf & vbCrLf & _
                "The amount of THIS Transaction exceeds " & vbCrLf & _
                "        " & Format(decAMOUNTTFORCTR, "Currency") & "")
        txtCTRREQUIRED.Text = "Y"
    End If



    '*****************************************************************************
    '********************  SET-UP FOR SQL TRANSACTIONS   *************************
    

    ' Build Select statement to query transaction over/under set limit
    strSQLSelect = "SELECT " & _
                    "       SUM(AMOUNTCASHIN) as SUMCASHIN, " & _
                    "       SUM(AMOUNTCASHOUT) AS SUMCASHOUT "

    strSQLFrom = "FROM TRANREG "


    '***********************************************************************
    '********************  MULTIPLE TRANSACTIONS   *************************

    '------------------------------------------------------------------------------------
    '1. Multiple cash transactions aggregating more than $10,000
    '   by the same individual,
    '   regardless of on whose behalf they were conducted
    '   a.  (e.g. John withdraws $6,000 and then withdraws $5,000)

    If CInt(rs.Fields("TEST_FOR_PERSONCONDUCTING")) <> 0 Then

        strSQLWhere = "WHERE " & _
                        "      TRANDATE >= '" & FormatDateTime(DateAdd("d", -(intDAYSTOSCANFORDEPOSITS), Now), vbShortDate) & "'" & _
                        "  AND TRANDATE <= '" & FormatDateTime(Now, vbShortDate) & "'" & _
                        "  AND PERSONCONDUCTINGFIRSTNAME = '" & txtPERSONCONDUCTINGFIRSTNAME.Text & "'" & _
                        "  AND PERSONCONDUCTINGLASTNAME = '" & txtPERSONCONDUCTINGLASTNAME.Text & "'"
    
    
        strSQL = strSQLSelect & strSQLFrom & strSQLWhere
    
        Set rs = New ADODB.Recordset
        
        With rs
            .ActiveConnection = Con
            .CursorLocation = adUseServer
            .CursorType = adOpenStatic
            .LockType = adLockReadOnly
        End With
    
        rs.Source = strSQL
        rs.Open
        
        If rs.RecordCount > 0 Then
    
            If rs.Fields("SUMCASHIN") <> "" Then CashIn = CDec(rs.Fields("SUMCASHIN"))
            If rs.Fields("SUMCASHOUT") <> "" Then CashOut = CDec(rs.Fields("SUMCASHOUT"))
    
            If (CashIn >= decAMOUNTTFORCTR) _
            Or (CashOut >= decAMOUNTTFORCTR) Then
                MsgBox "Please fill out a CTR! " & vbCrLf & vbCrLf & _
                        "'" & txtPERSONCONDUCTINGFIRSTNAME.Text & " " & txtPERSONCONDUCTINGLASTNAME & "'" & vbCrLf & _
                        " has made transactions Totaling more than " & vbCrLf & _
                        "        " & Format(decAMOUNTTFORCTR, "Currency") & vbCrLf & _
                        "   Cash IN =  " & Format(CashIn, "Currency") & vbCrLf & _
                        "   Cash OUT=  " & Format(CashOut, "Currency") & vbCrLf
                txtCTRREQUIRED.Text = "Y"
            End If
        Else
            MsgBox "I was unable to find any transactions.", vbInformation
        End If
    
        ' Close and Clean up objects
        rs.Close
        Set rs = Nothing
        
    End If
    
    '------------------------------------------------------------------------------------
    '2. Multiple cash transactions aggregating more than $10,000
    '   on behalf of the same party,
    '   regardless of who conducted the transactions
    '   a.  (e.g. John deposits $7,000 on behalf of Bill
    '       and then Jane deposits $8,000 on behalf of Bill) Analyzed by CIF

    If CInt(rs.Fields("TEST_FOR_ACCOUNT_TITLE")) <> 0 Then

        strSQLWhere = "WHERE " & _
                        "      TRANDATE >= '" & FormatDateTime(DateAdd("d", -(intDAYSTOSCANFORDEPOSITS), Now), vbShortDate) & "'" & _
                        "  AND TRANDATE <= '" & FormatDateTime(Now, vbShortDate) & "'" & _
                        "  AND ACCOUNTTITLE = '" & txtACCOUNTTITLE.Text & "'"
    
        strSQL = strSQLSelect & strSQLFrom & strSQLWhere
    
        Set rs = New ADODB.Recordset
        
        With rs
            .ActiveConnection = Con
            .CursorLocation = adUseServer
            .CursorType = adOpenStatic
            .LockType = adLockReadOnly
        End With
    
        rs.Source = strSQL
        rs.Open
    
        If rs.RecordCount > 0 Then
        
            If rs.Fields("SUMCASHIN") <> "" Then CashIn = CDec(rs.Fields("SUMCASHIN"))
            If rs.Fields("SUMCASHOUT") <> "" Then CashOut = CDec(rs.Fields("SUMCASHOUT"))
    
            If (CashIn >= decAMOUNTTFORCTR) _
            Or (CashOut >= decAMOUNTTFORCTR) Then
                MsgBox "Please fill out a CTR! " & vbCrLf & vbCrLf & _
                        " Transactions totaling more than " & vbCrLf & _
                        "        " & Format(decAMOUNTTFORCTR, "Currency") & vbCrLf & _
                        " have been conducted on behalf of  " & vbCrLf & _
                        "'" & txtACCOUNTTITLE.Text & "'" & vbCrLf & _
                        "   Cash IN =  " & Format(CashIn, "Currency") & vbCrLf & _
                        "   Cash OUT=  " & Format(CashOut, "Currency") & vbCrLf
                txtCTRREQUIRED.Text = "Y"
            End If
        Else
            MsgBox "I was unable to find any transactions.", vbInformation
        End If
    
        ' Close and Clean up objects
        rs.Close
        Set rs = Nothing
        
    End If
    
    
    '------------------------------------------------------------------------------------
    '3. Multiple cash deposits aggregating more than $10,000
    '   into the same CIF,
    '   regardless of who conducted the transactions
    '   a.  (e.g. John deposits $9,000 into Account X
    '       and then Jane deposits $4,000 into Account X).

    If CInt(rs.Fields("TEST_FOR_CIF")) <> 0 Then

        strSQLWhere = "WHERE " & _
                        "      TRANDATE >= '" & FormatDateTime(DateAdd("d", -(intDAYSTOSCANFORDEPOSITS), Now), vbShortDate) & "'" & _
                        "  AND TRANDATE <= '" & FormatDateTime(Now, vbShortDate) & "'" & _
                        "  AND CIF = '" & txtCIF.Text & "'"
    
        strSQL = strSQLSelect & strSQLFrom & strSQLWhere
        
        Set rs = New ADODB.Recordset
        
        With rs
            .ActiveConnection = Con
            .CursorLocation = adUseServer
            .CursorType = adOpenStatic
            .LockType = adLockReadOnly
        End With
    
        rs.Source = strSQL
        rs.Open
        
        If rs.RecordCount > 0 Then
        
                 If rs.Fields("SUMCASHIN") <> "" Then CashIn = CDec(rs.Fields("SUMCASHIN"))
                 If rs.Fields("SUMCASHOUT") <> "" Then CashOut = CDec(rs.Fields("SUMCASHOUT"))
    
                 If (CashIn >= decAMOUNTTFORCTR) _
                 Or (CashOut >= decAMOUNTTFORCTR) Then
                     MsgBox ("Please fill out a CTR! " & vbCrLf & vbCrLf & _
                             " Transactions totaling more than " & vbCrLf & _
                             "    " & Format(decAMOUNTTFORCTR, "Currency") & vbCrLf & _
                             " have been conducted for " & vbCrLf & _
                             "   CIF #  '" & txtCIF.Text & "'" & vbCrLf & vbCrLf & _
                             "   Cash IN =  " & CashIn & vbCrLf & _
                             "   Cash OUT=  " & CashOut & vbCrLf)
                     txtCTRREQUIRED.Text = "Y"
                 End If
             Else
                 MsgBox "I was unable to find any transactions.", vbInformation
             End If
    
             ' Close and Clean up objects
            rs.Close
            Set rs = Nothing
            
         If txtCTRREQUIRED.Text = "Y" Then
             subUpdateTransaction
             txtCTRNUMBER.SetFocus
         End If
         
    End If
    
 End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    
    If Mode = "Add" Then
        Result = MsgBox("You were Adding a record... this will clear what you have typed!  Are you sure you wish to display the selected item?", vbYesNo, "Select Transaction")
        If Result = vbNo Then
            Exit Sub
        End If
    End If
    
    subPopulateForm
    

End Sub

Private Sub subPopulateTRANLOCATIONSCombo()
        ' This procedure populates the list box on the
        ' form with a list of Categories from the
        ' database.

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With


    ' Build Select statement to query product information from the products
    ' table
    rs.Source = "SELECT TRANLOCATION FROM TRANLOCATIONS"
    rs.Open
    
    cbTRANLOCATION.Clear

    If rs.RecordCount > 0 Then

        ' Loop through the result set and add the category
        ' names to the combo box.
        For intLoop = 1 To rs.RecordCount
    
            cbTRANLOCATION.AddItem rs.Fields("TRANLOCATION")
            rs.MoveNext
            
        Next
    
    End If
    
    ' Close and Clean up objects
    rs.Close
    Set rs = Nothing
        
End Sub

Private Sub subPopulateTRANTYPECombo()
        ' This procedure populates the list box on the
        ' form with a list of Categories from the
        ' database.

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
    End With


    ' Build Select statement to query product information from the products
    ' table
    rs.Source = "SELECT TRANTYPE,CASHINOUT FROM TRANTYPES ORDER BY TRANTYPE"
    rs.Open
    
    cbTRANTYPE.Clear

    If rs.RecordCount > 0 Then

        ' Loop through the result set and add the category
        ' names to the combo box.
        For intLoop = 1 To rs.RecordCount
    
            cbTRANTYPE.AddItem rs.Fields("TRANTYPE")
            rs.MoveNext

        Next
    End If
    
    ' Close and Clean up objects
    rs.Close
    Set rs = Nothing
        
End Sub


Private Sub mebAMOUNTCASHIN_GotFocus()

    mebAMOUNTCASHIN.SelStart = 0
    mebAMOUNTCASHIN.SelLength = (Len(mebAMOUNTCASHIN.Text))
    

End Sub

Private Sub mebAMOUNTCASHOUT_Change()

    mebAMOUNTCASHOUT.SelStart = 0
    mebAMOUNTCASHOUT.SelLength = (Len(mebAMOUNTCASHOUT.Text))
    
End Sub

Private Sub txtCTRREQUIRED_GotFocus()
        'Make sure the user CANNOT edit the CTRREQUIRED Flag
        'by setting the focus back to the CTRNUMBER field.

        txtCTRNUMBER.SetFocus

End Sub


Private Sub subAddTransaction()
        ' This sub is used to add a product record to the database
        ' when the user clicks the Update button and the mode is set
        ' to ADD

        ' Validate form values.
        If Not IsValidForm() Then
            Exit Sub
        End If

    '*** GET THE NEXT AVAILABLE TRANSACTION RECID
    txtTRRECID.Text = GetNextControlNumber
    
    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With

        Dim TRRECID As String
        TRRECID = txtTRRECID.Text

        If mebAMOUNTCASHIN.Text = "" Then mebAMOUNTCASHIN.Text = "0"
        If mebAMOUNTCASHOUT.Text = "" Then mebAMOUNTCASHOUT.Text = "0"

            ' Build Insert statement to insert new product into the products table
            rs.Source = "INSERT TRANREG VALUES (" & _
                     txtTRRECID.Text & "," & _
                     PrepareStr(txtTRDATETIME.Text) & "," & _
                     PrepareStr(txtTRUSERID.Text) & "," & _
                     PrepareStr(txtTRANDATE.Text) & "," & _
                     PrepareStr(txtPERSONCONDUCTINGFIRSTNAME.Text) & "," & _
                     PrepareStr(txtPERSONCONDUCTINGLASTNAME.Text) & "," & _
                     txtCIF.Text & "," & _
                     txtACCOUNTNUM.Text & "," & _
                     PrepareStr(txtACCOUNTTITLE.Text) & "," & _
                     PrepareStr(cbTRANTYPE.Text) & "," & _
                     mebAMOUNTCASHIN.Text & "," & _
                     mebAMOUNTCASHOUT.Text & "," & _
                     "''" & "," & _
                     "''" & "," & _
                     PrepareStr(cbTRANLOCATION.Text) & "," & _
                     PrepareStr(txtCTRNUMBER.Text) & "," & _
                     "''" & "," & _
                     "''" & "," & _
                     "''" & _
                     ")"


            rs.Open
            
            ' Close and Clean up objects
'            rs.Close
            Set rs = Nothing
            
            ' Refresh Product List
            subPopulateListView

End Sub
    
    
Private Function PrepareStr(ByVal strValue As String) As String
    ' This function accepts a string and creates a string that can
    ' be used in a SQL statement by adding single quotes around
    ' it and handling empty values.
    If Trim(strValue) = "" Then
        PrepareStr = "NULL"
    Else
        PrepareStr = "'" & Replace(Trim(strValue), "'", "''") & "'"
        'Return Chr(34) & strValue.Trim() & Chr(34)
    End If
End Function


Private Function GetNextControlNumber() As Long
    ' This sub is used to update and existing record with values
    ' from the form.
    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
    End With


    'Build Select statement to get the First (should be the ONLY) record in the control table
    rs.Source = "SELECT * " & _
                 "FROM TRANCONFIG " & _
                 "WHERE  TCRECID = 1"

    rs.Open
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        'LOCK the record by setting the TCRECID equal to itself
        rs.Fields("TCRECID") = rs.Fields("TCRECID")
        
        'Return the Current RECID
        GetNextControlNumber = rs.Fields("TRNEXTRECID")
        
        'Increment the Next RECID
        rs.Fields("TRNEXTRECID") = rs.Fields("TRNEXTRECID") + 1

        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    

End Function


Private Function IsValidForm() As Boolean

    ' Check to make sure each field has a valid value
    If Trim(txtTRANDATE.Text) = "" Then
        MsgBox "Please enter a valid Transaction DATE.", vbExclamation, "Form Incomplete"
        txtTRANDATE.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If Trim(cbTRANTYPE.Text) = "" Then
        MsgBox "Please enter a valid TRANSACTION TYPE.", vbExclamation, "Form Incomplete"
        cbTRANTYPE.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If gCIF_REQUIRED And (Trim(txtCIF.Text) = "" Or (Not IsNumeric(txtCIF.Text))) Then
        MsgBox "Please enter a valid CIF #", vbExclamation, "Form Incomplete"
        txtCIF.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If gACCOUNT_REQUIRED And (Trim(txtACCOUNTNUM.Text) = "" Or (Not IsNumeric(txtACCOUNTNUM.Text))) Then
        MsgBox "Please enter a valid ACCOUNT NUMBER.", vbExclamation, "Form Incomplete"
        txtACCOUNTNUM.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If Trim(txtACCOUNTTITLE.Text) = "" Then
        MsgBox "Please enter a valid ACCOUNT TITLE.", vbExclamation, "Form Incomplete"
        txtACCOUNTTITLE.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If Trim(txtPERSONCONDUCTINGLASTNAME.Text) = "" Then
        MsgBox "Please enter a valid PERSON CONDUCTING Transation LAST Name.", vbExclamation, "Form Incomplete"
        txtPERSONCONDUCTINGFIRSTNAME.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If Trim(txtPERSONCONDUCTINGFIRSTNAME.Text) = "" Then
        MsgBox "Please enter a valid PERSON CONDUCTING Transation FIRST Name.", vbExclamation, "Form Incomplete"
        txtPERSONCONDUCTINGFIRSTNAME.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If mebAMOUNTCASHIN.Visible = True And (Trim(mebAMOUNTCASHIN.Text) = "" Or (Not IsNumeric(mebAMOUNTCASHIN.Text))) Then
        MsgBox "Please enter a valid CASH IN Amount.", vbExclamation, "Form Incomplete"
        mebAMOUNTCASHIN.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If mebAMOUNTCASHOUT.Visible = True And (Trim(mebAMOUNTCASHOUT.Text) = "" Or (Not IsNumeric(mebAMOUNTCASHOUT.Text))) Then
        MsgBox "Please enter a valid CASH OUT Amount.", vbExclamation, "Form Incomplete"
        mebAMOUNTCASHOUT.SetFocus
        IsValidForm = False
        Exit Function
    End If
    
    If Trim(cbTRANLOCATION.Text) = "" Then
        MsgBox "Please enter a valid TRANSACTION LOCATION.", vbExclamation, "Form Incomplete"
        mebAMOUNTCASHOUT.SetFocus
        IsValidForm = False
        Exit Function
    End If


    'Return TRUE If we got this far - form is OK
    IsValidForm = True

End Function


Private Sub subUpdateTransaction()
    ' This sub is used to update and existing record with values
    ' from the form.
    Dim strSQL As String
    Dim intRowsAffected As Integer

    ' Validate form values.
    If Not IsValidForm() Then
        Exit Sub
    End If


    If mebAMOUNTCASHIN.Text = "" Then mebAMOUNTCASHIN.Text = "0"
    If mebAMOUNTCASHOUT.Text = "" Then mebAMOUNTCASHOUT.Text = "0"


        ' Build update statement to update product table with data
        ' on form.
        strSQL = "UPDATE TRANREG SET" & _
                 " TRRECID = " & txtTRRECID.Text & _
                 " ,TRDATETIME = " & PrepareStr(txtTRDATETIME.Text) & _
                 " ,TRUSERID = " & PrepareStr(txtTRUSERID.Text) & _
                 " ,TRANDATE = " & PrepareStr(txtTRANDATE.Text) & _
                 " ,PERSONCONDUCTINGFIRSTNAME = " & PrepareStr(txtPERSONCONDUCTINGFIRSTNAME.Text) & _
                 " ,PERSONCONDUCTINGLASTNAME = " & PrepareStr(txtPERSONCONDUCTINGLASTNAME.Text) & _
                 " ,CIF = " & txtCIF.Text & _
                 " ,ACCOUNTNUM = " & txtACCOUNTNUM.Text & _
                 " ,ACCOUNTTITLE = " & PrepareStr(txtACCOUNTTITLE.Text) & _
                 " ,TRANTYPE = " & PrepareStr(cbTRANTYPE.Text) & _
                 " ,AMOUNTCASHIN = " & CDec(mebAMOUNTCASHIN.Text) & _
                 " ,AMOUNTCASHOUT = " & CDec(mebAMOUNTCASHOUT.Text) & _
                 " ,TRANLOCATION = " & PrepareStr(cbTRANLOCATION.Text) & _
                 " ,CTRNUMBER = " & PrepareStr(txtCTRNUMBER.Text) & _
                 " ,CTRREQUIRED = " & PrepareStr(txtCTRREQUIRED.Text) & _
                 " WHERE TRRECID = " & CLng(txtTRRECID.Text)

    Set Con = New ADODB.Connection
    Con.Open RegConnectionString
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    With rs
        .ActiveConnection = Con
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
    End With


    'Build Select statement to get the First (should be the ONLY) record in the control table
    rs.Source = strSQL

    rs.Open
        
    ' Close and Clean up objects
    Set rs = Nothing
    
End Sub

