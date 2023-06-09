VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmDocumentList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imaging101 DocumentList"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmDocumentListb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmDocumentListb.frx":0442
   ScaleHeight     =   10365
   ScaleWidth      =   11010
   Begin VB.CommandButton cmdResetFields 
      Caption         =   "&Reset Fields"
      Height          =   375
      Left            =   9240
      TabIndex        =   38
      Top             =   1912
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfigOrder 
      Caption         =   "Config/&Order"
      Height          =   375
      Left            =   9240
      TabIndex        =   37
      Top             =   2339
      Width           =   1695
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "S&can/Import"
      Height          =   420
      Left            =   9240
      TabIndex        =   36
      Top             =   2766
      Width           =   1695
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   660
      Left            =   9240
      Picture         =   "frmDocumentListb.frx":7E7C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdImaging101Batches 
      Caption         =   "Imaging101 Batches"
      Height          =   375
      Left            =   9240
      TabIndex        =   34
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtFolder 
      Height          =   315
      Left            =   2400
      TabIndex        =   32
      Top             =   2760
      Width           =   6135
   End
   Begin VB.TextBox txtFolderDescription 
      Height          =   315
      Left            =   2400
      TabIndex        =   31
      Top             =   3120
      Width           =   6135
   End
   Begin VB.ComboBox cmbFileroom 
      Height          =   315
      ItemData        =   "frmDocumentListb.frx":82BE
      Left            =   1800
      List            =   "frmDocumentListb.frx":82D1
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   6135
   End
   Begin VB.ComboBox cmbFilecabinet 
      Height          =   315
      ItemData        =   "frmDocumentListb.frx":8308
      Left            =   1800
      List            =   "frmDocumentListb.frx":831E
      TabIndex        =   19
      Top             =   1680
      Width           =   6135
   End
   Begin VB.ComboBox cmbDocumentType 
      Height          =   315
      ItemData        =   "frmDocumentListb.frx":835E
      Left            =   1800
      List            =   "frmDocumentListb.frx":8377
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CheckBox chkViewDocDetails 
      BackColor       =   &H0000C000&
      Caption         =   "View &Document Location Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   3480
      Width           =   3735
   End
   Begin VB.ComboBox cmbFileroomCondition 
      Height          =   315
      ItemData        =   "frmDocumentListb.frx":83CB
      Left            =   8040
      List            =   "frmDocumentListb.frx":83D5
      TabIndex        =   16
      Text            =   "And"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cmbFilecabinetCondition 
      Height          =   315
      ItemData        =   "frmDocumentListb.frx":83E2
      Left            =   8040
      List            =   "frmDocumentListb.frx":83EC
      TabIndex        =   15
      Text            =   "And"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox cmbDocumentTypeCondition 
      Height          =   315
      ItemData        =   "frmDocumentListb.frx":83F9
      Left            =   8040
      List            =   "frmDocumentListb.frx":8403
      TabIndex        =   14
      Text            =   "And"
      Top             =   2040
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   9720
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtFullPathName 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   8640
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   6030
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   6030
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   6030
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9330
      Visible         =   0   'False
      Width           =   6030
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmDocumentListb.frx":8410
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmDocumentListb.frx":8752
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmDocumentListb.frx":8A94
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmDocumentListb.frx":8DD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc AdodcFileRoom 
      Height          =   375
      Left            =   4320
      Top             =   9600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcFileRoom"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcFileCabinet 
      Height          =   375
      Left            =   4320
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcFileCabinet"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox mebDateFrom 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   21
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Bindings        =   "frmDocumentListb.frx":9118
      Height          =   2895
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5160
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mebDateThru 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   4920
      TabIndex        =   23
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8880
      TabIndex        =   39
      Top             =   480
      Width           =   1845
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Items Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   33
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "File Room"
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
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "File Cabinet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Document Date between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Doc Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "frmDocumentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean




Public Sub chkViewDocDetails_Click()
    
    If chkViewDocDetails.Value = 1 Then
        frmDocumentList.grdDataGrid.Columns.Item(0).Visible = True
        frmDocumentList.grdDataGrid.Columns.Item(1).Visible = True
        frmDocumentList.grdDataGrid.Columns.Item(2).Visible = True
        frmDocumentList.grdDataGrid.Columns.Item(3).Visible = True
    Else
        frmDocumentList.grdDataGrid.Columns.Item(0).Visible = False
        frmDocumentList.grdDataGrid.Columns.Item(1).Visible = False
        frmDocumentList.grdDataGrid.Columns.Item(2).Visible = False
        frmDocumentList.grdDataGrid.Columns.Item(3).Visible = False
    End If

End Sub



Private Sub cmdImaging101Batches_Click()
    Me.cmdImaging101Batches.Enabled = False
    frmImaging101BatchList.Show
End Sub


Private Sub cmdConfigOrder_Click()
    frmConfig.Show

End Sub


Private Sub cmdResetFields_Click()
    cmbFileroom = ""
    cmbFilecabinet = ""
    txtFolder = ""
    cmbDocumentType = ""
    mebDateFrom = ""
    mebDateThru = ""
    
End Sub

Private Sub cmdScan_Click()
''    Me.Hide
''    frmScan.Show
    frmImport.Show
End Sub

Private Sub Command1_Click()
End Sub


Private Sub DataCombo1_Click(Area As Integer)
'    AdodcFileRoom.RecordSource = "SELECT DISTINCT Fileroom FROM Documents"
'    AdodcFileRoom.Refresh

'    DataCombo1.DataSource = AdodcFileRoom
'    DataCombo1.DataMember = "Documents"
'    DataCombo1.DataField = "Fileroom"
'    DataCombo1.RowSource = AdodcFileRoom
'    DataCombo1.ListField = "Fileroom"
'    DataCombo1.BoundColumn = "Fileroom"
'    DataCombo1.Refresh
End Sub

Private Sub Form_Load()
  
    Dim RegConnectString As String
    Dim RegImaging101ConnectionType As String
    
    ' Get Database Connections settings from the registry
    On Error Resume Next
    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmDocumentList.Adodc1.ConnectionType", RegFileName)
    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmDocumentList.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
    On Error GoTo 0
    
    '*** Set SQL wildcard string
    RegConnectionWildcard = "%"
        
    
    '*** Connect to DB for Drop Down Lists
    
    AdodcFileRoom.ConnectionString = RegImaging101ConnectionString
    AdodcFileRoom.RecordSource = "SELECT DISTINCT Fileroom FROM 101Documents"
    AdodcFileRoom.Refresh
    
    
    '*** Connect to Document List DB
    
    Adodc1.ConnectionString = RegImaging101ConnectionString
    Adodc1.RecordSource = "SELECT * FROM 101Documents WHERE 0=1"
    Adodc1.Refresh
    frmDocumentList.grdDataGrid.Columns.Item(0).Visible = False
    frmDocumentList.grdDataGrid.Columns.Item(1).Visible = False
    frmDocumentList.grdDataGrid.Columns.Item(2).Visible = False
    frmDocumentList.grdDataGrid.Columns.Item(3).Visible = False
    
    
    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Left", RegFileName)
    Me.Width = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Width", RegFileName)
    Me.Height = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Height", RegFileName)
    Me.Caption = VBGetPrivateProfileString(RegAppname, "frmDocumentList.Caption", RegFileName)
    If Me.Caption = "" Then Me.Caption = "Document List"
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
    If gsecRightsAdminSystem = "Y" Then
        cmdConfigOrder.Enabled = True
    Else
        cmdConfigOrder.Enabled = False
    End If
    
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  Frame1.Width = Me.ScaleWidth
  grdDataGrid.Height = Me.ScaleHeight - 2600 - Frame1.Height
  grdDataGrid.Width = Me.Width - 300
  txtFullPathName.Top = Me.ScaleHeight - 300
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        CmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Top", Me.Top, RegFileName)
        Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Left", Me.Left, RegFileName)
        Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Width", Me.Width, RegFileName)
        Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Height", Me.Height, RegFileName)
        Result = WritePrivateProfileString(RegAppname, "frmDocumentList.Caption", Me.Caption, RegFileName)
    End If
  
    Screen.MousePointer = vbDefault
    frmMainMenu.Show

End Sub

Private Sub grdDataGrid_DblClick()
  
    grdDataGrid.Col = 2
    txtFullPathName = grdDataGrid.Text
    grdDataGrid.Col = 3
    txtFullPathName = frmConfig.txtRootDirToStoreObjects & txtFullPathName + "\" + grdDataGrid.Text
    
    DoEvents
 
    '*** Check if file Exists
    If Not subFileExists(txtFullPathName) Then
        Result = MsgBox("SORRY! I can't find file:" + vbNewLine + txtFullPathName + vbNewLine + "PLEASE CONTACT PC NETWORKS", vbCritical)
'        txtLOGOutputFilePath = Form1.txtOutputFilePath + "\" + txtFullPathName + ".LOG"
'        Open txtLOGOutputFilePath For Append As #4
'        Print #4, "Pass2 - Could Not Open either Original or Fixed file:  " + txtFullPathName
'        Close #4
        Exit Sub
    End If
    
    
    If frmConfig.chkRootDirForImagingApp = False Then
        '*** LOAD DOCUMENT INTO SPICER VIEWER
        Dim docContents As IDocContents
        Dim ActivePage As IActivePage
        Dim frmViewForm As ChildForm1
        
        Set frmViewForm = New ChildForm1
        
        frmViewForm.TextHeight (66)
        frmViewForm.TextWidth (132)
        
        frmViewForm.Caption = txtFullPathName
        frmViewForm.Show
    
        
        ' Close any open documents before opening a new one.
        ' Set the object variable for the IDocContents interface to the Document Control object
        Set docContents = frmViewForm.SpicerDoc1.object
        ' Close the document in the SpicerDoc1 control and
        ' Set CloseDocument to "True" to check if the document has been changed.
         docContents.CloseDocument False
        ' De-initialize the object variable
        Set docContents = Nothing
           
    '    ChildForm1.Show
        
        Set docContents = frmViewForm.SpicerDoc1.object
        Set ActivePage = frmViewForm.SpicerView1.object
        
        docContents.OpenFile txtFullPathName
        ActivePage.BindToDocumentControl frmViewForm.SpicerDoc1.object
        Set docContents = Nothing
        Set ActivePage = Nothing
        
    Else
        '*** TEMPORARY -- LOAD DOCUMENT INTO WINDOWS IMAGING
        intFileNameBegin = InStr(1, frmViewForm.Caption, "[") + 1
        txtLaunchString = Chr(34) & frmConfig.txtRootDirForImagingApp & Chr(34) & " " & Chr(34) & txtFullPathName & Chr(34)
        txtFrmID = Shell(txtLaunchString, vbNormalFocus)
    
    End If
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
''    PrimaryCLS.Resort ColIndex
End Sub



Private Sub PrimaryCLS_MoveComplete()
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(PrimaryCLS.AbsolutePosition)
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
'  PrimaryCLS.MoveLast
'  PrimaryCLS.AddNew
'  grdDataGrid.SetFocus

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
'  PrimaryCLS.Delete
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Public Sub cmdRefresh_Click()
    
    ' If DateTHRU is blank, fill the DateFROM
    If Trim(mebDateThru) = "" Then
        mebDateThru = mebDateFrom
    End If
  
    '***
    '*** SET UP THE "WHERE" CLAUSE FILTER STATEMENT
    '***
    
    txtfilterstatement = ""
    ' Check for FileRoom
    If Trim(frmDocumentList.cmbFileroom) <> "" Then
         txtfilterstatement = "Fileroom LIKE '%" + frmDocumentList.cmbFileroom + "%' "
    End If
    ' Check for FileCabinet
    If Trim(frmDocumentList.cmbFilecabinet) <> "" Then
        If Trim(txtfilterstatement) <> "" Then
                  txtfilterstatement = txtfilterstatement + " " + cmbFileroomCondition + " "
         End If
         txtfilterstatement = txtfilterstatement + " Filecabinet LIKE '%" + frmDocumentList.cmbFilecabinet + "%'"
    End If
    ' Check for DocumentType
    If Trim(frmDocumentList.cmbDocumentType) <> "" Then
        If Trim(txtfilterstatement) <> "" Then
                  txtfilterstatement = txtfilterstatement + " " + cmbFilecabinetCondition + " "
         End If
         txtfilterstatement = txtfilterstatement + " DocumentType LIKE '%" + frmDocumentList.cmbDocumentType + "%'"
    End If
    ' Check for DateFROM
    If Trim(frmDocumentList.mebDateFrom) <> "" Then
        If Trim(txtfilterstatement) <> "" Then
                  txtfilterstatement = txtfilterstatement + " " & cmbDocumentTypeCondition + " "
         End If
         txtfilterstatement = txtfilterstatement + " DocumentDate >= '" + Trim(Format(frmDocumentList.mebDateFrom, "####-##-##")) + "'"
    End If
    ' Check for DateTHRU
    If Trim(frmDocumentList.mebDateThru) <> "" Then
        If Trim(txtfilterstatement) <> "" Then
                  txtfilterstatement = txtfilterstatement + " And "
         End If
         txtfilterstatement = txtfilterstatement + " DocumentDate <= '" + Trim(Format(frmDocumentList.mebDateThru, "####-##-##")) + "'"
    End If
    ' Check for Folder
    If Trim(frmDocumentList.txtFolder) <> "" Then
        If Trim(txtfilterstatement) <> "" Then
                  txtfilterstatement = txtfilterstatement + " And "
         End If
         txtfilterstatement = txtfilterstatement + " Folder LIKE '%" + frmDocumentList.txtFolder + "%'"
    End If
    ' Check for FolderDescription
    If Trim(frmDocumentList.txtFolderDescription) <> "" Then
        If Trim(txtfilterstatement) <> "" Then
                  txtfilterstatement = txtfilterstatement + " And "
         End If
         txtfilterstatement = txtfilterstatement + " FolderDescription LIKE '%" + frmDocumentList.txtFolderDescription + "%'"
    End If
 
    If Trim(txtfilterstatement) <> "" Then
        txtfilterstatement = " WHERE " & txtfilterstatement
    End If
    
    '***
    '*** SET UP THE SELECT STATEMENT
    '***     INCLUDING THE "WHERE" AND "ORDER BY" CLAUSES
    
   Adodc1.RecordSource = "select BatchID, DocumentID , UNCFilePath, Filename, " & _
                        " Fileroom, Filecabinet, DocumentType, DocumentDate, PageCount, " & _
                        " Folder, FolderDescription, DocumentSubType, DateAdded, " & _
                        " DocumentExpireDate, DocumentNote, Field8, Field9, " & _
                        " Field10, Field11, Field12, Field13, Field14, Field15, " & _
                        " Field16, Field17, Field18, Field19, Field20 " & _
                        " FROM 101Documents " & txtfilterstatement & _
                        " ORDER BY " & _
                        frmConfig.cmbSort(0) + ", " + frmConfig.cmbSort(1) + ", " + frmConfig.cmbSort(2) + ", " + frmConfig.cmbSort(3)
       
   Adodc1.Refresh
   
   txtItemsFound = Adodc1.Recordset.RecordCount
   
   chkViewDocDetails_Click


   Exit Sub


RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
'  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  PrimaryCLS.Cancel
  SetButtons True
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

'  PrimaryCLS.Update
  SetButtons True
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  PrimaryCLS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  PrimaryCLS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub CmdNext_Click()
  On Error GoTo GoNextError

  PrimaryCLS.MoveNext
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  PrimaryCLS.MovePrevious
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  CmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub TabStrip1_Click()

End Sub

