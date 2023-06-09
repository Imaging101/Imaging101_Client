VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imaging101 Configuration"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   12960
   Begin TabDlg.SSTab sstabConfig 
      Height          =   9735
      Left            =   -120
      TabIndex        =   0
      Top             =   45
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   882
      TabMaxWidth     =   18
      BackColor       =   16777215
      MouseIcon       =   "frmConfig.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Application Definition"
      TabPicture(0)   =   "frmConfig.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmConfig.frx":0D38
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "OCR Options"
      TabPicture(2)   =   "frmConfig.frx":0D54
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Special Options"
      TabPicture(3)   =   "frmConfig.frx":0D70
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboSendEmailViaSMTP"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "frSMTPeMailFrame"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "frOCRFrame"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdUpdateSpecialOptions"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "frBarcodeFrame"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame7"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Shape2"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Security"
      TabPicture(4)   =   "frmConfig.frx":1A4A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frameRightsAssignments"
      Tab(4).Control(1)=   "chkRightsEnableCopyMode"
      Tab(4).Control(2)=   "chkRightsAdminSystem"
      Tab(4).Control(3)=   "cmbBatchDefaultApplication"
      Tab(4).Control(4)=   "txtSecurityApplicationRECID"
      Tab(4).Control(5)=   "cmdSecurityApplicationListGrant"
      Tab(4).Control(6)=   "lstSecurityApplicationSelectionList"
      Tab(4).Control(7)=   "cmdSecurityApplicationListRevoke"
      Tab(4).Control(8)=   "cmdSecurityApplicationListGrantRevoke"
      Tab(4).Control(9)=   "lstSecurityApplicationList"
      Tab(4).Control(10)=   "cmdSecurityRemove"
      Tab(4).Control(11)=   "cmdSecurityUpdate"
      Tab(4).Control(12)=   "cmdSecurityAddNew"
      Tab(4).Control(13)=   "cmdSecurityClearFields"
      Tab(4).Control(14)=   "txtPassword"
      Tab(4).Control(15)=   "lstUserList"
      Tab(4).Control(16)=   "txtUserName"
      Tab(4).Control(17)=   "txtUserID"
      Tab(4).Control(18)=   "txtSecurityRECID"
      Tab(4).Control(19)=   "cmdSecurityRefreshUsers"
      Tab(4).Control(20)=   "Frame4"
      Tab(4).Control(21)=   "chkUserGroup"
      Tab(4).Control(22)=   "cmdEditSearchTemplate"
      Tab(4).Control(23)=   "Frame5"
      Tab(4).Control(24)=   "Label31"
      Tab(4).Control(25)=   "Label33"
      Tab(4).Control(26)=   "Label27"
      Tab(4).Control(27)=   "Label26"
      Tab(4).Control(28)=   "Label25"
      Tab(4).Control(29)=   "Shape1"
      Tab(4).ControlCount=   30
      TabCaption(5)   =   "DB Maintenance"
      TabPicture(5)   =   "frmConfig.frx":2724
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Application Fields "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   135
         TabIndex        =   25
         Top             =   4680
         Width           =   12975
         Begin VB.TextBox txtFieldTypeHOLD 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   6960
            TabIndex        =   61
            Text            =   "txtFieldTypeHOLD"
            Top             =   1890
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cboFieldType 
            Height          =   315
            ItemData        =   "frmConfig.frx":2FFE
            Left            =   4560
            List            =   "frmConfig.frx":3000
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   1860
            Width           =   2265
         End
         Begin VB.TextBox txtFieldDescription 
            Height          =   285
            Left            =   4560
            TabIndex        =   59
            Top             =   1332
            Width           =   3735
         End
         Begin VB.TextBox txtFieldNameForOutput 
            Height          =   285
            Left            =   4560
            TabIndex        =   58
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txtFieldSize 
            Height          =   285
            Left            =   4560
            TabIndex        =   57
            Top             =   2175
            Width           =   2295
         End
         Begin VB.TextBox txtFieldName 
            Height          =   285
            Left            =   4560
            TabIndex        =   56
            Top             =   360
            Width           =   2955
         End
         Begin VB.TextBox txtFieldMask 
            Height          =   285
            Left            =   4560
            TabIndex        =   55
            Top             =   2580
            Width           =   3735
         End
         Begin VB.TextBox txtFieldFormat 
            Height          =   285
            Left            =   4560
            TabIndex        =   54
            Top             =   2865
            Width           =   3735
         End
         Begin VB.TextBox txtFieldLowValue 
            Height          =   285
            Left            =   4560
            TabIndex        =   53
            Top             =   3270
            Width           =   3735
         End
         Begin VB.TextBox txtFieldHighValue 
            Height          =   285
            Left            =   4560
            TabIndex        =   52
            Top             =   3555
            Width           =   3735
         End
         Begin VB.ComboBox cboFieldDefaultValue 
            Height          =   315
            ItemData        =   "frmConfig.frx":3002
            Left            =   4560
            List            =   "frmConfig.frx":301B
            TabIndex        =   51
            Top             =   3960
            Width           =   3735
         End
         Begin VB.TextBox txtFieldNameForInput 
            Height          =   285
            Left            =   4560
            TabIndex        =   50
            Top             =   765
            Width           =   3315
         End
         Begin VB.CommandButton cmdFieldsClear 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Clear &Input Fields"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3240
            Picture         =   "frmConfig.frx":3086
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   4440
            Width           =   2055
         End
         Begin VB.CommandButton cmdFieldDelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete Field"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   9480
            Picture         =   "frmConfig.frx":31D0
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   4440
            Width           =   2175
         End
         Begin VB.CommandButton cmdFieldUpdate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Update &Changes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7440
            Picture         =   "frmConfig.frx":375A
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   4440
            Width           =   2055
         End
         Begin VB.CommandButton cmdFieldAdd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Add New &Field"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5280
            Picture         =   "frmConfig.frx":38A4
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   4440
            Width           =   2175
         End
         Begin VB.CommandButton cmdFieldMoveDown 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            Picture         =   "frmConfig.frx":3BE6
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton cmdFieldMoveUp 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmConfig.frx":4028
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   4440
            Width           =   1455
         End
         Begin VB.TextBox txtFieldsRECID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   43
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox lstFields 
            Height          =   3960
            ItemData        =   "frmConfig.frx":446A
            Left            =   120
            List            =   "frmConfig.frx":446C
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   42
            Top             =   360
            Width           =   2895
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Field Options"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   8760
            TabIndex        =   26
            Top             =   240
            Width           =   4095
            Begin VB.CheckBox ckbFieldIsRequiredForCommit 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Required for COMMIT"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   40
               Top             =   2355
               Width           =   2535
            End
            Begin VB.CheckBox ckbFieldTableLookupOverridesDefault 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Table Lookup Overrides Default"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   39
               Top             =   3240
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.CheckBox ckbFieldIsForOutputOnly 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Prevent Manual Indexing"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   38
               Top             =   495
               Width           =   2535
            End
            Begin VB.CheckBox ckbFieldIsSticky 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Sticky Indexing"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   37
               ToolTipText     =   "Field Remains for each consecutive field until ovewritten."
               Top             =   240
               Value           =   1  'Checked
               Width           =   1680
            End
            Begin VB.CheckBox ckbFieldDropDownList 
               BackColor       =   &H00FFFFFF&
               Caption         =   "DropDown List on Retrieve"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   36
               ToolTipText     =   "Field Remains for each consecutive field until ovewritten."
               Top             =   750
               Width           =   2175
            End
            Begin VB.CheckBox ckbFieldIsRequiredForSplit 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Required for SPLIT"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   35
               Top             =   2610
               Width           =   2535
            End
            Begin VB.CheckBox ckbFieldSplitBatches 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Split Batches on This Field"
               Height          =   375
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   34
               Top             =   1005
               Width           =   2535
            End
            Begin VB.CheckBox ckbFieldRouteToBatchQueue 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Route to Batch QUEUE based on THIS Field"
               Height          =   270
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   33
               Top             =   1410
               Width           =   3615
            End
            Begin VB.ComboBox cboFieldSearchCondition 
               Height          =   315
               ItemData        =   "frmConfig.frx":446E
               Left            =   1470
               List            =   "frmConfig.frx":448D
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   3585
               Width           =   1335
            End
            Begin VB.CheckBox chkFieldDefaultForBarcodeOnly 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Default for Barcodes ONLY"
               Height          =   375
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   31
               Top             =   2880
               Width           =   2535
            End
            Begin VB.CheckBox ckbFieldRouteToBatchUser 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Route to Batch USER  based on THIS Field"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   30
               Top             =   1680
               Width           =   3615
            End
            Begin VB.CheckBox ckbFieldRouteToBatchManager 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Route to Batch MANAGER based on THIS Field"
               Height          =   255
               Left            =   120
               MaskColor       =   &H00000000&
               TabIndex        =   29
               Top             =   1950
               Width           =   3735
            End
            Begin VB.CheckBox ckbFieldDropDownListAlsoOnFiler 
               BackColor       =   &H00FFFFFF&
               Caption         =   "also on I101Filer"
               Height          =   255
               Left            =   2400
               MaskColor       =   &H00000000&
               TabIndex        =   28
               ToolTipText     =   "Field Remains for each consecutive field until ovewritten."
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox ckbHideForSearchIndex 
               BackColor       =   &H00FFFFFF&
               Caption         =   "HIDE for Search/Index"
               Height          =   255
               Left            =   2085
               MaskColor       =   &H00000000&
               TabIndex        =   27
               ToolTipText     =   "Field Remains for each consecutive field until ovewritten."
               Top             =   255
               Width           =   1905
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Search Condition"
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   3645
               Width           =   1335
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Label lblFieldSize 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   3240
            TabIndex        =   72
            Top             =   2268
            Width           =   264
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   3240
            TabIndex        =   71
            Top             =   1428
            Width           =   1212
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name for Output"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   3240
            TabIndex        =   70
            Top             =   1140
            Width           =   1044
         End
         Begin VB.Label lblFieldType 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   3240
            TabIndex        =   69
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label lblFieldName 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Field  Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   3240
            TabIndex        =   68
            Top             =   456
            Width           =   732
         End
         Begin VB.Label lblDefaultValue 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Default Value"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   7
            Left            =   3240
            TabIndex        =   67
            Top             =   4080
            Width           =   840
         End
         Begin VB.Label lblDefaultValue 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Input Mask"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   0
            Left            =   3240
            TabIndex        =   66
            Top             =   2676
            Width           =   708
         End
         Begin VB.Label lblDefaultValue 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Output Format"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   1
            Left            =   3240
            TabIndex        =   65
            Top             =   2952
            Width           =   924
         End
         Begin VB.Label lblLowValue 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Low Value"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   3
            Left            =   3240
            TabIndex        =   64
            Top             =   3360
            Width           =   648
         End
         Begin VB.Label lblHighValue 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "High Value"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   4
            Left            =   3240
            TabIndex        =   63
            Top             =   3648
            Width           =   684
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name for Input"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   192
            Left            =   3240
            TabIndex        =   62
            Top             =   852
            Width           =   948
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Applications "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   105
         TabIndex        =   172
         Top             =   600
         Width           =   12975
         Begin VB.ComboBox cboSiteID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            ItemData        =   "frmConfig.frx":44C1
            Left            =   10800
            List            =   "frmConfig.frx":44CB
            Style           =   2  'Dropdown List
            TabIndex        =   182
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdApplicationAdd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add New Application"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4680
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConfig.frx":44DF
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   3360
            Width           =   1920
         End
         Begin VB.CommandButton cmdApplicationUpdate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Update Changes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConfig.frx":4821
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   3360
            Width           =   1920
         End
         Begin VB.CommandButton cmdApplicationRemove 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Remove Application"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConfig.frx":496B
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   3360
            Width           =   1920
         End
         Begin VB.CommandButton cmdApplicationClear 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Clear Input Fields"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2760
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConfig.frx":4EF5
            Style           =   1  'Graphical
            TabIndex        =   178
            Top             =   3360
            Width           =   1920
         End
         Begin VB.CommandButton cmdApplicationEditDoctypes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit DocTypes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   10440
            Picture         =   "frmConfig.frx":503F
            Style           =   1  'Graphical
            TabIndex        =   177
            Top             =   3360
            Width           =   1920
         End
         Begin VB.CommandButton cmdRefreshApplications 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Re&fresh Applications"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Picture         =   "frmConfig.frx":53C9
            Style           =   1  'Graphical
            TabIndex        =   176
            Top             =   3360
            Width           =   2295
         End
         Begin VB.ListBox lstApplications 
            Height          =   2790
            ItemData        =   "frmConfig.frx":5753
            Left            =   120
            List            =   "frmConfig.frx":5755
            TabIndex        =   175
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtApplicationRECID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   9600
            TabIndex        =   174
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtApplicationName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4440
            TabIndex        =   173
            Top             =   120
            Width           =   5055
         End
         Begin TabDlg.SSTab sstabApplication 
            Height          =   3015
            Left            =   2760
            TabIndex        =   183
            Top             =   465
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   5318
            _Version        =   393216
            TabOrientation  =   3
            Tabs            =   5
            TabHeight       =   529
            BackColor       =   16777215
            ForeColor       =   8421376
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "MAIN"
            TabPicture(0)   =   "frmConfig.frx":5757
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label57"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label55"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label56"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label40"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label39"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label13"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label32"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "chkEnableSearchTemplates"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtMaxItemsToRetrieve"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txtApplicationBatchNameDelimiter"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "txtRouteMaxCount"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "chkAutoAdvanceOnSeparator"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "txtApplicationCommitBatchTo"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "chkApplicationIsReadOnly"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "chkApplicationIsActive"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "txtApplicationNotes"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "txtApplicationDescription"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "txtApplicationCommitBatchOption"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "chkSetUserAsBatchOwnerOnSPLIT"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "chkLogOpenedDocuments"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).ControlCount=   21
            TabCaption(1)   =   "Lookup"
            TabPicture(1)   =   "frmConfig.frx":5773
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label36"
            Tab(1).Control(1)=   "Label15"
            Tab(1).Control(2)=   "Label14"
            Tab(1).Control(3)=   "lblCaseIdCutoff"
            Tab(1).Control(4)=   "chkAutoLookupOnBatchLoad"
            Tab(1).Control(5)=   "chkLookupDBTableIsOnSQLServer"
            Tab(1).Control(6)=   "txtLookupDBWhereClause"
            Tab(1).Control(7)=   "cboLookupDBTableName"
            Tab(1).Control(8)=   "txtLookupDBConnectionString"
            Tab(1).Control(9)=   "txtCaseIdCutoff"
            Tab(1).Control(10)=   "cmdBackup"
            Tab(1).ControlCount=   11
            TabCaption(2)   =   "Fields"
            TabPicture(2)   =   "frmConfig.frx":578F
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cboFieldToAssignDocumentSubType"
            Tab(2).Control(1)=   "cboFieldToAssignDocumentGroup"
            Tab(2).Control(2)=   "cboFieldToAssignDocumentType"
            Tab(2).Control(3)=   "cboFieldToSelectAfterNextPageClick"
            Tab(2).Control(4)=   "cboFieldToSelectAfterDocListClick"
            Tab(2).Control(5)=   "cboFieldToSelectAfterLookupClick"
            Tab(2).Control(6)=   "Label22"
            Tab(2).Control(7)=   "Label38"
            Tab(2).Control(8)=   "Label35"
            Tab(2).Control(9)=   "Label34"
            Tab(2).Control(10)=   "Label42"
            Tab(2).Control(11)=   "Label43"
            Tab(2).ControlCount=   12
            TabCaption(3)   =   "Paths"
            TabPicture(3)   =   "frmConfig.frx":57AB
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label46"
            Tab(3).Control(1)=   "Label44"
            Tab(3).Control(2)=   "Label41"
            Tab(3).Control(3)=   "Label45"
            Tab(3).Control(4)=   "txtRootDirectoryPathForHtmlSource"
            Tab(3).Control(5)=   "txtRootDirectoryPathForBatches"
            Tab(3).Control(6)=   "txtRootDirectoryPathForImageAnnotations"
            Tab(3).Control(7)=   "txtRootDirectoryPathForImageArchive"
            Tab(3).ControlCount=   8
            TabCaption(4)   =   "FTP"
            TabPicture(4)   =   "frmConfig.frx":57C7
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "Label37"
            Tab(4).Control(1)=   "lblFTPDelimiter"
            Tab(4).Control(2)=   "lblFTPSelectField"
            Tab(4).Control(3)=   "lblFTPConfigureFileNaming"
            Tab(4).Control(4)=   "lblFTPUserID"
            Tab(4).Control(5)=   "lblFTPPassword"
            Tab(4).Control(6)=   "lblFTPSite"
            Tab(4).Control(7)=   "txtFTPPort"
            Tab(4).Control(8)=   "cmdTestFTP"
            Tab(4).Control(9)=   "cboFTPFileNameDelimiter(1)"
            Tab(4).Control(10)=   "cboFTPFileNameField(1)"
            Tab(4).Control(11)=   "cboFTPFileNameDelimiter(2)"
            Tab(4).Control(12)=   "cboFTPFileNameDelimiter(0)"
            Tab(4).Control(13)=   "cboFTPFileNameField(3)"
            Tab(4).Control(14)=   "cboFTPFileNameField(2)"
            Tab(4).Control(15)=   "cboFTPFileNameField(0)"
            Tab(4).Control(16)=   "txtFTPUserID"
            Tab(4).Control(17)=   "txtFTPPassword"
            Tab(4).Control(18)=   "txtFTPSite"
            Tab(4).ControlCount=   19
            Begin VB.CheckBox chkLogOpenedDocuments 
               Alignment       =   1  'Right Justify
               Caption         =   "LOG Opened Documents"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6000
               TabIndex        =   262
               Top             =   1425
               Width           =   2295
            End
            Begin VB.CheckBox chkSetUserAsBatchOwnerOnSPLIT 
               Alignment       =   1  'Right Justify
               Caption         =   "Use Logged-In User for SPLIT Batch Owner"
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
               Left            =   120
               TabIndex        =   224
               Top             =   2505
               Width           =   3615
            End
            Begin VB.CommandButton cmdBackup 
               Caption         =   "Backup"
               Height          =   255
               Left            =   -66840
               TabIndex        =   223
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox txtCaseIdCutoff 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   -66840
               TabIndex        =   222
               Text            =   "700000"
               Top             =   1440
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtFTPPort 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   -68040
               TabIndex        =   221
               Text            =   "21"
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdTestFTP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Test FTP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   -67200
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmConfig.frx":57E3
               Style           =   1  'Graphical
               TabIndex        =   220
               Top             =   360
               Width           =   1080
            End
            Begin VB.ComboBox txtApplicationCommitBatchOption 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               ItemData        =   "frmConfig.frx":60AD
               Left            =   5865
               List            =   "frmConfig.frx":60B7
               TabIndex        =   219
               Text            =   "Application Only"
               Top             =   1080
               Width           =   2175
            End
            Begin VB.ComboBox cboFTPFileNameDelimiter 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               ItemData        =   "frmConfig.frx":60E0
               Left            =   -69000
               List            =   "frmConfig.frx":6111
               Style           =   2  'Dropdown List
               TabIndex        =   218
               Top             =   1635
               Width           =   1695
            End
            Begin VB.ComboBox cboFTPFileNameField 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   -72960
               TabIndex        =   217
               Top             =   1635
               Width           =   3855
            End
            Begin VB.ComboBox cboFTPFileNameDelimiter 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               ItemData        =   "frmConfig.frx":61F3
               Left            =   -69000
               List            =   "frmConfig.frx":6224
               Style           =   2  'Dropdown List
               TabIndex        =   216
               Top             =   1920
               Width           =   1695
            End
            Begin VB.ComboBox cboFTPFileNameDelimiter 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               ItemData        =   "frmConfig.frx":6306
               Left            =   -69000
               List            =   "frmConfig.frx":6337
               Style           =   2  'Dropdown List
               TabIndex        =   215
               Top             =   1320
               Width           =   1695
            End
            Begin VB.ComboBox cboFTPFileNameField 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   -72960
               TabIndex        =   214
               Top             =   2265
               Width           =   3855
            End
            Begin VB.ComboBox cboFTPFileNameField 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   -72960
               TabIndex        =   213
               Top             =   1950
               Width           =   3855
            End
            Begin VB.ComboBox cboFTPFileNameField 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   -72960
               TabIndex        =   212
               Top             =   1320
               Width           =   3855
            End
            Begin VB.TextBox txtFTPUserID 
               Height          =   285
               Left            =   -73680
               TabIndex        =   211
               Top             =   720
               Width           =   3615
            End
            Begin VB.TextBox txtFTPPassword 
               Height          =   285
               Left            =   -68760
               TabIndex        =   210
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtFTPSite 
               Height          =   285
               Left            =   -73680
               TabIndex        =   209
               Top             =   360
               Width           =   4935
            End
            Begin VB.TextBox txtApplicationDescription 
               Height          =   285
               Left            =   1770
               TabIndex        =   208
               Top             =   120
               Width           =   6255
            End
            Begin VB.TextBox txtApplicationNotes 
               Height          =   525
               Left            =   1770
               MultiLine       =   -1  'True
               TabIndex        =   207
               Top             =   480
               Width           =   6255
            End
            Begin VB.CheckBox chkApplicationIsActive 
               Alignment       =   1  'Right Justify
               Caption         =   "Application Is Active?"
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
               Left            =   6000
               TabIndex        =   206
               Top             =   1680
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.CheckBox chkApplicationIsReadOnly 
               Alignment       =   1  'Right Justify
               Caption         =   "Application Is Read-Only?"
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
               Left            =   6000
               TabIndex        =   205
               Top             =   1920
               Width           =   2295
            End
            Begin VB.ComboBox txtApplicationCommitBatchTo 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               ItemData        =   "frmConfig.frx":6419
               Left            =   1755
               List            =   "frmConfig.frx":642F
               TabIndex        =   204
               Text            =   "Imaging101"
               Top             =   1095
               Width           =   3975
            End
            Begin VB.CheckBox chkAutoAdvanceOnSeparator 
               Alignment       =   1  'Right Justify
               Caption         =   "Automatically Advance to Next Image on Separator or Questionable?"
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
               Left            =   120
               TabIndex        =   203
               Top             =   1680
               Value           =   1  'Checked
               Width           =   5520
            End
            Begin VB.TextBox txtLookupDBConnectionString 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   -74790
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   202
               Top             =   360
               Width           =   7935
            End
            Begin VB.ComboBox cboLookupDBTableName 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmConfig.frx":647B
               Left            =   -73680
               List            =   "frmConfig.frx":647D
               TabIndex        =   201
               Top             =   1200
               Width           =   4215
            End
            Begin VB.TextBox txtRouteMaxCount 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   3630
               TabIndex        =   200
               Text            =   "4"
               Top             =   1935
               Width           =   495
            End
            Begin VB.TextBox txtLookupDBWhereClause 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   199
               Top             =   2040
               Width           =   8775
            End
            Begin VB.ComboBox cboFieldToSelectAfterLookupClick 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -71880
               TabIndex        =   198
               Top             =   1320
               Width           =   4575
            End
            Begin VB.ComboBox cboFieldToSelectAfterDocListClick 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -71880
               TabIndex        =   197
               Top             =   1635
               Width           =   4575
            End
            Begin VB.ComboBox cboFieldToSelectAfterNextPageClick 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -71880
               TabIndex        =   196
               Top             =   1950
               Width           =   4575
            End
            Begin VB.ComboBox cboFieldToAssignDocumentType 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -71880
               TabIndex        =   195
               Top             =   435
               Width           =   4575
            End
            Begin VB.ComboBox cboFieldToAssignDocumentGroup 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -71880
               TabIndex        =   194
               Top             =   120
               Width           =   4575
            End
            Begin VB.TextBox txtRootDirectoryPathForImageArchive 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74760
               TabIndex        =   193
               Top             =   360
               Width           =   7935
            End
            Begin VB.TextBox txtRootDirectoryPathForImageAnnotations 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74760
               TabIndex        =   192
               Top             =   960
               Width           =   7935
            End
            Begin VB.CheckBox chkLookupDBTableIsOnSQLServer 
               Alignment       =   1  'Right Justify
               Caption         =   "DB is on Microsoft SQL-Server"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -72120
               TabIndex        =   191
               Top             =   1560
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.TextBox txtRootDirectoryPathForBatches 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74760
               TabIndex        =   190
               Top             =   1560
               Width           =   7935
            End
            Begin VB.TextBox txtRootDirectoryPathForHtmlSource 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74760
               TabIndex        =   189
               Top             =   2160
               Width           =   7935
            End
            Begin VB.TextBox txtApplicationBatchNameDelimiter 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   7920
               TabIndex        =   188
               Text            =   "-"
               Top             =   2160
               Width           =   420
            End
            Begin VB.TextBox txtMaxItemsToRetrieve 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   3480
               TabIndex        =   187
               Text            =   "1000"
               Top             =   2280
               Width           =   735
            End
            Begin VB.CheckBox chkAutoLookupOnBatchLoad 
               Alignment       =   1  'Right Justify
               Caption         =   "Auto-Lookup on Batch Load"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -69360
               TabIndex        =   186
               Top             =   1200
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.ComboBox cboFieldToAssignDocumentSubType 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -71880
               TabIndex        =   185
               Top             =   750
               Width           =   4575
            End
            Begin VB.CheckBox chkEnableSearchTemplates 
               Alignment       =   1  'Right Justify
               Caption         =   "Enable Search Templates"
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
               Left            =   6000
               TabIndex        =   184
               Top             =   2520
               Width           =   2295
            End
            Begin VB.Label lblCaseIdCutoff 
               Caption         =   "Commit to Site B when CaseID >="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -69360
               TabIndex        =   253
               Top             =   1455
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.Label Label37 
               Caption         =   "FTP Port"
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
               Left            =   -68640
               TabIndex        =   252
               Top             =   390
               Width           =   615
            End
            Begin VB.Label lblFTPDelimiter 
               Caption         =   "Delimiter"
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
               Left            =   -69000
               TabIndex        =   251
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label lblFTPSelectField 
               Caption         =   "Select Field OR Enter Value to Use for File Name"
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
               Left            =   -72960
               TabIndex        =   250
               Top             =   1080
               Width           =   3255
            End
            Begin VB.Label lblFTPConfigureFileNaming 
               Caption         =   "CONFIGURE FILE NAMING"
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
               Left            =   -74640
               TabIndex        =   249
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label lblFTPUserID 
               Caption         =   "FTP UserID"
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
               Left            =   -74640
               TabIndex        =   248
               Top             =   750
               Width           =   975
            End
            Begin VB.Label lblFTPPassword 
               Caption         =   "FTP Password"
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
               Left            =   -69840
               TabIndex        =   247
               Top             =   750
               Width           =   1095
            End
            Begin VB.Label lblFTPSite 
               Caption         =   "FTP Site"
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
               Left            =   -74640
               TabIndex        =   246
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label32 
               Caption         =   "Commit Batches To:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   105
               TabIndex        =   245
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label13 
               Caption         =   "Application Notes"
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
               Left            =   120
               TabIndex        =   244
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Lookup Table Connection String"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74790
               TabIndex        =   243
               Top             =   120
               Width           =   3015
            End
            Begin VB.Label Label15 
               Caption         =   "Lookup Table "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   242
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label39 
               Caption         =   "Route Batch to Supervisor if Routed more than"
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
               Left            =   120
               TabIndex        =   241
               Top             =   1980
               Width           =   3450
            End
            Begin VB.Label Label40 
               Caption         =   "times."
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
               Left            =   4230
               TabIndex        =   240
               Top             =   1980
               Width           =   495
            End
            Begin VB.Label Label36 
               Caption         =   "Table Lookup Record Selection Condition (WHERE Clause)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   239
               Top             =   1800
               Width           =   4695
            End
            Begin VB.Label Label43 
               Caption         =   "Field to Assign Document GROUP to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   238
               Top             =   120
               Width           =   2895
            End
            Begin VB.Label Label42 
               Caption         =   "Field to Assign Document TYPE to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   237
               Top             =   435
               Width           =   2895
            End
            Begin VB.Label Label34 
               Caption         =   "Field to select AFTER Lookup"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   236
               Top             =   1320
               Width           =   2175
            End
            Begin VB.Label Label35 
               Caption         =   "Field to select AFTER DocType Click"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   -74760
               TabIndex        =   235
               Top             =   1635
               Width           =   2895
            End
            Begin VB.Label Label38 
               Caption         =   "Field to select AFTER Next Page Click"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74760
               TabIndex        =   234
               Top             =   1950
               Width           =   2895
            End
            Begin VB.Label Label45 
               Caption         =   "Full Directory Path for Image ARCHIVE:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   233
               Top             =   120
               Width           =   3255
            End
            Begin VB.Label Label41 
               Caption         =   "Full Directory Path for Image ANNOTATIONS:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   232
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label Label44 
               Caption         =   "Full Directory Path for Image BATCHES (This is the ""Default""... it can be changed in Scan Profile))"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   231
               Top             =   1320
               Width           =   7695
            End
            Begin VB.Label Label46 
               Caption         =   "Full Directory Path of HTML Source Code for EXPORT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   230
               Top             =   1920
               Width           =   7695
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Batch Name Delimiter"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6000
               TabIndex        =   229
               Top             =   2280
               Width           =   1815
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label55 
               Caption         =   "Maximum Items to Show on Retrieval Search"
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
               Left            =   120
               TabIndex        =   228
               Top             =   2295
               Width           =   3375
            End
            Begin VB.Label Label57 
               Caption         =   "( 0 = Unlimited)"
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
               Left            =   4245
               TabIndex        =   227
               Top             =   2310
               Width           =   1215
            End
            Begin VB.Label Label22 
               Caption         =   "Field to Assign Document SUB-TYPE to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   226
               Top             =   750
               Width           =   2895
            End
            Begin VB.Label Label11 
               Caption         =   "Application Description"
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
               Left            =   120
               TabIndex        =   225
               Top             =   120
               Width           =   1815
            End
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Application Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   2520
            TabIndex        =   254
            Top             =   150
            Width           =   2055
         End
      End
      Begin VB.Frame frameRightsAssignments 
         BackColor       =   &H00F4E0DB&
         Caption         =   "Rights Assignments by User / Application "
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
         Height          =   4095
         Left            =   -72480
         TabIndex        =   136
         Top             =   4680
         Width           =   9855
         Begin VB.CheckBox chkRightsAdvancedSearch 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Advanced Search"
            Height          =   255
            Left            =   7080
            TabIndex        =   167
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchAllowDocTypeEdit 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Edit DocTypes"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   166
            Top             =   2640
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsBatchFindRestrictToQueue 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Restrict to Queue"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4560
            TabIndex        =   165
            Top             =   3120
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchFindRestrictToOwner 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Restrict to Owner"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4560
            TabIndex        =   164
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchFindRestricted 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Restrict Batch Find"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4200
            TabIndex        =   163
            Top             =   2640
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchChangeQueue 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Change Queue"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   162
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchChangeOwner 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Change Owner"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   161
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkAllowModificationOfOrigDocs 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Allow modification of Original Documents in BATCH Mode [Edit]"
            Height          =   615
            Left            =   4200
            TabIndex        =   160
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsExport 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Export Doc"
            Height          =   255
            Left            =   7080
            TabIndex        =   159
            Top             =   3840
            Width           =   1935
         End
         Begin VB.CheckBox chkViewResetImagesOnFind 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Reset Viewer on Find"
            Height          =   255
            Left            =   7080
            TabIndex        =   158
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchAdministration 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Admininstration"
            Height          =   255
            Left            =   2520
            TabIndex        =   157
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsThumbnails 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Thumbnails"
            Height          =   255
            Left            =   7080
            TabIndex        =   156
            Top             =   3600
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsAnnotate 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Annotate"
            Height          =   255
            Left            =   7080
            TabIndex        =   155
            Top             =   3360
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsPrint 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Print"
            Height          =   255
            Left            =   7080
            TabIndex        =   154
            Top             =   3120
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsAdminApplication 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Application Admin"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   153
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox chkRightsBatchChangeOrder 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Change Order"
            Height          =   255
            Left            =   4200
            TabIndex        =   152
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchScan 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Scan"
            Height          =   255
            Left            =   2520
            TabIndex        =   151
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsBatchIndex 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Index"
            Height          =   255
            Left            =   2520
            TabIndex        =   150
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsImportFromFile 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Import From File"
            Height          =   255
            Left            =   2520
            TabIndex        =   149
            Top             =   2400
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsDeleteDocuments 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Delete Documents"
            Height          =   255
            Left            =   7080
            TabIndex        =   148
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsModifyIndexes 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Modify Indexes"
            Height          =   255
            Left            =   7080
            TabIndex        =   147
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchCommit 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Commit"
            Height          =   255
            Left            =   2520
            TabIndex        =   146
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsDeleteBatches 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Delete "
            Height          =   255
            Left            =   2520
            TabIndex        =   145
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsSendMail 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Send Mail"
            Height          =   255
            Left            =   7080
            TabIndex        =   144
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsLaunchDoc 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Launch Doc"
            Height          =   255
            Left            =   7080
            TabIndex        =   143
            Top             =   2640
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsScannerSettings 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Edit Scanner Settings"
            Height          =   255
            Left            =   4200
            TabIndex        =   142
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsBatchView 
            BackColor       =   &H00F4E0DB&
            Caption         =   "View"
            Height          =   255
            Left            =   2520
            TabIndex        =   141
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsBatchRoute 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Route"
            Height          =   255
            Left            =   2520
            TabIndex        =   140
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CheckBox chkRightsRetrieveImages 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Retrieve Images"
            Height          =   255
            Left            =   7080
            TabIndex        =   139
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkRightsFileDocsViaI101FILER 
            BackColor       =   &H00F4E0DB&
            Caption         =   "File Docs via I101FILER"
            Height          =   255
            Left            =   7080
            TabIndex        =   138
            Top             =   1920
            Width           =   2415
         End
         Begin VB.CheckBox chkRightsEditSearchTemplates 
            BackColor       =   &H00F4E0DB&
            Caption         =   "Edit Search Templates"
            Height          =   495
            Left            =   135
            TabIndex        =   137
            Top             =   1200
            Width           =   2160
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000000&
            BorderWidth     =   4
            X1              =   7080
            X2              =   9000
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00000000&
            BorderWidth     =   4
            X1              =   7065
            X2              =   8985
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            BorderWidth     =   4
            X1              =   2520
            X2              =   6360
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            BorderWidth     =   4
            X1              =   120
            X2              =   2040
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label8 
            BackColor       =   &H00F4E0DB&
            Caption         =   "VIEWER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7080
            TabIndex        =   171
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F4E0DB&
            Caption         =   "RETRIEVAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7080
            TabIndex        =   170
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackColor       =   &H00F4E0DB&
            Caption         =   "BATCH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   2520
            TabIndex        =   169
            Top             =   360
            Width           =   1692
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F4E0DB&
            Caption         =   "ADMINISTRATION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkRightsEnableCopyMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rights Copy Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07000&
         Height          =   255
         Left            =   -65280
         TabIndex        =   134
         ToolTipText     =   "Wjen Checked ON, allows quickly copying settings to other Applications and Users."
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox chkRightsAdminSystem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "System Administrator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07000&
         Height          =   255
         Left            =   -65280
         TabIndex        =   133
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cmbBatchDefaultApplication 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfig.frx":647F
         Left            =   -70320
         List            =   "frmConfig.frx":6481
         TabIndex        =   132
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtSecurityApplicationRECID 
         Height          =   345
         Left            =   -73365
         TabIndex        =   131
         Top             =   7185
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CheckBox cboSendEmailViaSMTP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Send Email Via SMTP instead of Outlook"
         Height          =   255
         Left            =   -68280
         TabIndex        =   130
         Top             =   960
         Width           =   4215
      End
      Begin VB.Frame frSMTPeMailFrame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SMTP eMail Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -68280
         TabIndex        =   109
         Top             =   1320
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox txtSMTPPOP3Host 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            TabIndex        =   118
            Top             =   2040
            Width           =   2655
         End
         Begin VB.TextBox txtSMTPPort 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            TabIndex        =   117
            Text            =   "25"
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtSMTPEmailSubject 
            Height          =   525
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   116
            Text            =   "frmConfig.frx":6483
            Top             =   4320
            Width           =   4695
         End
         Begin VB.TextBox txtSMTPAuthenticationUserID 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1200
            TabIndex        =   115
            Top             =   3360
            Width           =   3615
         End
         Begin VB.TextBox txtSMTPAuthenticationPassword 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1200
            TabIndex        =   114
            Top             =   3600
            Width           =   3615
         End
         Begin VB.CheckBox cboSMTPRequiresAuthentication 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Requires Authentication?"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   2880
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox cboSMTPUsePOP3Auth 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use POP3 Authentication?"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   3120
            Width           =   2535
         End
         Begin VB.TextBox txtSMTPEmailMessage 
            Height          =   1095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   111
            Text            =   "frmConfig.frx":649F
            Top             =   5280
            Width           =   4695
         End
         Begin VB.TextBox txtSMTPHost 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            TabIndex        =   110
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Port"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   129
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "But the Return Address and Display Name will be stored in the users' record in the Security Table as a Default."
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   120
            TabIndex        =   128
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Subject:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   127
            Top             =   4080
            Width           =   1935
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "The user will be prompted for his/her eMail Return Address, Display Name, Subject and Message each time. "
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "UserID"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   125
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "AUTHENTICATION Information"
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
            Index           =   12
            Left            =   120
            TabIndex        =   124
            Top             =   2640
            Width           =   3135
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   123
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Message:"
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
            Index           =   14
            Left            =   120
            TabIndex        =   122
            Top             =   5040
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP HOST (IP or Name)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   121
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "SERVER Information"
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
            Index           =   8
            Left            =   120
            TabIndex        =   120
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "POP3 HOST (IP or Name)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   119
            Top             =   2160
            Width           =   1935
         End
      End
      Begin VB.Frame frOCRFrame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OCR Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74640
         TabIndex        =   104
         Top             =   7560
         Visible         =   0   'False
         Width           =   6135
         Begin VB.CheckBox Check1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   105
            Text            =   """C:\Program Files\TextBridge Pro 98\Bin\tb.exe"""
            Top             =   600
            Width           =   5175
         End
         Begin VB.Label Label52 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable OCR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label51 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Full Path to OCR Application"
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
            Left            =   840
            TabIndex        =   107
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CommandButton cmdUpdateSpecialOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update Special Options Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -66120
         Picture         =   "frmConfig.frx":6539
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   8520
         Width           =   2775
      End
      Begin VB.Frame frBarcodeFrame 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Barcodes "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74670
         TabIndex        =   88
         Top             =   3165
         Width           =   6135
         Begin VB.TextBox txtBarcodeClipNumberOfCharacters 
            Height          =   285
            Left            =   4440
            TabIndex        =   93
            Top             =   2010
            Width           =   375
         End
         Begin VB.TextBox txtBarcodeClipBeginPosition 
            Height          =   285
            Left            =   3600
            TabIndex        =   92
            Top             =   2010
            Width           =   375
         End
         Begin VB.CheckBox chkUseBarcodeAsDocumentHeader 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Use Barcode as Document Header"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            ToolTipText     =   "Checking this ON will append non-Barcoded pages to the previously recognized barcode page."
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CheckBox chkDropLeadingZeroes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Drop Leading Zeroes "
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   1455
            Width           =   3015
         End
         Begin VB.TextBox txtBarcodeLicenseKey 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   89
            Top             =   570
            Width           =   3015
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "These settings affect the LOCAL Machine ONLY."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Leave Blank if NOT Applicable"
            Height          =   255
            Index           =   6
            Left            =   3600
            TabIndex        =   101
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Clip"
            Height          =   255
            Index           =   5
            Left            =   4080
            TabIndex        =   100
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Characters"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   99
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Begin with Character Position "
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   98
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Clip Barcode:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   97
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblBarcodeLicenseStatus 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode License Status"
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
            Left            =   2160
            TabIndex        =   96
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode License Status"
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
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode License Key"
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
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdSecurityApplicationListGrant 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grant"
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
         Left            =   -73560
         Picture         =   "frmConfig.frx":697B
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   5640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox lstSecurityApplicationSelectionList 
         Height          =   255
         ItemData        =   "frmConfig.frx":7645
         Left            =   -73320
         List            =   "frmConfig.frx":7647
         MultiSelect     =   1  'Simple
         OLEDragMode     =   1  'Automatic
         TabIndex        =   86
         Top             =   6480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdSecurityApplicationListRevoke 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Revoke"
         Enabled         =   0   'False
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
         Left            =   -73560
         Picture         =   "frmConfig.frx":7649
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   4920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdSecurityApplicationListGrantRevoke 
         BackColor       =   &H00FFFFFF&
         Caption         =   "          Grant/Revoke           Application Access"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         Picture         =   "frmConfig.frx":8313
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   8760
         Width           =   2415
      End
      Begin VB.ListBox lstSecurityApplicationList 
         Height          =   3960
         ItemData        =   "frmConfig.frx":8755
         Left            =   -74880
         List            =   "frmConfig.frx":8757
         OLEDropMode     =   1  'Manual
         TabIndex        =   83
         Top             =   4665
         Width           =   2295
      End
      Begin VB.CommandButton cmdSecurityRemove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Remove User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -64425
         Picture         =   "frmConfig.frx":8759
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   8790
         Width           =   1800
      End
      Begin VB.CommandButton cmdSecurityUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -66225
         Picture         =   "frmConfig.frx":8CE3
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   8790
         Width           =   1800
      End
      Begin VB.CommandButton cmdSecurityAddNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add New User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -67995
         Picture         =   "frmConfig.frx":9125
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   8790
         Width           =   1800
      End
      Begin VB.CommandButton cmdSecurityClearFields 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear Input Fields"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -70065
         Picture         =   "frmConfig.frx":9467
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   8790
         Width           =   1800
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -70440
         PasswordChar    =   "*"
         TabIndex        =   78
         Top             =   1620
         Width           =   3855
      End
      Begin VB.ListBox lstUserList 
         Height          =   3180
         ItemData        =   "frmConfig.frx":95B1
         Left            =   -74880
         List            =   "frmConfig.frx":95B3
         TabIndex        =   77
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70440
         TabIndex        =   76
         Top             =   1335
         Width           =   3855
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70440
         TabIndex        =   75
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtSecurityRECID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67800
         TabIndex        =   74
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSecurityRefreshUsers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Re&fresh Users"
         Height          =   495
         Left            =   -73800
         Picture         =   "frmConfig.frx":95B5
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E3CC6C&
         Caption         =   "Defaults by User / Application  "
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
         Height          =   2070
         Left            =   -72465
         TabIndex        =   13
         Top             =   2610
         Width           =   9810
         Begin VB.ComboBox cmbBatchDefaultQueue 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmConfig.frx":9B3F
            Left            =   2625
            List            =   "frmConfig.frx":9B41
            TabIndex        =   18
            Top             =   600
            Width           =   4935
         End
         Begin VB.ComboBox cmbUserSupervisor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmConfig.frx":9B43
            Left            =   2625
            List            =   "frmConfig.frx":9B45
            TabIndex        =   17
            Top             =   285
            Width           =   4935
         End
         Begin VB.ComboBox cmbBatchListOrder 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmConfig.frx":9B47
            Left            =   2625
            List            =   "frmConfig.frx":9B49
            TabIndex        =   16
            Top             =   915
            Width           =   4935
         End
         Begin VB.ComboBox cmbBatchMode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmConfig.frx":9B4B
            Left            =   2625
            List            =   "frmConfig.frx":9B55
            TabIndex        =   15
            Top             =   1230
            Width           =   2295
         End
         Begin VB.TextBox txtBatchQueueNotificationFrequency 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3345
            TabIndex        =   14
            Text            =   "0"
            Top             =   1530
            Width           =   495
         End
         Begin VB.Label Label30 
            BackColor       =   &H00E3CC6C&
            Caption         =   "Batch Default Queue"
            Height          =   255
            Left            =   720
            TabIndex        =   24
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label28 
            BackColor       =   &H00E3CC6C&
            Caption         =   "User's Supervisor"
            Height          =   255
            Left            =   705
            TabIndex        =   23
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label Label29 
            BackColor       =   &H00E3CC6C&
            Caption         =   "Batch List Order"
            Height          =   255
            Left            =   705
            TabIndex        =   22
            Top             =   945
            Width           =   1455
         End
         Begin VB.Label Label99 
            BackColor       =   &H00E3CC6C&
            Caption         =   "Batch Mode"
            Height          =   255
            Left            =   720
            TabIndex        =   21
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label48 
            BackColor       =   &H00E3CC6C&
            Caption         =   "Batch Queue Notification Frequency"
            Height          =   255
            Left            =   705
            TabIndex        =   20
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label49 
            BackColor       =   &H00E3CC6C&
            Caption         =   "Minutes (0 = Disable Notification, Max = 9999)"
            Height          =   255
            Left            =   3945
            TabIndex        =   19
            Top             =   1560
            Width           =   4215
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   9255
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   12735
         Begin VB.CommandButton cmdDBMaintenance 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Imaging101 Database Table &Editing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   7920
            Picture         =   "frmConfig.frx":9B67
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3720
            Width           =   3015
         End
         Begin VB.CommandButton cmdCommandPrompt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Command Prompt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   480
            Picture         =   "frmConfig.frx":A431
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   7200
            Width           =   1575
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "* * *   W A R N I N G  * * *"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3960
            TabIndex        =   12
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Changes made with these Utilities CANNOT Be Undone!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   855
            Left            =   3555
            TabIndex        =   11
            Top             =   1200
            Width           =   5175
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "BACKING UP your database PRIOR to using these Utilities is HIGHLY Recommended!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   3960
            TabIndex        =   10
            Top             =   2160
            Width           =   4335
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOTE: You may click this button more than once to open additional windows to allow editing multiple tables at once."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   2520
            TabIndex        =   9
            Top             =   4080
            Width           =   5175
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto-Launch Documents"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -74670
         TabIndex        =   3
         Top             =   960
         Width           =   6135
         Begin VB.ComboBox txtAutoLaunchTo 
            Height          =   315
            ItemData        =   "frmConfig.frx":ACFB
            Left            =   1740
            List            =   "frmConfig.frx":AD05
            TabIndex        =   260
            Text            =   "New Imaging101 Document Viewer"
            Top             =   915
            Width           =   4140
         End
         Begin VB.TextBox txtAutoLaunchFileTypes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label54 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto-Launch to"
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
            Left            =   255
            TabIndex        =   261
            Top             =   945
            Width           =   1530
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "File Extensions to Auto-Launch (PDF, MSG, etc.)"
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
            TabIndex        =   5
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.CheckBox chkUserGroup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "USER GROUP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07000&
         Height          =   255
         Left            =   -65280
         TabIndex        =   2
         ToolTipText     =   "Wjen Checked ON, allows quickly copying settings to other Applications and Users."
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdEditSearchTemplate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search &Templates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -71910
         Picture         =   "frmConfig.frx":AD49
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8790
         Width           =   1800
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F4E0DB&
         Caption         =   "Application Access"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07000&
         Height          =   5415
         Left            =   -75000
         TabIndex        =   135
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch Default Application"
         Height          =   255
         Left            =   -72240
         TabIndex        =   259
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   258
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
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
         Left            =   -72240
         TabIndex        =   257
         Top             =   1650
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UserID"
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
         Left            =   -72240
         TabIndex        =   256
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label25 
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
         Left            =   -72240
         TabIndex        =   255
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   9015
         Left            =   -74880
         Top             =   600
         Width           =   12495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   9135
         Left            =   -74880
         Top             =   495
         Width           =   12615
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adodcAppConn As ADODB.Connection
Public rsAppConn As ADODB.Recordset
Private Shape1ColorHold

    

Private Sub cboFieldToAssignDocumentSubType_DropDown()
    
    On Error Resume Next
    
    funcFillList cboFieldToAssignDocumentSubType, RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID, True, True

End Sub

Private Sub cboFieldToSelectAfterDocListClick_DropDown()

    On Error Resume Next
    
    funcFillList cboFieldToSelectAfterDocListClick, RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID & " AND FieldIsForOutputOnly <> '1'", False, True

End Sub

Private Sub cboFieldToSelectAfterLookupClick_DropDown()

    On Error Resume Next
    
    funcFillList cboFieldToSelectAfterLookupClick, RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID & " AND FieldIsForOutputOnly <> '1'", False, True

End Sub



Private Sub cboFieldToSelectAfterNextPageClick_DropDown()

    On Error Resume Next
    
    funcFillList cboFieldToSelectAfterNextPageClick, RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID & " AND FieldIsForOutputOnly <> '1'", False, True

End Sub

Private Sub cboFieldType_Click()

'    If cboFieldType.Text = "Text" Or cboFieldType.Text = "Notes" Then
'        txtFieldSize.Enabled = True
'    Else
'        txtFieldSize.Enabled = True
'        txtFieldSize = ""
'        txtFieldSize.Enabled = False
'    End If

    Select Case cboFieldType.Text
        Case "Text"
            txtFieldSize.Enabled = True
            lblDefaultValue(0).Visible = True
            lblDefaultValue(1).Visible = True
            txtFieldMask.Visible = True
            txtFieldFormat.Visible = True
            txtFieldMask.Text = ""
            txtFieldFormat.Text = ""
         Case "LongText"
            txtFieldSize.Enabled = True
            lblDefaultValue(0).Visible = False
            lblDefaultValue(1).Visible = False
            txtFieldMask.Visible = False
            txtFieldFormat.Visible = False
            txtFieldMask.Text = ""
            txtFieldFormat.Text = ""
        Case "Notes"
            txtFieldSize.Enabled = True
            txtFieldSize.Text = 255
            lblDefaultValue(0).Visible = False
            lblDefaultValue(1).Visible = False
            txtFieldMask.Visible = False
            txtFieldFormat.Visible = False
            txtFieldMask.Text = ""
            txtFieldFormat.Text = ""
        Case "Date"
            txtFieldSize.Enabled = False
            lblDefaultValue(0).Visible = True
            lblDefaultValue(1).Visible = True
            txtFieldMask.Visible = True
            txtFieldFormat.Visible = True
            txtFieldMask.Text = "##-##-####"
            txtFieldFormat.Text = "mm-dd-yyyy"
        Case "Currency"
            txtFieldSize.Enabled = False
            lblDefaultValue(0).Visible = True
            lblDefaultValue(1).Visible = True
            txtFieldMask.Visible = True
            txtFieldFormat.Visible = True
            txtFieldMask.Text = ""
            txtFieldFormat.Text = "###,###,###,##0.00"
        Case "Numeric"
            txtFieldSize.Enabled = False
            lblDefaultValue(0).Visible = True
            lblDefaultValue(1).Visible = True
            txtFieldMask.Visible = True
            txtFieldFormat.Visible = True
            txtFieldMask.Text = ""
            txtFieldFormat.Text = ""
       Case Else
            txtFieldSize = ""
            txtFieldSize.Enabled = False
            lblDefaultValue(0).Visible = False
            lblDefaultValue(1).Visible = False
             txtFieldMask.Visible = False
            txtFieldFormat.Visible = False
           txtFieldMask.Text = ""
            txtFieldFormat.Text = ""
    End Select

    
End Sub



Private Sub cboFTPFileNameField_DropDown(Index As Integer)

    On Error Resume Next
    
    funcFillList cboFTPFileNameField(Index), RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID, True, True

End Sub

Private Sub cboLookupDBTableName_DropDown()

    On Error GoTo ERROR_HANDLER
    
    cboLookupDBTableName.Clear
    
    Dim CatDB As ADOX.Catalog
    Set CatDB = New ADOX.Catalog
    
    'open the database
    RegGenericConnectionString = txtLookupDBConnectionString.Text
    CatDB.ActiveConnection = RegGenericConnectionString

    'Fill Tables Combo
    For inttableindex = 0 To CatDB.Tables.Count - 1
        ' Show User TABLES ONLY... Don't show System Tables or Queries
        If CatDB.Tables.item(inttableindex).Type = "TABLE" _
        Or CatDB.Tables.item(inttableindex).Type = "VIEW" Then
            cboLookupDBTableName.AddItem CatDB.Tables.item(inttableindex).name
        End If
    Next
    
    Set CatDB = Nothing

Exit Sub

ERROR_HANDLER:
    
    result = MsgBox("cboLookupDBTableName_DropDown: " & Err.Number & " - " & Err.Description & " - " & Err.Source & " [Line " & CStr(Erl) & "]", vbOKOnly, "Add Application Error")
    Err.Clear
    


End Sub



Private Sub cboFieldToAssignDocumentGroup_DropDown()
    
    On Error Resume Next
    
    funcFillList cboFieldToAssignDocumentGroup, RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID, True, True

End Sub



Private Sub cboFieldToAssignDocumentType_DropDown()
    
    On Error Resume Next
    
    funcFillList cboFieldToAssignDocumentType, RegImaging101ConnectionString, "I101Fields", "FieldNameForInput", "ApplicationRECID = " & txtApplicationRECID, True, True

End Sub



Private Sub cboSendEmailViaSMTP_Click()

    If cboSendEmailViaSMTP = vbChecked Then
        frSMTPeMailFrame.Visible = True
    Else
        frSMTPeMailFrame.Visible = False
    End If
    

End Sub



Private Sub cboSiteID_Click()

    'Only do after an Application is clicked
    If lstApplications.ListIndex > 0 And cboSiteID.Visible = True Then
        ' Re-Load Application information
        lstApplications_Click
    End If
    


End Sub



Private Sub chkSendEmailViaSMTP_Click()

    If chkSendEmailViaSMTP.Value = vbChecked Then
        frSMTPeMailFrame.Visible = True
    Else
        frSMTPeMailFrame.Visible = False
    End If

End Sub



Private Sub chkRightsBatchFindRestricted_Click()

    If chkRightsBatchFindRestricted = vbChecked Then
        chkRightsBatchFindRestrictToQueue.Enabled = True
        chkRightsBatchFindRestrictToOwner.Enabled = True
    Else
        chkRightsBatchFindRestrictToQueue.Enabled = False
        chkRightsBatchFindRestrictToOwner.Enabled = False
    End If

End Sub



Private Sub chkRightsEnableCopyMode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If chkRightsEnableCopyMode.Value = vbChecked Then
    
        Shape1ColorHold = Shape1.FillColor
        Shape1.FillColor = &HC0C0FF
        MsgBox "Rights Settings will NOT be Cleared or Loaded when other Users or Applications are selected." & vbCrLf & vbCrLf & _
                    "This will allow copying the settings quickly to other Users and/or Applications" & vbCrLf & vbCrLf & _
                    "CLICK [Update Settings] Button for EACH User / Application Selection Change" & vbCrLf & vbCrLf & _
                    "UN-CHECK to RESUME Normal operations.", vbOKOnly, "Rights Copy Mode"
                    
    Else
    
        Shape1.FillColor = Shape1ColorHold
        
    End If
    
End Sub

Private Sub ckbFieldIsForOutputOnly_Click()

    Dim bolFieldCleared As Boolean
    bolFieldCleared = False
    
    If ckbFieldIsForOutputOnly.Value = vbChecked Then
        If cboFieldToSelectAfterDocListClick.Text = txtFieldNameForInput.Text Then
            cboFieldToSelectAfterDocListClick.Text = ""
            bolFieldCleared = True
        End If
        
        If cboFieldToSelectAfterLookupClick.Text = txtFieldNameForInput.Text Then
            cboFieldToSelectAfterLookupClick.Text = ""
            bolFieldCleared = True
        End If
        
        If cboFieldToSelectAfterNextPageClick = txtFieldNameForInput.Text Then
            cboFieldToSelectAfterNextPageClick = ""
            bolFieldCleared = True
        End If
        
        If bolFieldCleared = True Then
            MsgBox "I have reset drop-downs that had this field selected as" & vbCrLf & "'Field to select AFTER...' the [Fields] tab.", vbInformation, "Field to select AFTER... cleared"
            'Update the changes to the Application "FieldToSelectAfter..." fields.
            cmdApplicationUpdate_Click
        End If
        
    End If
    
End Sub

Private Sub cmdApplicationAdd_Click()

    'Make sure RouteMaxCount is NOT blank
    If Trim(txtRouteMaxCount = "") Then
        txtRouteMaxCount = "3"
        MsgBox "Setting the 'Route Batch to Supervisor' to the Default value of '3'", vbOKOnly, "Default Route Max"
    End If
    
    '*** Declarations
1    Dim CatDB As ADOX.Catalog
2    Dim TabDB As ADOX.Table
3    Dim IdxDB As ADOX.Index
    
4    Dim intIndex As Integer

5    Dim con As ADODB.Connection
6    Dim rs As ADODB.Recordset
7    Dim ssql As String
8    Dim cmd As ADODB.Command
    
    
9    Set con = New ADODB.Connection
10    Set rs = New ADODB.Recordset
11    Set cmd = New ADODB.Command
    
12    con.mode = adModeReadWrite
      con.CommandTimeout = 120
      con.CommandTimeout = 600
      
13    rs.LOCKTYPE = adLockOptimistic
    
14    con.Open RegImaging101ConnectionString
    
    ' Begin Transaction
15    con.BeginTrans
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Applications where ApplicationName = '" & txtApplicationName & "'"
    rs.Open ssql, con
   
    ' Check if application already exists
    On Error Resume Next
    
    
100    If rs.BOF Then
    
            On Error GoTo ERROR_HANDLER
            
            '*************************************************
            '***  CREATE APPLICATION RECORD - BEGIN
101            rs.AddNew

            rs.Fields!ApplicationName = txtApplicationName
            rs.Fields!ApplicationDescription = txtApplicationDescription
            rs.Fields!ApplicationNotes = txtApplicationNotes
            rs.Fields!ApplicationIsActive = chkApplicationIsActive
            rs.Fields!ApplicationIsReadOnly = chkApplicationIsReadOnly
            rs.Fields!ApplicationCommitBatchTo = txtApplicationCommitBatchTo
            rs.Fields!ApplicationAutoAdvanceOnSeparator = chkAutoAdvanceOnSeparator
            rs.Fields!SetUserAsBatchOwnerOnSPLIT = chkSetUserAsBatchOwnerOnSPLIT
            rs.Fields!EnableSearchTemplates = chkEnableSearchTemplates
             '*** 2022-07-21 - ADD LogOpenedDocuments Field to I101Applications Table
            rs.Fields!LogOpenedDocuments = chkLogOpenedDocuments

            
            txtApplicationRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "ApplicationRECID")
            rs.Fields!ApplicationRECID = txtApplicationRECID
            
        'Save FTP Settings depending on selected SITE
        If cboSiteID = "Site A" Then
            rs.Fields!LookupDBConnectionString = txtLookupDBConnectionString
            rs.Fields!LookupDBTableName = cboLookupDBTableName
            rs.Fields!LookupDBTableIsOnSQLServer = chkLookupDBTableIsOnSQLServer
            rs.Fields!LookupDBWhereClause = txtLookupDBWhereClause
        Else
            rs.Fields!LookupDBConnectionString_B = txtLookupDBConnectionString
            rs.Fields!LookupDBTableName_B = cboLookupDBTableName
            rs.Fields!LookupDBTableIsOnSQLServer_B = chkLookupDBTableIsOnSQLServer
            rs.Fields!LookupDBWhereClause_B = txtLookupDBWhereClause
        End If

        rs.Fields!AutoLookupOnBatchLoad = chkAutoLookupOnBatchLoad

        rs.Fields!CaseIdCutoff = txtCaseIdCutoff
            
            rs.Fields!FieldToSelectAfterLookupClick = cboFieldToSelectAfterLookupClick
            rs.Fields!FieldToSelectAfterDocListClick = cboFieldToSelectAfterDocListClick
            rs.Fields!FieldToSelectAfterNextPageClick = cboFieldToSelectAfterNextPageClick
    
            rs.Fields!FieldToAssignDocumentGroup = cboFieldToAssignDocumentGroup
            rs.Fields!FieldToAssignDocumentType = cboFieldToAssignDocumentType
            rs.Fields!FieldToAssignDocumentSubType = cboFieldToAssignSubDocumentType
            
            rs.Fields!RootDirectoryPathForImageArchive = txtRootDirectoryPathForImageArchive
            rs.Fields!RootDirectoryPathForImageAnnotations = txtRootDirectoryPathForImageAnnotations
            rs.Fields!RootDirectoryPathForBatches = txtRootDirectoryPathForBatches
            rs.Fields!RootDirectoryPathForHtmlSource = txtRootDirectoryPathForHtmlSource
            
            rs.Fields!RouteMaxCount = txtRouteMaxCount
 
            rs.Fields!ApplicationBatchNameDelimiter = txtApplicationBatchNameDelimiter

            'Save FTP Settings depending on selected SITE
            If cboSiteID = "Site A" Then
                rs.Fields!FTPSite = txtFTPSite
                rs.Fields!ftpport = txtFTPPort
        
                rs.Fields!FTPUserID = txtFTPUserID
                rs.Fields!FTPPassword = txtFTPPassword
            Else
                rs.Fields!FTPSite_B = txtFTPSite
                rs.Fields!ftpport_B = txtFTPPort
        
                rs.Fields!FTPUserID_B = txtFTPUserID
                rs.Fields!FTPPassword_B = txtFTPPassword
            End If

            rs.Fields!MaxItemsToRetrieve = txtMaxItemsToRetrieve
            
            rs.Fields!ApplicationCommitBatchOption = txtApplicationCommitBatchOption
            
            rs.Fields!FTPFileNameField0 = cboFTPFileNameField(0)
            rs.Fields!FTPFileNameField1 = cboFTPFileNameField(1)
            rs.Fields!FTPFileNameField2 = cboFTPFileNameField(2)
            rs.Fields!FTPFileNameField3 = cboFTPFileNameField(3)
    
            rs.Fields!FTPFileNameDelimiter0 = Left(cboFTPFileNameDelimiter(0).Text, 1)
            rs.Fields!FTPFileNameDelimiter1 = Left(cboFTPFileNameDelimiter(1).Text, 1)
            rs.Fields!FTPFileNameDelimiter2 = Left(cboFTPFileNameDelimiter(2).Text, 1)
            
            
102            rs.Update
103            rs.Requery
104            rs.Close

            '***  CREATE APPLICATION RECORD - END
            '*************************************************
            
                '*************************************************
                '***  CREATE APPLICATION SQL TABLE - BEGIN
                
200                Set CatDB = New ADOX.Catalog
201                Set TabDB = New ADOX.Table
202                Set IdxDB = New ADOX.Index
                
                'open the database
203                CatDB.ActiveConnection = RegImaging101ConnectionString
                
                
204                With TabDB
                    .name = txtApplicationName   'set name
                    'add fields and specify datatype
                    .Columns.Append "DocumentRECID", adDouble
                    .Columns.Append "DocumentCommitDate", adDBTimeStamp
                    .Columns.Append "DocumentCommitUserID", adVarWChar, 50
                    .Columns.Append "DocumentNotes", adVarWChar, 250
                    .Columns.Append "DocumentLocked", adVarWChar, 2
                    .Columns.Append "DocumentLockedBy", adVarWChar, 50
                    .Columns.Append "DocumentLockedDate", adDBTimeStamp
                    .Columns.Append "DocumentLockExpDate", adDBTimeStamp
                    .Columns.Append "DocumentScanUserID", adVarWChar, 50
                    .Columns.Append "DocumentScanDate", adDBTimeStamp
                    .Columns.Append "DocumentIndexUserID", adVarWChar, 50
                    .Columns.Append "DocumentIndexDate", adDBTimeStamp
                    .Columns.Append "DocumentBatchRECID", adDouble
                    .Columns.Append "DocumentBatchName", adVarWChar, 250
                    .Columns.Append "DocumentImages", adInteger
                    .Columns.Append "DocumentPages", adInteger
                    .Columns.Append "BatchBoxNumber", adVarWChar, 50
                    
                    'set Allow Nulls = true
                    .Columns.item("DocumentCommitDate").Attributes = adColNullable
                    .Columns.item("DocumentCommitUserID").Attributes = adColNullable
                    .Columns.item("DocumentNotes").Attributes = adColNullable
                    .Columns.item("DocumentLocked").Attributes = adColNullable
                    .Columns.item("DocumentLockedBy").Attributes = adColNullable
                    .Columns.item("DocumentLockedDate").Attributes = adColNullable
                    .Columns.item("DocumentLockExpDate").Attributes = adColNullable
                    .Columns.item("DocumentScanUserID").Attributes = adColNullable
                    .Columns.item("DocumentScanDate").Attributes = adColNullable
                    .Columns.item("DocumentIndexUserID").Attributes = adColNullable
                    .Columns.item("DocumentIndexDate").Attributes = adColNullable
                    .Columns.item("DocumentBatchRECID").Attributes = adColNullable
                    .Columns.item("DocumentBatchName").Attributes = adColNullable
                    .Columns.item("DocumentImages").Attributes = adColNullable
                    .Columns.item("DocumentPages").Attributes = adColNullable
                    .Columns.item("BatchBoxNumber").Attributes = adColNullable
                    
                End With
                'add the table to database
                CatDB.Tables.Append TabDB
                
                Debug.Print "IMAGING DOCUMENT Table created..." & TabDB.name
                
                '*** CREATE INDEXES
                For intIndex = 0 To TabDB.Columns.Count - 1
                    'Only create index if the fieldname does not contain the word "notes"
                    If InStr(UCase(TabDB.Columns.item(intIndex).name), "NOTES") <= 0 Then
                        Set IdxDB = New ADOX.Index
                        IdxDB.name = TabDB.Columns.item(intIndex).name & "I"
                        IdxDB.Columns.Append TabDB.Columns.item(intIndex).name
                        ' Append the index to the table
                        TabDB.Indexes.Append IdxDB
                        Debug.Print "IMAGING Application Index Created... " & IdxDB.name
                        Set IdxDB = Nothing
                    End If
                Next
            
                '***  CREATE APPLICATION SQL TABLE - END
                '*************************************************
    
    
                '*************************************************
                '***  CREATE APPLICATION DETAIL SQL TABLE - BEGIN
                
                Set CatDB = New ADOX.Catalog
                Set TabDB = New ADOX.Table
                'open the database
                CatDB.ActiveConnection = RegImaging101ConnectionString
                
                'create new table object
                With TabDB
                    .name = txtApplicationName & "_Detail"   'set name
                    'add fields and specify datatype
                    .Columns.Append "DetailRECID", adDouble
                    .Columns.Append "DocumentRECID", adDouble
                    .Columns.Append "DetailOrder", adDouble
                    .Columns.Append "DetailCreatedDate", adDBTimeStamp
                    .Columns.Append "DetailSubdirectory", adVarWChar, 250
                    .Columns.Append "DetailFileName", adVarWChar, 250
                    .Columns.Append "DetailFileType", adVarWChar, 30
                    .Columns.Append "DetailRotation", adInteger
                End With
                'add the table to database
                CatDB.Tables.Append TabDB
                
                Debug.Print "IMAGING DETAIL Table created...", TabDB.name
            
                '*** CREATE INDEXES
                For intIndex = 0 To TabDB.Columns.Count - 1
                    'Only create index if the fieldname does not contain the word "notes"
                    If InStr(UCase(TabDB.Columns.item(intIndex).name), "NOTES") <= 0 Then
                        Set IdxDB = New ADOX.Index
                        IdxDB.name = TabDB.Columns.item(intIndex).name & "I"
                        IdxDB.Columns.Append TabDB.Columns.item(intIndex).name
                        ' Append the index to the table
                        TabDB.Indexes.Append IdxDB
                        Debug.Print "IMAGING Application Index Created... " & IdxDB.name
                        Set IdxDB = Nothing
                    End If
                Next
                
                '***  CREATE APPLICATION DETAIL SQL TABLE - END
                '*************************************************
    
    
                '*************************************************
                '***  CREATE BATCH PAGES SQL TABLE - BEGIN
                Set CatDB = New ADOX.Catalog
                Set TabDB = New ADOX.Table
                'open the database
                CatDB.ActiveConnection = RegImaging101BatchListConnectionString
                'create new table object
                With TabDB
                    .name = txtApplicationName & "_BatchPage"    'set name

                    'add Standard Batch fields and specify datatype
                    .Columns.Append "BatchRECID", adDouble
                    .Columns.Append "BatchPageRECID", adDouble
                    .Columns.Append "BatchPageFileName", adVarWChar, 250
                    .Columns.Append "BatchPageOrder", adDouble
                    .Columns.Append "BatchPageIndexed", adVarWChar, 2
                    .Columns.Append "BatchPageIsSeparator", adVarWChar, 2
                    .Columns.Append "BatchPageNote", adVarWChar, 250
                    .Columns.Append "BatchDocDesc", adVarWChar, 250
                    .Columns.Append "BatchPageStatus", adVarWChar, 250
                    .Columns.Append "BatchPageCommitDate", adDBTimeStamp
                    .Columns.Append "BatchPageCommitUser", adVarWChar, 50
                    .Columns.Append "BatchPageQCDate", adDBTimeStamp
                    .Columns.Append "BatchPageQCUser", adVarWChar, 50
                    .Columns.Append "BatchPageIndexDate", adDBTimeStamp
                    .Columns.Append "BatchPageIndexUser", adVarWChar, 50
                    .Columns.Append "BatchPagePageCount", adVarWChar, 50
                    .Columns.Append "BatchPageRotation", adInteger
                    .Columns.Append "CommitViaFTP", adVarWChar, 2
                   
                    'set Allow Nulls = true
                    .Columns.item("BatchPageFileName").Attributes = adColNullable
                    .Columns.item("BatchPageOrder").Attributes = adColNullable
                    .Columns.item("BatchPageIndexed").Attributes = adColNullable
                    .Columns.item("BatchPageIsSeparator").Attributes = adColNullable
                    .Columns.item("BatchPageIsSeparator").Attributes = adColNullable
                    .Columns.item("BatchPageNote").Attributes = adColNullable
                    .Columns.item("BatchDocDesc").Attributes = adColNullable
                    .Columns.item("BatchPageStatus").Attributes = adColNullable
                    .Columns.item("BatchPageCommitDate").Attributes = adColNullable
                    .Columns.item("BatchPageCommitUser").Attributes = adColNullable
                    .Columns.item("BatchPageQCDate").Attributes = adColNullable
                    .Columns.item("BatchPageQCUser").Attributes = adColNullable
                    .Columns.item("BatchPageIndexDate").Attributes = adColNullable
                    .Columns.item("BatchPageIndexUser").Attributes = adColNullable
                    .Columns.item("BatchPagePageCount").Attributes = adColNullable
                    .Columns.item("BatchPageRotation").Attributes = adColNullable
                    .Columns.item("CommitViaFTP").Attributes = adColNullable
                End With
                'add the table to database
                CatDB.Tables.Append TabDB
                DoEvents
                
                
                Debug.Print "BATCH Table created..." & TabDB.name
                DoEvents
                
                '*** CREATE INDEXES
                For intIndex = 0 To TabDB.Columns.Count - 1
                    'Only create index if the fieldname does not contain the word "notes"
                    If InStr(UCase(TabDB.Columns.item(intIndex).name), "NOTES") <= 0 Then
                        Set IdxDB = New ADOX.Index
                        IdxDB.name = TabDB.Columns.item(intIndex).name & "I"
                        IdxDB.Columns.Append TabDB.Columns.item(intIndex).name
                        ' Append the index to the table
                        TabDB.Indexes.Append IdxDB
                        Debug.Print "BATCH PAGES Index Created... " & IdxDB.name
                        Set IdxDB = Nothing
                    End If
                Next
                
    
                '***  CREATE BATCH PAGES SQL TABLE - END
                '*************************************************
                
                'Close connection and the recordset
                Set CatDB = Nothing
                Set TabDB = Nothing
                Set rs = Nothing
            
    '''''''           Next
    
    '''''''            '*************************************************
    '''''''            '***  OLD STYLE TOCREATE BATCH PAGES SQL TABLE INDEXES- BEGIN
    '''''''            '*** Create the INDEX for each Field by adding an "I" to the end of the fieldname
    '''''''            Conb.BeginTrans
    '''''''            ssqlb = "CREATE INDEX " & "BatchRECID" & "I" & " ON " & txtApplicationName & "_BatchPage"  & " (BatchRECID)"
    '''''''            rsb.Open ssqlb, Conb
    '''''''            '***  CREATE BATCH PAGES SQL TABLE INDEXES- END
    '''''''            '*************************************************
    
                cmdApplicationUpdate.Enabled = False
                cmdApplicationRemove.Enabled = True
            
            con.CommitTrans
        

    
        Else
            
            '*** APPLICATION ALREADY EXISTS
            con.RollbackTrans
            result = MsgBox("Sorry, This application already exists!", vbOKOnly, "Imaging101 - Add Application")
            
        End If
    
    'Commit Transaction
'    rs.Close
'    rsb.Close
    
    
'''''''        'Just a DUMMY LOOP to see if we can bypass an error caused by the DB not being
'''''''        '  ready after the Commit!
'''''''        For intDummyLoop = 1 To 2000000
'''''''            DoEvents
'''''''        Next
    
    con.Close
    Set con = Nothing
    
    'Refresh the Application List
    subLoadApplications
    
    'Find and Select the Application just added to begin adding fields
    For i = 0 To lstApplications.ListCount - 1
        If lstApplications.List(i) = txtApplicationName Then
            lstApplications.ListIndex = i
            lstApplications_Click
            Exit Sub
        End If
    Next
    
Exit Sub

ERROR_HANDLER:

    result = MsgBox("FRMCONFIG_ADD_APPLICATION_ERROR: " & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "* * * NOTE:  Application Names CANNOT begin with a NUMBER... " & vbCrLf & _
                                "* * * This is due to a SQL SERVER Limitation." & vbCrLf & vbCrLf & _
                                " - " & Err.Source & " [Line " & CStr(Erl) & "]", vbOKOnly, "Add Application Error")
    Err.Clear
    
    con.RollbackTrans
    con.Close
    Set con = Nothing


End Sub

Private Sub cmdApplicationClear_Click()

    txtApplicationName = ""
    txtApplicationDescription = ""
    txtApplicationNotes = ""
    txtApplicationCommitBatchTo = ""
    chkApplicationIsActive.Value = 0
    chkApplicationIsReadOnly.Value = 0
    chkAutoAdvanceOnSeparator.Value = 1
    chkSetUserAsBatchOwnerOnSPLIT.Value = 0
    
    txtApplicationName.Enabled = True
    cmdApplicationAdd.Enabled = True
    cmdApplicationRemove.Enabled = False
    cmdApplicationUpdate.Enabled = False
    
    txtRouteMaxCount = ""
    
    subLoadApplications
    txtApplicationName.SetFocus
    
End Sub

Private Sub cmdApplicationEditDoctypes_Click()

    frmDOCTYPES2.Show
    
End Sub

Private Sub cmdApplicationRemove_Click()
    Dim result As String
    result = InputBox("Type 'Remove ENTIRE Application INCLUDING Data'      you must type it with the proper case exactly as shown.   Click [OK] to proceed or Click [Cancel] if you DON'T want to remove it.", "Remove Application")
    If result = "Remove ENTIRE Application INCLUDING Data" Then
        result = MsgBox("Are you ABSOLUTELLY SURE you wish to REMOVE this Application?", vbYesNo)
        If result = vbYes Then
            result = MsgBox("LAST CHANCE!  Click [OK] to ZAP the Application " & vbCrLf & "including the DB Table, Fields, Indexes AND DOCUMENTS!" & vbCrLf & "or Click [Cancel] if you DON'T want to remove it.", vbOKCancel, "Last Chance")
            If result = vbOK Then
                '*******************************
                '*** ZAP Everything NOW!!!
                '*******************************
                
                '*** Declarations
                Dim CatDB As ADOX.Catalog
                Dim rs As ADODB.Recordset
                Dim con As ADODB.Connection
                Dim ssql As String
            
                Set con = New ADODB.Connection
                Set rs = New ADODB.Recordset
                
                con.mode = adModeReadWrite
                rs.LOCKTYPE = adLockPessimistic
                
                con.ConnectionString = RegImaging101ConnectionString
                con.ConnectionTimeout = 120
                con.CommandTimeout = 600
                
                con.Open RegImaging101ConnectionString
                
                'Begin SQL Transaction to make sure is doesn't zap the record if we
                '  get an error while deleting the DB TABLE
                con.BeginTrans
                
                'DELETE THE APPLICATION DATABASE TABLE
                    ssql = "DELETE from I101Applications where ApplicationRECID = " & txtApplicationRECID
                    rs.Open ssql, con
                
                        On Error Resume Next
                        
                        '*** DELETE THE APPLICATION DATABASE TABLE
                        Set CatDB = New ADOX.Catalog
                        CatDB.ActiveConnection = RegImaging101ConnectionString
                        CatDB.Tables.Delete (txtApplicationName)
                        Set CatDB = Nothing

                        '*** DELETE THE APPLICATION DETAIL DATABASE TABLE
                        Set CatDB = New ADOX.Catalog
                        CatDB.ActiveConnection = RegImaging101ConnectionString
                        CatDB.Tables.Delete (txtApplicationName & "_Detail")
                        Set CatDB = Nothing

                        '*** DELETE THE BATCH DATABASE TABLE
                        Set CatDB = New ADOX.Catalog
                        CatDB.ActiveConnection = RegImaging101BatchListConnectionString
                        CatDB.Tables.Delete (txtApplicationName & "_BatchPage")
                        Set CatDB = Nothing
                        
                        On Error GoTo 0

                ' Commit the Transaction
                If Err.Number = 0 Then
                    con.CommitTrans
                Else
                    con.RollbackTrans
                End If
                
                'Close connection and the recordset
''                rs.Close '*** Don't know why... but the DELETE automatically Closes the rs?
                Set CatDB = Nothing
                Set rs = Nothing
                con.Close
                Set con = Nothing
                
                cmdApplicationUpdate.Enabled = False
                cmdApplicationRemove.Enabled = True
                cmdApplicationClear_Click
                subLoadApplications
                
                cmdFieldsClear_Click
                lstFields.Clear

            End If
        End If
    End If
    
End Sub

Private Sub cmdApplicationUpdate_Click()

    Me.MousePointer = vbHourglass

    cmdApplicationUpdate.Enabled = False
    cmdApplicationRemove.Enabled = False

    'Make sure RouteMaxCount is NOT blank
    If Trim(txtRouteMaxCount = "") Then
        txtRouteMaxCount = "3"
        MsgBox "Setting the 'Route Batch to Supervisor' to the Default value of '3'", vbOKOnly, "Default Route Max"
    End If

    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con.mode = adModeReadWrite
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
    
    rs.LOCKTYPE = adLockOptimistic
    
    con.Open RegImaging101ConnectionString
    
    ' Begin Transaction
    con.BeginTrans
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Applications where ApplicationRECID = " & lstApplications.ItemData(lstApplications.ListIndex)
    rs.Open ssql, con
    
        rs.Fields!ApplicationName = txtApplicationName
        rs.Fields!ApplicationDescription = txtApplicationDescription
        rs.Fields!ApplicationNotes = txtApplicationNotes
        rs.Fields!ApplicationIsActive = chkApplicationIsActive
        rs.Fields!ApplicationIsReadOnly = chkApplicationIsReadOnly
        rs.Fields!ApplicationCommitBatchTo = txtApplicationCommitBatchTo
        rs.Fields!ApplicationAutoAdvanceOnSeparator = chkAutoAdvanceOnSeparator
        rs.Fields!SetUserAsBatchOwnerOnSPLIT = chkSetUserAsBatchOwnerOnSPLIT
        rs.Fields!EnableSearchTemplates = chkEnableSearchTemplates
         '*** 2022-07-21 - ADD LogOpenedDocuments Field to I101Applications Table
        rs.Fields!LogOpenedDocuments = chkLogOpenedDocuments

        'Save FTP Settings depending on selected SITE
        If cboSiteID = "Site A" Then
            rs.Fields!LookupDBConnectionString = txtLookupDBConnectionString
            rs.Fields!LookupDBTableName = cboLookupDBTableName
            rs.Fields!LookupDBTableIsOnSQLServer = chkLookupDBTableIsOnSQLServer
            rs.Fields!LookupDBWhereClause = txtLookupDBWhereClause
        Else
            rs.Fields!LookupDBConnectionString_B = txtLookupDBConnectionString
            rs.Fields!LookupDBTableName_B = cboLookupDBTableName
            rs.Fields!LookupDBTableIsOnSQLServer_B = chkLookupDBTableIsOnSQLServer
            rs.Fields!LookupDBWhereClause_B = txtLookupDBWhereClause
        End If

        rs.Fields!AutoLookupOnBatchLoad = chkAutoLookupOnBatchLoad
        
        rs.Fields!CaseIdCutoff = txtCaseIdCutoff
            
        rs.Fields!FieldToSelectAfterLookupClick = cboFieldToSelectAfterLookupClick
        rs.Fields!FieldToSelectAfterDocListClick = cboFieldToSelectAfterDocListClick
        rs.Fields!FieldToSelectAfterNextPageClick = cboFieldToSelectAfterNextPageClick
    
        rs.Fields!FieldToAssignDocumentGroup = cboFieldToAssignDocumentGroup
        rs.Fields!FieldToAssignDocumentType = cboFieldToAssignDocumentType
        rs.Fields!FieldToAssignDocumentSubType = cboFieldToAssignDocumentSubType
            
        rs.Fields!RootDirectoryPathForImageArchive = txtRootDirectoryPathForImageArchive
        rs.Fields!RootDirectoryPathForImageAnnotations = txtRootDirectoryPathForImageAnnotations
        rs.Fields!RootDirectoryPathForBatches = txtRootDirectoryPathForBatches
        rs.Fields!RootDirectoryPathForHtmlSource = txtRootDirectoryPathForHtmlSource
           
        rs.Fields!RouteMaxCount = txtRouteMaxCount
 
        rs.Fields!ApplicationBatchNameDelimiter = txtApplicationBatchNameDelimiter

        'Save FTP Settings depending on selected SITE
        If cboSiteID = "Site A" Then
            rs.Fields!FTPSite = txtFTPSite
            rs.Fields!ftpport = txtFTPPort
    
            rs.Fields!FTPUserID = txtFTPUserID
            rs.Fields!FTPPassword = txtFTPPassword
        Else
            rs.Fields!FTPSite_B = txtFTPSite
            rs.Fields!ftpport_B = txtFTPPort
    
            rs.Fields!FTPUserID_B = txtFTPUserID
            rs.Fields!FTPPassword_B = txtFTPPassword
        End If
    
        rs.Fields!MaxItemsToRetrieve = txtMaxItemsToRetrieve

        rs.Fields!ApplicationCommitBatchOption = txtApplicationCommitBatchOption
        
        rs.Fields!FTPFileNameField0 = cboFTPFileNameField(0)
        rs.Fields!FTPFileNameField1 = cboFTPFileNameField(1)
        rs.Fields!FTPFileNameField2 = cboFTPFileNameField(2)
        rs.Fields!FTPFileNameField3 = cboFTPFileNameField(3)

        rs.Fields!FTPFileNameDelimiter0 = Left(cboFTPFileNameDelimiter(0).Text, 1)
        rs.Fields!FTPFileNameDelimiter1 = Left(cboFTPFileNameDelimiter(1).Text, 1)
        rs.Fields!FTPFileNameDelimiter2 = Left(cboFTPFileNameDelimiter(2).Text, 1)

        
        
        rs.Update
            
        ' Commit Transaction
        con.CommitTrans
    
'        subLoadApplications
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    MsgBox "Application " & txtApplicationName & " Updated Successfully!", vbOKOnly
    cmdApplicationUpdate.Enabled = True
    cmdApplicationRemove.Enabled = True
    cmdApplicationUpdate.SetFocus
    Me.MousePointer = vbNormal

'    cmdApplicationUpdate.enabled = False
'    cmdApplicationRemove.enabled = True
'    'Clear Application Fields
'    cmdApplicationClear_Click
'    'Clear Fields Fields
'    cmdFieldsClear_Click
'    lstFields.Clear
'
'    subLoadApplications

End Sub







Private Sub cmdBackup_Click()

    If txtApplicationRECID = "" Then
        Exit Sub
    End If
    
    Dim dblBackupNumber As Double
    dblBackupNumber = InputBox("Enter Backup #", "Get TableLookup Backup Number", "1")
    dblBackupNumber = dblBackupNumber * 100000
    
        Dim rs As ADODB.Recordset
        Dim rsnew As ADODB.Recordset
        Dim con As ADODB.Connection
        Dim ssql As String
    
        
        Set con = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Set rsnew = New ADODB.Recordset
        
        con.mode = adModeReadWrite
        rs.LOCKTYPE = adLockOptimistic
        rsnew.LOCKTYPE = adLockOptimistic
        
        con.Open RegImaging101ConnectionString
        
        
            ssql = "Select * from I101TableLookupFields where ApplicationRECID = " & txtApplicationRECID
            rs.Open ssql, con
        
            rs.MoveFirst
            
        
            rsnew.Open "select *  from i101TableLookupFields where ApplicationRECID= " & txtApplicationRECID + dblBackupNumber, con
            
            If Not rsnew.EOF Then
                rsnew.MoveLast
'                result = MsgBox("A backup already exists for Application #" & ApplicationRECID & vbCrLf & "Do you want to OVERWRITE it?", vbYesNo, "Lookup Backup Exists")
                If result = vbYes Then
                    rsnew.Close
                    rsnew.Open "DELETE from i101TableLookupFields where ApplicationRECID =  " & txtApplicationRECID + dblBackupNumber, con
                    rsnew.Open "select  * from i101TableLookupFields where 1 = 2"
                    
                Else
                    Exit Sub
                End If
            End If
            
        While Not rs.EOF
            
            ' Begin Transaction
'            con.BeginTrans
            
            rsnew.AddNew
            
            For intIndex = 0 To rs.Fields.Count - 1
                
                Select Case rs.Fields(intIndex).name
                    Case "TableLookupRECID"
                            rsnew.Fields(intIndex) = rs.Fields(intIndex) + dblBackupNumber
                    Case "ApplicationRECID"
                            rsnew.Fields(intIndex) = rs.Fields(intIndex) + dblBackupNumber
                    Case Else
                            rsnew.Fields(intIndex) = rs.Fields(intIndex)
                    
                End Select
        
            Next
            
            rsnew.Update
            
            ' Commit Transaction
'            con.CommitTrans
        
            rs.MoveNext
        Wend
        
            rsnew.Close
        
        'Close connection and the recordset
        Set rs = Nothing
        con.Close
        Set con = Nothing


End Sub

Private Sub cmdCommandPrompt_Click()

    Call shelldoc("cmd")

End Sub

Private Sub cmdDBMaintenance_Click()
    ' Create each DB Maintenance Form as a NEW Instance
    '   to allow editing multiple tables simultaneously.
    Set frm101DBMaintConfig = New frm101DBMaint
    frm101DBMaintConfig.Show
End Sub



Private Sub cmdEditSearchTemplate_Click()

    frmImaging101SearchTemplate.Show

End Sub

Private Sub cmdFieldAdd_Click()

    If ckbFieldSplitBatches = vbChecked Then
        result = MsgBox("You have flagged Field < " & txtFieldName & " >" & vbCrLf & _
                        "as 'Split Batches on This Field'." & vbCrLf & _
                        "Only ONE Field per Application can be set to Split on." & vbCrLf & _
                        "   ARE YOU SURE? ", vbYesNo, "Set Split Field Confirmation")
        If result = vbNo Then
            MsgBox "Changes NOT Saved!", vbOKOnly
            Exit Sub
        End If
    End If

    If ckbFieldRouteToBatchQueue = vbChecked Then
        result = MsgBox("You have flagged Field < " & txtFieldName & " >" & vbCrLf & _
                        "as 'Route to Batch Queue based on This Field'." & vbCrLf & _
                        "Only ONE Field per Application can be set to ROUTE on." & vbCrLf & _
                        "   ARE YOU SURE? ", vbYesNo, "Set ROUTE Field Confirmation")
        If result = vbNo Then
            MsgBox "Changes NOT Saved!", vbOKOnly
            Exit Sub
        End If
    End If
    
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Dim rsf As ADODB.Recordset
    Dim Conf As ADODB.Connection
    Dim ssqlf As String

    Dim rsb As ADODB.Recordset
    Dim Conb As ADODB.Connection
    Dim ssqlb As String

    '*** Set Object Types
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
            
    Set Conf = New ADODB.Connection
    Set rsf = New ADODB.Recordset
            
    Set Conb = New ADODB.Connection
    Set rsb = New ADODB.Recordset
            
    '*** Set Connection Modes
    con.mode = adModeReadWrite
    Conf.mode = adModeReadWrite
    Conb.mode = adModeReadWrite
    
    '*** Set Lock Types
    rs.LOCKTYPE = adLockOptimistic
    rsf.LOCKTYPE = adLockOptimistic
    rsb.LOCKTYPE = adLockOptimistic
            
    '*** Set Connection Timeouts
    con.ConnectionTimeout = 120
    Conf.ConnectionTimeout = 120
    Conb.ConnectionTimeout = 120
    
    con.CommandTimeout = 600
    Conf.CommandTimeout = 600
    Conb.CommandTimeout = 600
    
    '*** OPEN Connections
    con.Open RegImaging101ConnectionString
    Conf.Open RegImaging101ConnectionString
    Conb.Open RegImaging101BatchListConnectionString
    
     'Begin SQL Transaction to make sure is doesn't zap the record if we
    '  get an error while deleting the DB TABLE
    con.BeginTrans
    Conf.BeginTrans
    Conb.BeginTrans
            
   'sql statement to select items on the drop down list
    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID & " AND FieldName = '" & txtFieldName & "'"
    rs.Open ssql, con
   
    

    On Error Resume Next
'    rs.MoveNext
    
    
    ' Check if FIELD already exists
'    If rs.Fields!FieldsRECID <> txtFieldsRECID Or rs.BOF Then
    If rs.RecordCount < 1 Then
        
        On Error GoTo ADD_FIELD_ERROR

        rs.AddNew
        txtFieldsRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "FieldsRECID")
        rs!FieldsRECID = txtFieldsRECID
        rs!ApplicationRECID = txtApplicationRECID
        rs!FieldName = txtFieldName
        rs!FieldNameForInput = txtFieldNameForInput
        rs!FieldNameForOutput = txtFieldNameForOutput
        rs!FieldDescription = txtFieldDescription
        rs!FieldType = cboFieldType.Text
        If txtFieldSize = "" Then txtFieldSize = 0
        rs!FieldSize = txtFieldSize
        rs!FieldFormat = txtFieldFormat
        rs!FieldMask = txtFieldMask
        rs!FieldLowValue = txtFieldLowValue
        rs!FieldHighValue = txtFieldHighValue
        rs!FieldDefaultValue = cboFieldDefaultValue
        rs!FieldAutoIncrField = ckbFieldAutoIncrField
        rs!FieldIsRequiredForCommit = ckbFieldIsRequiredForCommit
        rs!FieldIsRequiredForSplit = ckbFieldIsRequiredForSplit
        rs!FieldSplitBatches = ckbFieldSplitBatches
        
        rs!FieldRouteToBatchQueue = ckbFieldRouteToBatchQueue
        rs!FieldRouteToBatchUser = ckbFieldRouteToBatchUser
        rs!FieldRouteToBatchManager = ckbFieldRouteToBatchManager
        
        rs!FieldIsForOutputOnly = ckbFieldIsForOutputOnly
        rs!HideForSearchIndex = ckbHideForSearchIndex
        
        rs!FieldIsSticky = ckbFieldIsSticky
        rs!FieldDropDownList = ckbFieldDropDownList
        rs!FieldDropDownListAlsoOnFiler = ckbFieldDropDownListAlsoOnFiler
        rs!FieldDefaultForBarcodeOnly = chkFieldDefaultForBarcodeOnly
        '* ListCount is usually one more than the ListIndex of the last item.
        rs!FieldOrderBatch = lstFields.ListCount
        rs!FieldOrderDisplay = lstFields.ListCount
        rs!FieldSearchCondition = cboFieldSearchCondition
        rs!FieldTableLookupOverridesDefault = ckbFieldTableLookupOverridesDefault
        
        rs.Update
        
        '************************************************************************
        '***  CREATE SQL FIELD - BEGIN
        ssqlf = "ALTER TABLE " & txtApplicationName & " ADD " & txtFieldName & " "
        Select Case cboFieldType
            Case "Text"
                ssqlf = ssqlf & " nvarchar(" & txtFieldSize & ") "
            Case "LongText"
                ssqlf = ssqlf & " nvarchar(" & txtFieldSize & ") "
            Case "Boolean"
                ssqlf = ssqlf & " bit "
            Case "Numeric"
                ssqlf = ssqlf & " float "
            Case "Date"
                ssqlf = ssqlf & " datetime "
            Case "Currency"
                ssqlf = ssqlf & " money "
        End Select
        ssqlf = ssqlf & " NULL"

        rsf.Open ssqlf, Conf
        

        '*** 2020-06-10 - Jacob - MOVED code for ONLY ONE FIELD to AFTER the COMMIT
         
        
        '******************************************************************************
        ' Create the INDEX for each Field by adding an "I" to the end of the fieldname
        ssqlf = "CREATE INDEX " & txtFieldName & "I" & " ON " & txtApplicationName & " (" & txtFieldName & ")"
        rsf.Open ssqlf, Conf
        

        '************************************************************************
        '*** Create BATCH SQL Field
        ssqlb = "ALTER TABLE " & txtApplicationName & "_BatchPage" & " ADD " & txtFieldName & " "
        Select Case cboFieldType
            Case "Text"
                ssqlb = ssqlb & " nvarchar(" & txtFieldSize & ") "
            Case "LongText"
                ssqlb = ssqlb & " nvarchar(" & txtFieldSize & ") "
            Case "Boolean"
                ssqlb = ssqlb & " bit "
            Case "Numeric"
                ssqlb = ssqlb & " float "
            Case "Date"
                ssqlb = ssqlb & " datetime "
            Case "Currency"
                ssqlb = ssqlb & " money "
        End Select
        ssqlb = ssqlb & " NULL"
        
        rsb.Open ssqlb, Conb
            
''' I DON'T THINK WE NEED INDEXES ON THE BATCH FIELDS!!!
'''                ' Create the INDEX for each Field by adding an "I" to the end of the fieldname
'''                ssqlb = "CREATE INDEX " & txtFieldName & "I" & " ON " & txtApplicationName & "_BatchPage"  & " (" & txtFieldName & ")"
'''                rsb.Open ssqlb, Conb
            
            '***  CREATE SQL FIELD - END

        
        con.CommitTrans
        Conf.CommitTrans
        Conb.CommitTrans
        
        
        '*** 2020-06-10 - Jacob - MOVED code for ONLY ONE FIELD to AFTER the COMMIT and replaced with funcSaveFieldToDB
        '************************************************************************
        '*** Make Sure that Only ONE Field is flagged as
        '***    for each application.
        If ckbFieldSplitBatches = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldSplitBatches", 0)
        End If
        
        If ckbFieldRouteToBatchQueue = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldRouteToBatchQueue", 0)
        End If
        
        If ckbFieldRouteToBatchUser = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldRouteToBatchUser", 0)
        End If

        If ckbFieldRouteToBatchManager = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldRouteToBatchManager", 0)
        End If
        
        cmdFieldUpdate.Enabled = False
''        cmdFieldRemove.Enabled = False
        cmdFieldsClear_Click
        subLoadFields
       
        
    Else
        result = MsgBox("Sorry, Field '" & txtFieldName & "' already exists!", vbOKOnly, "Imaging101 - Add Field")
        
    End If
    
    'Close connection and the recordset
    'Again... for some reason the rs was already closed???
'    rs.Close
    Set rs = Nothing
    Set rsf = Nothing
    Set rsb = Nothing
    con.Close
    Set con = Nothing
    Set Conf = Nothing
    Set Conb = Nothing
    
    '* Refresh Fields List
    subLoadFields
    
    Exit Sub

ADD_FIELD_ERROR:
    con.RollbackTrans
    Conf.RollbackTrans
    Conb.RollbackTrans
    Set rs = Nothing
    Set rsb = Nothing
    Set rsf = Nothing
    con.Close
    Set con = Nothing
    Set Conf = Nothing
    Set Conb = Nothing
    
    MsgBox "ADD-FIELD ERROR: " & Err.Number & " - " & Err.Description & "[ TRANSACTION ROLLED BACK, Field NOT Added! ]"
    

End Sub

Private Sub cmdFieldDelete_Click()
    Dim result As String
    
    result = MsgBox("Are you Sure you wish to DELETE Field [" & txtFieldName & "]?" & vbCrLf & "This will delete ALL DATA for THIS FIELD and CANNOT be Undone!", vbYesNo, "Delete Field")
    If result <> vbYes Then
        MsgBox "Delete Cancelled...", vbOKOnly, "Delete Cancelled"
        Exit Sub
    End If
    
    result = MsgBox("LAST CHANCE..." & vbCrLf & "Are you Sure you wish to DELETE Field [" & txtFieldName & "]?" & vbCrLf & "You DO undestand this CANNOT be Undone!", vbYesNo, "Delete Field - LAST CHANCE")
    If result <> vbYes Then
        MsgBox "Delete Cancelled...", vbOKOnly, "Delete Cancelled"
        Exit Sub
    End If
    
    '*******************************
    '*** ZAP Field NOW!!!
    '*******************************
    
   On Error GoTo ADD_FIELD_ERROR
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Dim rsf As ADODB.Recordset
    Dim Conf As ADODB.Connection
    Dim ssqlf As String

    Dim rsb As ADODB.Recordset
    Dim Conb As ADODB.Connection
    Dim ssqlb As String

    '*** Set Object Types
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
            
    Set Conf = New ADODB.Connection
    Set rsf = New ADODB.Recordset
            
    Set Conb = New ADODB.Connection
    Set rsb = New ADODB.Recordset
            
    '*** Set Connection Modes
    con.mode = adModeReadWrite
    Conf.mode = adModeReadWrite
    Conb.mode = adModeReadWrite
    
    '*** Set Lock Types
    rs.LOCKTYPE = adLockOptimistic
    rsf.LOCKTYPE = adLockOptimistic
    rsb.LOCKTYPE = adLockOptimistic
            
    '*** OPEN Connections
    con.Open RegImaging101ConnectionString
    Conf.Open RegImaging101ConnectionString
    Conb.Open RegImaging101BatchListConnectionString
            
    'Begin SQL Transaction to make sure is doesn't zap the record if we
    '  get an error while deleting the DB TABLE
    con.BeginTrans
    Conf.BeginTrans
    Conb.BeginTrans
            
    
    'sql statement to select items on the drop down list
    ssql = "DELETE from I101Fields where FieldsRECID = " & txtFieldsRECID
    rs.Open ssql, con
    
    '***  DROP COLUMN FROM IMAGING TABLE - BEGIN
        ' DELETE the INDEX before Dropping the Field
        ssqlf = "DROP INDEX " & txtApplicationName & "." & txtFieldName & "I"
        rsf.Open ssqlf, con
        
        'sql statement
        ssqlf = "ALTER TABLE " & txtApplicationName & " DROP COLUMN " & txtFieldName
        rsf.Open ssqlf, con
    '***  DROP COLUMN - END

        '***  DROP COLUMN FROM BATCH TABLE - BEGIN
            
''' DID NOT CREATE INDEX FOR FIELDS ON BATCH TABLE
'''                    ' Create the INDEX for each Field by adding an "I" to the end of the fieldname
'''                    ssqlf = "DROP INDEX " & txtFieldName & "I" & " ON " & txtApplicationName & "_BatchPage"
'''                    rsf.Open ssqlf, Con
            
            'sql statement
            ssqlb = "ALTER TABLE " & txtApplicationName & "_BatchPage" & " DROP COLUMN " & txtFieldName
            rsb.Open ssqlb, Conb
        '***  DROP COLUMN - END
        
        
        
        con.CommitTrans
        Conf.CommitTrans
        Conb.CommitTrans

        'Close connection and the recordset
        'Again... for some reason the rs was already closed???
    '    rs.Close
        Set rs = Nothing
        Set rsf = Nothing
        Set rsb = Nothing
        con.Close
        Set con = Nothing
        Set Conf = Nothing
        Set Conb = Nothing
        
        cmdFieldUpdate.Enabled = False
        cmdFieldDelete.Enabled = True
        cmdFieldsClear_Click
        subLoadFields

    
Exit Sub

ADD_FIELD_ERROR:
    con.RollbackTrans
    Conf.RollbackTrans
    Conb.RollbackTrans
    Set rs = Nothing
    Set rsb = Nothing
    Set rsf = Nothing
    con.Close
    Set con = Nothing
    Set Conf = Nothing
    Set Conb = Nothing
    
    MsgBox "DELETE-FIELD ERROR: " & Err.Number & " - " & Err.Description & "[ TRANSACTION ROLLED BACK, Field NOT Deleted! ]"
    

End Sub


Private Sub cmdFieldsClear_Click()
 ' Clear Fields
        txtFieldsRECID = 0
        txtFieldName = ""
        txtFieldNameForInput = ""
        txtFieldNameForOutput = ""
        txtFieldDescription = ""
        cboFieldType.Text = "Text"
        txtFieldSize = "10"
        txtFieldFormat = ""
        txtFieldMask = ""
        cboFieldDefaultValue = ""
        txtFieldHighValue = ""
        txtFieldLowValue = ""
        cboFieldDefaultValue = ""
        ckbFieldAutoIncrField = "0"
        ckbFieldTableLookupOverridesDefault = "1"
        ckbFieldIsRequiredForCommit = "0"
        ckbFieldIsRequiredForSplit = "0"
        ckbFieldSplitBatches = "0"
        ckbFieldRouteToBatchQueue = "0"
        ckbFieldIsSticky = "1"
        ckbFieldDropDownList = "0"
        
        cmdFieldMoveUp.Enabled = False
        cmdFieldMoveDown.Enabled = False
       
        txtFieldName.Enabled = True
        txtFieldName.SetFocus
 
End Sub

Private Sub cmdFieldMoveDown_Click()
    
    cmdFieldMoveUp.Enabled = False
    cmdFieldMoveDown.Enabled = False

    Dim strHoldItemText As String
    Dim strHoldItemData As String
    Dim intHoldListIndex As Integer
    Dim intNewListIndex As Integer
    
    intHoldListIndex = lstFields.ListIndex
    lstFields.SetFocus
    lstFields.Selected(intHoldListIndex) = True
    
    If lstFields.ListIndex < lstFields.ListCount - 1 Then
        ' Hold variables
        strHoldItemText = lstFields.Text
        strHoldItemData = lstFields.ItemData(lstFields.ListIndex)
        ' Move it
        lstFields.RemoveItem (lstFields.ListIndex)
        intNewListIndex = intHoldListIndex + 1
        lstFields.AddItem strHoldItemText, intNewListIndex
        lstFields.ItemData(intNewListIndex) = strHoldItemData
        ' Re-focus
        lstFields.SetFocus
        lstFields.Selected(intNewListIndex) = True
        ' Update
        
        '*** Declarations
        Dim rs As ADODB.Recordset
        Dim con As ADODB.Connection
        Dim ssql As String
    
        Set con = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        con.mode = adModeReadWrite
        rs.LOCKTYPE = adLockOptimistic
        
        con.Open RegImaging101ConnectionString
        
        ' Begin Transaction
        con.BeginTrans
        
        For intIndex = 0 To lstFields.ListCount - 1
            'sql statement to select items on the drop down list
            ssql = "Select * from I101Fields where FieldsRECID = " & lstFields.ItemData(intIndex)
            rs.Open ssql, con
            
            rs!FieldOrderBatch = intIndex
            rs!FieldOrderDisplay = intIndex
    
            rs.Update
            rs.Close
        Next
        
        ' Commit Transaction
        con.CommitTrans
        
        'Close connection and the recordset
        Set rs = Nothing
        con.Close
        Set con = Nothing
           
        ' Re-focus
        lstFields.SetFocus
        lstFields.Selected(intNewListIndex) = True
            
        
    End If
    
    cmdFieldMoveUp.Enabled = True
    cmdFieldMoveDown.Enabled = True

End Sub
Private Sub cmdFieldMoveUp_Click()
    
    cmdFieldMoveUp.Enabled = False
    cmdFieldMoveDown.Enabled = False
    
    Dim strHoldItemText As String
    Dim strHoldItemData As String
    Dim intHoldListIndex As Integer
    Dim intNewListIndex As Integer
    
    intHoldListIndex = lstFields.ListIndex
    lstFields.SetFocus
    lstFields.Selected(intHoldListIndex) = True
    
    If lstFields.ListIndex > 0 Then
        ' Hold variables
        strHoldItemText = lstFields.Text
        strHoldItemData = lstFields.ItemData(lstFields.ListIndex)
        ' Move it
        lstFields.RemoveItem (lstFields.ListIndex)
        intNewListIndex = intHoldListIndex - 1
        lstFields.AddItem strHoldItemText, intNewListIndex
        lstFields.ItemData(intNewListIndex) = strHoldItemData
         ' Update
       
        '*** Declarations
        Dim rs As ADODB.Recordset
        Dim con As ADODB.Connection
        Dim ssql As String
    
        Set con = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        con.mode = adModeReadWrite
        rs.LOCKTYPE = adLockOptimistic
        
        con.Open RegImaging101ConnectionString
        
        ' Begin Transaction
        con.BeginTrans
        
        For intIndex = 0 To lstFields.ListCount - 1
            'sql statement to select items on the drop down list
            ssql = "Select * from I101Fields where FieldsRECID = " & lstFields.ItemData(intIndex)
            rs.Open ssql, con
            
            rs!FieldOrderBatch = intIndex
            rs!FieldOrderDisplay = intIndex
    
            rs.Update
            rs.Close
        Next
        
        ' Commit Transaction
        con.CommitTrans
        
        'Close connection and the recordset
        Set rs = Nothing
        con.Close
        Set con = Nothing
        
         ' Re-focus
        lstFields.SetFocus
        lstFields.Selected(intNewListIndex) = True
            
          
    End If
    
    cmdFieldMoveUp.Enabled = True
    cmdFieldMoveDown.Enabled = True
    
End Sub

Private Sub cmdFieldUpdate_Click()
    
    On Error GoTo ERROR_TRAP
    
    If ckbFieldSplitBatches = vbChecked Then
        result = MsgBox("You have flagged Field < " & txtFieldName & " >" & vbCrLf & _
                        "as 'Split Batches on This Field'." & vbCrLf & _
                        "Only ONE Field per Application can be set to ROUTE on." & vbCrLf & _
                        "   ARE YOU SURE? ", vbYesNo, "Set SPLIT Field Confirmation")
        If result = vbNo Then
            MsgBox "Changes NOT Saved!", vbOKOnly
            Exit Sub
        End If
    End If

    If ckbFieldRouteToBatchQueue = vbChecked Then
        result = MsgBox("You have flagged Field < " & txtFieldName & " >" & vbCrLf & _
                        "as 'Route to Batch Queue based on This Field'." & vbCrLf & _
                        "Only ONE Field per Application can be set to ROUTE on." & vbCrLf & _
                        "   ARE YOU SURE? ", vbYesNo, "Set ROUTE Field Confirmation")
        If result = vbNo Then
            MsgBox "Changes NOT Saved!", vbOKOnly
            Exit Sub
        End If
    End If
    
    If cboFieldType.Text <> txtFieldTypeHOLD Then
        MsgBox "YOU HAVE CHOSEN TO CHANGE THE FIELD TYPE! " & vbCrLf & vbCrLf _
            & "Please keep in mind that I will attempt to convert the existing values " & vbCrLf _
            & "from [" & txtFieldTypeHOLD & "] to [" & cboFieldType.Text & "]..." & vbCrLf & vbCrLf _
            & "However, it is possible that the Data contained in the " & vbCrLf _
            & "[" & txtFieldName & "] have values that cannot be converted." & vbCrLf & vbCrLf _
            & "If this is the case, the Update will fail.", vbInformation
    End If
        
    
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con.mode = adModeReadWrite
    rs.LOCKTYPE = adLockOptimistic
    
    con.Open RegImaging101ConnectionString
    
    ' SETUP FOR FIELD COLUMNS
    Dim rsf As ADODB.Recordset
    Dim Conf As ADODB.Connection
    Dim ssqlf As String

    Dim rsb As ADODB.Recordset
    Dim Conb As ADODB.Connection
    Dim ssqlb As String
    
    Set Conf = New ADODB.Connection
    Set rsf = New ADODB.Recordset
    
    Set Conb = New ADODB.Connection
    Set rsb = New ADODB.Recordset
                    
    Conf.mode = adModeReadWrite
    Conb.mode = adModeReadWrite
    
    rsf.LOCKTYPE = adLockOptimistic
    rsb.LOCKTYPE = adLockOptimistic
    
    Conf.Open RegImaging101ConnectionString
    Conb.Open RegImaging101BatchListConnectionString
    
    'Begin SQL Transaction to make sure is doesn't zap the record if we
    '  get an error while deleting the DB TABLE
    con.BeginTrans
    Conf.BeginTrans
    Conb.BeginTrans
    
    'sql statement to select items on the drop down list

        '*** 2020-06-10 - Jacob - MOVED code for ONLY ONE FIELD to AFTER the COMMIT and replaced with funcSaveFieldToDB

        
        If (cboFieldType.DataChanged = True) Or (txtFieldSize.DataChanged = True) Then
        
            
            '*** DROP THE INDEX ON THIS FIELD TO PREVENT ERRORS.
            ssqlf = "DROP INDEX " & txtApplicationName & "." & txtFieldName & "I"
            txtActionBeforeError = "Open ssqlf - " & ssqlf
            rsf.Open ssqlf, Conf
            
            
            
                ssqlf = ""
                ssqlb = ""
                '***  ALTER COLUMN - BEGIN
                    'sql statement
                    If cboFieldType.Text = "Text" Or cboFieldType.Text = "Notes" Then
                        ssqlf = "ALTER TABLE " & txtApplicationName & " ALTER COLUMN " & txtFieldName & " "
                        ssqlb = "ALTER TABLE " & txtApplicationName & "_BatchPage" & " ALTER COLUMN " & txtFieldName & " "
                    Else
                        ssqlf = "ALTER TABLE " & txtApplicationName & " ALTER COLUMN " & txtFieldName
                        ssqlb = "ALTER TABLE " & txtApplicationName & "_BatchPage" & " ALTER COLUMN " & txtFieldName
                    End If
                    
                Select Case cboFieldType
                    Case "Text"
                        ssqlf = ssqlf & " nvarchar(" & txtFieldSize & ") "
                        ssqlb = ssqlb & " nvarchar(" & txtFieldSize & ") "
                    Case "LongText"
                        ssqlf = ssqlf & " nvarchar(" & txtFieldSize & ") "
                        ssqlb = ssqlb & " nvarchar(" & txtFieldSize & ") "
                    Case "Boolean"
                        ssqlf = ssqlf & " bit "
                        ssqlb = ssqlb & " bit "
                    Case "Numeric"
                        ssqlf = ssqlf & " float "
                        ssqlb = ssqlb & " float "
                    Case "Date"
                        ssqlf = ssqlf & " datetime "
                        ssqlb = ssqlb & " datetime "
                    Case "Currency"
                        ssqlf = ssqlf & " money "
                        ssqlb = ssqlb & " money "
                End Select
            
            ssqlf = ssqlf & " NULL"
            ssqlb = ssqlb & " NULL"
                    
                    txtActionBeforeError = "Open ssqlf - " & ssqlf
                    rsf.Open ssqlf, Conf
                    
                    txtActionBeforeError = "Open ssqlb - " & ssqlb
                    rsb.Open ssqlb, Conb
                    
                '***  ALTER COLUMN - END
            
            
            ' RE-CREATE the INDEX for the modified Field by adding an "I" to the end of the fieldname
            ssqlf = "CREATE INDEX " & txtFieldName & "I" & " ON " & txtApplicationName & " (" & txtFieldName & ")"
            txtActionBeforeError = "Open ssqlf - " & ssqlf
            rsf.Open ssqlf, Conf
        
        End If
        
        
    '******************************************************************************
    '*** UPDATE the I101Fields Table
    
    ssql = "Select * from I101Fields where FieldsRECID = " & txtFieldsRECID
    rs.Open ssql, con
        
        rs!FieldsRECID = txtFieldsRECID
        rs!ApplicationRECID = txtApplicationRECID
        rs!FieldName = txtFieldName
        rs!FieldNameForInput = txtFieldNameForInput
        rs!FieldNameForOutput = txtFieldNameForOutput
        rs!FieldDescription = txtFieldDescription
        
        rs!FieldFormat = txtFieldFormat
        rs!FieldMask = txtFieldMask
        rs!FieldLowValue = txtFieldLowValue
        rs!FieldHighValue = txtFieldHighValue
        
        If Trim(cboFieldDefaultValue) = "" Then
            rs!FieldDefaultValue = vbNullString
        Else
            rs!FieldDefaultValue = cboFieldDefaultValue
        End If
        
        rs!FieldAutoIncrField = ckbFieldAutoIncrField
        rs!FieldIsRequiredForCommit = ckbFieldIsRequiredForCommit
        rs!FieldIsRequiredForSplit = ckbFieldIsRequiredForSplit
        
        rs!FieldSplitBatches = ckbFieldSplitBatches
        rs!FieldRouteToBatchQueue = ckbFieldRouteToBatchQueue
        rs!FieldRouteToBatchUser = ckbFieldRouteToBatchUser
        rs!FieldRouteToBatchManager = ckbFieldRouteToBatchManager
        
        rs!FieldIsForOutputOnly = ckbFieldIsForOutputOnly
        rs!HideForSearchIndex = ckbHideForSearchIndex
                
        rs!FieldIsSticky = ckbFieldIsSticky
        rs!FieldDropDownList = ckbFieldDropDownList
        rs!FieldDropDownListAlsoOnFiler = ckbFieldDropDownListAlsoOnFiler
        rs!FieldDefaultForBarcodeOnly = chkFieldDefaultForBarcodeOnly
        rs!FieldTableLookupOverridesDefault = ckbFieldTableLookupOverridesDefault

        rs!FieldOrderBatch = lstFields.ListIndex
        rs!FieldOrderDisplay = lstFields.ListIndex

        rs!FieldSearchCondition = cboFieldSearchCondition
        

        
        rs!FieldType = cboFieldType.Text
        If txtFieldSize = "" Then txtFieldSize = 10
        rs!FieldSize = txtFieldSize

        txtActionBeforeError = "rs.Update"
        rs.Update
        
        
        
        
        
    ' Commit Transactions
    con.CommitTrans
    Conf.CommitTrans
    Conb.CommitTrans
        
    'Close connection and the recordset
    'Again... for some reason the rs was already closed???
'    rs.Close
    Set rs = Nothing
    Set rsf = Nothing
    Set rsb = Nothing
    con.Close
    Set con = Nothing
    Set Conf = Nothing
    Set Conb = Nothing


        '*** 2020-06-10 - Jacob - MOVED code for ONLY ONE FIELD to AFTER the COMMIT and replaced with funcSaveFieldToDB
        '************************************************************************
        '*** Make Sure that Only ONE Field is flagged as
        '***    for each application.
        If ckbFieldSplitBatches = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldSplitBatches", 0)
        End If
        
        If ckbFieldRouteToBatchQueue = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldRouteToBatchQueue", 0)
        End If
        
        If ckbFieldRouteToBatchUser = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldRouteToBatchUser", 0)
        End If

        If ckbFieldRouteToBatchManager = vbChecked Then
            result = funcSaveFieldToDB(RegImaging101ConnectionString, "I101Fields", "ApplicationRECID = " & txtApplicationRECID & "   AND FieldsRECID <> " & txtFieldsRECID, "FieldRouteToBatchManager", 0)
        End If


    cmdFieldUpdate.Enabled = False
    cmdFieldDelete.Enabled = True
    cmdFieldsClear_Click
    subLoadFields

Exit Sub

ERROR_TRAP:
    On Error Resume Next
    
    result = MsgBox("Error cmdFieldUpdate_Click(): " & Err.Number & " - " & Err.Description & " - DURING ACTION: " & txtActionBeforeError, vbOK)
    
    If Conf.Errors.Count > 0 Then
        If Conf.Errors(0) <> "" Then
            result = MsgBox("SQL Error cmdFieldUpdate_Click(): " & Conf.Errors(0).Number & " - " & Conf.Errors(0).Description & " - DURING ACTION: " & txtActionBeforeError, vbOK)
        End If
    End If
    
    If Conb.Errors.Count > 0 Then
        If Conb.Errors(0) <> "" Then
            result = MsgBox("SQL Error cmdFieldUpdate_Click(): " & Conb.Errors(0).Number & " - " & Conb.Errors(0).Description & " - DURING ACTION: " & txtActionBeforeError, vbOK)
        End If
    End If
    
    If con.Errors.Count > 0 Then
        If con.Errors(0) <> "" Then
            result = MsgBox("SQL Error cmdFieldUpdate_Click(): " & con.Errors(0).Number & " - " & con.Errors(0).Description & " - DURING ACTION: " & txtActionBeforeError, vbOK)
        End If
    End If
    
    ' Rollback Transactions to prevent partial updates.
    Conf.RollbackTrans
    Conb.RollbackTrans
    con.RollbackTrans
    Err.Clear

'    Resume Next
End Sub







Private Sub cmdRefreshApplications_Click()

    subLoadApplications

End Sub

Private Sub cmdSecurityAddNew_Click()

    On Error GoTo ERROR_TRAP
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con.mode = adModeReadWrite
    rs.LOCKTYPE = adLockOptimistic
    
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Security where UserID = '" & txtUserID & "' OR UserName = '" & txtUserName & "'"
    rs.Open ssql, con
    
    On Error Resume Next
    
    
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        MsgBox "UserID [" & txtUserID & " - " & txtUserName & "] already exists." & vbCrLf & "UserID's & UserNames MUST be UNIQUE."
        Exit Sub
    End If
    
    
        rs.Fields!SecurityRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Control", "SecurityRECID")
        rs.Fields!UserID = txtUserID
        rs.Fields!username = txtUserName
        rs.Fields!Password = txtPassword
        rs.Fields!RightsAdminSystem = chkRightsAdminSystem
        rs.Fields!BatchDefaultApplication = cmbBatchDefaultApplication
        
        rs.Update
    
    
    'Close  recordset
    rs.Close


    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    '*** DON'T Allow Changing Name
    
    cmdSecurityUpdate.Enabled = True
    cmdSecurityAddNew.Enabled = True
    cmdSecurityRemove.Enabled = True

    If Trim(txtUserID) <> "" And lstSecurityApplicationList.ListCount = 0 Then
        result = MsgBox("No Applications have been assigned to this user!" & vbCrLf & "Please remember to GRANT rights", vbOKOnly)
    End If

    subSecurityLoadUserIDs
    
    'Select the Newly Added user
    funcFindItemInListBox lstUserList, txtUserName
    
    'Clear Application-Specific Fields
     subSecurityRightsClearFields

    
    
    'Select the Grant/Revoke option
    cmdSecurityApplicationListGrantRevoke_Click
    
Exit Sub
    
ERROR_TRAP:
    result = MsgBox("Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear


End Sub

Private Sub cmdSecurityApplicationListGrant_Click()

    subSecurityApplicationGrantAccess
    
    'Clear Application-Specific Fields
    subSecurityRightsClearFields


    
End Sub

Private Sub cmdSecurityApplicationListGrantRevoke_Click()

    If cmdSecurityApplicationListGrantRevoke.Caption = "          Grant/Revoke           Application Access" Then
    
        frameRightsAssignments.Visible = False
        cmdSecurityAddNew.Visible = False
        cmdSecurityClearFields.Visible = False
        cmdSecurityRemove.Visible = False
        cmdSecurityUpdate.Visible = False
        lstSecurityApplicationSelectionList.Visible = False

        cmdSecurityApplicationListGrant.Top = lstSecurityApplicationList.Top
        cmdSecurityApplicationListGrant.Left = lstSecurityApplicationList.Left + lstSecurityApplicationList.width
        
        cmdSecurityApplicationListRevoke.Top = cmdSecurityApplicationListGrant.Top + cmdSecurityApplicationListGrant.Height
        cmdSecurityApplicationListRevoke.Left = cmdSecurityApplicationListGrant.Left
        
        lstSecurityApplicationSelectionList.Top = lstSecurityApplicationList.Top
        lstSecurityApplicationSelectionList.Left = cmdSecurityApplicationListGrant.Left + cmdSecurityApplicationListGrant.width
        lstSecurityApplicationSelectionList.Height = lstSecurityApplicationList.Height
        lstSecurityApplicationSelectionList.width = lstSecurityApplicationList.width * 2
        
        DoEvents
        
        
        'Copy items from the main Application List
        lstSecurityApplicationSelectionList.Clear
        For i = 0 To lstApplications.ListCount - 1
            lstApplications.ListIndex = i
            lstSecurityApplicationSelectionList.AddItem lstApplications.Text
            lstSecurityApplicationSelectionList.ItemData(lstSecurityApplicationSelectionList.ListCount - 1) = lstApplications.ItemData(lstApplications.ListIndex)
        Next
        
        lstSecurityApplicationSelectionList.Visible = True
        cmdSecurityApplicationListGrant.Visible = True
        cmdSecurityApplicationListRevoke.Visible = True
        cmdSecurityApplicationListGrantRevoke.Caption = "End Grant/Revoke"
        
        
    Else
        lstSecurityApplicationSelectionList.Visible = False
        cmdSecurityApplicationListGrant.Visible = False
        cmdSecurityApplicationListRevoke.Visible = False
        
        cmdSecurityApplicationListGrantRevoke.Caption = "          Grant/Revoke           Application Access"
        frameRightsAssignments.Visible = True
        cmdSecurityAddNew.Visible = True
        cmdSecurityClearFields.Visible = True
        cmdSecurityRemove.Visible = True
        cmdSecurityUpdate.Visible = True
        
    End If

End Sub



Private Sub cmdSecurityApplicationListRevoke_Click()

       
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    If Trim(lstSecurityApplicationList.Text = "") Or Trim(txtSecurityRECID) = "" Then
        Exit Sub
    End If
    
'''    con.Errors.Clear
    
    On Error GoTo ERROR_TRAP

    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    
'*** 2021-11-17 - Jacob - Disabled check for record existing.  Not sure why, but it was not returning record found,
'                                              even though the record DID exist.
'                                            Also changed the delete method for I101SecurityApplications.
'''
''''*** Changed the Load to work with Security
''''    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
'''    rs.Source = ""
'''    rs.Source = rs.Source & "SELECT  ApplicationRECID"
'''    rs.Source = rs.Source & " FROM  I101SecurityApplications"
'''    rs.Source = rs.Source & " WHERE I101SecurityApplications.ApplicationRECID = " & lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListIndex)
'''    rs.Source = rs.Source & "   AND I101SecurityApplications.SecurityRECID = " & txtSecurityRECID
'''
'''    rs.CursorLocation = adUseServer
'''    rs.CursorType = adOpenDynamic
'''    rs.LOCKTYPE = adLockOptimistic
'''
'''
'''
'''    rs.Open
'''
'''    If rs.RecordCount > 0 Then
'''        rs.MoveFirst
'''    Else
'''        MsgBox "SORRY!   I was NOT able  to find this Record in the I101SecurityApplications Table." & vbCrLf & _
'''                        "ApplicationRECID=" & lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListIndex) & vbCrLf & _
'''                        "txtSecurityRECID=" & txtSecurityRECID, vbExclamation, "Record Not Found in I101SecurityApplications Table"
'''        Exit Sub
'''    End If
'''
'''    rs.Delete
'''
'''    'Close connection and the recordset
'''    rs.Close
    
    'DELETE the User Rights Record
    ssql = "DELETE FROM I101SecurityRoleApp WHERE " & _
                "SecurityRECID = " & txtSecurityRECID & _
                 " AND  ApplicationRECID = " & lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListIndex)
    rs.Open ssql, con
    
    'DELETE the User Application Granted Record
    ssql = "DELETE FROM I101SecurityApplications WHERE " & _
                "SecurityRECID = " & txtSecurityRECID & _
                 " AND  ApplicationRECID = " & lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListIndex)
    rs.Open ssql, con
    
    'Close connection and the recordset
'    rs.Close
    
    Set rs = Nothing
    con.Close
    Set con = Nothing

    'Refresh the Application List, etc.
    lstUserList_Click
    
Exit Sub
    
ERROR_TRAP:

    result = MsgBox("cmdSecurityApplicationListRevoke_Click() ERROR: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
    
End Sub

Private Sub cmdSecurityClearFields_Click()

        txtSecurityRECID = ""
        txtUserID = ""
        txtUserName = ""
        txtPassword = ""
        
        subSecurityRightsClearFields
        
        lstSecurityApplicationList.Clear
        
        txtUserID.SetFocus

End Sub

Private Sub subSecurityRightsClearFields()

    'Clear Application-Specific Fields ONLY if NOT in Copy Mode
    If chkRightsEnableCopyMode.Value = vbUnchecked Then

        cmbBatchMode = ""
        cmbUserSupervisor = ""
        cmbBatchListOrder = ""
        cmbBatchDefaultQueue = ""
'        cmbBatchDefaultApplication = ""
        
'        chkRightsAdminSystem = vbUnchecked
        chkRightsAdminApplication = vbUnchecked
        chkRightsBatchScan = vbUnchecked
        chkRightsBatchIndex = vbUnchecked
        chkRightsImportFromFile = vbUnchecked
        chkRightsImportFromEcapture = vbUnchecked
        chkRightsDeleteDocuments = vbUnchecked
        chkRightsModifyIndexes = vbUnchecked
        chkRightsBatchCommit = vbUnchecked
        chkRightsDeleteBatches = vbUnchecked
        chkRightsSendMail = vbUnchecked
        chkRightsLaunchDoc = vbUnchecked
        chkRightsScannerSettings = vbUnchecked
        chkRightsBatchView = vbUnchecked
        chkRightsBatchRoute = vbUnchecked
        chkRightsBatchChangeOrder = vbUnchecked
        chkRightsRetrieveImages = vbUnchecked
        
        chkViewResetImagesOnFind = vbUnchecked
        
        chkAllowModificationOfOrigDocs = vbUnchecked
        
        chkRightsBatchFindRestricted = vbUnchecked
        chkRightsBatchFindRestrictToQueue = vbUnchecked
        chkRightsBatchFindRestrictToOwner = vbUnchecked
        chkRightsBatchChangeQueue = vbUnchecked
        chkRightsBatchChangeOwner = vbUnchecked

        chkRightsBatchAllowDocTypeEdit = vbUnchecked
        
        chkRightsBatchAdministration = vbUnchecked
'        chkRightsDocPackage = vbUnchecked
        chkRightsPrint = vbUnchecked
        chkRightsAnnotate = vbUnchecked
        chkRightsThumbnails = vbUnchecked
        chkRightsExport = vbUnchecked
        
    End If
        
        
End Sub
Private Sub cmdSecurityConfigure_Click()
    frm101Security.Show
End Sub



Private Sub cmdSecurityRefreshUsers_Click()

    subSecurityLoadUserIDs

End Sub



Private Sub cmdSecurityRemove_Click()

    On Error GoTo ERROR_TRAP

    result = MsgBox("Are you SURE you wish to DELETE user [" & txtUserID & "] from the system?   This action is irreversible.", vbYesNo, "Delete User")
    If result = vbNo Then
        Exit Sub
    End If

    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con.mode = adModeReadWrite
    rs.LOCKTYPE = adLockOptimistic
    
    con.Open RegImaging101ConnectionString
    
    'DELETE the User Rights Record
    ssql = "Delete from I101SecurityRoleApp where SecurityRECID = " & txtSecurityRECID
    rs.Open ssql, con
    
    'DELETE the User Record
    ssql = "Delete from i101SecurityApplications where SecurityRECID = " & txtSecurityRECID
    rs.Open ssql, con
    
    'DELETE the User Record
    ssql = "Delete from I101Security where SecurityRECID = " & txtSecurityRECID
    rs.Open ssql, con
    
    
    
    'Close connection and the recordset
'    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    
    cmdSecurityUpdate.Enabled = True
    cmdSecurityAddNew.Enabled = True
    cmdSecurityRemove.Enabled = True

    '*** CLEAR all the fields
    cmdSecurityClearFields_Click
    subSecurityLoadUserIDs
    
Exit Sub
    
    
ERROR_TRAP:
    funcQuickMessage "SHOW", "cmdSecurityRemove_Click ERROR: " & Err.Number & " - " & Err.Description
    Err.Clear


End Sub

Private Sub cmdSecurityUpdate_Click()
    
    If Trim(cmbBatchDefaultApplication) = "" Then
        funcQuickMessage "SHOW", "Please select a 'Batch Default Application.'"
        cmbBatchDefaultApplication.SetFocus
        Exit Sub
    End If
        

    On Error GoTo ERROR_TRAP
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con.mode = adModeReadWrite
    rs.LOCKTYPE = adLockOptimistic
    
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Security where SecurityRECID = " & txtSecurityRECID
    rs.Open ssql, con
    
    rs.MoveFirst
    
    
'        rs.Fields!SecurityRECID = funcGetNextControlNumber(RegImaging101ConnectionString, "I101Security", "SecurityRECID")
        rs.Fields!UserID = txtUserID
        rs.Fields!username = txtUserName
        rs.Fields!Password = txtPassword
'        rs.Fields!BatchMode = cmbBatchMode
'        rs.Fields!UserSupervisor = cmbUserSupervisor
'        rs.Fields!BatchListOrder = cmbBatchListOrder
'        rs.Fields!BatchDefaultQueue = cmbBatchDefaultQueue
        rs.Fields!BatchDefaultApplication = cmbBatchDefaultApplication
'        rs.Fields!BatchQueueNotificationFrequency = txtBatchQueueNotificationFrequency
'
        rs.Fields!RightsAdminSystem = chkRightsAdminSystem
'        rs.Fields!RightsAdminApplication = chkRightsAdminApplication
'        rs.Fields!RightsBatchScan = chkRightsBatchScan
'        rs.Fields!RightsBatchIndex = chkRightsBatchIndex
'        rs.Fields!RightsBatchAdministration = chkRightsBatchAdministration
'        rs.Fields!RightsImportFromFile = chkRightsImportFromFile
'        rs.Fields!RightsImportFromEcapture = chkRightsImportFromEcapture
'        rs.Fields!RightsDeleteDocuments = chkRightsDeleteDocuments
'        rs.Fields!RightsModifyIndexes = chkRightsModifyIndexes
'        rs.Fields!RightsBatchCommit = chkRightsBatchCommit
'        rs.Fields!RightsDeleteBatches = chkRightsDeleteBatches
'        rs.Fields!RightsSendMail = chkRightsSendMail
'        rs.Fields!RightsLaunchDoc = chkRightsLaunchDoc
'        rs.Fields!RightsPrint = chkRightsPrint
'        rs.Fields!RightsAnnotate = chkRightsAnnotate
'        rs.Fields!RightsThumbnails = chkRightsThumbnails
'        rs.Fields!RightsScannerSettings = chkRightsScannerSettings
'        rs.Fields!RightsBatchView = chkRightsBatchView
'        rs.Fields!RightsBatchRoute = chkRightsBatchRoute
'        rs.Fields!RightsBatchChangeOrder = chkRightsBatchChangeOrder
'        rs.Fields!RightsRetrieveImages = chkRightsRetrieveImages
'        rs.Fields!RightsDocPackage = chkRightsDocPackage
'        rs.Fields!RightsExport = chkRightsExport
'
'        rs.Fields!ViewResetImagesOnFind = chkViewResetImagesOnFind
'
'        rs.Fields!AllowModificationOfOrigDocs = chkAllowModificationOfOrigDocs
'
'        rs.Fields!RightsBatchFindRestricted = chkRightsBatchFindRestricted
'        rs.Fields!RightsBatchFindRestrictToQueue = chkRightsBatchFindRestrictToQueue
'        rs.Fields!RightsBatchFindRestrictToOwner = chkRightsBatchFindRestrictToOwner
'        rs.Fields!RightsBatchChangeQueue = chkRightsBatchChangeQueue
'        rs.Fields!RightsBatchChangeOwner = chkRightsBatchChangeOwner
'
'        rs.Fields!RightsBatchAllowDocTypeEdit = chkRightsBatchAllowDocTypeEdit
'
    rs.Update
    
    'Close the Recordset
    rs.Close
    
    
    
    ssql = "Select * from I101SecurityRoleApp where SecurityRECID = " & _
            txtSecurityRECID & _
            "AND ApplicationRECID = " & _
            txtSecurityApplicationRECID
    rs.Open ssql, con
    
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
    End If
    
   
        rs.Fields!SecurityRoleRECID = 0
        rs.Fields!SecurityRECID = txtSecurityRECID
        rs.Fields!ApplicationRECID = txtSecurityApplicationRECID
        
        rs.Fields!BatchMode = cmbBatchMode
        rs.Fields!UserSupervisor = cmbUserSupervisor
        rs.Fields!BatchListOrder = cmbBatchListOrder
        rs.Fields!BatchDefaultQueue = cmbBatchDefaultQueue
'        rs.Fields!BatchDefaultApplication = cmbBatchDefaultApplication
        rs.Fields!BatchQueueNotificationFrequency = txtBatchQueueNotificationFrequency

'        rs.Fields!RightsAdminSystem = chkRightsAdminSystem
        rs.Fields!RightsAdminApplication = chkRightsAdminApplication
        rs.Fields!RightsBatchScan = chkRightsBatchScan
        rs.Fields!RightsBatchIndex = chkRightsBatchIndex
        rs.Fields!RightsBatchAdministration = chkRightsBatchAdministration
        rs.Fields!RightsImportFromFile = chkRightsImportFromFile
        rs.Fields!RightsImportFromEcapture = chkRightsImportFromEcapture
        rs.Fields!RightsDeleteDocuments = chkRightsDeleteDocuments
        rs.Fields!RightsModifyIndexes = chkRightsModifyIndexes
        rs.Fields!RightsBatchCommit = chkRightsBatchCommit
        rs.Fields!RightsDeleteBatches = chkRightsDeleteBatches
        rs.Fields!RightsSendMail = chkRightsSendMail
        rs.Fields!RightsLaunchDoc = chkRightsLaunchDoc
        rs.Fields!RightsPrint = chkRightsPrint
        rs.Fields!RightsAnnotate = chkRightsAnnotate
        rs.Fields!RightsThumbnails = chkRightsThumbnails
        rs.Fields!RightsScannerSettings = chkRightsScannerSettings
        rs.Fields!RightsBatchView = chkRightsBatchView
        rs.Fields!RightsBatchRoute = chkRightsBatchRoute
        rs.Fields!RightsBatchChangeOrder = chkRightsBatchChangeOrder
        rs.Fields!RightsRetrieveImages = chkRightsRetrieveImages
'        rs.Fields!RightsDocPackage = chkRightsDocPackage
        rs.Fields!RightsExport = chkRightsExport
        
        rs.Fields!ViewResetImagesOnFind = chkViewResetImagesOnFind
        
        rs.Fields!AllowModificationOfOrigDocs = chkAllowModificationOfOrigDocs
        
        rs.Fields!RightsBatchFindRestricted = chkRightsBatchFindRestricted
        rs.Fields!RightsBatchFindRestrictToQueue = chkRightsBatchFindRestrictToQueue
        rs.Fields!RightsBatchFindRestrictToOwner = chkRightsBatchFindRestrictToOwner
        rs.Fields!RightsBatchChangeQueue = chkRightsBatchChangeQueue
        rs.Fields!RightsBatchChangeOwner = chkRightsBatchChangeOwner

        rs.Fields!RightsBatchAllowDocTypeEdit = chkRightsBatchAllowDocTypeEdit
        
        rs.Fields!RightsAdvancedSearch = chkRightsAdvancedSearch
               
        rs.Fields!RightsFileDocsViaI101FILER = chkRightsFileDocsViaI101FILER
        
        '*** 2020-05-15 - Jacob - Added chkRightsEditSearchTemplates
        rs.Fields!RightsEditSearchTemplates = chkRightsEditSearchTemplates
        

    rs.Update
    
    'Close the Recordset
    rs.Close
    
    
    
    Set rs = Nothing
    'Close the Connection
    con.Close
    Set con = Nothing
    
    
    
    
    cmdSecurityUpdate.Enabled = True
    cmdSecurityAddNew.Enabled = True
    cmdSecurityRemove.Enabled = True

    
Exit Sub
    
    
ERROR_TRAP:
    result = MsgBox("Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear




End Sub



Private Sub cmdTestFTP_Click()

        Dim strTempDirectory As String
        Dim strFTPUploadSourceFilePath As String
        Dim strFTPUploadDestinationFileName As String

        strTempDirectory = funcGetTempDir()
        
        frmFTP.Show

        frmFTP.lblRESPONSE.Caption = "CREATING TEST FILE."
        frmFTP.lblRESPONSE.Refresh

        'BUILD DESTINATION FILE NAME
        strFTPUploadSourceFilePath = strTempDirectory & "FTPTEST.TXT"
        strFTPUploadDestinationFileName = "FTPTEST_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".TXT"

        Open strFTPUploadSourceFilePath For Output As #1
        Print #1, "TEXT FILE FOR FTP CONNECTION TEST"
        Close #1

''        'UPLOAD FILE TO THE FTP SITE
''        txtActionBeforeError = "FTP Site= " & txtFTPSite & "  Command= PUT " & txtFTPUserID & ", " & txtFTPPassword & ", " & strFTPUploadSourceFilePath & ", " & strFTPUploadDestinationFileName & ", " & False
''        funcWriteToDebugLog Me.name, txtActionBeforeError
''        frmFTP.FTPFile txtFTPSite, "PUT", txtFTPUserID, txtFTPPassword, strFTPUploadSourceFilePath, strFTPUploadDestinationFileName, True
''
''        'DELETE THE TEST FILE FROM THE FTP SITE
''        frmFTP.FTPFile txtFTPSite, "DELETE", txtFTPUserID, txtFTPPassword, strFTPUploadDestinationFileName, "", False
''
''        frmFTP.lblRESPONSE.Caption = "TEST COMPLETE... PLEASE CHECK FOR ERRORS."
''        frmFTP.lblRESPONSE.Refresh

        Dim bolFTPErrorOccured As Boolean
        bolFTPErrorOccured = funcFTPPutFile(txtFTPSite, CInt(txtFTPPort), txtFTPUserID, txtFTPPassword, strFTPUploadSourceFilePath, strFTPUploadDestinationFileName)
        
        frmFTP.lblRESPONSE.Caption = "TEST COMPLETE... PLEASE CHECK FOR ERRORS."
        frmFTP.lblRESPONSE.Refresh

        
End Sub

Private Sub cmdUpdateSpecialOptions_Click()

    result = WritePrivateProfileString(RegAppname, "frmConfig.txtBarcodeLicenseKey", frmConfig.txtBarcodeLicenseKey, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmConfig.chkDropLeadingZeroes", frmConfig.chkDropLeadingZeroes, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmConfig.chkUseBarcodeAsDocumentHeader", frmConfig.chkUseBarcodeAsDocumentHeader, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmConfig.txtBarcodeClipBeginPosition", frmConfig.txtBarcodeClipBeginPosition, RegFileName)
    result = WritePrivateProfileString(RegAppname, "frmConfig.txtBarcodeClipNumberOfCharacters", frmConfig.txtBarcodeClipNumberOfCharacters, RegFileName)
    
    
    
        '***********************************************************
        '*** Save SMTP Settings
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim ssql As String
        
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        conn.ConnectionString = RegImaging101ConnectionString
        conn.ConnectionTimeout = 120
        conn.mode = adModeReadWrite
        conn.Open
        
        ssql = "SELECT * FROM I101Control WHERE ID = 1"
        
        With rs
            .ActiveConnection = conn
            .CursorLocation = adUseServer
            .CursorType = adOpenDynamic
            .LOCKTYPE = adLockOptimistic
            .Source = ssql
        End With
    
        rs.Open

        'Return the Found Value
        rs("SendEmailViaSMTP") = cboSendEmailViaSMTP
        rs("SMTPHost") = txtSMTPHost
        rs("SMTPPOP3Host") = txtSMTPPOP3Host
        rs("SMTPRequiresAuthentication") = cboSMTPRequiresAuthentication
        rs("SMTPUsePOP3Auth") = cboSMTPUsePOP3Auth
        rs("SMTPAuthenticationUserID") = txtSMTPAuthenticationUserID
        rs("SMTPAuthenticationPassword") = txtSMTPAuthenticationPassword
        rs("SMTPDefaultEmailSubject") = txtSMTPEmailSubject
        rs("SMTPDefaultEmailMessage") = txtSMTPEmailMessage
        rs("SMTPPort") = txtSMTPPort
        
        rs("AutoLaunchFileTypes") = frmConfig.txtAutoLaunchFileTypes
        
        '2021-11-09 - Jacob - Added New Field for selecting how to Auto-Launch documents.
        rs("AutoLaunchTo") = frmConfig.txtAutoLaunchTo

        rs.Update
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

    
    
    subCheckBarCodeStatus
    
End Sub



Private Sub Combo1_Change()

End Sub



Private Sub Form_Load()
''    Set PrimaryCLS = New clsDocuments
''    frmImaging101Retrieve.grdDataGrid.DataMember = "Primary"
''    Set frmImaging101Retrieve.grdDataGrid.DataSource = PrimaryCLS
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    If bolDebug Then
        Me.Caption = Me.Caption & " - DEBUG"
    End If

    On Error Resume Next
        
        
    
    '*** SET DEFAULT BUTTON PROPERTIES
    cmdApplicationAdd.Enabled = False
    cmdApplicationClear.Enabled = False
    cmdApplicationRemove.Enabled = False
    cmdApplicationUpdate.Enabled = False
    cmdFieldAdd.Enabled = False
    cmdFieldsClear.Enabled = False
    cmdFieldDelete.Enabled = False
    cmdFieldUpdate.Enabled = False
    cmdFieldMoveUp.Enabled = False
    cmdFieldMoveDown.Enabled = False
    
    cboLookupDBTableName.Enabled = False
    cboFieldToSelectAfterLookupClick.Enabled = False
    cboFieldToSelectAfterDocListClick.Enabled = False
    
    
    
    '*** OPEN DATA SOURCE
    
''    If optJetVersion(0).Value = True Then
''        adoAppConn.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & gFileSpec & ";"
''    ElseIf optJetVersion(1).Value = True Then
''        adoAppConn.Open "PROVIDER=Microsoft.Jet.OLEDB.3.6;Data Source=" & gFileSpec & ";"
''    ElseIf optJetVersion(2).Value = True Then
''        adoAppConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & gFileSpec & ";"
''    End If
''
'''''''    ' Get Database Connections settings from the registry
'''''''    On Error Resume Next
'''''''    RegImaging101ConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionType", RegFileName)
'''''''    RegImaging101ConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101Retrieve.Adodc1.ConnectionString." & RegImaging101ConnectionType, RegFileName)
'''''''    RegImaging101BatchListConnectionType = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionType", RegFileName)
'''''''    RegImaging101BatchListConnectionString = VBGetPrivateProfileString("DATABASE", "frmImaging101BatchList.AdodcImaging101Batch.ConnectionString." & RegImaging101BatchListConnectionType, RegFileName)
'''''''    On Error GoTo 0
    
    subLoadApplications
    
    ' Load FieldTypes
    cboFieldType.AddItem "Text"
    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adVarWChar
    cboFieldType.AddItem "LongText"
    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adVarWChar
    cboFieldType.AddItem "Boolean"
    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adBoolean
    cboFieldType.AddItem "Numeric"
    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adDouble
    cboFieldType.AddItem "Date"
    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adDBTimeStamp
'''''''    cboFieldType.AddItem "Notes"
'''''''    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adLongVarWChar
    cboFieldType.AddItem "Currency"
    cboFieldType.ItemData(cboFieldType.ListCount - 1) = adCurrency
    ' Set Default FieldType
    cboFieldType.Text = "Text"
    txtFieldSize.Text = 10
    
    '*** Force the Site to A
    cboSiteID.Text = "Site A"
    
    '*** ADD BATCH QUEUES TO THE cboFieldDefaultValue FIELD, Don't clear the list.
    funcFillList Me.cboFieldDefaultValue, RegImaging101ConnectionString, "I101BatchQueues", "BatchQueue", "", True, False

    
''    Set rsSchema = adoAppConn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
''    Me.Screen.MousePointer = vbHourglass

    ' Set the Selected TAB to be the FIRST (i.e.- "Application Definition")
    sstabConfig.Tab = 0
    sstabApplication.Tab = 0
    
    ' HIDE Tabs we DON'T need.
    sstabConfig.TabVisible(1) = False
    sstabConfig.TabVisible(2) = False

    txtApplicationCommitBatchoption_Click
    
End Sub

Private Sub subLoadApplications()
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
        

    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "SELECT ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.Open ssql, con
    
    lstApplications.Clear  ' Reset the List
    
    Do Until rs.EOF
        lstApplications.AddItem rs!ApplicationName
        lstApplications.ItemData(lstApplications.ListCount - 1) = rs!ApplicationRECID  'Adds lastnames to list box
        rs.MoveNext
    Loop
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

End Sub

Private Sub subLoadSecurityApplications()

    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
'*** Changed the Load to work with Security
'    rs.Source = "Select ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.Source = ""
    rs.Source = rs.Source & "Select * "
    rs.Source = rs.Source & " FROM I101Applications, I101SecurityApplications"
    rs.Source = rs.Source & " WHERE I101Applications.ApplicationRECID = I101SecurityApplications.ApplicationRECID And I101SecurityApplications.SecurityRECID = " & txtSecurityRECID
    rs.Source = rs.Source & " ORDER BY ApplicationName"
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    lstSecurityApplicationList.Clear
       
       
       rs.Open
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    For intIndex = 0 To rs.RecordCount - 1
        lstSecurityApplicationList.AddItem rs.Fields!ApplicationName
        lstSecurityApplicationList.ItemData(intIndex) = rs.Fields!ApplicationRECID
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    'Select the FIRST Application in the List
    If lstSecurityApplicationList.ListCount > 0 Then
        lstSecurityApplicationList.ListIndex = 0
        lstSecurityApplicationList.Selected(0) = True
    Else
        subSecurityRightsClearFields
    End If
    
    
End Sub

Private Sub subLoadFields()
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select FieldName, FieldsRECID FROM I101Fields WHERE ApplicationRECID = " & txtApplicationRECID & " ORDER BY FieldOrderBatch "
    rs.Open ssql, con
    
    lstFields.Clear  ' Reset the List
    
    Do Until rs.EOF
        lstFields.AddItem rs!FieldName
        lstFields.ItemData(lstFields.ListCount - 1) = rs!FieldsRECID
        rs.MoveNext
    Loop
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Trim(txtUserID) <> "" And lstSecurityApplicationList.ListCount = 0 Then
        result = MsgBox("No Applications have been assigned to this user!" & vbCrLf & "Are you sure you wish to exit?", vbYesNo)
        If result = vbNo Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMainMenu.Show
    frmMainMenu.WindowState = vbNormal
    frmMainMenu.SetFocus

End Sub









Private Sub lstApplications_Click()
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    
    con.ConnectionTimeout = 120
    con.CommandTimeout = 600
        

    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Applications where ApplicationRECID = " & lstApplications.ItemData(lstApplications.ListIndex)
    rs.Open ssql, con
    
    On Error Resume Next
    
        txtApplicationName = rs.Fields!ApplicationName & ""
        txtApplicationRECID = rs.Fields!ApplicationRECID
        txtApplicationDescription = rs.Fields!ApplicationDescription & ""
        txtApplicationNotes = rs.Fields!ApplicationNotes & ""
        chkApplicationIsActive = rs.Fields!ApplicationIsActive & ""
        chkApplicationIsReadOnly = rs.Fields!ApplicationIsReadOnly & ""
        txtApplicationCommitBatchTo = rs.Fields!ApplicationCommitBatchTo & ""
        chkAutoAdvanceOnSeparator = rs.Fields!ApplicationAutoAdvanceOnSeparator & ""
        chkSetUserAsBatchOwnerOnSPLIT = rs.Fields!SetUserAsBatchOwnerOnSPLIT
        chkEnableSearchTemplates = rs.Fields!EnableSearchTemplates
        
        '*** ONLY Show the SiteID and CaseIdCutofffor TTC... for now...
        If txtApplicationName = "TTC" Then
            cboSiteID.Visible = True
            lblCaseIdCutoff.Visible = True
            txtCaseIdCutoff.Visible = True
            'Don't change the Site
            'Make the FTP File naming fields Invisible.  They are Hard-Coded for TTC
            lblFTPConfigureFileNaming.Visible = False
            lblFTPSelectField.Visible = False
            lblFTPDelimiter.Visible = False
            cboFTPFileNameField(0).Visible = False
            cboFTPFileNameField(1).Visible = False
            cboFTPFileNameField(2).Visible = False
            cboFTPFileNameField(3).Visible = False
            cboFTPFileNameDelimiter(0).Visible = False
            cboFTPFileNameDelimiter(1).Visible = False
            cboFTPFileNameDelimiter(2).Visible = False
        Else
            cboSiteID.Visible = False
            lblCaseIdCutoff.Visible = False
            txtCaseIdCutoff.Visible = False
            'Force the site to A
            cboSiteID = "Site A"
            'Make the FTP File naming fields Visible
            lblFTPConfigureFileNaming.Visible = True
            lblFTPSelectField.Visible = True
            lblFTPDelimiter.Visible = True
            cboFTPFileNameField(0).Visible = True
            cboFTPFileNameField(1).Visible = True
            cboFTPFileNameField(2).Visible = True
            cboFTPFileNameField(3).Visible = True
            cboFTPFileNameDelimiter(0).Visible = True
            cboFTPFileNameDelimiter(1).Visible = True
            cboFTPFileNameDelimiter(2).Visible = True
        End If
        
        If cboSiteID = "Site A" Then
            txtLookupDBConnectionString = rs.Fields!LookupDBConnectionString & ""
            cboLookupDBTableName = rs.Fields!LookupDBTableName & ""
            chkLookupDBTableIsOnSQLServer = rs.Fields!LookupDBTableIsOnSQLServer & ""
            txtLookupDBWhereClause = rs.Fields!LookupDBWhereClause & ""
            '
            txtFTPSite = rs.Fields!FTPSite & ""
            txtFTPPort = rs.Fields!ftpport & ""
            txtFTPUserID = rs.Fields!FTPUserID & ""
            txtFTPPassword = rs.Fields!FTPPassword & ""
        Else
            txtLookupDBConnectionString = rs.Fields!LookupDBConnectionString_B & ""
            cboLookupDBTableName = rs.Fields!LookupDBTableName_B & ""
            chkLookupDBTableIsOnSQLServer = rs.Fields!LookupDBTableIsOnSQLServer_B & ""
            txtLookupDBWhereClause = rs.Fields!LookupDBWhereClause_B & ""
            '
            txtFTPSite = rs.Fields!FTPSite_B & ""
            txtFTPPort = rs.Fields!ftpport_B & ""
            txtFTPUserID = rs.Fields!FTPUserID_B & ""
            txtFTPPassword = rs.Fields!FTPPassword_B & ""
        End If
        
        chkAutoLookupOnBatchLoad = rs.Fields!AutoLookupOnBatchLoad & ""
        
        txtCaseIdCutoff = rs.Fields!CaseIdCutoff & ""

        cboFieldToSelectAfterLookupClick = rs.Fields!FieldToSelectAfterLookupClick & ""
        cboFieldToSelectAfterDocListClick = rs.Fields!FieldToSelectAfterDocListClick & ""
        cboFieldToSelectAfterNextPageClick = rs.Fields!FieldToSelectAfterNextPageClick & ""

        cboFieldToAssignDocumentGroup = rs.Fields!FieldToAssignDocumentGroup & ""
        cboFieldToAssignDocumentType = rs.Fields!FieldToAssignDocumentType & ""
         cboFieldToAssignDocumentSubType = rs.Fields!FieldToAssignDocumentSubType & ""
       
        txtRootDirectoryPathForImageArchive = rs.Fields!RootDirectoryPathForImageArchive & ""
        txtRootDirectoryPathForImageAnnotations = rs.Fields!RootDirectoryPathForImageAnnotations & ""
        txtRootDirectoryPathForBatches = rs.Fields!RootDirectoryPathForBatches & ""
        txtRootDirectoryPathForHtmlSource = rs.Fields!RootDirectoryPathForHtmlSource & ""
        
        txtRouteMaxCount = rs.Fields!RouteMaxCount & ""
        
        txtApplicationBatchNameDelimiter = rs.Fields!ApplicationBatchNameDelimiter
    
        
        txtMaxItemsToRetrieve = rs.Fields!MaxItemsToRetrieve & ""
        
        txtApplicationCommitBatchOption = rs.Fields!ApplicationCommitBatchOption & ""
        
        cboFTPFileNameField(0) = rs.Fields!FTPFileNameField0 & ""
        cboFTPFileNameField(1) = rs.Fields!FTPFileNameField1 & ""
        cboFTPFileNameField(2) = rs.Fields!FTPFileNameField2 & ""
        cboFTPFileNameField(3) = rs.Fields!FTPFileNameField3 & ""

        funcFindItemInComboBoxPartial cboFTPFileNameDelimiter(0), rs.Fields!FTPFileNameDelimiter0 & ""
        funcFindItemInComboBoxPartial cboFTPFileNameDelimiter(1), rs.Fields!FTPFileNameDelimiter1 & ""
        funcFindItemInComboBoxPartial cboFTPFileNameDelimiter(2), rs.Fields!FTPFileNameDelimiter2 & ""
        
        

    On Error GoTo ERROR_TRAP
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    '*** DON'T Allow Changing Name
    txtApplicationName.Enabled = False
    
    cmdApplicationUpdate.Enabled = True
    cmdApplicationAdd.Enabled = False
    cmdApplicationRemove.Enabled = True

    cmdFieldAdd.Enabled = True
    
    cboLookupDBTableName.Enabled = True
    cboFieldToSelectAfterLookupClick.Enabled = True
    cboFieldToSelectAfterDocListClick.Enabled = True
    
    subLoadFields
    cmdFieldsClear_Click
    
Exit Sub
    
ERROR_TRAP:

End Sub

Private Sub lstFields_Click()
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    On Error GoTo ERROR_TRAP
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Fields where ApplicationRECID = " & txtApplicationRECID & " AND FieldName = '" & lstFields.List(lstFields.ListIndex) & "'"
    rs.Open ssql, con
    
        txtFieldsRECID = rs.Fields!FieldsRECID
        txtFieldName = rs.Fields!FieldName
        txtFieldNameForInput = rs.Fields!FieldNameForInput
        txtFieldNameForOutput = rs.Fields!FieldNameForOutput
        txtFieldDescription = rs.Fields!FieldDescription
        txtFieldSize = rs.Fields!FieldSize
        cboFieldType.Text = rs.Fields!FieldType
        txtFieldTypeHOLD.Text = rs.Fields!FieldType

        txtFieldFormat = rs.Fields!FieldFormat
        txtFieldMask = rs.Fields!FieldMask
        cboFieldDefaultValue = rs.Fields!FieldDefaultValue
        txtFieldHighValue = rs.Fields!FieldHighValue
        txtFieldLowValue = rs.Fields!FieldLowValue
        cboFieldDefaultValue = rs!FieldDefaultValue
        ckbFieldAutoIncrField = rs!FieldAutoIncrField
        ckbFieldIsRequiredForCommit = rs!FieldIsRequiredForCommit
        ckbFieldIsRequiredForSplit = rs!FieldIsRequiredForSplit
        ckbFieldSplitBatches = rs!FieldSplitBatches
        
         If IsNull(rs!FieldRouteToBatchQueue) Then
            ckbFieldRouteToBatchQueue = 0
        Else
             ckbFieldRouteToBatchQueue = rs!FieldRouteToBatchQueue & ""
        End If
        
         If IsNull(rs!FieldRouteToBatchUser) Then
            ckbFieldRouteToBatchUser = 0
        Else
            ckbFieldRouteToBatchUser = rs!FieldRouteToBatchUser & ""
        End If
            
         
         If IsNull(rs!FieldRouteToBatchManager) Then
            ckbFieldRouteToBatchManager = 0
        Else
            ckbFieldRouteToBatchManager = rs!FieldRouteToBatchManager & ""
        End If
        
        ckbFieldIsSticky = rs!FieldIsSticky
        
        ckbFieldIsForOutputOnly = rs!FieldIsForOutputOnly
        ckbHideForSearchIndex = rs!HideForSearchIndex
        
        ckbFieldDropDownList = rs!FieldDropDownList
        ckbFieldDropDownListAlsoOnFiler = rs!FieldDropDownListAlsoOnFiler

        chkFieldDefaultForBarcodeOnly = rs!FieldDefaultForBarcodeOnly
        ckbFieldTableLookupOverridesDefault = rs!FieldTableLookupOverridesDefault
        If rs!FieldSearchCondition = "" Then
            cboFieldSearchCondition = "Contains"
        Else
            cboFieldSearchCondition = rs!FieldSearchCondition
        End If
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    '*** DON'T Allow Changing Name
    txtFieldName.Enabled = False
    
    cmdFieldUpdate.Enabled = True
    cmdFieldAdd.Enabled = True
    cmdFieldDelete.Enabled = True
    cmdFieldsClear.Enabled = True
    cmdFieldMoveUp.Enabled = True
    cmdFieldMoveDown.Enabled = True
    
    'Reset the "DataChanged" properties so it doesn't trigger the ALTER SQL Statement unnecesarily
    cboFieldType.DataChanged = False
    txtFieldSize.DataChanged = False
    
Exit Sub
    
ERROR_TRAP:
    ' 94 = Invalid use of Null
    If Err.Number = 94 Then
        Err.Clear
        Resume Next
    End If
    result = MsgBox("Error: " & Err.Number & " - " & Err.Description, vbOK)
    Err.Clear
End Sub

Private Sub lstSecurityApplicationList_Click()

        
        If Trim(lstSecurityApplicationList.Text) <> "" Then
        
            cmdSecurityApplicationListRevoke.Enabled = True
            txtSecurityApplicationRECID = lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListIndex)
            
            'If Enable COPY MODE is checked, Do NOT load rights or Clear Fields)
            If chkRightsEnableCopyMode.Value = vbUnchecked Then
                subLoadSecuritRoleApp
            End If   'chkRightsEnableCopyMode.Value = vbUnchecked
            
        Else
            cmdSecurityApplicationListRevoke.Enabled = False
        End If
        
    
End Sub

Private Sub lstSecurityApplicationSelectionList_DblClick()
    
    subSecurityApplicationGrantAccess

End Sub


Private Sub lstUserList_Click()

    On Error GoTo ERROR_HANDLER
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "Select * from I101Security where SecurityRECID = " & lstUserList.ItemData(lstUserList.ListIndex)
    rs.Open ssql, con
    
    
        txtSecurityRECID = rs.Fields!SecurityRECID & ""
        txtUserID = rs.Fields!UserID & ""
        txtUserName = rs.Fields!username & ""
        txtPassword = rs.Fields!Password & ""
        
'        cmbBatchMode = rs.Fields!BatchMode & ""
'        cmbUserSupervisor = rs.Fields!UserSupervisor & ""
'        cmbBatchListOrder = rs.Fields!BatchListOrder & ""
'        cmbBatchDefaultQueue = rs.Fields!BatchDefaultQueue & ""
        cmbBatchDefaultApplication = rs.Fields!BatchDefaultApplication & ""
'        txtBatchQueueNotificationFrequency = rs.Fields!BatchQueueNotificationFrequency & ""
'
        chkRightsAdminSystem = rs.Fields!RightsAdminSystem & ""
'        chkRightsAdminApplication = rs.Fields!RightsAdminApplication & ""
'        chkRightsBatchScan = rs.Fields!RightsBatchScan & ""
'        chkRightsBatchIndex = rs.Fields!RightsBatchIndex & ""
'        chkRightsBatchAdministration = rs.Fields!RightsBatchAdministration & ""
'        chkRightsImportFromFile = rs.Fields!RightsImportFromFile & ""
'        chkRightsImportFromEcapture = rs.Fields!RightsImportFromEcapture & ""
'        chkRightsDeleteDocuments = rs.Fields!RightsDeleteDocuments & ""
'        chkRightsModifyIndexes = rs.Fields!RightsModifyIndexes & ""
'        chkRightsBatchCommit = rs.Fields!RightsBatchCommit & ""
'        chkRightsDeleteBatches = rs.Fields!RightsDeleteBatches & ""
'        chkRightsSendMail = rs.Fields!RightsSendMail & ""
'        chkRightsLaunchDoc = rs.Fields!RightsLaunchDoc & ""
'        chkRightsPrint = rs.Fields!RightsPrint & ""
'        chkRightsAnnotate = rs.Fields!RightsAnnotate & ""
'        chkRightsThumbnails = rs.Fields!RightsThumbnails & ""
'        chkRightsScannerSettings = rs.Fields!RightsScannerSettings & ""
'        chkRightsBatchView = rs.Fields!RightsBatchView & ""
'        chkRightsBatchRoute = rs.Fields!RightsBatchRoute & ""
'        chkRightsBatchChangeOrder = rs.Fields!RightsBatchChangeOrder & ""
'        chkRightsRetrieveImages = rs.Fields!RightsRetrieveImages & ""
'        chkViewResetImagesOnFind = rs.Fields!ViewResetImagesOnFind & ""
'        chkRightsDocPackage = rs.Fields!RightsDocPackage & ""
'        chkRightsExport = rs.Fields!RightsExport & ""
'
'        chkAllowModificationOfOrigDocs = rs.Fields!AllowModificationOfOrigDocs & ""
'
'        chkRightsBatchFindRestricted = rs.Fields!RightsBatchFindRestricted & ""
'        chkRightsBatchFindRestrictToQueue = rs.Fields!RightsBatchFindRestrictToQueue & ""
'        chkRightsBatchFindRestrictToOwner = rs.Fields!RightsBatchFindRestrictToOwner & ""
'        chkRightsBatchChangeQueue = rs.Fields!RightsBatchChangeQueue & ""
'        chkRightsBatchChangeOwner = rs.Fields!RightsBatchChangeOwner & ""
'
'        chkRightsBatchAllowDocTypeEdit = rs.Fields!RightsBatchAllowDocTypeEdit & ""

    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    'Load List of Applications
    subLoadSecurityApplications
    
'    'Load Security Rights for the selected User/Role/Application
'    subLoadSecuritRoleApp
    
    cmdSecurityApplicationListGrantRevoke.Enabled = True
    
    cmdSecurityClearFields.Enabled = True
    cmdApplicationAdd.Enabled = False
    cmdApplicationRemove.Enabled = True

    cmdFieldAdd.Enabled = True
        
        
Exit Sub

ERROR_HANDLER:

    MsgBox "lstUserList_Click ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    Resume Next

End Sub



Private Sub sstabConfig_Click(PreviousTab As Integer)
    
    Select Case sstabConfig.TabCaption(sstabConfig.Tab)
        Case "Security"
            subSecurityInitializeForm
    
        Case "Special Options"
            subCheckBarCodeStatus
            subGetSMTPeMailSettings
            
        Case "FTP"
            subGetFTPSettings
            
            
    End Select
    
End Sub


Private Sub txtApplicationCommitBatchOption_Change()

    Call txtApplicationCommitBatchoption_Click
    
End Sub

Private Sub txtApplicationCommitBatchoption_Click()

    '*** Enable/Disable FTP User & Password fields
    If txtApplicationCommitBatchTo.Text = "TTC" _
    Or txtApplicationCommitBatchOption = "FTP Only" _
    Or txtApplicationCommitBatchOption = "Application & FTP" Then

        'SHOW the FTP Tab
        sstabApplication.TabVisible(4) = True

'        lblFTPSite.Visible = True
'        txtFTPSite.Visible = True
'        lblFTPUserID.Visible = True
'        txtFTPUserID.Visible = True
'        lblFTPPassword.Visible = True
'        txtFTPPassword.Visible = True
        
''        lblFTPSite.enabled = True
''        txtFTPSite.enabled = True
''        lblFTPUserID.enabled = True
''        txtFTPUserID.enabled = True
''        lblFTPPassword.enabled = True
''        txtFTPPassword.enabled = True
''
''        txtFTPSite.BackColor = vbWhite
''        txtFTPUserID.BackColor = vbWhite
''        txtFTPPassword.BackColor = vbWhite
        
    Else
    
        'HIDE the FTP Tab
        sstabApplication.TabVisible(4) = False
    
'        lblFTPSite.Visible = False
'        txtFTPSite.Visible = False
'        lblFTPUserID.Visible = False
'        txtFTPUserID.Visible = False
'        lblFTPPassword.Visible = False
'        txtFTPPassword.Visible = False
        
''        lblFTPSite.enabled = False
''        txtFTPSite.enabled = False
''        lblFTPUserID.enabled = False
''        txtFTPUserID.enabled = False
''        lblFTPPassword.enabled = False
''        txtFTPPassword.enabled = False
''
''        txtFTPSite.BackColor = vbGrayed
''        txtFTPUserID.BackColor = vbGrayed
''        txtFTPPassword.BackColor = vbGrayed
        
    End If
    
End Sub

Private Sub txtApplicationMaxItemsToRetrieve_Change()

End Sub

Private Sub txtApplicationMaxItemsToRetrieve_Validate(Cancel As Boolean)
    'remove commas
    txtMaxItemsToRetrieve = Replace(rs.Fields!MaxItemsToRetrieve, ",", "")
    If txtMaxItemsToRetrieve.Text > "999999" Then
        MsgBox "Maximum = 999999... Setting to Unlimited (0)", vbOKOnly, "Set max items to retrieve"
        txtMaxItemsToRetrieve.Text = 0
    End If
End Sub


Private Sub txtApplicationCommitBatchTo_Click()

            '*** 2021-06-14 - Jacob - Added Option for Commit to Imaging101 AutoImport, but Without FTP Option
            If txtApplicationCommitBatchTo = "Imaging101AutoImport" Then
            
                txtApplicationCommitBatchOption.Enabled = False
                txtApplicationCommitBatchOption.Text = "Application Only"
                
            End If
            
End Sub

Private Sub txtApplicationName_Change()

    If Len(txtApplicationName.Text) > 0 Then
        cmdApplicationAdd.Enabled = True
        cmdApplicationClear.Enabled = True
        If Len(txtApplicationName) > 30 Then
            result = MsgBox("Application Name must be no more than 30 Characters, including only Letters, Numbers or Underscores!", vbOKCancel)
            txtApplicationName = Left(txtApplicationName, 30)
            Exit Sub
        End If
    Else
        cmdApplicationAdd = False
    End If
    
End Sub

Private Sub txtApplicationName_LostFocus()
    
    txtApplicationName = Trim(UCase(Replace(txtApplicationName, Chr(32), "_", , , vbBinaryCompare)))
    
End Sub


Private Sub txtBatchQueueNotificationFrequency_Validate(Cancel As Boolean)
        
    Select Case Trim(txtBatchQueueNotificationFrequency.Text)
        Case ""
            txtBatchQueueNotificationFrequency.Text = "0"
        Case Is > "9999"
            txtBatchQueueNotificationFrequency = "9999"
    End Select
 

End Sub

Private Sub txtFieldFormat_LostFocus()
    subFieldCheckFormat
End Sub

Private Sub subFieldCheckFormat()
    
'    If (Len(txtFieldMask.Text) <> Len(txtFieldFormat.Text)) Then
'        MsgBox "The length of your MASK and FORMAT are different... This could result in unpredictable storage or display of values.", vbInformation
'    End If
'    If (Len(txtFieldMask.Text) <> Len(txtFieldSize.Text)) Then
'        MsgBox "The length of your MASK is different than the Field Size you defined... This could result in unpredictable storage or display of values.", vbInformation
'    End If
'    If (Len(txtFieldFormat.Text) <> Len(txtFieldSize.Text)) Then
'        MsgBox "The length of your FORMAT is different than the Field Size you defined... This could result in unpredictable storage or display of values.", vbInformation
'    End If
   
End Sub
Private Sub txtFieldMask_LostFocus()
    subFieldCheckFormat
End Sub

Private Sub txtFieldName_LostFocus()

    txtFieldName = Trim(Replace(txtFieldName, Chr(32), "_", , , vbBinaryCompare))

End Sub

Private Sub txtFieldNameForOutput_GotFocus()

    If Trim(txtFieldNameForOutput.Text) = "" Then
        txtFieldNameForOutput.Text = txtFieldNameForInput.Text
    End If
    
End Sub

Private Sub txtFieldSize_LostFocus()
    If txtFieldSize = "" Then
        result = MsgBox("Must enter a field size!", vbOKOnly)
        If result = vbCancel Then
            Exit Sub
        End If
        txtFieldSize.SetFocus
    End If
End Sub

Private Sub subSecurityInitializeForm()

    '*** 2023-02-09 - Jacob - Added On Error Resume Next
    On Error Resume Next
    
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    '*** CLEAR the USERLIST & Combos
    lstUserList.Clear
    cmbBatchDefaultApplication.Clear
    cmbBatchDefaultQueue.Clear
    cmbBatchListOrder.Clear
    cmbUserSupervisor.Clear

    '***************************************
    '*** LOAD USER ID's
        
    subSecurityLoadUserIDs
    
    '***************************************
    '*** LOAD BATCH QUEUES LIST DROP-DOWN
        
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchQueues ORDER BY BatchQueue"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    'Add a Blank value to allow clearing the BatchOwner
'    cmbBatchQueue.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchDefaultQueue.AddItem rs.Fields!BatchQueue
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
    If cmbBatchDefaultQueue.ListCount > 0 Then
        cmbBatchDefaultQueue.ListIndex = cmbBatchDefaultQueue.TopIndex
    End If
    
    
    
    
    '***************************************
    '*** LOAD APPLICATIONS DROP-DOWN
        
    '*** Declarations
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'sql statement to select items on the drop down list
    ssql = "SELECT ApplicationName, ApplicationRECID from I101Applications ORDER BY ApplicationName"
    rs.Open ssql, con
    
    cmbBatchDefaultApplication.Clear  ' Reset the List
    
    Do Until rs.EOF
        cmbBatchDefaultApplication.AddItem rs!ApplicationName
        rs.MoveNext
    Loop
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    '****************************
    
    
    '***************************************
    '*** LOAD BATCH LIST ORDER DROP-DOWN
        
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101BatchListOrder ORDER BY BatchListOrder"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    con.Errors.Clear
    
    rs.Open
    rs.MoveFirst
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbBatchListOrder.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        cmbBatchListOrder.AddItem rs.Fields!BatchListOrder
        rs.MoveNext
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing

    '****************************
    


End Sub

Private Sub subSecurityLoadUserIDs()

    '*** CLEAR the USERLIST & Combos
    lstUserList.Clear
    cmbUserSupervisor.Clear
    
    '*************************************************************
    '*** LOAD UserID's   - BEGIN
    
    txtActionBeforeError = "Connect to Imaging101 DB"
    
    Set con = New ADODB.Connection
    con.Open RegImaging101ConnectionString

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = con
    
    rs.Source = "Select * from I101Security ORDER BY UserName"
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LOCKTYPE = adLockReadOnly
    
    rs.Open
    
    txtActionBeforeError = "Populate UserID List"
    
    'Add a Blank value to allow clearing the BatchOwner
    cmbUserSupervisor.AddItem ""
    For intIndex = 0 To rs.RecordCount - 1
        lstUserList.AddItem rs.Fields("UserName")
        lstUserList.ItemData(lstUserList.ListCount - 1) = rs.Fields("SecurityRECID")
        cmbUserSupervisor.AddItem rs.Fields("UserName")
        rs.MoveNext
        DoEvents
    Next
        
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    
    '*** LOAD UserID's   - END
    '*************************************************************

End Sub



Private Sub lstSecurityApplicationList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'When an item is Dragged from lstSecurityApplicationSelectionList to lstSecurityApplicationList
    subSecurityApplicationGrantAccess
    
    
End Sub


Private Sub subSecurityApplicationGrantAccess()
' drag and drop subroutine for moving items from listbox lstSecurityApplicationSelectionList to lstSecurityApplicationList
' this routine was copied from a VB programming Web site

Dim i As Long ' counter
Dim j As Long ' counter
Dim intDuplicate As Integer

' copy the values from lstSecurityApplicationSelectionList to lstSecurityApplicationList
' here you can add them in order starting at the top
    For i = 0 To lstSecurityApplicationSelectionList.ListCount - 1
        If lstSecurityApplicationSelectionList.Selected(i) = True Then
            ' we don't want duplicate entries so for each new item, run
            ' through the existing selected list and flag if already there
            intDuplicate = 0
            For j = 0 To lstSecurityApplicationList.ListCount - 1
                If lstSecurityApplicationSelectionList.List(i) = lstSecurityApplicationList.List(j) Then
                    intDuplicate = 1
                    Exit For
                End If
            Next j
            ' if not a duplicate, add to list
            If intDuplicate = 0 Then
                lstSecurityApplicationList.AddItem lstSecurityApplicationSelectionList.List(i)
                lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListCount - 1) = lstSecurityApplicationSelectionList.ItemData(i)
                            
                                            
                Dim con As ADODB.Connection
                Dim cmd As ADODB.Command
                
                Set con = New ADODB.Connection
                con.Open RegImaging101ConnectionString
                
                Set cmd = New ADODB.Command
                Set cmd.ActiveConnection = con
                
                
                Err.Clear
                con.Errors.Clear
                
                'Insert one record for each Application for the existing user
                cmd.CommandText = "INSERT INTO i101SecurityApplications (ApplicationRECID, SecurityRECID) SELECT Applicationrecid, " & txtSecurityRECID & " FROM I101Applications WHERE ApplicationName = '" & lstSecurityApplicationSelectionList.List(i) & "'"
                txtActionBeforeError = cmd.CommandText
                cmd.Execute , , adCmdText
                            
                con.Close
                
                Set con = Nothing
                Set cmd = Nothing
                            
            End If
        End If
    Next i
    
' remove the values from lstSecurityApplicationSelectionList
' but here you must remove them from the bottom up
''   For i = lstSecurityApplicationSelectionList.ListCount - 1 To 0 Step -1
''       If lstSecurityApplicationSelectionList.Selected(i) = True Then
''           lstSecurityApplicationSelectionList.RemoveItem (i)
''       End If
''   Next i
End Sub

Private Sub subCheckBarCodeStatus()

    On Error Resume Next
    
    frmConfig.txtBarcodeLicenseKey = VBGetPrivateProfileString(RegAppname, "frmConfig.txtBarcodeLicenseKey", RegFileName)
    frmConfig.chkDropLeadingZeroes = CInt(VBGetPrivateProfileString(RegAppname, "frmConfig.chkDropLeadingZeroes", RegFileName))
    frmConfig.chkUseBarcodeAsDocumentHeader = CInt(VBGetPrivateProfileString(RegAppname, "frmConfig.chkUseBarcodeAsDocumentHeader", RegFileName))
    frmConfig.txtBarcodeClipBeginPosition = VBGetPrivateProfileString(RegAppname, "frmConfig.txtBarcodeClipBeginPosition", RegFileName)
    frmConfig.txtBarcodeClipNumberOfCharacters = VBGetPrivateProfileString(RegAppname, "frmConfig.txtBarcodeClipNumberOfCharacters", RegFileName)

    '*** Validate the Barcode License Key
    bolBarcodeLicenseValidated = funcValidateBarCodeLicense
    
    If bolBarcodeLicenseValidated Then
        frmConfig.lblBarcodeLicenseStatus.Caption = "VALID LICENSE"
        lblBarcodeLicenseStatus.ForeColor = vbGreen
    Else
        frmConfig.lblBarcodeLicenseStatus.Caption = "LICENSE NOT VALIDATED"
        lblBarcodeLicenseStatus.ForeColor = vbRed
    End If
    
End Sub

Private Sub subGetSMTPeMailSettings()

        '***********************************************************
        '*** Save SMTP Settings
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim ssql As String
        
        Set conn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        On Error GoTo ERROR_HANDLER
        
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
        cboSendEmailViaSMTP = rs("SendEmailViaSMTP") & ""
        txtSMTPHost = rs("SMTPHost") & ""
        txtSMTPPOP3Host = rs("SMTPPOP3Host") & ""
        cboSMTPRequiresAuthentication = rs("SMTPRequiresAuthentication") & ""
        cboSMTPUsePOP3Auth = rs("SMTPUsePOP3Auth") & ""
        txtSMTPAuthenticationUserID = rs("SMTPAuthenticationUserID") & ""
        txtSMTPAuthenticationPassword = rs("SMTPAuthenticationPassword") & ""
        txtSMTPEmailSubject = rs("SMTPDefaultEmailSubject") & ""
        txtSMTPEmailMessage = rs("SMTPDefaultEmailMessage") & ""
        txtSMTPPort = rs("SMTPPort") & ""
        
        txtAutoLaunchFileTypes = rs("AutoLaunchFiletypes") & ""
        
        '2021-11-09 - Jacob - Added New Field for selecting how to Auto-Launch documents.
        frmConfig.txtAutoLaunchTo = rs("AutoLaunchTo") & ""

        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

Exit Sub

ERROR_HANDLER:

    If Err.Number = 13 Then  ' Type Mismatch
        Resume Next
    Else
        MsgBox "subGetSMTPeMailSettings ERROR: " & Err.Number & vbCrLf & Err.Description
    End If
    

End Sub


Private Sub subLoadSecuritRoleApp()

    'LOAD Security settings by Role/User/Application

    On Error GoTo ERROR_HANDLER
    
    '*** Declarations
    Dim rs As ADODB.Recordset
    Dim con As ADODB.Connection
    Dim ssql As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open RegImaging101ConnectionString
    
    'SQL Statement to Select the Security settings by Role/User/Application
    ssql = "Select * from I101SecurityRoleApp where SecurityRECID = " & _
            txtSecurityRECID & _
            " AND ApplicationRECID = " & _
            lstSecurityApplicationList.ItemData(lstSecurityApplicationList.ListIndex)
    rs.Open ssql, con
    
    If rs.BOF = True Or rs.EOF = True Then
        
        subSecurityRightsClearFields
        
    Else
    
        cmbBatchMode = rs.Fields!BatchMode & ""
        cmbUserSupervisor = rs.Fields!UserSupervisor & ""
        cmbBatchListOrder = rs.Fields!BatchListOrder & ""
        cmbBatchDefaultQueue = rs.Fields!BatchDefaultQueue & ""
'        cmbBatchDefaultApplication = rs.Fields!BatchDefaultApplication & ""
        txtBatchQueueNotificationFrequency = rs.Fields!BatchQueueNotificationFrequency & ""
        
'        chkRightsAdminSystem = rs.Fields!RightsAdminSystem & ""
        chkRightsAdminApplication = rs.Fields!RightsAdminApplication & ""
        chkRightsBatchScan = rs.Fields!RightsBatchScan & ""
        chkRightsBatchIndex = rs.Fields!RightsBatchIndex & ""
        chkRightsBatchAdministration = rs.Fields!RightsBatchAdministration & ""
        chkRightsImportFromFile = rs.Fields!RightsImportFromFile & ""
        chkRightsImportFromEcapture = rs.Fields!RightsImportFromEcapture & ""
        chkRightsDeleteDocuments = rs.Fields!RightsDeleteDocuments & ""
        chkRightsModifyIndexes = rs.Fields!RightsModifyIndexes & ""
        chkRightsBatchCommit = rs.Fields!RightsBatchCommit & ""
        chkRightsDeleteBatches = rs.Fields!RightsDeleteBatches & ""
        chkRightsSendMail = rs.Fields!RightsSendMail & ""
        chkRightsLaunchDoc = rs.Fields!RightsLaunchDoc & ""
        chkRightsPrint = rs.Fields!RightsPrint & ""
        chkRightsAnnotate = rs.Fields!RightsAnnotate & ""
        chkRightsThumbnails = rs.Fields!RightsThumbnails & ""
        chkRightsScannerSettings = rs.Fields!RightsScannerSettings & ""
        chkRightsBatchView = rs.Fields!RightsBatchView & ""
        chkRightsBatchRoute = rs.Fields!RightsBatchRoute & ""
        chkRightsBatchChangeOrder = rs.Fields!RightsBatchChangeOrder & ""
        chkRightsRetrieveImages = rs.Fields!RightsRetrieveImages & ""
        chkViewResetImagesOnFind = rs.Fields!ViewResetImagesOnFind & ""
'        chkRightsDocPackage = rs.Fields!RightsDocPackage & ""
        chkRightsExport = rs.Fields!RightsExport & ""

        chkAllowModificationOfOrigDocs = rs.Fields!AllowModificationOfOrigDocs & ""
        
        chkRightsBatchFindRestricted = rs.Fields!RightsBatchFindRestricted & ""
        chkRightsBatchFindRestrictToQueue = rs.Fields!RightsBatchFindRestrictToQueue & ""
        chkRightsBatchFindRestrictToOwner = rs.Fields!RightsBatchFindRestrictToOwner & ""
        chkRightsBatchChangeQueue = rs.Fields!RightsBatchChangeQueue & ""
        chkRightsBatchChangeOwner = rs.Fields!RightsBatchChangeOwner & ""
        
        chkRightsBatchAllowDocTypeEdit = rs.Fields!RightsBatchAllowDocTypeEdit & ""
        
        chkRightsAdvancedSearch = rs.Fields!RightsAdvancedSearch & ""

        chkRightsFileDocsViaI101FILER = rs.Fields!RightsFileDocsViaI101FILER

       chkRightsEditSearchTemplates = rs.Fields!RightsEditSearchTemplates
        
    End If
    
    'Close connection and the recordset
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
        
Exit Sub

ERROR_HANDLER:

    MsgBox "subLoadSecuritRoleApp ERROR: " & Err.Number & vbCrLf & Err.Description & vbCrLf & sErrMessage
    Resume Next

End Sub

Private Sub subGetFTPSettings()



End Sub
