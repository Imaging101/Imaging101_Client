VERSION 5.00
Begin VB.Form frmImaging101ExportSimple 
   Caption         =   "Export Options"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmImaging101ExportSimple"
   ScaleHeight     =   2730
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optExportFormaTIF 
      Caption         =   "TIF"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1050
      Width           =   975
   End
   Begin VB.OptionButton optExportFormatPDF 
      Caption         =   "PDF"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   1050
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdExportSelected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Export"
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
         Height          =   735
         Left            =   840
         Picture         =   "frmImaging101ExportSimple.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Export Selected Documents"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
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
         Left            =   0
         Picture         =   "frmImaging101ExportSimple.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
      Begin VB.PictureBox picImaging101Logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6480
         Picture         =   "frmImaging101ExportSimple.frx":0E54
         ScaleHeight     =   375
         ScaleWidth      =   1695
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   6480
         TabIndex        =   4
         Top             =   480
         Width           =   1125
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Export Format"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1020
      Width           =   1575
   End
End
Attribute VB_Name = "frmImaging101ExportSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    gstrExportFormat = "CANCEL"
    Unload Me
    
End Sub

Private Sub cmdExportSelected_Click()

    If optExportFormaPDF = True Then
        gstrExportFormat = "PDF"
    Else
        gstrExportFormat = "TIF"
    End If
        
    result = WritePrivateProfileString(RegAppname, "frmImaging101ExportSimple.ExportFormat", gstrExportFormat, RegFileName)
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    ' Get Position settings from the registry
    On Error Resume Next
    Me.Top = VBGetPrivateProfileString(RegAppname, "frmImaging101ExportSimple.Top", RegFileName)
    Me.Left = VBGetPrivateProfileString(RegAppname, "frmImaging101ExportSimple.Left", RegFileName)
    gstrExportFormat = VBGetPrivateProfileString(RegAppname, "frmImaging101ExportSimple.ExportFormat", RegFileName)

    If gstrExportFormat = "PDF" Then
        optExportFormaPDF.Value = True
    Else
        optExportFormaTIF.Value = True
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'Save Form settings to the registry, except if minimized (i.e.- Top or Left < 0)
    If Me.Top >= 0 And Me.Left >= 0 Then
        result = WritePrivateProfileString(RegAppname, "frmImaging101ExportSimple.Top", Me.Top, RegFileName)
        result = WritePrivateProfileString(RegAppname, "frmImaging101ExportSimple.Left", Me.Left, RegFileName)
    End If


End Sub
