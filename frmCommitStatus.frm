VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommitStatus 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Commit Status"
   ClientHeight    =   6540
   ClientLeft      =   960
   ClientTop       =   315
   ClientWidth     =   7650
   ControlBox      =   0   'False
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
   ScaleHeight     =   6540
   ScaleWidth      =   7650
   Begin MSComctlLib.ProgressBar ProgressBarCopyPageToArchive 
      Height          =   330
      Left            =   4050
      TabIndex        =   25
      Top             =   840
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtPagesDoNotFile 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   19
      Top             =   4050
      Width           =   1095
   End
   Begin VB.TextBox txtPagesCommitted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4095
      TabIndex        =   17
      Top             =   1635
      Width           =   1095
   End
   Begin VB.CommandButton cmdCloseBatch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close Batch"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4065
      Picture         =   "frmCommitStatus.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5550
      Width           =   2895
   End
   Begin VB.CommandButton cmdStayOnBatch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Stay On This Batch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   345
      Picture         =   "frmCommitStatus.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5550
      Width           =   2775
   End
   Begin VB.TextBox txtPagesTotalSkipped 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5115
      TabIndex        =   14
      Top             =   2250
      Width           =   1095
   End
   Begin VB.TextBox txtPagesFailedValidation 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   13
      Top             =   4770
      Width           =   1095
   End
   Begin VB.TextBox txtPagesRequiredButEmpty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   12
      Top             =   4410
      Width           =   1095
   End
   Begin VB.TextBox txtPagesQuestionable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   11
      Top             =   3690
      Width           =   1095
   End
   Begin VB.TextBox txtPagesSeparator 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   10
      Top             =   3330
      Width           =   1095
   End
   Begin VB.TextBox txtPagesPreviouslyCommitted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   9
      Top             =   2970
      Width           =   1095
   End
   Begin VB.TextBox txtPagesProcessed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4095
      TabIndex        =   8
      Top             =   1275
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
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
      Left            =   5760
      Picture         =   "frmCommitStatus.frx":0B14
      ScaleHeight     =   375
      ScaleWidth      =   1575
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      TabIndex        =   22
      Top             =   15
      Width           =   7575
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         ForeColor       =   &H00C07000&
         Height          =   225
         Left            =   5760
         TabIndex        =   24
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMIT SUMMARY"
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
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Copying Page to Archive Progress"
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
      Height          =   255
      Left            =   495
      TabIndex        =   26
      Top             =   855
      Width           =   3375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DO-NOT-FILE Pages Skipped:"
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
      Left            =   1725
      TabIndex        =   20
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages Committed:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1095
      TabIndex        =   18
      Top             =   1635
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pages Skipped:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1350
      TabIndex        =   7
      Top             =   2310
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason Pages Skipped (if any)"
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
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2655
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages with Failed Validation Skipped:"
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
      Left            =   1725
      TabIndex        =   5
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages with FIELDS REQUIRED but Empty:"
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
      Left            =   1725
      TabIndex        =   4
      Top             =   4440
      Width           =   3195
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "QUESTIONABLE Pages Skipped:"
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
      Left            =   1725
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SEPARATOR Pages Skipped:"
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
      Left            =   1725
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages PREVIOUSLY Committed:"
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
      Left            =   1725
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pages Processed:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1095
      TabIndex        =   0
      Top             =   1275
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   22875
      Left            =   0
      Picture         =   "frmCommitStatus.frx":11A7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmCommitStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub MakeTopMost()

    On Error Resume Next
    
    ' Make this window the top most window so that it appears in
    ' front of the ecIndex window
    Dim lReturn As Long
    
    ' The SetWindowPos API call will do this.
    lReturn = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Private Sub cmdCloseBatch_Click()
    
    On Error Resume Next
    
    If funcIsFormLoaded2("frmIndex") Then
        'This form will be closed by the frmIndex Unload
        Unload frmIndex
    End If
    
    
End Sub

Private Sub cmdStayOnBatch_Click()

    On Error Resume Next
    
    Me.Hide
    
    frmIndex.Show
    DoEvents
    frmIndex.subDisplayOrHideForms
    DoEvents

    
    Unload Me

End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision


End Sub

Private Sub Form_Unload(Cancel As Integer)

     On Error Resume Next
    
   'SAFE way of saying: Set Me = Nothing
    Dim Form As Form
    For Each Form In Forms
            If Form Is Me Then
                    Set Form = Nothing
                    Exit For
            End If
    Next Form

End Sub

