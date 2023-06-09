VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00C07000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4530
      Left            =   360
      TabIndex        =   0
      Top             =   345
      Width           =   7080
      Begin VB.TextBox strConnectionStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   165
         TabIndex        =   4
         Text            =   "Connection Status"
         Top             =   2385
         Width           =   6615
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6240
         Top             =   3960
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connecting to Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   285
         TabIndex        =   5
         Top             =   1920
         Width           =   6495
      End
      Begin VB.Image Image2 
         Height          =   645
         Left            =   2400
         Picture         =   "frmSplash.frx":000C
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label lblDisclaimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WARNING: ..."
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
         Height          =   585
         Left            =   480
         TabIndex        =   3
         Top             =   3840
         Width           =   6015
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version"
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
         Height          =   225
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   3885
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00FFFFFF&
         Caption         =   "App Description"
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
         Height          =   810
         Left            =   720
         TabIndex        =   1
         Top             =   2880
         Width           =   5445
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    
    Me.Caption = "Startup " & App.Title
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    lblTitle.Caption = App.Title
    lblDescription.Caption = "Imaging101 is an amazingly powerful document scanning, routing, indexing and retrieval solution.  The Paperless office is only a Click away."
    lblDisclaimer.Caption = "Copyright 2002 - " & Format(Now(), "yyyy") & " by Imaging101, Inc.   Please do not duplicate this software without a valid license.  This software is protected by US and International law."
    
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Timer1.Enabled = False
    
End Sub

Private Sub lblMessage_Change()
        '*** 2022-11-07 - Jacob - Added Write to DEBUG LOG
        funcWriteToDebugLog Me.name, lblMessage
End Sub

Private Sub Timer1_Timer()

'    Me.strConnectionStatus.Text = frmImaging101Winsock.strConnectionStatus.Text
    DoEvents
    
End Sub



