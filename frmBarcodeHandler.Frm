VERSION 5.00
Object = "{ED512BE6-6629-4FB4-953D-D0C353847163}#1.0#0"; "ImagXpr7.dll"
Object = "{72734C46-9CBF-4499-B315-24170670045C}#1.0#0"; "PegasusImaging.ActiveX.SSXBC4.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBarcodeHandler 
   Caption         =   "Detecting Barcodes using the Picture Property"
   ClientHeight    =   9552
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9552
   ScaleWidth      =   7380
   Begin ImagXpr7Ctl.ImagXpress img 
      Height          =   3975
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   6855
      _ExtentX        =   12086
      _ExtentY        =   7006
      ErrStr          =   "05022C5607440F593D3450859C300185"
      ErrCode         =   1661232004
      ErrInfo         =   -2143093037
      Persistence     =   -1  'True
      _cx             =   12086
      _cy             =   7006
      AutoSize        =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SaveTransparencyColor=   0
      OLEDropMode     =   0
      SaveTIFFCompression=   0
      SaveTransparent =   0
      SaveJPEGProgressive=   0   'False
      SaveJPEGGrayscale=   0   'False
      SaveJPEGLumFactor=   10
      SaveJPEGChromFactor=   10
      SaveJPEGSubSampling=   2
      ViewAntialias   =   -1  'True
      ViewDithered    =   -1  'True
      AlignH          =   1
      AlignV          =   1
      LoadRotated     =   0
      JPEGEnhDecomp   =   -1  'True
      WMFConvert      =   0   'False
      ProcessImageID  =   1
      OwnDIB          =   -1  'True
      FileTimeout     =   10000
      AsyncPriority   =   0
      LZWPassword     =   ""
      ViewUpdate      =   -1  'True
      TWAINProductName=   ""
      TWAINProductFamily=   ""
      TWAINManufacturer=   ""
      TWAINVersionInfo=   ""
      SavePDFCompression=   8
      ViewProgressive =   0   'False
      SaveTIFFByteOrder=   0
      FTPUserName     =   ""
      FTPPassword     =   ""
      ProxyServer     =   ""
      SaveEXIFThumbnailSize=   0
      SaveLJPPrediction=   1
      PDFSwapBlackandWhite=   0   'False
      SaveTIFFRowsPerStrip=   0
      TIFFIFDOffset   =   0
      ViewGrayMode    =   0
      SaveWSQQuant    =   1
      DisplayError    =   0   'False
      EvalMode        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Barcode Type:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   6855
      Begin VB.OptionButton radBarcode1D 
         Caption         =   "&1D (Standard)"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton radBarcodePatchCode 
         Caption         =   "&Patch Code"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton radBarcodeTypePDF417 
         Caption         =   "PDF-&417"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton radBarcodeTypeDataMatrix 
         Caption         =   "&Data Matrix"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton radBarcodeTypePostnet 
         Caption         =   "P&ostNet"
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Detect Barcodes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   8
      Top             =   8760
      Width           =   2292
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   240
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "9"
   End
   Begin PegasusImagingActiveXSSXBarcode4Ctl.SSXBarcode SS 
      Left            =   915
      Top             =   8745
      _ExtentX        =   847
      _ExtentY        =   847
      ErrStr          =   "1CD8CA89D5B4A85DF44C60AC1806F163"
      ErrCode         =   1661232004
      ErrInfo         =   -1180790406
      _cx             =   847
      _cy             =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarcodeType     =   0
      MaxBarcodes     =   100
      BarcodeAppendCheckSum=   0   'False
      BarcodeInkColor =   1
      BarcodeOrientation=   0
      ReturnPossibleBarcodes=   0   'False
      MakeBarcodeStyle=   9
      MakeBarcodeWidth=   600
      MakeBarcodeHeight=   600
      MakeBarcodeValue=   ""
      MakeBarcodeBarSize=   3
      MakeBarcodeShowValue=   3
      MakeBarcodeBWidth=   50
      MakeBarcodeBHeight=   50
      MakeBarcodeAutoSize=   -1  'True
      MakeBarcodeValueGap=   5
      MakeBarcodeValueOGap=   5
      MakeBarcodeMinHeight=   100
      MakeBarcodeValueLRGap=   0
      MakeBarcodeUPCNotchPercent=   10
      MakeBarcodeAdjustBarSize=   -1  'True
      WriteChecksumChar=   0
      BeginProperty MakeBarcodeValueOFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageSource     =   0
      EvaluationMode  =   0
      LicenseEdition  =   0
      RaiseExceptions =   -1  'True
      ErrorLevel      =   0
      Debug           =   0   'False
      DebugLogFile    =   "c:\SSXBarcode4.log"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1110
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmBarcodeHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*** Declare Pegasus Imaging UNLOCK for Bar-Code License
Private Declare Sub BC_UnlockControl Lib "PegasusImaging.ActiveX.SSXBC4.dll" Alias "UnlockControl" (ByVal pw1 As Long, ByVal pw2 As Long, ByVal pw3 As Long, ByVal pw4 As Long)
Private Declare Sub ControlUnlock Lib "ImagXpr7" Alias "PS_Unlock" (ByVal pw1 As Long, ByVal pw2 As Long, ByVal pw3 As Long, ByVal pw4 As Long)

Private Enum EnumBarcodeType
    enumDefault = 0
    enumPDF417 = &H200000
    enumPatchCode = &H80000
    enumDataMatrix = &H400000
    enumPostNet = &H100000
End Enum




Public Function funcProcessBarcode(strFilePath As String) As String
    
    
    Dim result As String
    Dim i As Long
    Dim strBiTonalFilePath As String
    
    
    If Trim(strFilePath) = "" Then
        MsgBox "funcProcessBarcode ERROR:  No FilePath Provided!"
        Exit Function
    End If
    
    'Handle the BiTonal/Binary version of the image file
    strBiTonalFilePath = Replace(strFilePath, "color", "binary")
    strBiTonalFilePath = Replace(strBiTonalFilePath, "jpg", "tif")

    With img
'        .FileName = strFilePath
        .FileName = strBiTonalFilePath
'        .width = 6900
'        .Height = 4000
    End With

    
    If img.hDIB = 0 Then
        MsgBox ("No file loaded")
    Else
        Me.MousePointer = vbHourglass
        DoEvents
        
        On Error GoTo BCError
        Call SetBarcodeType
        
        With SS
            .ImageSource = 1   ' Use Picture property for the image source
            .Picture = img.Picture
            .AnalyzeBarcode
            If .NumBarcodes > 0 And .SSError = 0 Then
                For i = 0 To SS.NumBarcodes - 1
                    SS.GetBarcode i
                    result = result & "Barcode #" & i & _
                    "   Value = " & SS.BarcodeResult & _
                    "   Type = " & SS.BarcodeCodeName & Chr$(13)
                Next i
                
                Debug.Print result
'                MsgBox result
                funcProcessBarcode = SS.BarcodeResult
            Else
'                MsgBox "No barcodes were found.  Error = " & SS.SSError & ".  " & SS.SSErrorMsg, , _
                    "Smartscan Xpress Barcode Sample"
                funcProcessBarcode = "0"
            End If
        End With
        'Clear the Picture property
        SS.Picture = LoadPicture("")
        
        Me.MousePointer = vbDefault
        DoEvents
        End If

Exit Function

BCError:
    MsgBox "ERROR: " & Err.Number & " - " & Err.Description
    
    Resume Next
    
End Function

Private Sub Command1_Click()

    Dim result As String
    Dim i As Long
    Dim strBiTonalFilePath As String
    
    

    
    If img.hDIB = 0 Then
        MsgBox ("No file loaded")
    Else
        Me.MousePointer = vbHourglass
        DoEvents
        
        On Error GoTo BCError
        Call SetBarcodeType
        
        With SS
            .ImageSource = 1   ' Use Picture property for the image source
            .Picture = img.Picture
            .AnalyzeBarcode
            If .NumBarcodes > 0 And .SSError = 0 Then
                For i = 0 To SS.NumBarcodes - 1
                    SS.GetBarcode i
                    result = result & "Barcode #" & i & _
                    "   Value = " & SS.BarcodeResult & _
                    "   Type = " & SS.BarcodeCodeName & Chr$(13)
                Next i
                
                Debug.Print result
                MsgBox result
            Else
                MsgBox "No barcodes were found.  Error = " & SS.SSError & ".  " & SS.SSErrorMsg, , _
                    "Smartscan Xpress Barcode Sample"
            End If
        End With
        'Clear the Picture property
        SS.Picture = LoadPicture("")
        
        Me.MousePointer = vbDefault
        DoEvents
        End If

Exit Sub

BCError:
    MsgBox "ERROR: " & Err.Number & " - " & Err.Description
    
    Resume Next

End Sub

Private Sub Form_Load()
    Top = (Screen.Height - Height) / 2
    Left = (Screen.width - width) / 2
    
    ' You need to get an Integrator Key from Pegasus Software
    ' in order to distribute your application.  The key shown below
    ' is for illustration purposes only.
    'SS.IntegratorKey 123456789, 987654321, 123454321, 12345
    
    '*** UNLOCK Pegasus Imaging Bar-Code License
'    BC_UnlockControl 1530704668, 554306, 968718533, 11370
    SS.UnlockRuntime 1530704668, 554306, 968718533, 11370
    
    '*** UNLOCK Pegasus Imaging ImagXpress License
    ControlUnlock 1908228613, 373670336, 1341974545, 18755



End Sub

Private Sub mnuAbout_Click()
    SS.AboutBox
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOpen_Click()
On Error GoTo PressedCancel
    With cd
        .CancelError = True
        .DialogTitle = "Open 1-bit Monochrome Image File"
        .Filter = "TIFF (*.tif)|*.tif|BMP (*.bmp)|*.bmp"
        .InitDir = App.Path & "\..\..\..\..\..\..\Common\Images\"
        .Action = 1
    End With
    With img
        .FileName = cd.FileName
        If .FileName = "" Then
            MsgBox "Please choose a file"
        End If
        .width = 6900
        .Height = 4000
    End With
    Exit Sub
PressedCancel:
    MsgBox "Error opening file.  ImagXpress error: " & img.ImagError
    Resume Next
End Sub

Private Sub SetBarcodeType()
    With SS
        If radBarcode1D.Value = True Then
            .BarcodeType = EnumBarcodeType.enumDefault
        ElseIf radBarcodePatchCode.Value = True Then
            .BarcodeType = EnumBarcodeType.enumPatchCode
        ElseIf radBarcodeTypeDataMatrix.Value = True Then
            .BarcodeType = EnumBarcodeType.enumDataMatrix
        ElseIf radBarcodeTypePDF417.Value = True Then
            .BarcodeType = EnumBarcodeType.enumPDF417
        ElseIf radBarcodeTypePostnet.Value = True Then
            .BarcodeType = EnumBarcodeType.enumPostNet
        Else
            .BarcodeType = EnumBarcodeType.enumDefault
        End If
    End With
End Sub
